//! Guards `go/internal/capi/lib/*/libdxpdf.a` against drifting from the
//! Rust source that produced them.
//!
//! The Go bindings' prebuilt static libraries are committed (unlike
//! `go/internal/capi/dxpdf.h`, they can't be *regenerated* by a test —
//! nothing here can cross-compile all four platforms), so the only thing a
//! test can check is *provenance*: a hash of every input that can change
//! the compiled bytes, stamped into `go/internal/capi/lib/SOURCE_HASH` the
//! last time the libraries were rebuilt. A mismatch means someone changed
//! `src/`, `Cargo.toml`, `Cargo.lock` or the pinned toolchain without also
//! rebuilding and recommitting `libdxpdf.a` for every platform.
//!
//! Deliberately **not** `std::collections::hash_map::DefaultHasher`: its own
//! docs say the algorithm "is not guaranteed to be stable across different
//! versions of the standard library," which would make a bare rustc bump —
//! with no source change at all — report every committed library stale.
//! FNV-1a is hand-rolled instead: a fixed, unspecified-by-nobody algorithm,
//! which is all a staleness check needs (this is provenance tracking, not a
//! security boundary — a hash collision here just means a real bug slips
//! through, not that someone can forge one).

use std::fs;
use std::path::{Path, PathBuf};

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
}

/// The platforms `.github/workflows/ci.yml`'s `go-bindings` job and
/// AGENTS.md's Go-bindings note both name as supported.
const EXPECTED_PLATFORMS: &[&str] = &["darwin_arm64", "darwin_amd64", "linux_amd64", "linux_arm64"];

struct Fnv1a(u64);

impl Fnv1a {
    const OFFSET_BASIS: u64 = 0xcbf29ce484222325;
    const PRIME: u64 = 0x100000001b3;

    fn new() -> Self {
        Fnv1a(Self::OFFSET_BASIS)
    }

    fn write(&mut self, bytes: &[u8]) {
        for &byte in bytes {
            self.0 ^= u64::from(byte);
            self.0 = self.0.wrapping_mul(Self::PRIME);
        }
    }
}

/// Every file that can change the compiled `capi` staticlib's bytes or
/// behavior: the whole crate's source, its exact resolved dependency graph,
/// and the compiler that builds it. Deliberately excludes `tests/`,
/// `benches/`, and `go/` itself — none of those reach the shipped library.
fn source_files(root: &Path) -> Vec<PathBuf> {
    let mut files = vec![
        root.join("Cargo.toml"),
        root.join("Cargo.lock"),
        root.join("rust-toolchain.toml"),
    ];
    collect_rs_files(&root.join("src"), &mut files);
    files.sort();
    files
}

fn collect_rs_files(dir: &Path, out: &mut Vec<PathBuf>) {
    let entries =
        fs::read_dir(dir).unwrap_or_else(|e| panic!("reading directory {}: {e}", dir.display()));
    for entry in entries {
        let path = entry.expect("reading directory entry").path();
        if path.is_dir() {
            collect_rs_files(&path, out);
        } else if path.extension().is_some_and(|ext| ext == "rs") {
            out.push(path);
        }
    }
}

fn hash_source_tree(root: &Path) -> u64 {
    let mut hasher = Fnv1a::new();
    for path in source_files(root) {
        let bytes = fs::read(&path).unwrap_or_else(|e| panic!("reading {}: {e}", path.display()));
        // The path itself is part of the hash so that renaming or moving a
        // file (without touching its content) still counts as a change.
        let relative = path.strip_prefix(root).unwrap().to_string_lossy();
        hasher.write(relative.as_bytes());
        hasher.write(&bytes);
    }
    hasher.0
}

/// A platform ships either one `libdxpdf.a`, or — when the whole archive
/// would exceed GitHub's 100 MB file limit (linux/amd64, linux/arm64) —
/// `libdxpdf_part1.a`, `libdxpdf_part2.a`, ... with no gap in the numbering
/// (see `scripts/split_capi_lib.py` and `go/internal/capi/lib/README.md`).
/// Never both forms for the same platform: a leftover unsplit `libdxpdf.a`
/// alongside parts would make `go/cgo_*.go`'s `#cgo LDFLAGS` ambiguous
/// about which one it's actually linking.
fn platform_library_is_present(dir: &Path) -> Result<(), String> {
    let whole = dir.join("libdxpdf.a");
    let mut part_count = 0usize;
    loop {
        let candidate = dir.join(format!("libdxpdf_part{}.a", part_count + 1));
        if !candidate.is_file() {
            break;
        }
        part_count += 1;
    }

    match (whole.is_file(), part_count) {
        (true, 0) => Ok(()),
        (false, n) if n > 0 => Ok(()),
        (true, n) if n > 0 => Err(format!(
            "{} has both a whole libdxpdf.a and {n} libdxpdf_partN.a files — remove one form",
            dir.display()
        )),
        _ => Err(format!(
            "{} has neither libdxpdf.a nor libdxpdf_part1.a — every platform in \
             EXPECTED_PLATFORMS must have a committed library (see AGENTS.md's \
             Go-bindings note)",
            dir.display()
        )),
    }
}

#[test]
fn every_supported_platform_has_a_committed_library() {
    let lib_dir = repo_root().join("go/internal/capi/lib");
    for platform in EXPECTED_PLATFORMS {
        if let Err(message) = platform_library_is_present(&lib_dir.join(platform)) {
            panic!("{message}");
        }
    }
}

/// Set `UPDATE_CAPI_LIB_HASH=1` after rebuilding `libdxpdf.a` for every
/// platform to restamp `SOURCE_HASH`, then commit both together.
#[test]
fn committed_libs_match_current_source() {
    let root = repo_root();
    let hash_path = root.join("go/internal/capi/lib/SOURCE_HASH");
    let current_hash = format!("{:016x}", hash_source_tree(&root));

    if std::env::var_os("UPDATE_CAPI_LIB_HASH").is_some() {
        fs::write(&hash_path, format!("{current_hash}\n")).expect("writing SOURCE_HASH");
        return;
    }

    let committed_hash = fs::read_to_string(&hash_path)
        .unwrap_or_else(|e| panic!("{} is missing or unreadable: {e}", hash_path.display()));

    assert_eq!(
        current_hash,
        committed_hash.trim(),
        "go/internal/capi/lib/*/libdxpdf.a are stale relative to the current \
         Rust source (src/, Cargo.toml, Cargo.lock, rust-toolchain.toml). \
         Rebuild libdxpdf.a for every platform in EXPECTED_PLATFORMS, commit \
         them, then run `UPDATE_CAPI_LIB_HASH=1 cargo test --test \
         capi_lib_freshness` and commit the refreshed SOURCE_HASH."
    );
}
