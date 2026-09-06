//! Packaging invariants for the Debian package (issue #92).
//!
//! The `.deb` itself can only be built and inspected on a Debian host — that
//! job belongs to `scripts/verify_deb.py` and the container jobs in CI. What
//! *can* be checked on every platform, in `cargo test --all`, is the source
//! material those jobs consume: the man page and the `[package.metadata.deb]`
//! table. Both are the kind of thing that rots silently — a new CLI flag or a
//! renamed file breaks the package at release time, weeks after the change that
//! caused it. These tests move that failure to the commit that causes it.

use std::path::{Path, PathBuf};
use std::process::Command;

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
}

/// The man page as plain text, with roff's escaping removed.
///
/// A man page writes a literal hyphen as `\-` (an unescaped `-` is a
/// typographic hyphen that breaks copy-paste and `man -K` search), so
/// `\-\-image\-dpi` is the correct spelling of `--image-dpi` on disk. Dropping
/// every backslash is enough to search for flag names without teaching this
/// test the rest of roff.
fn man_page_text() -> String {
    let path = repo_root().join("man/dxpdf.1");
    let raw = std::fs::read_to_string(&path)
        .unwrap_or_else(|e| panic!("{} is missing or unreadable: {e}", path.display()));
    raw.replace('\\', "")
}

/// Every flag `dxpdf --help` advertises, as it appears in the help text:
/// long flags (`--image-dpi`) and short ones (`-o`).
fn flags_from_help() -> Vec<String> {
    let out = Command::new(env!("CARGO_BIN_EXE_dxpdf"))
        .arg("--help")
        .output()
        .expect("failed to run the dxpdf binary");
    assert!(
        out.status.success(),
        "`dxpdf --help` exited with {:?}",
        out.status
    );
    let help = String::from_utf8(out.stdout).expect("--help output is not UTF-8");

    let mut flags = Vec::new();
    let bytes: Vec<char> = help.chars().collect();
    let mut i = 0;
    while i < bytes.len() {
        // A flag starts at a `-` that is not preceded by a word character, so
        // the hyphen inside `image-dpi` doesn't start a second match.
        let boundary = i == 0 || !bytes[i - 1].is_alphanumeric() && bytes[i - 1] != '-';
        if bytes[i] == '-' && boundary {
            let start = i;
            i += 1;
            if i < bytes.len() && bytes[i] == '-' {
                i += 1;
            }
            let name_start = i;
            while i < bytes.len() && (bytes[i].is_alphanumeric() || bytes[i] == '-') {
                i += 1;
            }
            if i > name_start {
                let flag: String = bytes[start..i].iter().collect();
                if !flags.contains(&flag) {
                    flags.push(flag);
                }
            }
        } else {
            i += 1;
        }
    }
    flags
}

#[test]
fn man_page_documents_every_cli_flag() {
    let man = man_page_text();
    let flags = flags_from_help();

    // A guard on the guard: if the scrape ever returns nothing, the assertion
    // loop below passes vacuously and the test stops testing anything.
    assert!(
        flags.len() >= 4,
        "expected at least -o/--output, --image-dpi, --help, --version from `--help`, got {flags:?}"
    );

    let missing: Vec<&String> = flags.iter().filter(|f| !man.contains(f.as_str())).collect();
    assert!(
        missing.is_empty(),
        "man/dxpdf.1 does not document {missing:?} — `dxpdf --help` advertises {flags:?}"
    );
}

#[test]
fn man_page_states_the_real_dpi_bounds() {
    let man = man_page_text();
    // `MIN_CLI_IMAGE_DPI` / `MAX_CLI_IMAGE_DPI` / `DEFAULT_IMAGE_DPI`. The CLI
    // *rejects* out-of-range values rather than clamping them (see the comment
    // on `parse_image_dpi` in src/main.rs), so the documented range is the
    // difference between a working invocation and an error.
    for bound in ["220", "2400"] {
        assert!(
            man.contains(bound),
            "man/dxpdf.1 never mentions {bound}, so it cannot be stating the --image-dpi range"
        );
    }
}

#[test]
fn deb_metadata_is_complete_and_points_at_files_that_exist() {
    let manifest = std::fs::read_to_string(repo_root().join("Cargo.toml")).unwrap();
    // `Table`, not `Value`: `FromStr for Value` parses a bare TOML *value*, so a
    // whole manifest comes back as "unexpected content" the moment it reaches
    // the newline after `[package]` (which it read as an inline array).
    let manifest: toml::Table = manifest.parse().expect("Cargo.toml is not valid TOML");

    let deb = manifest
        .get("package")
        .and_then(|p| p.get("metadata"))
        .and_then(|m| m.get("deb"))
        .expect("Cargo.toml has no [package.metadata.deb] — `cargo deb` cannot build a package");

    // `maintainer` has no default: Cargo.toml declares no `authors`, so
    // cargo-deb has nothing to fall back on and refuses to build without it.
    for field in ["maintainer", "section", "priority", "assets", "changelog"] {
        assert!(
            deb.get(field).is_some(),
            "[package.metadata.deb] is missing `{field}`"
        );
    }

    // `changelog` and `license-file` are paths cargo-deb reads directly rather
    // than through `assets`, so the asset loop below never sees them.
    for key in ["changelog", "license-file"] {
        let path = match &deb[key] {
            toml::Value::String(s) => s.clone(),
            // `license-file = ["LICENSE", "0"]` — path first, lines-to-skip second.
            toml::Value::Array(a) => a[0].as_str().unwrap_or_default().to_string(),
            other => panic!("[package.metadata.deb] {key} has unexpected type {other:?}"),
        };
        assert!(
            repo_root().join(&path).is_file(),
            "[package.metadata.deb] {key} points at `{path}`, which does not exist"
        );
    }

    let assets = deb["assets"].as_array().expect("`assets` must be an array");
    let mut sources = Vec::new();
    for asset in assets {
        let row = asset
            .as_array()
            .expect("each asset is [source, dest, mode]");
        let source = row[0].as_str().expect("asset source must be a string");
        sources.push(source.to_string());

        // `target/…` is a build product and only exists after `cargo build
        // --release`; everything else is committed and must be there now.
        if !source.starts_with("target/") {
            let path = repo_root().join(source);
            assert!(
                path.is_file(),
                "[package.metadata.deb] ships `{source}`, which does not exist"
            );
        }
    }

    // The binary and the man page are what make this a usable package: Debian
    // Policy §12.1 requires a manual page for every program in /usr/bin.
    assert!(
        sources.iter().any(|s| s == "target/release/dxpdf"),
        "the package must ship the dxpdf binary, got {sources:?}"
    );
    let man_dest = assets
        .iter()
        .find(|a| a[0].as_str() == Some("man/dxpdf.1"))
        .map(|a| a[1].as_str().unwrap_or_default().to_string())
        .expect("the package must ship man/dxpdf.1");
    assert_eq!(
        man_dest, "usr/share/man/man1/",
        "the man page must land in section 1's directory or `man dxpdf` will not find it"
    );

    // The crate is `crate-type = ["rlib", "cdylib", "staticlib"]`, so `cargo
    // build --release` also emits `libdxpdf.so` and `libdxpdf.a` — the bodies
    // of the PyO3 extension module and the Go bindings' cgo target, neither
    // with a soname, headers or an ABI promise. cargo-deb's *default* asset
    // set picks up C-ABI libraries, so the explicit list above exists to keep
    // both out of /usr/lib. `verify_deb.py` proves they stayed out; this
    // proves nobody added them back by hand.
    assert!(
        !sources.iter().any(|s| s.contains("libdxpdf")),
        "the cdylib must not be packaged: {sources:?}"
    );
}

#[test]
fn debian_changelog_leads_with_the_current_version() {
    // `apt changelog dxpdf` shows this, and a changelog whose newest entry is
    // for a version nobody is running is worse than none. Nothing generates
    // this file, so the only thing keeping it current is this assertion firing
    // on the commit that bumps the version.
    let changelog = std::fs::read_to_string(repo_root().join("debian/changelog"))
        .expect("debian/changelog is missing — [package.metadata.deb] points at it");
    let first = changelog.lines().next().unwrap_or_default();

    // `dxpdf (0.4.0-1) unstable; urgency=medium`
    let entry = first
        .split_once('(')
        .and_then(|(name, rest)| rest.split_once(')').map(|(v, _)| (name.trim(), v)))
        .unwrap_or_else(|| panic!("first changelog line is not a Debian entry header: {first:?}"));
    assert_eq!(entry.0, "dxpdf", "changelog names the wrong package");

    let upstream = entry.1.split('-').next().unwrap();
    assert_eq!(
        upstream,
        env!("CARGO_PKG_VERSION"),
        "debian/changelog's newest entry is {:?}, but the crate is at {} — \
         add an entry (`dch -v {}-1`) before releasing",
        entry.1,
        env!("CARGO_PKG_VERSION"),
        env!("CARGO_PKG_VERSION"),
    );
}

#[test]
fn lintian_overrides_name_tags_without_context() {
    // Lintian prints a tag as `dxpdf: embedded-library expat [usr/bin/dxpdf]`,
    // and pasting that straight back in is the obvious way to write an
    // override — but the context is a property of one build. An override
    // carrying it matched exactly on arm64 and came back as
    // `mismatched-override` on amd64, which failed the release. A bare tag name
    // matches every instance, and an override matching nothing is only a note.
    let path = repo_root().join("debian/dxpdf.lintian-overrides");
    let overrides = std::fs::read_to_string(&path).expect("lintian overrides file is missing");

    for (n, line) in overrides.lines().enumerate() {
        let line = line.trim();
        if line.is_empty() || line.starts_with('#') {
            continue;
        }
        let tag = line.strip_prefix("dxpdf: ").unwrap_or_else(|| {
            panic!(
                "{}:{} is not `dxpdf: <tag>`: {line:?}",
                path.display(),
                n + 1
            )
        });
        assert!(
            !tag.is_empty()
                && tag
                    .chars()
                    .all(|c| c.is_ascii_lowercase() || c.is_ascii_digit() || c == '-'),
            "{}:{} carries context after the tag name: {line:?} — write `dxpdf: {}` alone, \
             or it will mismatch on another architecture",
            path.display(),
            n + 1,
            tag.split_whitespace().next().unwrap_or(tag),
        );
    }
}

#[test]
fn verify_deb_script_is_executable_python() {
    // CI runs this as `python3 scripts/verify_deb.py`; a missing file there
    // fails the packaging job several minutes into a Skia build.
    let script = repo_root().join("scripts/verify_deb.py");
    assert!(
        Path::new(&script).is_file(),
        "scripts/verify_deb.py is missing — CI's packaging job calls it"
    );
}
