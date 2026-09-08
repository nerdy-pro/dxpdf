//! Guards `go/internal/capi/dxpdf.h` against drifting from `src/capi.rs`.
//!
//! The header is generated (`scripts/generate_capi_header.sh`) but committed,
//! the same tradeoff `src/i18n/data/icu_data.blob` and the man page make:
//! Go's build has no step that could regenerate it, so it has to be checked
//! in, and a checked-in generated file only stays honest if something
//! re-derives it and complains on mismatch. Requires `cbindgen` on `PATH`
//! (`cargo install cbindgen`) — skipped with a loud notice rather than
//! failed when it's absent, since unlike the man page or deb metadata this
//! needs an external tool `cargo test --all` cannot assume every contributor
//! has installed. CI installs it explicitly so the check is real there.

use std::path::PathBuf;
use std::process::Command;

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
}

#[test]
fn committed_header_matches_a_fresh_cbindgen_run() {
    let root = repo_root();

    if Command::new("cbindgen").arg("--version").output().is_err() {
        eprintln!(
            "skipping committed_header_matches_a_fresh_cbindgen_run: \
             `cbindgen` not found on PATH (`cargo install cbindgen`)"
        );
        return;
    }

    let committed = std::fs::read_to_string(root.join("go/internal/capi/dxpdf.h"))
        .expect("go/internal/capi/dxpdf.h is missing — run scripts/generate_capi_header.sh");

    let fresh_path = root.join("target").join("capi_header_freshness_check.h");
    let status = Command::new("cbindgen")
        .current_dir(&root)
        .args([
            "--config",
            "cbindgen.toml",
            "--output",
            fresh_path.to_str().unwrap(),
            "src/capi.rs",
        ])
        .status()
        .expect("failed to run cbindgen");
    assert!(status.success(), "cbindgen exited with {status:?}");

    let fresh = std::fs::read_to_string(&fresh_path).expect("cbindgen did not write its output");
    assert_eq!(
        committed, fresh,
        "go/internal/capi/dxpdf.h is stale — re-run scripts/generate_capi_header.sh and commit the result"
    );
}
