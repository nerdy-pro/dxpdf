#!/usr/bin/env bash
# Regenerate go/internal/capi/dxpdf.h, the C header for the Go bindings' C
# ABI (src/capi.rs, enabled by the `capi` feature).
#
# Deliberately points cbindgen at src/capi.rs directly rather than the crate
# root — see cbindgen.toml's own comment for why (pointing at the whole
# crate walks every `pub` item in every module, most of which this crate
# never meant as a stable C API). No nightly toolchain is needed: this
# doesn't use cbindgen's `expand` (macro-expansion) path, since capi.rs has
# no macros for it to expand.
#
# Usage:
#
#     cargo install cbindgen
#     scripts/generate_capi_header.sh
#
# Regenerate and commit the header whenever capi.rs's public surface
# changes. tests/capi_header.rs re-runs this and diffs the result, so a
# stale committed header fails `cargo test` rather than silently drifting.
set -euo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")/.."

cbindgen --config cbindgen.toml --output go/internal/capi/dxpdf.h src/capi.rs
