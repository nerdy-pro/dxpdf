# Prebuilt static libraries

`libdxpdf.a` per supported `GOOS_GOARCH`, built from `src/capi.rs` (the
crate's `capi` feature) via `cargo build --release --features capi`. Skia
is statically linked in via `skia-safe`'s `embed-freetype` feature, which is
why each one is tens of MB.

**Committed on purpose**, so `go get github.com/nerdy-pro/dxpdf/go` works
with no separate fetch step — see `go/README.md`'s Install section. The
tradeoff this accepts: every platform's binary ships in every `go get`
regardless of which one you actually need, and git never shrinks once a
version lands in history. Measured against the two actual limits this
bumps into: the Go module proxy caps a module zip at 500 MB total, and all
four platforms together (~420 MB) fit under that; GitHub caps a single git
blob at 100 MB, which linux/amd64 and linux/arm64 (~120-127 MB unsplit)
don't — see "Splitting" below for how those two clear it.

**Not Git LFS**, on purpose, despite it being GitHub's own suggested fix
for an oversized file: `go get`'s module fetch reads raw git blobs the way
the Go module proxy builds its zip, not through an LFS-aware checkout, so
an LFS-tracked file resolves to its tiny pointer stub instead of real
content unless the fetching machine has `git-lfs` installed *and* the
fetch path happens to invoke it — confirmed broken by golang/go's own
issue tracker (#47241, #39720), not just untested here. That would silently
defeat the one thing committing these libraries is for.

**Kept honest by `tests/capi_lib_freshness.rs`** (in the main repo, runs as
part of `cargo test --all`): it hashes every input that can change these
bytes — `src/`, `Cargo.toml`, `Cargo.lock`, `rust-toolchain.toml` — and
compares against `SOURCE_HASH`, so a source change that isn't followed by a
rebuild-and-recommit here fails CI rather than silently shipping a stale
library.

## Regenerating

For each platform:

```sh
cargo build --release --features capi [--target <triple>]
cp target/[<triple>/]release/libdxpdf.a go/internal/capi/lib/<os>_<arch>/
```

Then, for linux/amd64 and linux/arm64 specifically, split the result (see
"Splitting" below) — their unsplit archive is over GitHub's 100 MB limit.

Once every platform is rebuilt (and the two Linux ones split):

```sh
UPDATE_CAPI_LIB_HASH=1 cargo test --test capi_lib_freshness
```

and commit the refreshed libraries alongside the new `SOURCE_HASH`. Rebuild
*all four* together — the hash covers the whole source tree, not a single
platform, so restamping it after updating only one silently vouches for the
other three being current when they aren't.

## Splitting (linux/amd64, linux/arm64 only)

Neither symbol-stripping (~3-7% smaller, measured on both platforms) nor
LTO (which made the archive *larger* — it embeds LLVM bitcode for a
downstream linker to exploit, but cgo's linker is Go's own and never will)
gets these two anywhere near under 100 MB. The size is ~2,150 object files
spread across Skia/HarfBuzz/ICU/fontcull with no single component worth
cutting, so `scripts/split_capi_lib.py` splits the *archive*, not the
source: an `ar` file is just a container of `.o` members, and a linker
resolving `-lfoo` doesn't care whether all of a library's members live in
one archive or several.

```sh
python3 scripts/split_capi_lib.py go/internal/capi/lib/linux_amd64/libdxpdf.a
python3 scripts/split_capi_lib.py go/internal/capi/lib/linux_arm64/libdxpdf.a
```

Requires GNU `ar`/`ranlib` — run on Linux or in a Linux container matching
the archive's own architecture (macOS's `ar` operates on Mach-O, not ELF).
Replaces the unsplit `libdxpdf.a` with `libdxpdf_part1.a`,
`libdxpdf_part2.a`, ... in the same directory — see the script's own
module doc for why splitting works at all, and why the matching
`go/cgo_linux_*.go` files link every part inside a single
`-Wl,--start-group`/`--end-group` (not optional: GNU ld resolves archives
in one left-to-right pass each by default, and splitting one archive
scatters mutually-referencing object files across the pieces). Verified
end-to-end before being written up here — see AGENTS.md's Go-bindings note.
