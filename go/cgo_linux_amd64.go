//go:build linux && amd64

package dxpdf

// Verified: built the capi staticlib and ran `go test ./... -race` for real
// inside a `rust:1.95-bookworm` container running under QEMU emulation
// (linux/amd64, ~8x slower to compile than the native linux/arm64 run, but
// otherwise identical) — same system-library flags as cgo_linux_arm64.go,
// and both are now confirmed rather than one copied from the other on the
// assumption that the list doesn't depend on architecture, only OS.
//
// libdxpdf is split into several archives (scripts/split_capi_lib.py) —
// see go/internal/capi/lib/README.md for why (linux/amd64's single archive
// is ~121 MB, over GitHub's 100 MB file limit, and neither stripping nor
// LTO closes anywhere near that gap). `-Wl,--start-group`/`--end-group` is
// required, not decorative: splitting one archive scatters mutually
// referencing object files across the parts (a class defined in part 1,
// a caller of one of its methods in part 2), and GNU ld only resolves
// archives in one left-to-right pass per archive by default — the group
// makes it keep re-scanning until nothing new resolves. Verified by
// actually linking and running `go test ./... -race` against the split
// archives, not assumed from how `--start-group` is documented to behave.

/*
#cgo LDFLAGS: -L${SRCDIR}/internal/capi/lib/linux_amd64 -Wl,--start-group -ldxpdf_part1 -ldxpdf_part2 -ldxpdf_part3 -Wl,--end-group -lfontconfig -lstdc++ -lm -ldl -lpthread
*/
import "C"
