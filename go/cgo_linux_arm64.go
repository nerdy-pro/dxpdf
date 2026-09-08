//go:build linux && arm64

package dxpdf

// Verified: built the capi staticlib and ran `go test ./... -race` for real
// inside a native-arm64 `rust:1.95-bookworm` container, linking against
// exactly this flag list with no adjustment needed. cgo_linux_amd64.go
// carries the same system-library list, also verified — see its comment.
//
// libdxpdf is split into several archives (scripts/split_capi_lib.py) and
// linked inside a single `-Wl,--start-group`/`--end-group` — see
// cgo_linux_amd64.go's comment for why both are necessary, and
// go/internal/capi/lib/README.md for the full picture.

/*
#cgo LDFLAGS: -L${SRCDIR}/internal/capi/lib/linux_arm64 -Wl,--start-group -ldxpdf_part1 -ldxpdf_part2 -ldxpdf_part3 -Wl,--end-group -lfontconfig -lstdc++ -lm -ldl -lpthread
*/
import "C"
