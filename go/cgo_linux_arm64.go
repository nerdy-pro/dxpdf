//go:build linux && arm64

package dxpdf

// Verified: built the capi staticlib and ran `go test ./... -v` for real
// inside a `rust:1.95-bookworm` container (linux/arm64), linking against
// exactly this flag list with no adjustment needed. cgo_linux_amd64.go
// carries the same list, unverified on that architecture — see its comment.

/*
#cgo LDFLAGS: -L${SRCDIR}/internal/capi/lib/linux_arm64 -ldxpdf -lfontconfig -lstdc++ -lm -ldl -lpthread
*/
import "C"
