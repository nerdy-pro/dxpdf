//go:build linux && amd64

package dxpdf

// Same flags as cgo_linux_arm64.go, which were verified by actually
// building the capi staticlib and running `go test ./... -v` inside a
// `rust:1.95-bookworm` container (linux/arm64) — copied here rather than
// re-derived since the system-library list doesn't depend on the
// architecture, only the OS (AGENTS.md's documented Linux dependencies,
// libfontconfig1-dev/libfreetype-dev, plus dpkg-shlibdeps' own finding in
// scripts/verify_deb.py that `embed-freetype` needs no dynamic FreeType).
// Not itself run on real amd64 hardware; flag if it doesn't link.

/*
#cgo LDFLAGS: -L${SRCDIR}/internal/capi/lib/linux_amd64 -ldxpdf -lfontconfig -lstdc++ -lm -ldl -lpthread
*/
import "C"
