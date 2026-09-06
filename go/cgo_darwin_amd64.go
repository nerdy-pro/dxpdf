//go:build darwin && amd64

package dxpdf

// Same flags as cgo_darwin_arm64.go, which were verified by actually
// linking and running `go test ./... -race` on darwin/arm64 — copied here
// rather than re-derived because the frameworks Skia's CPU/PDF backend
// needs (font matching, color management, no GPU) don't vary by Apple
// silicon vs. Intel. The committed darwin_amd64 libdxpdf.a itself was
// cross-compiled from darwin/arm64 (Apple's toolchain supports this
// directly — same OS, same SDK) and symbol-checked with `nm` for all five
// `dxpdf_*` C entry points, but not itself executed: nothing here can run
// an x86_64 macOS binary. Flag if it doesn't actually link/run on real
// amd64 hardware (see AGENTS.md's Go-bindings note).

/*
#cgo LDFLAGS: -L${SRCDIR}/internal/capi/lib/darwin_amd64 -ldxpdf -framework CoreFoundation -framework CoreGraphics -framework CoreText -framework AppKit -lc++
*/
import "C"
