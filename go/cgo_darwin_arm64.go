//go:build darwin && arm64

package dxpdf

/*
#cgo LDFLAGS: -L${SRCDIR}/internal/capi/lib/darwin_arm64 -ldxpdf -framework CoreFoundation -framework CoreGraphics -framework CoreText -framework AppKit -lc++
*/
import "C"
