// Package dxpdf converts DOCX bytes to PDF bytes, via cgo against dxpdf's
// C ABI (../src/capi.rs, the crate's `capi` feature).
//
// This mirrors the Python bindings' shape (python/dxpdf/__init__.py) one for
// one — Convert/ConvertWithOptions/ConvertFile/ConvertFileWithOptions map to
// Python's convert/convert_file — so the two can be read side by side.
// Unlike Python, there is no C-side ConvertFile: file I/O is done here in
// pure Go on top of ConvertWithOptions, keeping the cgo boundary to a single
// call. See src/capi.rs's module doc for why, and for the concurrency and
// memory-ownership invariants this package relies on.
//
// Building a program that imports this package requires CGO_ENABLED=1 and a
// C compiler. The prebuilt static library for every supported GOOS/GOARCH is
// committed under internal/capi/lib/ (see that directory's own note on why,
// and tests/capi_lib_freshness.rs in the main repo for how staleness against
// the Rust source is caught), so no separate fetch step is needed. See
// README.md.
package dxpdf

/*
#include <stdlib.h>
#include "internal/capi/dxpdf.h"
*/
import "C"

import (
	"errors"
	"fmt"
	"os"
	"unsafe"
)

// DefaultImageDPI is the target resolution (pixels per inch) embedded
// raster images are downsampled to when not overridden, matching the Rust
// crate's own DEFAULT_IMAGE_DPI.
const DefaultImageDPI float32 = 220.0

// MinImageDPI is the floor ConvertWithOptions clamps a requested image DPI
// to, matching the Rust crate's own MIN_IMAGE_DPI.
const MinImageDPI float32 = 1.0

// Convert converts DOCX bytes to PDF bytes using [DefaultImageDPI].
func Convert(docxBytes []byte) ([]byte, error) {
	return ConvertWithOptions(docxBytes, DefaultImageDPI)
}

// ConvertWithOptions converts DOCX bytes to PDF bytes. imageDPI sets the
// target resolution (pixels per inch) embedded raster images are
// downsampled to; values below [MinImageDPI] (including non-finite ones)
// are clamped up to it.
func ConvertWithOptions(docxBytes []byte, imageDPI float32) ([]byte, error) {
	var docxPtr *C.uint8_t
	if len(docxBytes) > 0 {
		docxPtr = (*C.uint8_t)(unsafe.Pointer(&docxBytes[0]))
	}

	var outPDF C.DxpdfBuffer
	var outErr *C.char

	code := C.dxpdf_convert(docxPtr, C.uintptr_t(len(docxBytes)), C.float(imageDPI), &outPDF, &outErr)
	if code != 0 {
		defer C.dxpdf_free_error(outErr)
		message := "dxpdf: conversion failed"
		if outErr != nil {
			message = C.GoString(outErr)
		}
		return nil, errors.New(message)
	}
	defer C.dxpdf_free_buffer(outPDF)
	return C.GoBytes(unsafe.Pointer(outPDF.data), C.int(outPDF.len)), nil
}

// ConvertFile converts a DOCX file to a PDF file using [DefaultImageDPI].
func ConvertFile(input, output string) error {
	return ConvertFileWithOptions(input, output, DefaultImageDPI)
}

// ConvertFileWithOptions converts a DOCX file to a PDF file. imageDPI is as
// in [ConvertWithOptions].
func ConvertFileWithOptions(input, output string, imageDPI float32) error {
	docxBytes, err := os.ReadFile(input)
	if err != nil {
		return fmt.Errorf("dxpdf: failed to read %s: %w", input, err)
	}
	pdfBytes, err := ConvertWithOptions(docxBytes, imageDPI)
	if err != nil {
		return err
	}
	if err := os.WriteFile(output, pdfBytes, 0o644); err != nil {
		return fmt.Errorf("dxpdf: failed to write %s: %w", output, err)
	}
	return nil
}
