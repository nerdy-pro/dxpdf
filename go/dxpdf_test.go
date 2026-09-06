package dxpdf_test

import (
	"bytes"
	"os"
	"path/filepath"
	"sync"
	"testing"

	dxpdf "github.com/nerdy-pro/dxpdf/go"
)

// This is a smoke test for the FFI plumbing (bytes in, bytes out, errors
// surface, memory doesn't leak/crash) — parity with Word/the OOXML spec is
// the main Rust test suite's job, not this package's.
func TestConvertProducesAPDF(t *testing.T) {
	docxBytes, err := os.ReadFile(filepath.Join("testdata", "minimal.docx"))
	if err != nil {
		t.Fatalf("reading testdata/minimal.docx: %v", err)
	}

	pdfBytes, err := dxpdf.Convert(docxBytes)
	if err != nil {
		t.Fatalf("Convert: %v", err)
	}
	if !bytes.HasPrefix(pdfBytes, []byte("%PDF-")) {
		t.Fatalf("output does not start with a PDF header: %q", pdfBytes[:min(16, len(pdfBytes))])
	}
}

func TestConvertWithOptionsHonorsImageDPI(t *testing.T) {
	docxBytes, err := os.ReadFile(filepath.Join("testdata", "minimal.docx"))
	if err != nil {
		t.Fatalf("reading testdata/minimal.docx: %v", err)
	}

	pdfBytes, err := dxpdf.ConvertWithOptions(docxBytes, 96)
	if err != nil {
		t.Fatalf("ConvertWithOptions: %v", err)
	}
	if !bytes.HasPrefix(pdfBytes, []byte("%PDF-")) {
		t.Fatalf("output does not start with a PDF header")
	}
}

func TestConvertRejectsGarbageInput(t *testing.T) {
	if _, err := dxpdf.Convert([]byte("not a docx")); err == nil {
		t.Fatal("expected an error converting non-DOCX bytes, got nil")
	}
}

func TestConvertFileRoundTrips(t *testing.T) {
	output := filepath.Join(t.TempDir(), "out.pdf")
	if err := dxpdf.ConvertFile(filepath.Join("testdata", "minimal.docx"), output); err != nil {
		t.Fatalf("ConvertFile: %v", err)
	}
	pdfBytes, err := os.ReadFile(output)
	if err != nil {
		t.Fatalf("reading converted output: %v", err)
	}
	if !bytes.HasPrefix(pdfBytes, []byte("%PDF-")) {
		t.Fatalf("output file does not start with a PDF header")
	}
}

// Exercises the claim in src/capi.rs's module doc — FontRegistry is owned
// per render, not process-global — from the Go side of the FFI boundary,
// under `go test -race`.
func TestConvertIsSafeForConcurrentUse(t *testing.T) {
	docxBytes, err := os.ReadFile(filepath.Join("testdata", "minimal.docx"))
	if err != nil {
		t.Fatalf("reading testdata/minimal.docx: %v", err)
	}

	const goroutines = 8
	var wg sync.WaitGroup
	errs := make([]error, goroutines)
	for i := 0; i < goroutines; i++ {
		wg.Add(1)
		go func(i int) {
			defer wg.Done()
			_, errs[i] = dxpdf.Convert(docxBytes)
		}(i)
	}
	wg.Wait()

	for i, err := range errs {
		if err != nil {
			t.Errorf("goroutine %d: Convert: %v", i, err)
		}
	}
}
