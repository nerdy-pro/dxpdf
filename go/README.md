# dxpdf (Go bindings)

Go bindings for [dxpdf](https://github.com/nerdy-pro/dxpdf), a fast
DOCX-to-PDF converter powered by Skia. This package is a thin cgo layer over
the same Rust engine the CLI and Python package use — see
[`src/capi.rs`](../src/capi.rs) for the C ABI it binds to.

## Requirements

- `CGO_ENABLED=1` and a C compiler (cgo requirement).
- One of linux/amd64, linux/arm64, darwin/amd64, darwin/arm64 — not yet
  Windows.
- No tagged releases of this module yet, so `go get .../go@vX.Y.Z` won't
  resolve. Plain `go get github.com/nerdy-pro/dxpdf/go` (or `@main`, or a
  commit SHA to pin) works fine — Go falls back to a pseudo-version off the
  default branch tip.

## Install

The prebuilt static library for every supported platform is committed
under `internal/capi/lib/` (see that directory's own note on why, and
`tests/capi_lib_freshness.rs` in the main repo for how staleness against
the Rust source is caught), so there's no separate fetch step:

```sh
go get github.com/nerdy-pro/dxpdf/go
```

Or, working from a clone of the main repo:

```sh
cd go
go build ./...
```

## Usage

```go
package main

import (
	"log"

	"github.com/nerdy-pro/dxpdf/go"
)

func main() {
	pdfBytes, err := dxpdf.Convert(docxBytes)
	if err != nil {
		log.Fatal(err)
	}
	_ = pdfBytes

	// Or work with files directly, optionally overriding the embedded-image DPI:
	if err := dxpdf.ConvertFileWithOptions("input.docx", "output.pdf", 300); err != nil {
		log.Fatal(err)
	}
}
```

See [`dxpdf.go`](dxpdf.go) for the full API
(`Convert`/`ConvertWithOptions`/`ConvertFile`/`ConvertFileWithOptions`),
which mirrors the Python package's `convert`/`convert_file` one for one.
