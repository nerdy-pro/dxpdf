# dxpdf (Go bindings)

Go bindings for [dxpdf](https://github.com/nerdy-pro/dxpdf), a fast
DOCX-to-PDF converter powered by Skia. This package is a thin cgo layer over
the same Rust engine the CLI and Python package use — see
[`src/capi.rs`](../src/capi.rs) for the C ABI it binds to.

## Requirements

- `CGO_ENABLED=1` and a C compiler (cgo requirement).
- The prebuilt static library for your platform, fetched once via `go
  generate` (see below). Currently built for linux/amd64, linux/arm64,
  darwin/amd64 and darwin/arm64 — not yet Windows.

## Install

```sh
go get github.com/nerdy-pro/dxpdf/go
cd $(go env GOMODCACHE)/github.com/nerdy-pro/dxpdf/go@<version> && go generate ./...
```

Or, working from a clone of the main repo:

```sh
cd go
go generate ./...   # downloads internal/capi/lib/<os>_<arch>/libdxpdf.a
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
