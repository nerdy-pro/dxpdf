# dxpdf — Fast DOCX to PDF Converter in Rust

**Convert Microsoft Word DOCX files to PDF without Microsoft Office, LibreOffice, or any cloud API.**

dxpdf is an open-source, standalone DOCX-to-PDF conversion engine written in Rust and powered by [Skia](https://skia.org). It reads `.docx` files and produces high-fidelity PDF output — preserving text formatting, tables, images, headers, footers, hyperlinks, and page layout. Available as a CLI tool, a Rust library, and a Python package.

[![Crates.io](https://img.shields.io/crates/v/dxpdf)](https://crates.io/crates/dxpdf)
[![Documentation](https://img.shields.io/docsrs/dxpdf)](https://docs.rs/dxpdf)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Built by [nerdy.pro](https://nerdy.pro).

---

## Key Features

- **Blazing fast** — converts multi-page documents in under 100 ms on modern hardware
- **High fidelity** — Flutter-inspired measure → layout → paint pipeline with pixel-accurate baseline positioning
- **Type-safe** — compile-time dimensional type system (`Twips`, `Pt`, `Emu`) prevents unit mixing bugs
- **Standalone** — no Office installation, no LibreOffice, no external services needed
- **Cross-platform** — runs natively on macOS, Linux, and Windows
- **Three interfaces** — use as a CLI tool, Rust library (`use dxpdf;`), or Python package (`import dxpdf`)
- **ISO 29500 compliant** — validated against the Office Open XML specification

## Installation

### Command-Line Tool

```bash
cargo install dxpdf
```

### Rust Library

Add to your `Cargo.toml`:

```toml
[dependencies]
dxpdf = "0.2.5"
```

### Python Package

```bash
pip install dxpdf
```

## Usage

### CLI — Convert DOCX to PDF from the Terminal

```bash
dxpdf input.docx                  # produces input.pdf
dxpdf input.docx -o output.pdf    # specify output path
```

### Rust — Convert DOCX to PDF Programmatically

```rust
let docx_bytes = std::fs::read("document.docx")?;
let pdf_bytes = dxpdf::convert(&docx_bytes)?;
std::fs::write("output.pdf", &pdf_bytes)?;
```

You can also inspect or transform the parsed document model before conversion:

```rust
use dxpdf::{docx, model, render};

let document = docx::parse(&std::fs::read("document.docx")?)?;

for block in &document.blocks {
    match block {
        model::Block::Paragraph(p) => { /* inspect paragraph content */ }
        model::Block::Table(t) => { /* inspect table structure */ }
    }
}

let pdf_bytes = render::render(&document)?;
```

### Python — Convert DOCX to PDF in Python

```python
import dxpdf

# Bytes in, bytes out
pdf_bytes = dxpdf.convert(open("input.docx", "rb").read())

# File path to file path
dxpdf.convert_file("input.docx", "output.pdf")
```

## Supported DOCX Features

dxpdf handles the most common DOCX features found in real-world business documents, reports, and forms:

| Category | Features |
|---|---|
| **Text formatting** | Bold, italic, underline, highlighting, font size/family/color, character spacing, superscript/subscript, run shading |
| **Paragraphs** | Alignment (left/center/right), spacing (before/after/line with auto/exact/atLeast), indentation, tab stops, paragraph borders, paragraph shading |
| **Tables** | Column widths, cell margins (3-level cascade), merged cells (gridSpan + vMerge), row heights, borders, cell shading, nested tables |
| **Images** | Inline images (PNG, JPEG, BMP, WebP), floating/anchored images with alignment and percentage-based positioning |
| **Styles** | Paragraph and character styles, `basedOn` inheritance, document defaults, theme fonts |
| **Headers & footers** | Text, images, page numbers via PAGE/NUMPAGES field codes |
| **Lists** | Bullets, decimal, lower/upper letter, lower/upper roman numbering with counter tracking |
| **Hyperlinks** | Clickable PDF link annotations with URL resolution |
| **Page layout** | Multiple page sizes/margins, section breaks, portrait and landscape orientation |
| **Pagination** | Automatic page breaking, word wrapping, line spacing modes, floating image text flow |

## Performance Benchmarks

Benchmarked on Apple M3 Max with `hyperfine` (20 runs, 3 warmup):

| Document type | Pages | Conversion time | Memory usage |
|---|---|---|---|
| Short form with tables and images | 2 | **48 ms** | 20 MB |
| Multi-page report | 7 | **52 ms** | 24 MB |
| Image-heavy document (60+ images) | 24 | **353 ms** | 76 MB |

dxpdf processes most business documents in under 100 ms, making it suitable for batch processing, server-side conversion, and CI/CD pipelines.

## Building from Source

### Prerequisites

- Rust toolchain (1.70+)
- `clang` (required by `skia-safe` for building Skia bindings)
- **Linux only**: `libfontconfig1-dev` and `libfreetype-dev`

  ```bash
  sudo apt-get install -y libfontconfig1-dev libfreetype-dev
  ```

### Build

```bash
cargo build --release
```

The release binary will be at `target/release/dxpdf`.

### Run Tests

```bash
cargo test
```

## Architecture

dxpdf follows a **measure → layout → paint** pipeline inspired by Flutter's rendering model:

```
DOCX (ZIP) → Parse → Document Model → Measure → Layout → Paint → PDF
             Twips/Emu/HalfPoints       ←── Pt throughout ──→   Skia
```

Type-safe dimensions flow through the entire pipeline: OOXML units (`Twips`, `Emu`, `HalfPoints`) in the parsed model, `Pt` (typographic points) in layout, and `f32` only at the Skia rendering boundary.

Each layout element (paragraphs, table cells, headers/footers) goes through three phases:

1. **Measure** — collect text fragments, fit lines, produce draw commands with relative coordinates
2. **Layout** — assign absolute positions, handle page breaks, distribute heights (e.g., vertically merged cells)
3. **Paint** — emit draw commands at final positions (shading → content → borders)

### Module Overview

| Module | Purpose |
|---|---|
| `dimension` | Type-safe dimensional units (`Twips`, `HalfPoints`, `EighthPoints`, `Emu`, `Pt`) with compile-time unit safety |
| `geometry` | Spatial types (`Offset`, `Size`, `Rect`, `EdgeInsets`, `LineSegment`) — generic over unit, with Skia interop |
| `model` | Algebraic data types representing the full document tree (`Document`, `Block`, `Inline`, etc.) |
| `docx` | DOCX ZIP extraction, event-driven XML parser, style and numbering resolution |
| `render/layout` | Measure → layout → paint pipeline: fragment-based line fitting, paragraph layout, three-pass table layout, header/footer handling |
| `render/painter` | Skia canvas operations for PDF output |
| `render/fonts` | Font resolution with metric-compatible substitution (e.g., Calibri → Carlito, Cambria → Caladea) |

## OOXML Feature Coverage

Validated against ISO 29500 (Office Open XML). **37 features fully implemented, 9 partial, 13 planned.**

<details>
<summary>Full feature matrix (click to expand)</summary>

### Text Formatting (w:rPr)

| Feature | Status |
|---|---|
| Bold, italic | ✅ with toggle support |
| Underline | ✅ font-proportional stroke width |
| Font size, family, color | ✅ |
| Superscript/subscript | ✅ |
| Character spacing | ✅ |
| Run shading | ✅ |
| Strikethrough | ⚠️ parsed, not yet rendered |
| Highlighting | ✅ full ST_HighlightColor palette |
| Caps, smallCaps | ❌ |
| Shadow, outline, emboss, imprint | ❌ |
| Hidden text | ❌ |

### Paragraph Properties (w:pPr)

| Feature | Status |
|---|---|
| Alignment (left, center, right) | ✅ |
| Alignment (justify) | ⚠️ parsed, renders left-aligned |
| Spacing before/after, line spacing | ✅ auto/exact/atLeast |
| Indentation (left, right, first-line, hanging) | ✅ |
| Tab stops (left) | ✅ |
| Tab stops (center, right) | ✅ |
| Tab stops (decimal) | ⚠️ rendered as left-aligned |
| Paragraph shading | ✅ |
| Paragraph borders | ✅ with adjacent border merging, `w:space` offset |
| Keep with next, widow/orphan control | ❌ |

### Styles

| Feature | Status |
|---|---|
| Paragraph styles, character styles | ✅ |
| `basedOn` inheritance | ✅ |
| Document defaults, theme fonts | ✅ |

### Tables

| Feature | Status |
|---|---|
| Grid columns, cell widths (dxa) | ✅ |
| Cell widths (pct, auto) | ⚠️ fall back to grid |
| Cell margins (3-level cascade) | ✅ |
| Merged cells (gridSpan, vMerge) | ✅ |
| Row heights | ✅ min / ⚠️ exact treated as min |
| Table borders (per-cell, per-table) | ✅ |
| Border styles (single) | ✅ |
| Border styles (double, dashed, dotted) | ⚠️ render as single |
| Cell shading (solid) | ✅ |
| Cell shading (patterns), vertical alignment | ❌ |
| Nested tables | ✅ |

### Images

| Feature | Status |
|---|---|
| Inline images | ✅ PNG, JPEG, BMP, WebP |
| Floating images | ✅ offset, align, wp14:pctPos |
| Wrap modes | ✅ none/square/tight/through |
| VML images | ❌ |

### Page Layout

| Feature | Status |
|---|---|
| Page size and orientation | ✅ |
| Page margins (all 6) | ✅ |
| Section breaks (nextPage) | ✅ |
| Section breaks (continuous) | ✅ continues on current page |
| Section breaks (even, odd) | ⚠️ treated as nextPage |
| Multi-column, page borders, doc grid | ❌ |

### Headers & Footers

| Feature | Status |
|---|---|
| Default header/footer | ✅ |
| First page, even/odd, per-section | ❌ |

### Lists

| Feature | Status |
|---|---|
| Bullet, decimal, letter, roman | ✅ |
| Multi-level lists | ⚠️ levels parsed, nesting limited |

### Fields

| Feature | Status |
|---|---|
| PAGE, NUMPAGES | ✅ |
| Hyperlinks | ✅ clickable PDF annotations |
| Unknown fields | ✅ cached value fallback |
| TOC, MERGEFIELD, DATE | ❌ |

### Other

| Feature | Status |
|---|---|
| Footnotes/endnotes | ❌ warned |
| Comments, tracked changes | ❌ / ⚠️ |
| Text boxes, shapes, SmartArt, charts | ❌ |
| RTL text, automatic hyphenation | ❌ |

</details>

## Dependencies

| Crate | Purpose |
|---|---|
| [`quick-xml`](https://crates.io/crates/quick-xml) | Event-driven XML parsing |
| [`zip`](https://crates.io/crates/zip) | DOCX ZIP archive reading |
| [`skia-safe`](https://crates.io/crates/skia-safe) | PDF rendering, text measurement, link annotations |
| [`clap`](https://crates.io/crates/clap) | CLI argument parsing |
| [`thiserror`](https://crates.io/crates/thiserror) | Error types |
| [`log`](https://crates.io/crates/log) + [`env_logger`](https://crates.io/crates/env_logger) | Logging for unsupported features (`RUST_LOG=warn`) |
| [`pyo3`](https://crates.io/crates/pyo3) (optional) | Python bindings via maturin |

## Frequently Asked Questions

### How do I convert a DOCX file to PDF?

Install dxpdf with `cargo install dxpdf`, then run `dxpdf input.docx`. The PDF will be created in the same directory. You can also specify an output path with `-o output.pdf`.

### Does dxpdf require Microsoft Office or LibreOffice?

No. dxpdf is a standalone converter that reads DOCX files directly and renders PDF output using Skia. No Office installation or external service is needed.

### Can I use dxpdf as a library in my Rust or Python project?

Yes. In Rust, add `dxpdf` as a dependency and call `dxpdf::convert(&docx_bytes)`. In Python, install with `pip install dxpdf` and call `dxpdf.convert(bytes)` or `dxpdf.convert_file("input.docx", "output.pdf")`.

### What DOCX features are supported?

dxpdf supports text formatting, paragraphs, tables (including nested and merged cells), inline and floating images, styles with inheritance, headers/footers, lists, hyperlinks, section breaks, and automatic pagination. See the full [feature matrix](#ooxml-feature-coverage) above.

### How fast is dxpdf?

On Apple M3 Max, dxpdf converts a typical multi-page business document in under 100 ms. A 24-page image-heavy document takes about 350 ms. It is designed for batch processing and server-side use.

### What platforms does dxpdf support?

dxpdf runs on macOS, Linux, and Windows. On Linux, you need `libfontconfig1-dev` and `libfreetype-dev` installed.

## Contributing

Contributions are welcome. Please open an issue before submitting large PRs.

Built by [nerdy.pro](https://nerdy.pro).

## License

MIT
