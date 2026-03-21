# dxpdf

**A fast, lightweight DOCX-to-PDF converter written in Rust, powered by Skia.**

Convert Microsoft Word `.docx` files to high-fidelity PDF documents — from the command line or as a Rust library. No Microsoft Office, LibreOffice, or cloud API required.

Built by [nerdy.pro](https://nerdy.pro).

## Why dxpdf?

- **Fast** — converts a 24-page document in ~115ms on Apple Silicon
- **Accurate** — Flutter-inspired measure→layout→paint pipeline with pixel-level fidelity
- **Standalone** — no external dependencies beyond Skia; no Office installation needed
- **Cross-platform** — runs on macOS, Linux, and Windows
- **Dual-use** — works as a CLI tool, Rust library (`use dxpdf;`), or Python package (`import dxpdf`)

## Quick Start

### Install

```bash
cargo install dxpdf
```

### Convert a file

```bash
dxpdf input.docx                  # outputs input.pdf
dxpdf input.docx -o output.pdf    # specify output path
```

### Use as a library

```rust
let docx_bytes = std::fs::read("document.docx")?;
let pdf_bytes = dxpdf::convert(&docx_bytes)?;
std::fs::write("output.pdf", &pdf_bytes)?;
```

You can also inspect or transform the parsed document model:

```rust
use dxpdf::{parse, model};

let document = parse::parse(&std::fs::read("document.docx")?)?;

for block in &document.blocks {
    match block {
        model::Block::Paragraph(p) => { /* ... */ }
        model::Block::Table(t) => { /* ... */ }
    }
}

let pdf_bytes = dxpdf::convert_document(&document)?;
```

### Use from Python

```bash
pip install dxpdf
```

```python
import dxpdf

# Bytes in, bytes out
pdf_bytes = dxpdf.convert(open("input.docx", "rb").read())

# File to file
dxpdf.convert_file("input.docx", "output.pdf")
```

## What's Supported

dxpdf handles the most common DOCX features used in real-world business documents:

| Category | Features |
|---|---|
| **Text** | Bold, italic, underline, font size/family/color, character spacing, superscript/subscript, run shading |
| **Paragraphs** | Alignment (left/center/right), spacing (before/after/line with auto/exact/atLeast), indentation, tab stops, paragraph borders, paragraph shading |
| **Tables** | Column widths, cell margins (3-level cascade), merged cells (gridSpan + vMerge with height distribution), row heights, borders, cell shading, nested tables |
| **Images** | Inline (PNG/JPEG/BMP/WebP), floating/anchored with alignment and percentage-based positioning |
| **Styles** | Paragraph styles, character styles (including built-in Hyperlink), `basedOn` inheritance, document defaults, theme fonts |
| **Headers/Footers** | Text, images, page numbers (PAGE/NUMPAGES field codes) |
| **Lists** | Bullets, decimal, lower/upper letter, lower/upper roman with counter tracking |
| **Hyperlinks** | Clickable PDF link annotations with URL resolution from relationships |
| **Sections** | Multiple page sizes/margins, section breaks, portrait/landscape |
| **Layout** | Automatic pagination, word wrapping at spaces and hyphens, line spacing modes, floating image text flow |

## Building from Source

### Prerequisites

- Rust toolchain (1.70+)
- `clang` (required by `skia-safe` for building Skia bindings)

### Build

```bash
cargo build --release
```

The release binary will be at `target/release/dxpdf`.

## Architecture

dxpdf follows a **measure→layout→paint** pipeline inspired by Flutter's rendering model:

```
DOCX (ZIP) → Parse → Document Model (ADT) → Measure → Layout → Paint → Skia PDF
```

Each layout element (paragraphs, table cells, headers/footers) goes through three phases:

1. **Measure** — collect fragments, fit lines, produce draw commands with relative coordinates
2. **Layout** — assign absolute positions, handle page breaks, distribute heights (e.g., vMerge spans)
3. **Paint** — emit draw commands at final positions (shading → content → borders)

### Modules

| Module | Description |
|---|---|
| `model` | Algebraic data types representing the document tree (`Document`, `Block`, `Inline`, etc.) |
| `parse` | DOCX ZIP extraction, event-driven XML parser, style/numbering resolution |
| `render/layout` | Measure→layout→paint pipeline: `fragment` (shared line fitting), `paragraph`, `table` (three-pass), `header_footer` |
| `render/painter` | Skia canvas operations for PDF output |
| `render/fonts` | Font resolution: tries requested font first, falls back to metric-compatible substitutes |
| `units` | OOXML unit conversions (twips, EMUs, half-points) — spec-defined constants only |

## Running Tests

```bash
cargo test
```

The test suite includes **104 unit tests** and **9 integration tests** covering layout, tables, lists, floats, headers/footers, hyperlinks, superscript/subscript, field codes, and end-to-end conversion.

Visual regression tests compare rendered PDFs against Word-generated references using pixel matching (see [VISUAL_COMPARISON.md](VISUAL_COMPARISON.md)).

## OOXML Feature Coverage

Validated against ISO 29500 (Office Open XML). **34 features fully implemented, 6 partial, 15 not implemented.**

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
| Strikethrough | ❌ |
| Highlighting | ❌ |
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
| Tab stops (center, right, decimal) | ⚠️ parsed, render as left |
| Paragraph shading | ✅ |
| Paragraph borders | ✅ with adjacent border merging |
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
| Section breaks (continuous, even, odd) | ⚠️ treated as nextPage |
| Multi-column, page borders, doc grid | ❌ |

### Headers/Footers

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

## Performance

Benchmarked on Apple M3 Max with `hyperfine` (20 runs, 3 warmup):

| Metric | Value |
|---|---|
| Mean conversion time | **113 ms** |
| Peak memory (RSS) | **19 MB** |
| Test document | 3-page form with 11 tables, 2 images, 2 sections |

See [BENCHMARKS.md](BENCHMARKS.md) for full history.

## Dependencies

| Crate | Purpose |
|---|---|
| `quick-xml` | Event-driven XML parsing |
| `zip` | DOCX ZIP archive reading |
| `skia-safe` | PDF rendering, text measurement, link annotations |
| `clap` | CLI argument parsing |
| `thiserror` | Error types |
| `log` + `env_logger` | Warnings for unsupported features (`RUST_LOG=warn`) |
| `pyo3` (optional) | Python bindings via `maturin` |

## Contributing

Contributions are welcome. Please open an issue before submitting large PRs.

Built by [nerdy.pro](https://nerdy.pro).

## License

MIT
