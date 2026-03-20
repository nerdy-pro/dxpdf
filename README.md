# docx-pdf

A lightweight Rust binary that parses DOCX files and renders them to PDF using Skia.

## Features

- Parses DOCX (Office Open XML) files from ZIP archives
- Extracts text content, formatting, and table structure
- Renders to PDF via Skia with font styling, colors, and layout
- Simple CLI interface
- Cross-platform (macOS, Linux, Windows)

## Building

### Prerequisites

- Rust toolchain (1.70+)
- `clang` (required by `skia-safe` for building Skia bindings)

### Build

```bash
cargo build --release
```

The release binary will be at `target/release/docx-pdf`.

## Usage

```bash
# Basic conversion (outputs to same name with .pdf extension)
docx-pdf input.docx

# Specify output path
docx-pdf input.docx -o output.pdf
```

### As a library

```rust
use docx_pdf;

let docx_bytes = std::fs::read("document.docx")?;
let pdf_bytes = docx_pdf::convert(&docx_bytes)?;
std::fs::write("output.pdf", &pdf_bytes)?;
```

You can also work with the parsed document model directly:

```rust
use docx_pdf::{parse, model, render};

let docx_bytes = std::fs::read("document.docx")?;
let document = parse::parse(&docx_bytes)?;

// Inspect or transform the document model
for block in &document.blocks {
    match block {
        model::Block::Paragraph(p) => { /* ... */ }
        model::Block::Table(t) => { /* ... */ }
    }
}

let pdf_bytes = docx_pdf::convert_document(&document)?;
```

## Architecture

The converter follows a three-phase pipeline:

```
DOCX (ZIP) → Parse → Document Model (ADT) → Layout → Draw Commands → Skia PDF
```

### Modules

| Module | Description |
|---|---|
| `model` | Algebraic data types representing the document tree (`Document`, `Block`, `Inline`, etc.) |
| `parse/archive` | Extracts `word/document.xml` from the DOCX ZIP archive |
| `parse/xml` | Event-driven XML parser (state machine) that builds the document model |
| `render/layout` | Converts the document model into positioned draw commands with word wrapping and pagination |
| `render/painter` | Translates draw commands into Skia canvas operations to produce PDF output |
| `error` | Unified error type across all modules |

### Document Model

The core ADT uses two sum types as extension points:

- **`Block`** = `Paragraph` | `Table`
- **`Inline`** = `TextRun` | `LineBreak` | `Tab`

Tables are recursive — `TableCell` contains `Vec<Block>`, mirroring the OOXML spec.

## Running Tests

```bash
cargo test
```

The test suite includes unit tests for the model, XML parser, and layout engine, as well as integration tests that create DOCX files in-memory and verify end-to-end conversion.

## Known Limitations

### Parsing

- **Styles**: Only direct formatting (`w:rPr`, `w:pPr`) is supported. Document-level styles from `word/styles.xml` are not resolved, so text relying on named styles (e.g., "Heading 1") will render with default formatting.
- **Images**: Embedded images (`w:drawing`, `w:pict`) are not extracted or rendered.
- **Headers/Footers**: `word/header.xml` and `word/footer.xml` are not parsed.
- **Lists**: Numbered and bulleted lists (`w:numPr`) are not recognized. List items render as plain paragraphs without bullets or numbering.
- **Footnotes/Endnotes**: Not supported.
- **Comments and tracked changes**: Ignored.
- **Hyperlinks**: `w:hyperlink` elements are not parsed; linked text renders as plain text.
- **Section breaks and page size overrides**: Not handled. All pages use the default US Letter size (8.5" x 11").
- **Fields and form controls**: Merge fields, checkboxes, and other form elements are not rendered.

### Rendering

- **Text measurement**: Character widths are estimated (0.5 x font size) rather than measured via Skia font metrics. This means line wrapping and alignment are approximate, especially with proportional fonts.
- **Table column widths**: All columns are evenly distributed across the page width. Column width specifications from the DOCX are ignored.
- **Merged cells**: Row and column spans in tables are not handled.
- **Tab stops**: Tabs render as four spaces rather than honoring defined tab stop positions.
- **Line spacing**: Only basic before/after/line spacing is supported. Exact and at-least spacing modes are not distinguished.
- **Page margins**: Uses fixed 1-inch margins. Per-section margin overrides from the DOCX are not read.
- **Fonts**: Falls back to Helvetica if the specified font is not available on the system. Font embedding and subsetting depend on Skia's PDF backend behavior.
- **Right-to-left text**: Not supported.
- **Hyphenation**: No automatic hyphenation; words break only at spaces.

## Dependencies

| Crate | Purpose |
|---|---|
| `quick-xml` | Fast, event-driven XML parsing |
| `zip` | Reading DOCX ZIP archives |
| `skia-safe` | PDF rendering via Skia |
| `clap` | CLI argument parsing |
| `thiserror` | Ergonomic error types |

## License

MIT
