# docx-pdf

A lightweight Rust binary that parses DOCX files and renders them to PDF using Skia.

## Features

- Parses DOCX (Office Open XML) files from ZIP archives
- Text with formatting: bold, italic, underline, font size, font family, color
- Paragraph properties: alignment (left, center, right, justify), spacing, indentation
- Tables with borders and cell content
- Inline images (PNG, JPEG, and other formats supported by Skia)
- Floating/anchored images with text wrapping (wrap tight, wrap square)
- Multiple sections with different page sizes and margins (e.g., portrait + landscape in one document)
- Page size and margin support via `w:sectPr` (page size from `w:pgSz`, margins from `w:pgMar`)
- Word wrapping and automatic pagination
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
| `parse/archive` | Extracts `word/document.xml`, relationships, and media files from the DOCX ZIP archive |
| `parse/xml` | Event-driven XML parser (state machine) that builds the document model |
| `render/layout` | Converts the document model into positioned draw commands with word wrapping, pagination, and float-aware text flow |
| `render/painter` | Translates draw commands into Skia canvas operations to produce PDF output |
| `error` | Unified error type across all modules |

### Document Model

The core ADT uses two sum types as extension points:

- **`Block`** = `Paragraph` | `Table`
- **`Inline`** = `TextRun` | `LineBreak` | `Tab` | `Image`

Tables are recursive — `TableCell` contains `Vec<Block>`, mirroring the OOXML spec. Paragraphs can carry `FloatingImage` elements for anchored images with text wrapping. Type-safe newtypes (`RelId`, `FormatHint`) prevent accidental misuse of string fields.

## Supported Features

### Parsing

- **Text runs** with direct formatting: bold, italic, underline, font size (`w:sz`), font family (`w:rFonts`), color (`w:color`)
- **Paragraph properties**: alignment (`w:jc`), spacing before/after/line (`w:spacing`), indentation with first-line and hanging (`w:ind`)
- **Tables**: rows, cells, nested content (cells can contain paragraphs or nested tables)
- **Inline images**: `w:drawing` > `wp:inline` with dimensions from `wp:extent` and image data via relationship IDs
- **Floating images**: `w:drawing` > `wp:anchor` with position offsets (`wp:positionH`, `wp:positionV`) and text wrapping (`wrapTight`, `wrapSquare`, `wrapThrough`)
- **Section properties**: `w:sectPr` both mid-document (in `w:pPr`) and final (in `w:body`), including page size (`w:pgSz` with width, height, orientation) and page margins (`w:pgMar`)
- **Relationships**: `word/_rels/document.xml.rels` parsed to resolve image references
- **Media files**: all files under `word/media/` extracted and linked to their relationship IDs

### Rendering

- **Text**: rendered via Skia with correct font family, size, bold/italic style, and color
- **Underlines**: drawn as lines beneath text
- **Images**: PNG, JPEG, and other formats decoded and rendered via `skia_safe::Image::from_encoded`
- **Tables**: drawn with cell borders and text content
- **Word wrapping**: text fragments split at spaces and wrapped to fit available width
- **Pagination**: automatic page breaks when content exceeds the page content area
- **Floating image layout**: text flow adjusts around anchored images (left-anchored floats shift text to the right)
- **Multi-section documents**: each section can have its own page size and margins; section breaks trigger new pages with updated dimensions

## Running Tests

```bash
cargo test
```

The test suite includes unit tests for the model, XML parser, and layout engine, as well as integration tests that create DOCX files in-memory and verify end-to-end conversion.

## Known Limitations

### Parsing

- **Styles**: Only direct formatting (`w:rPr`, `w:pPr`) is supported. Document-level styles from `word/styles.xml` are not resolved, so text relying on named styles (e.g., "Heading 1") will render with default formatting.
- **Headers/Footers**: `word/header.xml` and `word/footer.xml` are not parsed.
- **Lists**: Numbered and bulleted lists (`w:numPr`) are not recognized. List items render as plain paragraphs without bullets or numbering.
- **Footnotes/Endnotes**: Not supported.
- **Comments and tracked changes**: Ignored.
- **Hyperlinks**: `w:hyperlink` elements are not parsed; linked text renders as plain text.
- **Fields and form controls**: Merge fields, checkboxes, and other form elements are not rendered.
- **Legacy images**: VML images (`w:pict`) are not supported; only DrawingML (`w:drawing`) is handled.
- **Floating image positioning**: Only left-anchored floats with `relativeFrom="margin"` are fully supported. Right-anchored and complex positioning are not handled.

### Rendering

- **Text measurement**: Character widths are estimated (0.5x font size for letters, 0.25x for spaces) rather than measured via Skia font metrics. Line wrapping and alignment are approximate, especially with proportional fonts.
- **Table column widths**: All columns are evenly distributed across the page width. Column width specifications from the DOCX are ignored.
- **Merged cells**: Row and column spans in tables are not handled.
- **Tab stops**: Tabs render as four spaces rather than honoring defined tab stop positions.
- **Line spacing**: Only basic before/after/line spacing is supported. Exact and at-least spacing modes are not distinguished.
- **Fonts**: Falls back to Helvetica if the specified font is not available on the system. Font embedding and subsetting depend on Skia's PDF backend behavior.
- **Right-to-left text**: Not supported.
- **Hyphenation**: No automatic hyphenation; words break only at spaces.
- **Whitespace collapsing**: Runs of more than 2 consecutive spaces are collapsed, which handles manual-alignment whitespace from Word but may affect intentionally spaced content.

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
