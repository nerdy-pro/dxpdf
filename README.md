# docx-pdf

A lightweight Rust binary that parses DOCX files and renders them to PDF using Skia.

## Features

- Parses DOCX (Office Open XML) files from ZIP archives
- Text with formatting: bold, italic, underline, font size, font family, color
- Paragraph properties: alignment, spacing, indentation, tab stops
- Line breaks (`<w:br/>`) and tab characters (`<w:tab/>`) inside runs
- Tables with per-column widths, cell formatting, borders, and dynamic row heights
- Inline images (PNG, JPEG, and other formats supported by Skia)
- Floating/anchored images with text wrapping
- Multiple sections with different page sizes and margins (e.g., portrait + landscape)
- Document defaults from `word/styles.xml` (font size, font family, paragraph spacing)
- Skia-based text measurement for accurate line wrapping and alignment
- Automatic pagination with page breaks
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
DOCX (ZIP) → Parse → Document Model (ADT) → Layout (Skia metrics) → Draw Commands → Skia PDF
```

### Modules

| Module | Description |
|---|---|
| `model` | Algebraic data types representing the document tree (`Document`, `Block`, `Inline`, etc.) |
| `parse/archive` | Extracts `word/document.xml`, relationships, media files, settings, and style defaults from the DOCX ZIP |
| `parse/xml` | Event-driven XML parser (state machine) that builds the document model |
| `render/layout` | Converts the document model into positioned draw commands using Skia font metrics for text measurement, with word wrapping, pagination, and float-aware text flow |
| `render/painter` | Translates draw commands into Skia canvas operations (`draw_str`, `draw_line`, `draw_image_rect`) to produce PDF output |
| `error` | Unified error type across all modules |

### Document Model

The core ADT uses two sum types as extension points:

- **`Block`** = `Paragraph` | `Table`
- **`Inline`** = `TextRun` | `LineBreak` | `Tab` | `Image`

Tables are recursive — `TableCell` contains `Vec<Block>`, mirroring the OOXML spec. Paragraphs can carry `FloatingImage` elements for anchored images with text wrapping. Type-safe newtypes (`RelId`, `FormatHint`) prevent accidental misuse of string fields.

## Supported Features

### Parsing

- **Text runs** with direct formatting: bold (`w:b`), italic (`w:i`), underline (`w:u`), font size (`w:sz` in half-points), font family (`w:rFonts` — tries `ascii` then `hAnsi`), color (`w:color` as 6-digit hex)
- **Toggle properties**: `w:b`, `w:i` support `val="false"` / `val="0"` to disable
- **Line breaks**: `<w:br/>` inside runs parsed as `Inline::LineBreak`, forcing a line break in layout
- **Tab characters**: `<w:tab/>` inside runs parsed as `Inline::Tab`, advancing to the next tab stop
- **Paragraph properties**: alignment (`w:jc` — left/start, center, right/end, both/justify), spacing before/after/line (`w:spacing`), indentation with left, right, first-line, and hanging (`w:ind`), custom tab stops (`w:tabs` > `w:tab` with `val` and `pos`)
- **Tables**: `w:tbl` with rows (`w:tr`), cells (`w:tc`), column grid (`w:tblGrid` > `w:gridCol`), cell widths (`w:tcW` with `type="dxa"` only — percentage widths fall back to grid)
- **Self-closing paragraphs**: `<w:p/>` parsed as empty paragraphs (produce blank lines)
- **Inline images**: `w:drawing` > `wp:inline` with dimensions from `wp:extent` (cx/cy in EMUs) and image data via `a:blip r:embed` relationship IDs
- **Floating images**: `w:drawing` > `wp:anchor` with horizontal/vertical position offsets (`wp:positionH`, `wp:positionV`, `wp:posOffset`), and text wrapping mode (`wrapTight`, `wrapSquare`, `wrapThrough`, `wrapNone` with `wrapText` attribute)
- **Section properties**: `w:sectPr` both mid-document (inside `w:pPr`) and final (inside `w:body`), including page size (`w:pgSz` with `w:w`, `w:h`) and page margins (`w:pgMar` with top, right, bottom, left)
- **Relationships**: `word/_rels/document.xml.rels` parsed to map relationship IDs to target file paths
- **Media files**: all files under `word/media/` extracted from the ZIP and linked to their relationship IDs
- **Document defaults** from `word/styles.xml` `docDefaults`:
  - Default font size (`rPrDefault` > `w:sz`)
  - Default font family (`rPrDefault` > `w:rFonts`)
  - Default paragraph spacing (`pPrDefault` > `w:spacing` — before, after, line)
- **Default tab stop interval** from `word/settings.xml` (`w:defaultTabStop`)
- **Text normalization**: literal newlines/carriage returns in `<w:t>` content are stripped (they are XML formatting artifacts; actual line breaks use `<w:br/>`)

### Rendering

- **Text**: rendered via Skia with the correct font family, size, bold/italic style, and color. Font fallback chain: requested family -> Helvetica -> system default
- **Text measurement**: uses Skia's `Font::measure_str()` for accurate text widths and `Font::metrics()` for line heights (ascent + descent + leading), replacing the earlier character-width estimation
- **Underlines**: drawn as 0.5pt lines 2pt beneath the text baseline
- **Images**: PNG, JPEG, BMP, WebP, and other formats decoded via `skia_safe::Image::from_encoded` and rendered with `draw_image_rect`. Images are scaled to fit within the content area
- **Tables**: per-column widths from `w:tblGrid` (scaled proportionally to page width), dynamic row heights computed from cell content, black 0.5pt borders on all cell edges, full paragraph layout within cells (formatting, alignment, images, tabs, line breaks)
- **Table cell alignment**: paragraph alignment (left, center, right) honored within cells
- **Word wrapping**: text split into separate word and space fragments; line breaks occur at space boundaries. Trailing spaces excluded from line width for correct alignment
- **Forced line breaks**: `<w:br/>` forces a new line within a paragraph
- **Tab stops**: tabs advance to the next custom tab stop position, or fall back to the document's default tab interval (typically 0.5 inch)
- **Floating image layout**: anchored images drawn at their offset position; text flow narrows around left-anchored floats. Floats in table cells are centered within the cell if the offset exceeds cell width
- **Pagination**: automatic page breaks when content exceeds the page content area
- **Multi-section documents**: each section can have its own page size and margins; section breaks trigger new pages with updated dimensions. Section configs are pre-collected and applied in order
- **Document defaults**: paragraphs without explicit spacing/font settings fall back to the document-level defaults from `word/styles.xml`
- **Whitespace collapsing**: runs of more than 2 consecutive spaces are collapsed to prevent manual-alignment whitespace from Word from disrupting layout

## Running Tests

```bash
cargo test
```

The test suite includes unit tests for the model, XML parser, layout engine (including tab stop resolution), and integration tests that create DOCX files in-memory and verify end-to-end conversion.

## Known Limitations

### Parsing

- **Named styles**: Only direct formatting (`w:rPr`, `w:pPr`) and document-level defaults (`docDefaults`) are supported. Named styles from `word/styles.xml` (e.g., "Heading 1", "Normal") are not resolved, so text that relies on style inheritance for its formatting will render with document defaults only.
- **Headers/Footers**: `word/header.xml` and `word/footer.xml` are not extracted or rendered.
- **Lists**: Numbered and bulleted lists (`w:numPr`, `w:abstractNum`) are not recognized. List items render as plain paragraphs without bullets, numbers, or indentation.
- **Footnotes/Endnotes**: Not supported.
- **Comments and tracked changes**: Ignored entirely.
- **Hyperlinks**: `w:hyperlink` elements are not parsed; linked text renders as unstyled plain text.
- **Fields and form controls**: Merge fields (`w:fldChar`), checkboxes, dropdowns, and other form elements are not rendered.
- **Legacy images**: VML images (`w:pict`, `v:imagedata`) are not supported; only DrawingML (`w:drawing`) is handled.
- **Text boxes and shapes**: Drawing shapes (`wsp:`, `v:shape`) and text boxes are not parsed.
- **SmartArt and charts**: `dgm:` and `c:chart` elements are not supported.
- **Strikethrough and other text effects**: `w:strike`, `w:dstrike`, `w:shadow`, `w:outline`, `w:emboss`, `w:imprint` are not parsed.
- **Background/highlight colors**: Run highlighting (`w:highlight`) and shading (`w:shd`) on paragraphs, runs, and cells are not rendered.
- **Paragraph borders**: `w:pBdr` is not parsed.
- **Multi-column layouts**: `w:cols` section properties are parsed but not used for layout.

### Rendering

- **Justify alignment**: Parsed from `w:jc val="both"` but not applied — text renders left-aligned. Implementing justify requires distributing extra space across word gaps.
- **Tab stop types**: Left, Center, Right, and Decimal tab stop types are parsed into the model but all tab stops are treated as left-aligned in layout. Center/Right/Decimal alignment at the stop position is not implemented.
- **Merged cells**: Row spans (`w:vMerge`) and column spans (`w:gridSpan`) in tables are not handled. Merged cells render as separate cells.
- **Table borders from DOCX**: All table borders are hardcoded to black 0.5pt. Per-cell and per-table border styles, colors, and widths from the DOCX (`w:tcBorders`, `w:tblBorders`) are ignored.
- **Cell shading/background**: Cell background colors (`w:shd`) are not rendered.
- **Cell vertical alignment**: Vertical alignment within cells (`w:vAlign`) is not applied; content is always top-aligned.
- **Cell margins/padding**: Cell padding is hardcoded to 4pt. Per-cell margins from `w:tcMar` are ignored.
- **Floating image positioning**: Only left-anchored floats with `relativeFrom="margin"` are handled. Right-anchored, centered, and complex multi-float positioning are not supported.
- **Line spacing modes**: Only the basic `w:line` value in twips is used. Exact and at-least spacing modes (`w:lineRule="exact"` / `"atLeast"`) are not distinguished from auto spacing.
- **Fonts**: System fonts are used via Skia's font manager. If a font specified in the DOCX is not installed on the system, Helvetica (or the system default) is used. Font embedding and subsetting depend on Skia's PDF backend.
- **Right-to-left text**: Not supported.
- **Hyphenation**: No automatic hyphenation; words break only at space boundaries.
- **Kerning and ligatures**: `w:kern` and `w14:ligatures` are parsed by Skia's font engine but not explicitly controlled.

## Dependencies

| Crate | Purpose |
|---|---|
| `quick-xml` | Fast, event-driven XML parsing |
| `zip` | Reading DOCX ZIP archives |
| `skia-safe` | PDF rendering and text measurement via Skia |
| `clap` | CLI argument parsing |
| `thiserror` | Ergonomic error types |

## License

MIT
