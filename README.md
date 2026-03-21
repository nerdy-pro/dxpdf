# docx-pdf

A lightweight Rust binary that parses DOCX files and renders them to PDF using Skia.

## Features

- Parses DOCX (Office Open XML) files from ZIP archives
- Text with formatting: bold, italic, underline, font size, font family, color, character spacing
- Paragraph properties: alignment, spacing (auto/exact/atLeast), indentation, tab stops, shading
- Named style resolution from `word/styles.xml` with `basedOn` inheritance chains
- Line breaks (`<w:br/>`) and tab characters (`<w:tab/>`) inside runs
- Tables with per-column widths, cell margins, merged cells (`gridSpan`/`vMerge`), dynamic row heights, cell shading, borders
- Three-pass table layout: measure→layout→paint with vMerge height distribution
- Inline images (PNG, JPEG, and other formats supported by Skia)
- Floating/anchored images with text wrapping and alignment (left/center/right)
- Headers and footers with images and text
- Numbered and bulleted lists (decimal, lower/upper letter, lower/upper roman, bullet)
- Multiple sections with different page sizes and margins (e.g., portrait + landscape)
- Document defaults from `word/styles.xml` (font size, font family, paragraph spacing, cell margins, table borders)
- Theme font resolution from `word/theme/theme1.xml`
- Automatic font substitution for proprietary fonts (Calibri → Carlito, etc.)
- Skia-based text measurement for accurate line wrapping and alignment
- Flutter-inspired measure→layout→paint pipeline for all elements
- Automatic pagination with page breaks
- Centralized unit conversion system (`units.rs`) — spec-defined constants only
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

The converter follows a measure→layout→paint pipeline inspired by Flutter's rendering model:

```
DOCX (ZIP) → Parse → Document Model (ADT) → Measure → Layout → Paint → Skia PDF
```

Each layout element (paragraphs, table cells, headers/footers) goes through three phases:
1. **Measure**: Collect fragments, fit lines, produce draw commands with relative y-coordinates
2. **Layout**: Assign absolute positions, handle page breaks, distribute heights (e.g., vMerge spans)
3. **Paint**: Emit draw commands at final positions (shading → content → borders)

### Modules

| Module | Description |
|---|---|
| `units` | Centralized unit conversion constants and helpers (twips, EMUs, points) — single source of truth for all numeric constants |
| `model` | Algebraic data types representing the document tree (`Document`, `Block`, `Inline`, etc.) |
| `parse/archive` | Extracts `word/document.xml`, relationships, media files, settings, style defaults, and theme fonts from the DOCX ZIP |
| `parse/xml` | Event-driven XML parser (state machine) split into submodules: `helpers`, `properties`, `drawing`, `section` |
| `render/fonts` | Font substitution table mapping proprietary fonts to metric-compatible open-source alternatives |
| `render/layout` | Measure→layout→paint pipeline, split into: `measurer` (Skia font metrics), `fragment` (shared fragment collection, line fitting, and `measure_lines` — the single source of truth for fragment→command conversion), `paragraph` (paragraph layout with floats and page breaks), `table` (three-pass table layout: measure cells, distribute vMerge heights, paint), `header_footer` (header/footer rendering using shared `measure_lines`) |
| `render/painter` | Translates draw commands into Skia canvas operations (`draw_text`, `draw_line`, `draw_image`) to produce PDF output |
| `error` | Unified error type across all modules |

### Document Model

The core ADT uses two sum types as extension points:

- **`Block`** = `Paragraph` | `Table`
- **`Inline`** = `TextRun` | `LineBreak` | `Tab` | `Image`

Tables are recursive — `TableCell` contains `Vec<Block>`, mirroring the OOXML spec. Paragraphs can carry `FloatingImage` elements for anchored images with text wrapping. Type-safe newtypes (`RelId`, `FormatHint`) with `From` trait impls prevent accidental misuse of string fields. `Document` implements `Default` for clean construction with sensible defaults.

## Supported Features

### Parsing

- **Text runs** with direct formatting: bold (`w:b`), italic (`w:i`), underline (`w:u`), font size (`w:sz` in half-points), font family (`w:rFonts` — tries `ascii` then `hAnsi`), color (`w:color` as 6-digit hex), character spacing (`w:spacing w:val` in twips), run shading (`w:shd w:fill`)
- **Toggle properties**: `w:b`, `w:i` support `val="false"` / `val="0"` to disable
- **Line breaks**: `<w:br/>` inside runs parsed as `Inline::LineBreak`, forcing a line break in layout
- **Tab characters**: `<w:tab/>` inside runs parsed as `Inline::Tab`, advancing to the next tab stop
- **Paragraph properties**: alignment (`w:jc` — left/start, center, right/end, both/justify), spacing before/after/line (`w:spacing`), indentation with left, right, first-line, and hanging (`w:ind`), custom tab stops (`w:tabs` > `w:tab` with `val` and `pos`)
- **Tables**: `w:tbl` with rows (`w:tr`), cells (`w:tc`), column grid (`w:tblGrid` > `w:gridCol`), cell widths (`w:tcW` with `type="dxa"` only — percentage widths fall back to grid)
- **Merged cells**: horizontal spans (`w:gridSpan`) and vertical merges (`w:vMerge` with `val="restart"` / continue)
- **Row heights**: `w:trHeight` parsed as minimum row height (content can grow beyond it)
- **Cell margins**: per-table defaults (`w:tblCellMar`) and per-cell overrides (`w:tcMar`) with top/right/bottom/left in twips
- **Table borders**: `w:tblBorders` at the table level and `w:tcBorders` at the cell level, with `w:val` (style), `w:sz` (width in eighths of a point), `w:color` (hex or `auto`). Supports `top`/`bottom`/`left`/`right`/`insideH`/`insideV`
- **Cell shading**: `w:shd` with `w:fill` as hex color (e.g., `D9E2F3` for light blue)
- **Self-closing paragraphs**: `<w:p/>` parsed as empty paragraphs (produce blank lines)
- **Inline images**: `w:drawing` > `wp:inline` with dimensions from `wp:extent` (cx/cy in EMUs) and image data via `a:blip r:embed` relationship IDs
- **Floating images**: `w:drawing` > `wp:anchor` with horizontal/vertical position offsets (`wp:positionH`, `wp:positionV`, `wp:posOffset`), and text wrapping mode (`wrapTight`, `wrapSquare`, `wrapThrough`, `wrapNone` with `wrapText` attribute)
- **Section properties**: `w:sectPr` both mid-document (inside `w:pPr`) and final (inside `w:body`), including page size (`w:pgSz` with `w:w`, `w:h`) and page margins (`w:pgMar` with top, right, bottom, left)
- **Relationships**: `word/_rels/document.xml.rels` parsed to map relationship IDs to target file paths
- **Media files**: all files under `word/media/` extracted from the ZIP and linked to their relationship IDs
- **Theme fonts**: minor (body) and major (heading) font names extracted from `word/theme/theme1.xml` (`a:minorFont` / `a:majorFont` > `a:latin typeface`)
- **Document defaults** from `word/styles.xml`:
  - Default font size and family from `docDefaults/rPrDefault`
  - Default paragraph spacing from `docDefaults/pPrDefault`
  - Default cell margins from the "Normal Table" style's `tblCellMar`
  - Default table cell paragraph spacing from the "Table Grid" style's `pPr/spacing`
  - Default table borders from the "Table Grid" style's `tblBorders`
- **Default tab stop interval** from `word/settings.xml` (`w:defaultTabStop`)
- **Text normalization**: literal newlines/carriage returns in `<w:t>` content are stripped (they are XML formatting artifacts; actual line breaks use `<w:br/>`)
- **Named styles**: resolved from `word/styles.xml` with `basedOn` inheritance. Paragraph styles (alignment, spacing, indentation) and run styles (bold, italic, font size/family, color) are merged field-by-field with direct formatting winning
- **Headers and footers**: `w:headerReference`/`w:footerReference` in section properties resolved to `word/headerN.xml`/`word/footerN.xml`, with separate relationship files for images
- **Lists/numbering**: `w:numPr` references resolved via `word/numbering.xml` with `abstractNum`→`num` mapping. Supported formats: bullet, decimal, lowerLetter, upperLetter, lowerRoman, upperRoman. Level text, indentation, and hanging indent parsed
- **Paragraph shading**: `w:shd` with `w:fill` on paragraphs renders as a background rect covering the text area (excluding before/after spacing)

### Rendering

- **Text**: rendered via Skia with the correct font family, size, bold/italic style, and color
- **Font substitution**: when a font is not installed, metric-compatible open-source alternatives are tried automatically (e.g., Calibri → Carlito, Cambria → Caladea, Arial → Liberation Sans, Times New Roman → Liberation Serif). Falls back to Helvetica if no substitute is found. Full substitution table in `render/fonts.rs`
- **Text measurement**: uses Skia's `Font::measure_str()` for accurate text widths and `Font::metrics()` for line heights (ascent + descent + leading)
- **Underlines**: drawn beneath the text baseline
- **Images**: PNG, JPEG, BMP, WebP, and other formats decoded via `skia_safe::Image::from_encoded` and rendered with `draw_image_rect`. Images are scaled to fit within the content area
- **Tables**: per-column widths from `w:tblGrid` (scaled proportionally to page width), dynamic row heights computed from cell content, full paragraph layout within cells (formatting, alignment, images, tabs, line breaks)
- **Table borders**: resolved from per-cell overrides (`w:tcBorders`), table-level defaults (`w:tblBorders`), or document style defaults, with correct inner/outer edge distinction (`insideH`/`insideV` vs `top`/`bottom`/`left`/`right`). Border width, color, and style (`single`/`none`) are honored
- **Table cell alignment**: paragraph alignment (left, center, right) honored within cells
- **Merged cells**: `gridSpan` cells span multiple grid columns with correct widths; `vMerge` continuation cells skip content rendering and suppress top borders
- **Cell margins**: resolved per-cell (per-cell override > table default > document default) with separate top/right/bottom/left values
- **Cell shading**: background fill colors from `w:shd` rendered as filled rectangles behind cell content
- **Table cell spacing**: table style paragraph spacing (e.g., `after=0` from Table Grid style) applied inside cells, distinct from document-level paragraph spacing
- **Word wrapping**: text split into fragments at spaces and hyphens; line breaks occur at space and hyphen boundaries. Trailing spaces excluded from line width for correct alignment
- **Forced line breaks**: `<w:br/>` forces a new line within a paragraph
- **Tab stops**: tabs advance to the next custom tab stop position, or fall back to the document's default tab interval
- **Floating image layout**: anchored images drawn at their offset position; text flow narrows around left-anchored floats. Floats in table cells are centered within the cell if the offset exceeds cell width
- **Pagination**: automatic page breaks when content exceeds the page content area
- **Adjacent tables**: consecutive tables with no content between them render seamlessly without extra spacing
- **Character spacing**: `w:spacing w:val` expands or condenses inter-character spacing; each character drawn individually when non-zero
- **Line spacing**: `lineRule="auto"` applies the value as a multiplier (240=1.0×) to the font's natural line height; `lineRule="exact"` fixes the line height; `lineRule="atLeast"` sets a minimum. When no `w:line` is set, the font's natural line height from Skia metrics is used
- **Headers and footers**: rendered on every page from the document's default header/footer. Float images with alignment (left/center/right, top/center/bottom) supported. Header extent pushes body content down if it overflows margin_top
- **Lists**: bullet and numbered list labels rendered at the hanging indent position. Counters increment per list item. Supported: bullet, decimal, lower/upper letter, lower/upper roman
- **After-table spacing**: applies the document-default paragraph `spacing.after` after each table (skipped between adjacent tables), matching Word behavior
- **Multi-section documents**: each section can have its own page size and margins; section breaks trigger new pages with updated dimensions. Section configs are pre-collected and applied in order
- **Document defaults**: paragraphs without explicit spacing/font settings fall back to the document-level defaults from `word/styles.xml`
- **Whitespace handling**: `xml:space="preserve"` spaces are measured at their actual width; long space runs naturally overflow the line and wrap, matching Word's behavior. Leading spaces at line boundaries are skipped to prevent blank lines
- **Render order**: cell shading → cell content (text, images) → cell borders, ensuring borders are always visible on top of content

## Running Tests

```bash
cargo test
```

The test suite includes 86 unit tests and 9 integration tests covering:

- **Layout engine**: measure_lines, fit_fragments, resolve_line_height, paragraph spacing, indentation, alignment (center/right), page breaks, paragraph shading across pages
- **Tables**: borders (2x2, gridSpan, alignment, tcW vs grid), vMerge (skip content, 3-row distribution, multi-column), cell margins, cell shading, empty/single-cell tables, table split across pages, after-table spacing
- **Lists**: bullet label rendering, decimal counter increment
- **Floats**: float adjustment shifts text
- **Headers/footers**: header on every page, footer at bottom
- **Unit conversions**: twips, EMU, signed variants
- **XML parser**: paragraphs, runs, formatting, images, tables, sections, spacing, tabs
- **Integration**: end-to-end DOCX→PDF conversion, error handling

Visual regression tests compare rendered PDFs against Word-generated references using pixel matching (see [VISUAL_COMPARISON.md](VISUAL_COMPARISON.md)).

## Known Limitations

### Parsing

- **Footnotes/Endnotes**: Not supported.
- **Comments and tracked changes**: Ignored entirely.
- **Hyperlinks**: `w:hyperlink` with `r:id` is parsed and rendered as clickable PDF link annotations. Internal anchor links (`w:anchor`) are not yet supported.
- **Fields and form controls**: Merge fields (`w:fldChar`), checkboxes, dropdowns, and other form elements are not rendered. Page number fields (`PAGE`, `NUMPAGES`) in headers/footers are not evaluated.
- **Legacy images**: VML images (`w:pict`, `v:imagedata`) are not supported; only DrawingML (`w:drawing`) is handled.
- **Text boxes and shapes**: Drawing shapes (`wsp:`, `v:shape`) and text boxes are not parsed.
- **SmartArt and charts**: `dgm:` and `c:chart` elements are not supported.
- **Strikethrough and other text effects**: `w:strike`, `w:dstrike`, `w:shadow`, `w:outline`, `w:emboss`, `w:imprint`, superscript/subscript (`w:vertAlign`) are not parsed.
- **Run highlighting**: `w:highlight` is not rendered (run shading `w:shd` is supported).
- **Paragraph borders**: `w:pBdr` is not parsed.
- **Multi-column layouts**: `w:cols` section properties are parsed but not used for layout.
- **Per-section headers/footers**: Only the document-default header/footer is rendered; section-specific overrides (first page, even/odd) are not supported.

### Rendering

- **Justify alignment**: Parsed from `w:jc val="both"` but not applied — text renders left-aligned. Implementing justify requires distributing extra space across word gaps.
- **Tab stop types**: Left, Center, Right, and Decimal tab stop types are parsed into the model but all tab stops are treated as left-aligned in layout. Center/Right/Decimal alignment at the stop position is not implemented.
- **Table border styles**: Only `single` borders are rendered as solid lines. `double`, `dashed`, `dotted` styles are parsed but all render as solid lines.
- **Cell shading patterns**: Only solid fill colors are rendered. Shading patterns (e.g., `val="pct25"`) are not supported — only `val="clear"` with a `fill` color.
- **Cell vertical alignment**: Vertical alignment within cells (`w:vAlign`) is not applied; content is always top-aligned.
- **Floating image distance**: Float-to-text gap uses a fixed 4pt default; `wp:distL`/`wp:distR`/`wp:distT`/`wp:distB` attributes are not parsed yet.
- **Floating image positioning**: `wp14:pctPosVOffset`/`wp14:pctPosHOffset` percentage-based positioning is supported (relative to page dimensions). `relativeFrom` variants other than `page` are not yet distinguished.
- **Fonts**: System fonts are used via Skia's font manager with automatic substitution for common proprietary fonts. If neither the requested font nor any substitute is installed, Helvetica is used. Font embedding and subsetting depend on Skia's PDF backend.
- **Right-to-left text**: Not supported.
- **Hyphenation**: No automatic hyphenation. Words break at spaces and existing hyphens, but no new hyphenation points are inserted.
- **Kerning and ligatures**: `w:kern` and `w14:ligatures` are parsed by Skia's font engine but not explicitly controlled.

## Dependencies

| Crate | Purpose |
|---|---|
| `quick-xml` | Fast, event-driven XML parsing |
| `zip` | Reading DOCX ZIP archives |
| `skia-safe` | PDF rendering and text measurement via Skia |
| `clap` | CLI argument parsing |
| `thiserror` | Ergonomic error types |

## Benchmarks

See [BENCHMARKS.md](BENCHMARKS.md) for conversion performance data.

## License

MIT
