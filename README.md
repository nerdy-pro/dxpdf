# docx-pdf

A lightweight Rust binary that parses DOCX files and renders them to PDF using Skia.

## Features

- Parses DOCX (Office Open XML) files from ZIP archives
- Text with formatting: bold, italic, underline, font size, font family, color, character spacing, superscript/subscript
- Paragraph properties: alignment, spacing (auto/exact/atLeast), indentation, tab stops, shading
- Named style resolution from `word/styles.xml` with `basedOn` inheritance chains
- Line breaks (`<w:br/>`) and tab characters (`<w:tab/>`) inside runs
- Tables with per-column widths, cell margins, merged cells (`gridSpan`/`vMerge`), dynamic row heights, cell shading, borders
- Three-pass table layout: measure→layout→paint with vMerge height distribution
- Inline images (PNG, JPEG, and other formats supported by Skia)
- Floating/anchored images with text wrapping, alignment, and percentage-based positioning (`wp14:pctPosVOffset`)
- Hyperlinks rendered as clickable PDF link annotations
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
- Warnings for unsupported DOCX features (via `log` crate, controlled by `RUST_LOG`)
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

- **Text runs** with direct formatting: bold (`w:b`), italic (`w:i`), underline (`w:u`), font size (`w:sz` in half-points), font family (`w:rFonts` — tries `ascii` then `hAnsi`), color (`w:color` as 6-digit hex), character spacing (`w:spacing w:val` in twips), run shading (`w:shd w:fill`), superscript/subscript (`w:vertAlign val="superscript|subscript"`)
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
- **Hyperlinks**: `w:hyperlink` with `r:id` resolved to URLs via relationships; text inside retains its formatting
- **Percentage-based image positioning**: `wp14:pctPosHOffset`/`wp14:pctPosVOffset` for Word 2010+ documents

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
- **Superscript/subscript**: `w:vertAlign` renders text at 58% font size with baseline shifted up (superscript) or down (subscript), while maintaining the original line height
- **Hyperlinks**: clickable PDF link annotations via Skia's `annotate_rect_with_url()`, covering the text bounding box
- **Unsupported feature warnings**: logs a warning (once per feature per document) for VML images, field codes, footnotes/endnotes, tracked deletions, strikethrough, and other unsupported elements
- **Render order**: cell shading → cell content (text, images) → cell borders, ensuring borders are always visible on top of content

## Running Tests

```bash
cargo test
```

The test suite includes 98 unit tests and 9 integration tests covering:

- **Layout engine**: measure_lines, fit_fragments, resolve_line_height, paragraph spacing, indentation, alignment (center/right), page breaks, paragraph shading across pages
- **Tables**: borders (2x2, gridSpan, alignment, tcW vs grid), vMerge (skip content, 3-row distribution, multi-column), cell margins, cell shading, empty/single-cell tables, table split across pages, after-table spacing
- **Lists**: bullet label rendering, decimal counter increment
- **Floats**: float adjustment shifts text, percentage-based positioning
- **Headers/footers**: header on every page, footer at bottom
- **Hyperlinks**: link annotations produced / not produced
- **Superscript/subscript**: font reduction and baseline shift (up/down)
- **Unit conversions**: twips, EMU, signed variants
- **XML parser**: paragraphs, runs, formatting, images, tables, sections, spacing, tabs, vertAlign, pctPosOffset, hyperlinks
- **Integration**: end-to-end DOCX→PDF conversion, error handling

Visual regression tests compare rendered PDFs against Word-generated references using pixel matching (see [VISUAL_COMPARISON.md](VISUAL_COMPARISON.md)).

## OOXML Feature Coverage

Validated against ISO 29500 (OOXML). **67 features fully implemented, 18 partial, 34 not implemented.**

### A. Text Formatting (w:rPr)

| Feature | Status |
|---|---|
| Bold, italic | ✅ with toggle support (`val="false"`) |
| Underline | ✅ font-proportional stroke width (thicker for bold) |
| Font size, family, color | ✅ |
| Superscript/subscript (`w:vertAlign`) | ✅ 58% font size, baseline shifted |
| Character spacing (`w:spacing`) | ✅ per-character advance expansion |
| Run shading (`w:shd`) | ✅ |
| Strikethrough (`w:strike`, `w:dstrike`) | ❌ |
| Highlighting (`w:highlight`) | ❌ |
| Caps, smallCaps | ❌ |
| Shadow, outline, emboss, imprint | ❌ |
| Hidden text (`w:vanish`) | ❌ |

### B. Paragraph Properties (w:pPr)

| Feature | Status |
|---|---|
| Alignment (left, center, right) | ✅ |
| Alignment (justify) | ⚠️ parsed, renders left-aligned |
| Spacing before/after | ✅ |
| Line spacing (auto/exact/atLeast) | ✅ |
| Indentation (left, right, first-line, hanging) | ✅ |
| Tab stops (left type) | ✅ |
| Tab stops (center, right, decimal) | ⚠️ parsed, all render as left |
| Paragraph shading | ✅ excludes before/after spacing |
| Paragraph borders (`w:pBdr`) | ✅ top, bottom, left, right |
| Keep with next, keep lines together | ❌ |
| Widow/orphan control | ❌ |

### C. Styles

| Feature | Status |
|---|---|
| Paragraph styles (`w:pStyle`) | ✅ |
| Character styles (`w:rStyle`) | ✅ including built-in Hyperlink |
| `basedOn` inheritance | ✅ field-by-field merging |
| Document defaults (`docDefaults`) | ✅ |
| Theme fonts | ✅ minor/major from `theme1.xml` |

### D. Tables

| Feature | Status |
|---|---|
| Grid columns, cell widths (dxa) | ✅ |
| Cell widths (pct, auto) | ⚠️ fall back to grid |
| Cell margins (per-cell/table/doc default) | ✅ three-level cascade |
| Merged cells (gridSpan, vMerge) | ✅ with height distribution |
| Row heights (minimum) | ✅ |
| Row heights (exact) | ⚠️ treated as minimum |
| Table borders (per-cell, per-table) | ✅ |
| Border styles (single) | ✅ |
| Border styles (double, dashed, dotted) | ⚠️ parsed, all render as single |
| Cell shading (solid) | ✅ |
| Cell shading (patterns) | ❌ |
| Cell vertical alignment (`w:vAlign`) | ❌ |
| Table alignment | ❌ |
| Nested tables | ✅ recursive |
| Table width (`w:tblW`) | ❌ |

### E. Images

| Feature | Status |
|---|---|
| Inline images (`wp:inline`) | ✅ PNG, JPEG, BMP, WebP via Skia |
| Floating images (`wp:anchor`) | ✅ |
| Position (offset, align, `wp14:pctPos`) | ✅ |
| Wrap (none, square, tight, through) | ✅ basic flow adjustment |
| Wrap (topAndBottom) | ❌ |
| Distance from text (`distL/R/T/B`) | ⚠️ fixed 4pt default |
| VML images (`w:pict`) | ❌ |

### F. Page Layout

| Feature | Status |
|---|---|
| Page size and orientation | ✅ |
| Page margins (top, right, bottom, left, header, footer) | ✅ |
| Section breaks (nextPage) | ✅ |
| Section breaks (continuous, even, odd) | ⚠️ treated as nextPage |
| Multi-column layout | ❌ |
| Page borders | ❌ |
| Document grid (`linePitch`) | ❌ |

### G. Headers/Footers

| Feature | Status |
|---|---|
| Default header/footer | ✅ with images, text, field codes |
| First page header/footer | ❌ |
| Even/odd page header/footer | ❌ |
| Per-section headers/footers | ❌ |

### H. Lists/Numbering

| Feature | Status |
|---|---|
| Bullet lists | ✅ custom characters |
| Numbered (decimal, letter, roman) | ✅ |
| List indentation (left, hanging) | ✅ |
| Multi-level lists | ⚠️ levels parsed, nesting limited |
| List restart numbering | ⚠️ start value parsed, restart not implemented |

### I. Fields

| Feature | Status |
|---|---|
| PAGE, NUMPAGES | ✅ with run properties from surrounding context |
| Hyperlinks (`w:hyperlink`) | ✅ clickable PDF link annotations |
| Unknown fields | ✅ cached value used as fallback |
| TOC, INDEX, MERGEFIELD, DATE | ❌ |

### J. Other

| Feature | Status |
|---|---|
| Footnotes/endnotes | ❌ warned |
| Comments | ❌ ignored |
| Tracked insertions | ⚠️ content passes through |
| Tracked deletions | ⚠️ warned, deleted text may appear |
| Text boxes, shapes | ❌ |
| SmartArt, charts, equations | ❌ |
| Embedded objects | ❌ |
| Right-to-left text | ❌ |
| Automatic hyphenation | ❌ words break at spaces and hyphens only |
| Fonts | ⚠️ system fonts with substitution table; Helvetica fallback |

## Dependencies

| Crate | Purpose |
|---|---|
| `quick-xml` | Fast, event-driven XML parsing |
| `zip` | Reading DOCX ZIP archives |
| `skia-safe` | PDF rendering, text measurement, and link annotations via Skia |
| `clap` | CLI argument parsing |
| `thiserror` | Ergonomic error types |
| `log` | Logging facade for unsupported feature warnings |
| `env_logger` | Log output controlled via `RUST_LOG` environment variable |

## Benchmarks

See [BENCHMARKS.md](BENCHMARKS.md) for conversion performance data.

## License

MIT
