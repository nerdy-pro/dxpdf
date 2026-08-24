# dxpdf — Fast DOCX to PDF Converter in Rust

**Convert Microsoft Word DOCX files to PDF without Microsoft Office, LibreOffice, or any cloud API.**

dxpdf is an open-source, standalone DOCX-to-PDF conversion engine written in Rust and powered by [Skia](https://skia.org). It reads `.docx` files and produces high-fidelity PDF output — preserving text formatting, tables, images, headers, footers, hyperlinks, and page layout. Available as a CLI tool, a Rust library, and a Python package.

[![Crates.io](https://img.shields.io/crates/v/dxpdf)](https://crates.io/crates/dxpdf)
[![Documentation](https://img.shields.io/docsrs/dxpdf)](https://docs.rs/dxpdf)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Built by [nerdy.pro](https://nerdy.pro).

---

## Key Features

- **Fast** — business documents convert in 55–170 ms depending on how their fonts resolve, a 171-page document in about 420 ms
- **High fidelity** — parse → resolve → layout → subset → paint pipeline with pixel-accurate baseline positioning
- **Compact output** — embedded fonts are subsetted to the glyphs actually used, so PDFs stay small
- **Type-safe** — compile-time dimensional type system (`Twips`, `Pt`, `Emu`) prevents unit mixing bugs
- **Standalone** — no Office installation, no LibreOffice, no external services needed
- **Cross-platform** — runs natively on macOS, Linux, and Windows
- **Three interfaces** — use as a CLI tool, Rust library (`use dxpdf;`), or Python package (`import dxpdf`)
- **Unicode-aware** — grapheme-correct segmentation, plus full-color emoji including ZWJ, skin-tone, keycap and flag sequences shaped through Skia's HarfBuzz
- **Internationalised** — UAX #14 line breaking (including Thai, Lao, Khmer and Burmese), UAX #9 bidirectional text, and CLDR-driven numbers and dates that follow the document's own `w:lang`
- **Packaged** — `cargo install`, `pip install`, or a `.deb` for Debian and Ubuntu
- **Tolerant** — documents written by tools other than Word still parse: a schema-invalid repeated element resolves the way Word resolves it rather than failing the file
- **ISO 29500 compliant** — validated against the Office Open XML specification

## Installation

### Command-Line Tool

```bash
cargo install dxpdf
```

### Debian / Ubuntu

Every [release](https://github.com/nerdy-pro/dxpdf/releases) ships a `.deb` for
`amd64` and `arm64`:

```bash
curl -LO https://github.com/nerdy-pro/dxpdf/releases/download/v0.5.1/dxpdf_0.5.1-1_amd64.deb
sudo apt install ./dxpdf_0.5.1-1_amd64.deb
```

Installs `dxpdf` to `/usr/bin` with a `dxpdf(1)` man page, and recommends
`fonts-liberation2` — metric-compatible with Arial, Times New Roman and Courier
New, so documents that ask for them lay out the way Word does.

Built on Debian 12, so it installs on Debian 12 and 13, Ubuntu 24.04 and newer,
and their derivatives. Not yet in the Debian archive itself
([#92](https://github.com/nerdy-pro/dxpdf/issues/92)): that needs a Skia that
Debian packages, and there is not one.

### Rust Library

Add to your `Cargo.toml`:

```toml
[dependencies]
dxpdf = "0.5.1"
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
dxpdf input.docx --image-dpi 300  # embed images at 300 DPI (default 220; range 1–2400)
```

Embedded raster images are downsampled to `--image-dpi` pixels per inch
(default **220**, matching Word). Raise it for print-quality output (e.g. `300`)
or lower it for smaller files (e.g. `96`); images are never upsampled past their
source resolution.

### Rust — Convert DOCX to PDF Programmatically

```rust
let docx_bytes = std::fs::read("document.docx")?;
let pdf_bytes = dxpdf::convert(&docx_bytes)?;
std::fs::write("output.pdf", &pdf_bytes)?;
```

To customize rendering — e.g. the embedded-image resolution (default 220 DPI) —
use `convert_with_options`:

```rust
use dxpdf::RenderOptions;

let options = RenderOptions::default().with_image_dpi(300.0);
let pdf_bytes = dxpdf::convert_with_options(&docx_bytes, &options)?;
```

You can also inspect or transform the parsed document model before conversion:

```rust
use dxpdf::{docx, model, render};

let document = docx::parse(&std::fs::read("document.docx")?)?;

for block in &document.body {
    match block {
        model::Block::Paragraph(p) => { /* inspect paragraph content */ }
        model::Block::Table(t) => { /* inspect table structure */ }
        model::Block::SectionBreak(props) => { /* inspect section properties */ }
    }
}

let pdf_bytes = render::render(document, &dxpdf::RenderOptions::default())?;
```

### Python — Convert DOCX to PDF in Python

```python
import dxpdf

# Bytes in, bytes out
pdf_bytes = dxpdf.convert(open("input.docx", "rb").read())

# File path to file path
dxpdf.convert_file("input.docx", "output.pdf")

# Customize embedded-image resolution (default 220 DPI)
pdf_bytes = dxpdf.convert(open("input.docx", "rb").read(), image_dpi=300)
dxpdf.convert_file("input.docx", "output.pdf", image_dpi=300)
```

## Supported DOCX Features

dxpdf handles the most common DOCX features found in real-world business documents, reports, and forms:

| Category | Features |
|---|---|
| **Text formatting** | Bold, italic, underline, highlighting, font size/family/color, character spacing, character scaling, superscript/subscript, run shading, run borders |
| **Paragraphs** | Alignment (left/center/right/justify/distribute), spacing (before/after/line with auto/exact/atLeast), indentation, tab stops (left/center/right/decimal/bar) incl. absolute-position tabs, paragraph borders, paragraph shading |
| **Tables** | Column widths, cell margins (3-level cascade), merged cells (gridSpan + vMerge), row heights, borders (single and double), cell shading, table styles with conditional formatting, nested tables, floating tables, row splitting across pages |
| **Images** | Inline images (PNG, JPEG, GIF, BMP, WebP, and single-bitmap EMF), floating/anchored images with alignment, wrapping, cropping and percentage-based positioning |
| **Styles** | Paragraph and character styles, `basedOn` inheritance, document defaults, theme fonts |
| **Fonts** | Embedded DOCX fonts, metric-compatible substitution, and subsetting so only used glyphs are embedded |
| **Text & emoji** | Grapheme-correct segmentation; full-color emoji including ZWJ, modifier, keycap and flag sequences, GSUB-shaped through Skia's HarfBuzz |
| **Shapes & text boxes** | DrawingML and VML shapes, shape text bodies with insets, anchoring and autofit, custom geometry with guide formulas |
| **Headers & footers** | Text, images, page numbers via PAGE/NUMPAGES field codes |
| **Lists** | Multi-level numbering — bullets, decimal, lower/upper letter, lower/upper roman, ordinal and spelled-out text — with counter tracking and picture bullets |
| **Navigation** | Clickable PDF link annotations with URL resolution, bookmarks and internal cross-references as named destinations, and a PDF outline built from heading levels |
| **Page layout** | Multiple page sizes/margins, section breaks, multi-column sections, portrait and landscape orientation |
| **Pagination** | Automatic page breaking, paragraph splitting across pages with keep-lines and widow/orphan control, word wrapping, line spacing modes, footnotes, endnotes, floating image text flow |
| **Internationalisation** | UAX #14 line breaking incl. the scripts written without spaces (Thai, Lao, Khmer, Burmese); UAX #9 bidirectional text with rule L4 mirroring; `w:lang`-driven decimal separators, DATE/TIME field pictures, and numbers spelled out in English, German, French and Spanish |

## Performance Benchmarks

Measured on Apple M3 Max with `hyperfine` (30 runs, 5 warmup) at **v0.5.1**,
against fixtures committed in `test-files/` so the numbers are reproducible.
Times are rounded to 5 ms — run-to-run spread on a normally loaded machine is
around ±10 ms, so smaller differences are not meaningful:

| Fixture | Pages | Input | Conversion time | Peak RSS |
|---|---|---|---|---|
| `sample-docx-files-sample3` | 3 | 34 KB | **170 ms** | 54 MB |
| `sample-docx-files-sample-4` | 7 | 10 KB | **170 ms** | 51 MB |
| `sample-docx-files-sample1` | 9 | 1.3 MB | **55 ms** | 40 MB |
| `sample-docx-files-sample4` | 171 | 14 MB | **420 ms** | 145 MB |

**Font resolution, not document size, decides what a conversion costs.** Notice
that the 9-page fixture converts in a third of the time the 3-page one does,
though it carries forty times the input. The difference is entirely in how its
fonts resolve.

The font registry is built in tiers and lazily. A document whose every font is
embedded or already present on the host never reaches the expensive tier —
`sample1` spends **~4 ms** there. One that has to fall back to the host
metadata index, matching on PostScript and style names, pays **~120–185 ms**,
once, and that then dominates everything else it does: on `sample3` the
registry is roughly five times parse, layout, subsetting and painting put
together. (Per-phase figures from `RUST_LOG=debug`, single runs with logging
on, so read them as proportions rather than as timings.)

So the useful question for a batch workload is not how large the documents are
but whether they name fonts the host has. Documents written by Word normally
embed or name available faces and land on the fast side; the slow side is worth
measuring for yourself before sizing anything.

To measure your own workload, run `cargo bench` for the Criterion suites, or
use the release binary with `RUST_LOG=debug` for a per-phase breakdown of
parse, resolve, registry, layout, subset and paint.

dxpdf is designed for batch processing, server-side conversion, and CI/CD
pipelines.

## Building from Source

### Prerequisites

- Rust 1.95.0 — pinned via `rust-toolchain.toml`, so `rustup` selects it automatically
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

The `subset-fonts` feature (font subsetting) is on by default; build with
`--no-default-features` to skip it.

### Run Tests

```bash
cargo test --all
```

## Architecture

dxpdf follows a **parse → resolve → layout → subset → paint** pipeline, with a measure-then-position model inspired by Flutter's rendering approach:

```
DOCX (ZIP) → Parse → Document Model → Resolve → Layout → Subset → Paint → PDF
             Twips/Emu/HalfPoints        ←──── Pt throughout ────→      Skia
```

Type-safe dimensions flow through the entire pipeline: OOXML units (`Twips`, `Emu`, `HalfPoints`) are `i64`-backed in the parsed model so they round-trip losslessly, layout works in `Pt` (typographic points), and raw `f32` appears only at the Skia rendering boundary.

1. **Parse** — declarative serde schemas over the DOCX XML parts, producing an immutable document model
2. **Resolve** — flatten the style cascade, split sections, pre-load images, generate shape geometry
3. **Layout** — measure text, fit lines, and position content into pages; runs first so total page count is known before headers/footers resolve PAGE/NUMPAGES
4. **Subset** — reduce each embedded typeface to the glyphs actually painted
5. **Paint** — emit draw commands in order (shading → content → borders) through Skia's PDF backend

### Module Overview

| Module | Purpose |
|---|---|
| `model::dimension` | Type-safe OOXML units (`Twips`, `HalfPoints`, `EighthPoints`, `Emu`, `ThousandthPercent`) with compile-time unit safety; the `Pt` rendering unit lives in `render::dimension` |
| `model::geometry` | Spatial types (`Offset`, `Size`, `Rect`, `EdgeInsets`, `PartialEdgeInsets`) — generic over unit, and free of any Skia dependency; `render::geometry` holds the `Pt`-specialized equivalents incl. `PtLineSegment` |
| `model` | Algebraic data types representing the full document tree (`Document`, `Block`, `Inline`, etc.) |
| `docx` | DOCX ZIP extraction, declarative serde-based XML parser for document, styles, numbering, theme, VML and DrawingML parts |
| `field` | OOXML field instruction parser (PAGE, NUMPAGES, HYPERLINK, TOC, …) |
| `render/resolve` | Style-cascade flattening, section splitting, image pre-loading, DrawingML shape geometry |
| `render/layout` | Fragment-based line fitting, paragraph layout, three-pass table layout, section stacking and pagination, header/footer handling |
| `render/subset` | Codepoint collection and per-typeface font subsetting before paint |
| `render/emoji` | Color-emoji pipeline — cluster classification, host typeface resolution, GSUB shaping through Skia's HarfBuzz, rasterization |
| `render/fonts` | Font resolution with embedded-font priority and metric-compatible substitution (e.g., Calibri → Carlito, Cambria → Caladea) |
| `render/painter` | Skia canvas operations for PDF output |

## OOXML Feature Coverage

Validated against ISO 29500 (Office Open XML). **77 entries fully implemented, 12 partial, 10 not yet supported.**

<details>
<summary>Full feature matrix (click to expand)</summary>

### Text Formatting (w:rPr)

| Feature | Status |
|---|---|
| Bold, italic | ✅ with toggle support |
| Underline | ✅ font-proportional stroke width |
| Font size, family, color | ✅ |
| Superscript/subscript | ✅ |
| Character spacing | ✅ §17.3.2.35 applied per UAX #29 grapheme cluster, so a combining mark is never separated from its base |
| Character scaling (`w:w` horizontal compression/expansion) | ✅ |
| Run shading | ✅ |
| Strikethrough | ⚠️ parsed, not yet rendered |
| Highlighting | ✅ full ST_HighlightColor palette |
| Caps, smallCaps | ⚠️ parsed, not applied at layout |
| Shadow, outline, emboss, imprint | ❌ |
| Hidden text (`w:vanish`) | ⚠️ a hidden run is removed before layout — its text, tabs and breaks take no space and the text either side closes up — resolved through the §17.7.2 cascade, so a character style can hide and `w:val="0"` can un-hide. Two gaps: a `w:sym`, `w:drawing` or `w:pict` in a hidden run still draws, and a hidden paragraph **mark** does not merge its paragraph into the next |
| Run borders (`w:bdr`) | ✅ |

### Paragraph Properties (w:pPr)

| Feature | Status |
|---|---|
| Alignment (left, center, right) | ✅ |
| Alignment (justify) | ✅ |
| Alignment (distribute) | ✅ §17.3.1.13 spare width shared between UAX #29 grapheme clusters, never inside one; not applied to a run that is shaped (see *Complex-script shaping*) |
| Spacing before/after, line spacing | ✅ auto/exact/atLeast |
| Indentation (left, right, first-line, hanging) | ✅ |
| Tab stops (left) | ✅ |
| Tab stops (center, right) | ✅ |
| Tab stops (decimal) | ✅ §17.18.85 zone anchored on the separator, which follows `w:lang` |
| Tab stops (bar) | ✅ §17.18.85 draws a vertical rule; does not position text |
| Tab leaders | ✅ §17.3.1.38 drawn in the formatting in effect at the tab |
| Absolute position tabs (`w:ptab`) | ✅ §17.3.1.30 left/center/right, margin-relative |
| Paragraph shading | ✅ |
| Paragraph borders | ✅ with adjacent border merging, `w:space` offset |
| Keep with next | ✅ incl. chain pre-flight and page-fill |
| Keep lines together | ✅ §17.3.1.14 |
| Widow/orphan control | ✅ §17.3.1.44 |
| Paragraph splitting across pages | ✅ per-page re-fit around floats, per-segment borders |

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
| Row heights (atLeast, exact) | ✅ §17.4.81 both rules honored |
| Table borders (per-cell, per-table) | ✅ incl. §17.4.66 conflict resolution |
| Border styles (single, double) | ✅ §17.4.38 double drawn as two sub-rules |
| Border styles (the other 24) | ⚠️ approximated by a solid line of the declared width and colour; warned once per style |
| Cell shading (solid) | ✅ |
| Cell shading (patterns) | ❌ parsed, fill colour only |
| Table styles, conditional formatting | ✅ §17.7.6 wholeTable, row/column bands, first/last row and column |
| Floating tables (`tblpPr`) | ✅ §17.4.58 anchors, spillover, `tblOverlap` |
| Vertical alignment (top / center / bottom) | ✅ incl. vMerge-aware bottom alignment |
| Row splitting across page breaks | ✅ §17.4.1 row content split at legal cut points; `cantSplit` honored |
| Repeating header rows | ✅ §17.4.49 |
| Nested tables | ✅ |

### Images

| Feature | Status |
|---|---|
| Inline images | ✅ PNG, JPEG, GIF, BMP, WebP via Skia |
| EMF images | ⚠️ single embedded bitmap (`EMR_STRETCHDIBITS`/`EMR_BITBLT`); full GDI record replay unsupported |
| WMF, SVG images | ❌ detected, not decoded |
| Image cropping (`a:srcRect`) | ✅ §20.1.10.48 |
| Floating images | ✅ offset, align, wp14:pctPos, page-parity mirroring |
| Wrap modes (none, square, topAndBottom) | ✅ |
| Wrap modes (tight, through) | ⚠️ approximated by the bounding box; no polygon-aware line fitting |
| VML images and shapes (`w:pict`) | ✅ inline and floating |
| `mc:AlternateContent` branch selection | ✅ MCE §M.1.2 |

### Page Layout

| Feature | Status |
|---|---|
| Page size and orientation | ✅ |
| Page margins (all 6) | ✅ |
| Section breaks (nextPage) | ✅ |
| Section breaks (continuous) | ✅ continues on current page |
| Section breaks (even, odd, nextColumn) | ⚠️ treated as nextPage |
| Multi-column sections | ✅ incl. splitting across unequal-width columns |
| Page borders (`w:pgBorders`) | ✅ line styles (single/double, others approximated solid), `offsetFrom`, `display`, `zOrder`; art borders declined |
| Document grid (`w:docGrid`) | ✅ line pitch for `lines`/`linesAndChars`/`snapToChars` — lines snap to whole grid slots, centered; `default`/absent type stays off; character pitch not applied |

### Headers & Footers

| Feature | Status |
|---|---|
| Default header/footer | ✅ |
| First page, even/odd, per-section | ✅ |

### Lists

| Feature | Status |
|---|---|
| Bullet, decimal, letter, roman | ✅ |
| Ordinal, cardinalText, ordinalText | ✅ §17.9.27 spelled out in English, German, French and Spanish (`Eins`, `Vingt et un`, `Veintiuno`, `Erste`, `1er`, `1.º`); other languages fall back to digits |
| Non-Latin sequences | ✅ §17.18.59 — Cyrillic, full-width/Devanagari/Thai/ideographic digits, circled and parenthesised decimals, kana (`aiueo`, `iroha`, both widths), hangul (`ganada`, `chosung`), Hebrew/Arabic/Devanagari/Thai alphabets, Chicago footnote symbols, heavenly stems, earthly branches and the sexagenary cycle, Hebrew and abjad numerals. A level whose §17.9.3 `w:rPr` names a covering font, as Word writes, uses it; one that does not now falls back per glyph to a host face that covers the sequence (see below) |
| Counting-system formats | ❌ §17.18.59 `chineseCounting`, `japaneseCounting`, `koreanCounting`, `thaiCounting`, `bahtText`, … render as decimal — each spells the number out in its own language rather than substituting digits |
| Picture bullets | ✅ §17.9.21 |
| Multi-level lists | ✅ `%1`–`%9` templates, per-level counters and resets, §17.9.8 `isLgl` |

### Fields

| Feature | Status |
|---|---|
| PAGE, NUMPAGES | ✅ evaluated per page |
| DATE, TIME | ✅ §17.16.4.2 evaluated against the `\@` picture at the moment of the render, with month, weekday and AM/PM names taken from the paragraph's own `w:lang`; a field naming no picture gets that locale's short date or time |
| Hyperlinks | ✅ clickable PDF annotations |
| All other fields | ✅ Word's cached result text is rendered, so a TOC or MERGEFIELD written by Word displays correctly but is not recomputed |
| Field instruction parser (`dxpdf::field`) | ✅ ~20 instructions parsed and evaluable as a library — REF, PAGEREF, SEQ, IF, MERGEFIELD, DOCPROPERTY, SYMBOL and more — but only PAGE, NUMPAGES, DATE and TIME are wired into rendering |

### Other

| Feature | Status |
|---|---|
| Footnotes | ✅ §17.11.23 separator, per-page reservation, split-aware |
| Endnotes | ✅ §17.11.2 roman superscript marks, collected at document end |
| Color emoji (ZWJ, modifier, keycap, flag sequences) | ✅ host-resolved color typeface, cross-run cluster reassembly, GSUB-shaped via Skia's HarfBuzz |
| Complex-script shaping — cursive joining | ✅ §17.3.2.30 a run whose script has positional forms (Unicode `Joining_Type` — Arabic, Syriac, N'Ko, Mongolian, Adlam …) is shaped through Skia's HarfBuzz; everything else keeps the cmap path unchanged |
| Complex-script shaping — Indic reordering | ❌ needs the spacing unit to become the shaped cluster, not just a new call site |
| Language (`w:lang`) | ⚠️ §17.3.2.20 drives the decimal-tab separator, the DATE/TIME picture names and the picture-less date and time defaults (all from CLDR, region-aware — `de-CH` and `de-DE` disagree correctly) and number-word spelling (English, German, French, Spanish; every other language gets digits) |
| Font subsetting | ✅ codepoint-driven, with shapeability validation |
| Per-glyph font fallback | ⚠️ a codepoint the resolved face cannot draw is drawn from a host face that can, chosen per UAX #29 grapheme cluster and carried through subsetting, so `ASCII ① ア` renders in full. Which face the host offers is its choice, so output is host-dependent (as color emoji already is), and no `w:lang` hint is passed yet — Han text may be given a face for the wrong language's glyph shapes. A codepoint no host face covers still draws nothing |
| Comments, tracked changes | ❌ |
| DrawingML fills, strokes, outer shadow | ⚠️ solid fills, strokes incl. dash patterns, and outer shadow; gradient and blip fills, blur, glow, reflection and soft edge are not rendered |
| DrawingML preset geometry | ⚠️ `line` and `rect`; `custGeom` fully evaluated incl. guide formulas |
| Text boxes (shape text bodies) | ✅ insets, vertical anchoring, `vertOverflow` clipping, `normAutofit` shrink |
| SmartArt, charts | ❌ |
| Bookmarks and internal cross-references | ✅ `w:bookmarkStart` → PDF named destinations; internal hyperlinks → GoTo link annotations |
| PDF outline sidebar (`/Outlines`) | ✅ §17.3.1.19 `w:outlineLvl` → structure-element headers; levels 7–9 clamp to `H6` (ISO 32000-1 stops there) and headings in headers, footers and notes are excluded |
| Line breaking | ✅ UAX #14 via ICU4X, per paragraph rather than per run, so a token split across `<w:r>` boundaries still breaks where the algorithm says. The four scripts UAX #14 hands to "complex context analysis" (Thai, Lao, Khmer, Burmese) get LSTM word boundaries; a token no rule may break is cut at the container edge rather than overflowing it |
| Bidirectional text (`w:bidi`, `w:rtl`) | ✅ §17.3.1.6 / §17.3.2.30 UAX #9 levels resolved per paragraph, reordered per line, with rule L4 mirroring; `w:jc` and `w:ind` resolve against the base direction |
| Bidirectional tab stops and numbering labels | ❌ §17.3.1.37 stop positions are not mirrored under `w:bidi`, so a line reorders within each tab-delimited segment and a label before its suffix tab stays at the left |
| `w:bidiVisual` (mirrored table columns) | ❌ not parsed |
| Automatic hyphenation | ❌ |

</details>

## Dependencies

| Crate | Purpose |
|---|---|
| [`quick-xml`](https://crates.io/crates/quick-xml) + [`serde`](https://crates.io/crates/serde) | Declarative XML parsing via serde deserializers |
| [`zip`](https://crates.io/crates/zip) | DOCX ZIP archive reading |
| [`skia-safe`](https://crates.io/crates/skia-safe) | PDF rendering, text measurement, link annotations, and HarfBuzz emoji shaping via the `textlayout` feature |
| [`unicode-segmentation`](https://crates.io/crates/unicode-segmentation), [`unicode-properties`](https://crates.io/crates/unicode-properties), [`unicode-normalization`](https://crates.io/crates/unicode-normalization) | Grapheme clusters, emoji properties, NFC normalization |
| [`unicode-bidi`](https://crates.io/crates/unicode-bidi) + [`unicode-bidi-mirroring`](https://crates.io/crates/unicode-bidi-mirroring) | UAX #9 embedding levels and rule L4 mirroring. Their Unicode tables are compiled in, so they add no locale data |
| [`unicode-joining-type`](https://crates.io/crates/unicode-joining-type) | Which scripts need shaping to be legible at all — the predicate that keeps HarfBuzz off Latin |
| `icu_*` ([`icu_segmenter`](https://crates.io/crates/icu_segmenter), [`icu_decimal`](https://crates.io/crates/icu_decimal), [`icu_datetime`](https://crates.io/crates/icu_datetime), [`icu_calendar`](https://crates.io/crates/icu_calendar), …) | ICU4X: UAX #14 line breaking, region-aware decimal separators, and localized date/time picture names. Locale data ships as one trimmed blob loaded through [`icu_provider_blob`](https://crates.io/crates/icu_provider_blob), not as each crate's built-in `compiled_data` |
| [`fontcull`](https://crates.io/crates/fontcull) (optional) | Font subsetting — `subset-fonts` feature, on by default |
| `fontcull-skrifa`, `fontcull-write-fonts`, `fontcull-read-fonts` + [`kurbo`](https://crates.io/crates/kurbo) | OpenType table reading and writing — baking a variable-font instance's coordinates into the bytes the PDF embeds |
| [`clap`](https://crates.io/crates/clap) | CLI argument parsing |
| [`thiserror`](https://crates.io/crates/thiserror) | Error types |
| [`log`](https://crates.io/crates/log) + [`env_logger`](https://crates.io/crates/env_logger) | Logging for unsupported features (`RUST_LOG=warn`) |
| [`rustc-hash`](https://crates.io/crates/rustc-hash) | Fast hasher for the per-render measurement cache |
| [`bitflags`](https://crates.io/crates/bitflags) | Compact flag sets in the document model |
| [`pyo3`](https://crates.io/crates/pyo3) (optional) | Python bindings via maturin |

## Frequently Asked Questions

### How do I convert a DOCX file to PDF?

Install dxpdf with `cargo install dxpdf`, then run `dxpdf input.docx`. The PDF will be created in the same directory. You can also specify an output path with `-o output.pdf`.

### Does dxpdf require Microsoft Office or LibreOffice?

No. dxpdf is a standalone converter that reads DOCX files directly and renders PDF output using Skia. No Office installation or external service is needed.

### Can I use dxpdf as a library in my Rust or Python project?

Yes. In Rust, add `dxpdf` as a dependency and call `dxpdf::convert(&docx_bytes)`. In Python, install with `pip install dxpdf` and call `dxpdf.convert(bytes)` or `dxpdf.convert_file("input.docx", "output.pdf")`.

### What DOCX features are supported?

dxpdf supports text formatting, paragraphs, tables (including nested, merged and floating tables with conditional formatting), inline and floating images, shapes and text boxes, styles with inheritance, headers/footers, multi-level lists, hyperlinks and a navigable PDF outline, footnotes and endnotes, section breaks, and automatic pagination. See the full [feature matrix](#ooxml-feature-coverage) above.

Notable gaps: Indic reordering, mirrored tab stops under `w:bidi`, automatic hyphenation, tracked changes and comments, and SmartArt and charts.

### How fast is dxpdf?

On an Apple M3 Max the committed fixtures convert in 55–170 ms, and a 171-page, 14 MB document in about 420 ms. Document size matters less than you would expect: what dominates a small conversion is how its fonts resolve, since a document naming faces the host has to look up in its metadata index pays roughly 120–185 ms once, where one whose fonts are embedded or already present pays about 4 ms. See [Performance Benchmarks](#performance-benchmarks) for measured figures and how to benchmark your own workload.

### What platforms does dxpdf support?

dxpdf runs on macOS, Linux, and Windows. On Linux, you need `libfontconfig1-dev` and `libfreetype-dev` installed.

## Used By

- <img src="https://www.google.com/s2/favicons?domain=nerdy.pro&sz=32" width="16" height="16" alt=""> [nerdy.pro](https://nerdy.pro)
- <img src="https://www.google.com/s2/favicons?domain=formtastic.de&sz=32" width="16" height="16" alt=""> [formtastic.de](https://formtastic.de)

## Contributing

Contributions are welcome. Please open an issue before submitting large PRs.

Build commands and project conventions are in [`AGENTS.md`](AGENTS.md).

Before opening a PR, run what CI runs:

```bash
cargo fmt --all -- --check
cargo clippy --all-targets -- -D warnings
cargo test --all
```

Built by [nerdy.pro](https://nerdy.pro).

## License

MIT
