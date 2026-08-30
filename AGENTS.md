# AGENTS.md

Instructions for AI coding agents working in this repository. `AGENTS.md` is the
tool-neutral filename, so every agent reads the same source of truth — `CLAUDE.md`
is a pointer to this file, not a second copy. Add guidance here, never there.

## Project

**dxpdf** — a fast DOCX-to-PDF converter in Rust, powered by Skia. Three interfaces: CLI tool, Rust library, and Python package (via PyO3/maturin).

## Build & Test Commands

```bash
cargo build                    # Debug build
cargo build --release          # Release build
cargo test --all               # Run all tests
cargo test <test_name>         # Run a single test by name
cargo bench                    # Run Criterion benchmarks
cargo clippy --all-targets -- -D warnings   # Lint (CI enforces zero warnings)
RUSTDOCFLAGS="-D warnings" cargo doc --no-deps   # Doc links (CI enforces zero warnings)
cargo fmt --all -- --check     # Format check
cargo fmt --all                # Auto-format
```

System dependencies (Linux): `libfontconfig1-dev`, `libfreetype-dev`. Requires `clang` for Skia. Toolchain is pinned to 1.95.0 via `rust-toolchain.toml`.

**Cargo features**: `subset-fonts` (default, via `fontcull`) runs the font-subsetting pass; `python` gates the PyO3 bindings. Build with `--no-default-features` to skip subsetting (the pass is `#[cfg(feature = "subset-fonts")]`-gated in `render_with_font_mgr`).

Benchmarking: `cargo bench` for Criterion benchmarks (`benches/convert_bench.rs`, `benches/parse_bench.rs`). `RUST_LOG=debug` for per-phase timing from CLI; `RUST_LOG=warn` surfaces unsupported-feature warnings logged by parse/layout.

CLI usage: `cargo run -- input.docx [-o output.pdf] [--image-dpi 300]` (release binary: `dxpdf`). `--image-dpi` sets the resolution embedded raster images are downsampled to — default 220 (matching Word), valid range 1–2400.

Python bindings: `maturin develop --features python` builds and installs into the active venv. The `python` feature is gated in `Cargo.toml`.

Note: `Cargo.toml` excludes dev-only paths — `test-files/`, `scripts/`, `output/` — from the published crate. Any local-only scratch corpora are excluded there too, so nothing local can be published by accident.

## Architecture

The converter follows a **parse → resolve → layout → (subset) → paint** pipeline, orchestrated in `src/lib.rs::convert()` (parse + render) and `src/render/mod.rs::render_with_font_mgr()` (resolve → layout → subset → paint).

1. **Parse** (`src/docx/`) — Declarative XML parsing of DOCX (ZIP of XML) via serde schemas on `quick_xml::de`. `zip.rs` handles ZIP extraction; `relationships.rs` parses rels. `parse/primitives/` holds shared schema atoms (unit wrappers, `HexColor`, `OnOff`, `ST_*` enum catalog). `parse/properties/` holds `PPr`/`RPr`/`TblPr`/`SectPr` schemas shared across body, styles, and numbering. Part-specific schemas live under `body.rs`+`body_schema.rs`, `drawing/` (DrawingML — `anchor`, `inline`, `picture`, `shape`, `fill`, `stroke`, `geometry`), `styles.rs`, `numbering.rs`, `theme/`, `notes.rs`, `settings.rs`, `vml/`. Each schema type is `pub(crate)` and suffixed `Xml`; `From<XxxXml> for ModelType` is the XML→domain seam. Outputs an immutable `Document` model.

2. **Model** (`src/model/`) — Pure data types with no parsing logic. `types/` contains the ADT: `Document` → `Vec<Block>` (`Paragraph | Table | SectionBreak`) → `Vec<Inline>`. `Inline` has 17 variants — text and drawing (`TextRun`, `Image`, `Pict`, `Symbol`, `AlternateContent`), fields (`Field`, `FieldChar`, `InstrText`), notes (`FootnoteRef`, `EndnoteRef`, `FootnoteRefMark`, `EndnoteRefMark`, `Separator`, `ContinuationSeparator`), and navigation (`Hyperlink`, `BookmarkStart`, `BookmarkEnd`). `dimension.rs` and `geometry.rs` provide the type-safe unit system. `src/field/` contains the OOXML field instruction parser (PAGE, TOC, HYPERLINK, etc.).

3. **Resolve** (`src/render/resolve/`) — Flattens style inheritance, splits sections, extracts font families, pre-loads images, resolves conditional formatting and colors. `shape_geometry/` generates DrawingML preset/custom shape paths (guide-formula evaluation under `guides.rs`). Produces a `ResolvedDocument` with fully-resolved styles and sections.

4. **Layout** (`src/render/layout/`) — Measures text with Skia font metrics (`measurer.rs`) and fits content into pages. `build/` orchestrates the constraint cascade: page → section → table → cell → paragraph (`block.rs`, `table.rs`, `floating.rs`, `convert.rs`, `list_label.rs`). `fragment/` breaks inline content into measurable units for line fitting, using the `unicode-*` crates for grapheme and script handling; emoji clusters are shaped GSUB-aware through Skia's own HarfBuzz (`render/emoji/shape.rs`). Three passes then run over the **finished** fragment vector, in a fixed order that `build/block.rs::resolve_paragraph_bidi` owns: `bidi.rs` (UAX #9 levels, splitting at level boundaries) → `fallback.rs` (issue #139 per-glyph font fallback, splitting at coverage boundaries) → `shape.rs` (which runs last because it re-measures against the resolved typeface, so it must see the family fallback may have changed). `paragraph/` handles line emission and paragraph borders. `table/` handles 3-pass table layout (`measure.rs` → `grid.rs` → `emit.rs`, with `borders.rs` for border resolution and `split.rs` for row splitting across pages). `section/` stacks blocks into pages: `stacker.rs` is the shared vertical-flow core used by *both* page and table-cell layout, while `layout.rs` owns the page-level algorithm (`layout_section`, keepNext chains, paragraph splitting, columns, footnotes). `float.rs` handles text wrapping around floating images. `header_footer.rs` renders headers/footers in a second pass (after total page count is known). Outputs `Vec<LayoutedPage>` of positioned `DrawCommand`s.

5. **Subset** (`src/render/subset/`, default `subset-fonts` feature) — Between layout and paint: `collect.rs` walks draw commands recording **codepoint** usage per resolved typeface (keyed by `TypefaceId`, so substituted and direct requests for the same face merge), `apply.rs` subsets each typeface via `fontcull`, splices the original `name` table back in, validates that every kept codepoint still shapes to a non-`.notdef` glyph, and swaps the bytes into the `FontRegistry`. Every failure mode is an explicit `SubsetOutcome` variant; a typeface that can't be subsetted keeps its original bytes.

6. **Paint** (`src/render/painter.rs`) — Iterates draw commands and emits PDF bytes via `skia_safe::pdf`. This is the only f32/Skia boundary. `skia_conv.rs` handles Pt-to-Skia conversions. `emf.rs` handles EMF (Enhanced Metafile) image rendering. `emoji/` is a separate color-emoji pipeline (UAX #29 / UTS #51 cluster classification in `cluster.rs`, host-OS color-typeface resolution in `resolve.rs`, GSUB shaping via Skia's HarfBuzz in `shape.rs`, Skia raster rasterization with a per-render cache in `raster.rs`).

### Key Design Patterns

- **Type-safe dimensions** (`src/model/dimension.rs`): Generic `Dimension<U>` parameterized by a unit marker (grep `impl Unit` for the full list — `Twips`, `Emu`, `HalfPoints`, `Pt`, etc.). `i64` storage for lossless OOXML round-tripping; `Pt` is the `f32` rendering unit. Prevents accidental unit mixing at compile time.

- **Generic geometry** (`src/model/geometry.rs`): `Offset<U>`, `Size<U>`, `Rect<U>`, `EdgeInsets<U>` parameterized over dimension units.

- **Spec-faithful ADT modeling**: All parsed values use typed enums/structs per OOXML spec sections. No raw strings for enumerated attributes — each gets a Rust enum. Typed identifiers (`RelId`, `StyleId`, `VmlShapeId`, `BookmarkId`) prevent mixing. Catch-all branches log warnings for unparsed elements; invalid enum values produce parse errors.

- **Two-pass rendering**: Layout runs first to determine total page count, then headers/footers are rendered in a second pass so PAGE/NUMPAGES fields resolve correctly.

- **Font resolution** (`src/render/fonts/`): A request is a name plus two **tri-state** §17.7.2 toggles (`Toggle::{Absent, Off, On}`), not a name plus a `FontStyle` — `Absent` asks for no weight, which is what lets a face name keep its own. `catalog.rs` turns the host font system and the DOCX's embedded fonts into one list of `FaceRecord`s, reading each face's own `name`/`OS/2`/`fvar`/`STAT` through the hand-written readers in `opentype/` (one table at a time via `copy_table_data`, never `to_font_data`). `resolve.rs` is a **pure function** over a request and a catalogue, running eight steps in order: embedded face, embedded family, host family, host face name, other metadata alias, parsed family+weight-word, metric-compatible substitute, host default. Everything down to step 5 is evidence the font asserts about itself; step 6 is the first guess. Ambiguous names are reported, not guessed. `resolve_exact`/`resolve_system_only` are the narrow variants the emoji pipeline needs. Because a request cannot know its own text, coverage is a *separate* question answered after layout has fragments to ask about — `layout/fragment/fallback.rs`, issue #139. It carries its answer as a **family name**, because `DrawCommand::Text` holds a name and both the painter and `subset::collect` re-resolve from it; `pin_system_face` is what makes that name authoritative, since a host's last-resort face (macOS `.LastResort`) is not reachable by name at all. `FontRegistry` is the single source of truth for typeface bytes and is owned **per render** — the subset pass mutates it in place after layout, so a process-wide (`thread_local!`) typeface cache would leak subsetted faces across documents and must not be reintroduced; the same rule binds the catalogue.

- **Text shaping & emoji**: Grapheme and script handling uses `unicode-segmentation`/`unicode-properties`/`unicode-normalization`; emoji clusters are shaped through **Skia's HarfBuzz** (`skia-safe`'s `textlayout` feature) driven by a `Typeface`, never by extracted font bytes — `Typeface::to_font_data()` on a 183 MB emoji font costs ~549 MB of unreleasable RSS, which is why the pure-Rust shaper it replaced is gone; color emoji is handled by the dedicated `render/emoji/` pipeline via typed ADTs (no font-name allowlists, no bundled emoji fonts — it resolves the host OS color-emoji typeface at render time).

## OOXML Reference

**There is no reference directory.** `docs/` was removed deliberately, page by page: a prose page describing behaviour drifts from the behaviour, and the second copy is the one that goes stale. WHY the engine makes a choice — which is generally not re-derivable from the source — belongs in the **module doc or the comment at the site that makes the choice**, where it is next to the thing it explains and moves when that moves. The "No doc yet" list below is now simply the entry-point list.

Working notes — designs, profiling analyses, branch reviews — are kept **local and uncommitted** (`/plans/`, gitignored), because they describe a point in time rather than current behaviour. Nothing tracked may link to them, and no code comment may cite them: a fresh clone does not have them. That cuts both ways, and it is the rule most easily broken by accident: **anything in `/plans/` that must outlive the work has to be moved out before the file is deleted** — into the code it describes, into this file, or into a GitHub issue. A note that exists only there is one `rm -rf` from gone, and nothing will warn you.

### Known-unimplemented work

Open engineering units are tracked as GitHub issues, not here — this file goes stale the moment one closes. Everything that is *not* a tracked unit is recorded where it applies: each ambiguity ECMA-376 cannot settle is stated in a comment at the site that makes the choice, saying what the choice is, why the spec does not decide it, and what evidence would. Grep for "Word reference render" to find them. Where a capability boundary is deliberate, the code says so at the boundary rather than deferring to the tracker — `SubsetOutcome::VariableInstanceNotBaked` states why a variable instance cannot be baked into embedded PDF bytes and names its two candidate routes; `register_embedded` states which faces of an embedded collection a given platform will open; `src/render/fonts/request.rs` states why `Toggle::Off` and `Toggle::Absent` select the same face today. One larger question is a decision rather than a gap: whether to take on a CLDR/ICU dependency for the i18n gaps tracked in issue #124.

### Decided — do not redo

Work deliberately *not* done. It is here rather than at a site because there is no site: no code was written, so a comment has nowhere to live. Reopen any of it with evidence, not with reasoning — each was closed against a measurement.

**Rejected optimisations.** Layout is 1–3.5 ms on 25 of 33 corpus documents, and that is the budget every one of these competed for: font-family interning · `Vec::with_capacity` seeding in the hot builders · `format!` per footnote number (which would also have cost an `itoa` dependency) · the per-run `RunProperties` clone. Reopen only with a profile showing layout is the bottleneck for the workload in question. `PTabLeader`'s pass-through enum is a separate "no": a distinct per-spec-type enum is what the spec-faithful ADT convention prescribes, so the extra layer is intentional — unlike `PTabAlignment`, which earns its enum by driving distinct layout math.

**Investigated and rejected.** Sharing `PackageContents` parts to remove the last package→media image copy: a parse-only probe peaks at 135 MB on the largest corpus document against 217 MB for the full render, so that duplicate never sets the peak. The keepNext double-layout and the floating-table double-measure: both measured in microseconds, and not worth the pagination risk.

**Inherited § citations are suspect until checked.** The retired findings cited `a:bodyPr/@vertOverflow` as §20.1.10.85, but [`drawing.rs`](src/model/types/drawing.rs) already annotates §20.1.10.85 as `ST_TextWrappingType` — the `wrap` attribute. Both cannot be right, and the conflict was resolved by *not* citing the disputed number: the code names the attribute (`a:bodyPr/@vertOverflow`, `ST_TextVertOverflowType`), which is unambiguous. Confirm any § against the spec before adding it, and never inherit one from a document.

**Worth knowing.** The `textlayout` feature costs binary size — the release binary is 15.9 → 28.2 MB (+77%) for ICU plus SkShaper/HarfBuzz. The same Skia build independently improved PDF font embedding: corpus output fell 28.98 → 27.22 MB (−6.1%), one document by −88%. That is Skia's PDF backend, not this engine's subset pass; don't attribute it here.

**When a fix is written from the symptom rather than the spec, it tends to be wrong.** The last attempt to do so prescribed *clipping text Word draws* (`@vertOverflow`). Three times a written plan was itself wrong — the MCE ADT, the `bar` tab's semantics, and a locale unit's premise that its ambiguity blocked implementation — and writing the tests from the spec first is what exposed it each time. Treat any plan, including one in this file, as a starting point to verify rather than a specification to implement.

**No doc yet** — start from the module docs at these entry points: character spacing and distributed alignment (`src/render/spacing.rs` — §17.3.2.35 and §17.3.1.13 share one unit, the UAX #29 grapheme cluster; the module doc says why it is that and not a shaped cluster), color-emoji pipeline (`src/render/emoji/mod.rs`), parse/serde schemas (`src/docx/parse/`, the `XxxXml` → domain seam), text shaping & fragments (`src/render/layout/fragment/`), per-glyph font fallback (`src/render/layout/fragment/fallback.rs` — why the fallback is carried as a name and not a resolved face, and why the early-out is load-bearing), paint & PDF emission (`src/render/painter.rs`), EMF images (`src/render/emf.rs`), numbering & list labels (`src/docx/parse/numbering.rs`, `src/render/layout/build/list_label.rs`), VML fallback (`src/docx/parse/vml/`).

## Test Organization

- **Unit tests**: `#[cfg(test)]` modules within source files.
- **Integration tests** (`tests/`): `integration.rs` (in-memory DOCX build + parse), `parse_test_files.rs` (parse real DOCX files from `test-files/`), `render_integration.rs` (layout + rendering validation), `emoji_e2e.rs` (color-emoji pipeline end-to-end), `header_footer_selection.rs` and `header_part_rels.rs` (header/footer resolution), `serde_spike.rs` (mixed-content parsing), `table_border_conflict.rs` (§17.4.66 nil-vs-none, conflict resolution), `table_row_height.rs` (§17.4.80/§17.4.84 row heights plus every malformed `w:vMerge` shape whose cell content used to vanish — a `restart` with nothing continuing it, a `continue` with no `restart` above it, and the two that only a **spanning** cell can express: a `continue` beside the column a `gridSpan` restart anchors on, and one below a `continue` wider than its own restart. Written with the bare `<w:vMerge/>` spelling, which is the only one the corpus contains), `table_auto_width.rs` (§17.4.63/§17.4.52 — how wide a `w:type="auto"` table may get; the guard is drawn at the paper edge and the file records why the obvious clamp to the text column is refuted by 40 tables of real Word output in the corpus), `table_style_whole_table.rs` (§17.7.6 `wholeTable` cascade), `table_style_cascade.rs` (§17.7.2/§17.7.4.3/§17.7.4.17 — what a table *style* declares reaching the table, asserted as parity against the same property written directly on the `<w:tbl>`; and, for the six `CT_TblPrBase` elements [MS-OI29500] §2.1.250(a)/§2.1.249(a) say Word does not read from a style, the *inverse* parity — declaring one there changes nothing while the same element on the `<w:tbl>` still applies. The band sizes are the case that runs the other way, §2.1.164(a), and `build_table` holds the whole split), `table_conditional_grid.rs` (§17.7.6/§17.4.16 — conditional regions follow the **grid column**, not the cell's index in its row, which `w:gridSpan`/`w:gridBefore` separate; the oracle is Word's own `w:cnfStyle` in `sample-docx-files-sample1.docx`, asserted against the fixture by `render::resolve::conditional`'s unit tests), `table_grid_seating.rs` (§17.4.48/§17.4.71 — the grid must have a column for every cell, and what `seat_every_cell` does when it does not. Both halves are pinned: a grid that *can* seat every cell is scaled and otherwise untouched, which is all 398 corpus tables and so the trap-detector for the repair's gate; and a grid that cannot grows rather than laying the overrun cell out at zero width. The deliberate non-repairs are pinned too — a row *shorter* than its grid, and a `w:gridAfter` overrunning it, both of which lose no content and are left alone), `table_geometry_sizing.rs` (§17.4.63/§17.4.48/§17.4.80 — the width `w:tblW` resolves to *on the page*: a `dxa` width scales the declared grid rather than truncating it, a `pct` one is that fraction of the width offered, and the full-width cell-margin extension applies to a left-aligned table and not to a centred one, which is the whole of `extends_for_alignment` and is asserted as the 10.8 pt between the same table both ways. Plus `w:trHeight` from the section layer: `exact` is a height whatever the row holds — the overflow it does not clip is left unasserted, being an open question — and `atLeast` is bracketed by two equalities, indistinguishable from `exact` under its minimum and from no rule at all over it), `table_geometry_paint.rs` (§17.18.2/§17.4.83 — the same geometry once it reaches the page rather than a `TableSlice`: a `w:val="double"` edge arrives as two lines of `w:sz / 3`, that far apart, asserted against the same table drawn `single` so the band is a *relation* between the styles and not four literals — and, incidentally, so a failure is about the style rather than about `w:sz`'s eighths of a point; and `w:vAlign` places content 0, half and all of a row's spare height, measured across two `hRule="exact"` rows of different heights holding identical content, so the content height cancels and no glyph metric has to be known. A row that *is* its content is the control: every other assertion is a difference, which a renderer that always bottom-aligned would satisfy), `table_border_corners.rs` (§17.4.38/§17.4.66 — the square where a vertical border crosses a horizontal one, which neither spec settles and which three separate reports have found unpainted. Asserts the *property* rather than any one shape of it: over every page of every committed fixture, no junction square is painted by nobody. It knows nothing about which cell a rect came from, which is the point — that knowledge is what each of the three defects had too little of. Two more tests run the same audit over the untracked `test-cases/` corpus and over the third reporter's own document, and skip when it is absent), `table_cell_content_box.rs` (§17.4.39/§17.4.66 against §17.4.41/§17.4.42 — where a cell's content box begins, and that the row is measured from *that* box. A border is drawn inside the cell, so the inset is `max(border, margin)` per side; the reported defect was that only the two horizontal sides were charged to the measurement while all four were charged to the placement, so a row whose top border outweighed its top cell margin overflowed its own box by the difference and its last line was painted through the bottom border. Every assertion is a difference between two renders of one document, so no glyph metric is pinned: `w:sz` 4→48 is 5.5pt, and that is the only literal. The controls are what keep the rule honest — a margin as thick as the border absorbs it, a bottom border grows the box without moving the content, and a bottom-aligned cell is lifted clear of it), `table_shading_seams.rs` (§17.4.33 — the pale hairline two abutting same-colour fills leave under a rasterizer that anti-aliases each one separately. Not a spec question: the ideal geometry of the pair is that of the single rect covering both, and only the raster differs. The audit is over the command stream — no two **consecutive** rects of one colour share an edge — because such a pair is always safe to fuse, and `coalesce_abutting_rects` fuses it. The file records what that leaves open, measured: §17.3.2.32 run shading interleaves fill/text/fill/text, so its pairs are not adjacent and settling them needs horizontal bounds on `DrawCommand`), `floating_table_pagination.rs` (§17.4.57 anchor/spillover termination), `font_resolution.rs` (§17.8 face resolution against the committed fixture fonts), `font_fallback.rs` (issue #139 per-glyph fallback — asserted structurally, never by face name, since which face covers a script is a property of the host).
- **Test helpers**: `make_docx()` and `simple_docx()` in `tests/integration.rs` build minimal in-memory DOCX archives.
- **Visual diffing**: `scripts/compare_pdfs.py` diffs generated PDFs against references. `scripts/verify_wheel.py` checks that FreeType is embedded in built wheels (run by the CI wheel job). `scripts/make_font_fixtures.py` rebuilds the font fixtures under `test-files/fonts/` (needs `fonttools`). **`scripts/verify_docx.py` checks that a `.docx` is a sound OPC package** — run it on any hand-built fixture before committing. This engine's parser is deliberately tolerant and will happily read a package Word refuses to open: three `issue-165-*` fixtures were rejected by Word with "unreadable content" while dxpdf, `textutil` and the whole test suite saw nothing wrong, over a single `.rels` part declaring the relationship-*type* URI (`.../officeDocument/2006/relationships`) as its `xmlns` instead of the relationship-*part* one (`.../package/2006/relationships`). `scripts/make_font_fallback_fixture.py` rebuilds `test-files/issue-139-minimal.docx`, `scripts/make_issue165_fixtures.py` the three issue #165 probes, and `scripts/make_hidden_text_fixture.py` the `w:vanish` fixture.

## Public API

- **Rust**: `convert(&[u8])` uses default options; `convert_with_options(&[u8], &RenderOptions)` is the full entry point. `RenderOptions` is a builder (`with_image_dpi`) with `DEFAULT_IMAGE_DPI = 220.0`; non-finite or non-positive requests are clamped up to `MIN_IMAGE_DPI`.
- **Python** (`--features python`, built with maturin via `pyproject.toml`): `convert(docx_bytes, image_dpi=220)` and `convert_file(input, output, image_dpi=220)`. Type stubs and the `py.typed` marker live in `python/dxpdf/`.

## Working in this repo

**Test corpus** — `test-files/` holds the committed DOCX fixtures, and is the corpus to use for reproductions and regression work:

| File | Exercises |
|---|---|
| `sample-docx-files-sample1`…`sample4` | General documents — text, tables, images, sections. `sample4` (14 MB) is the large-document/perf case |
| `sample-docx-files-sample-4`…`sample-6` | Small focused samples |
| `font_scaling.docx` | Font sizing and scaling |
| `sample-emoji.docx` | Color-emoji pipeline |
| `fonts/*.ttf`, `fonts/*.ttc` | §17.8 face resolution — built by `scripts/make_font_fixtures.py`, exercised by `tests/font_resolution.rs`. Regenerate rather than hand-edit; the build is deterministic |
| `comment-reference.docx` | A real Word package with `<w:commentRangeStart/End>` and `<w:commentReference>` plus `word/comments.xml` (and the modern-comments sibling parts, which stay unread) — since issue #154 the comment renders: range wash, narrow-margin balloon, Cyrillic author. Exercised by `tests/comments.rs` |
| `tracked-changes.docx`, `tracked-changes-final.docx` | Issue #154 — one body, two `w:revisionView` states: unaccepted `<w:ins>`/`<w:del>` by two authors (palette pin), a control paragraph, and a §17.3.2.37/§17.3.2.9 strike/dstrike paragraph that must stay struck in *both* views — the discriminator between a revision mark and strike formatting. Built by `scripts/make_tracked_changes_fixtures.py`; exercised by `tests/tracked_changes.rs` |
| `comments.docx`, `comments-hidden.docx` | Issue #154 — two comments by two authors (one with no initials), a multi-paragraph balloon, a range spanning a paragraph boundary, a control paragraph, and a 1.5in right margin for balloon room; the `-hidden` twin adds `w:revisionView w:comments="0"` and must draw none of it. Built by `scripts/make_comments_fixtures.py`; exercised by `tests/comments.rs` |
| `russian-numbering.docx` | `russianUpper` list numbering (А, Б, В…) |
| `numbering-direct-indent.docx` | Direct paragraph `w:ind` overriding the numbering level's indentation (suffix-tab position) |
| `centered-numbered-heading.docx` | Numbered heading with `jc=center` (suffix tab must not suppress alignment) |
| `issue-126-minimal.docx` | The reporter's document from issue #126 — a paragraph before an explicit page break. Renders 4 pages, matching the LibreOffice reference attached to that issue |
| `issue-139-minimal.docx` | The reproduction from issue #139 — one paragraph naming **no font anywhere**, mixing `① ア ๑` (which the spec-fallback face cannot draw) against `א` (which it can, so it is the control that must not move). Built by `scripts/make_font_fallback_fixture.py`; exercised by `tests/font_fallback.rs` |
| `issue-165-vmerge.docx`, `issue-165-cellspacing.docx`, `issue-165-floatv.docx` | The three issue #165 probes — documents authored so that each candidate reading of an ECMA-376 ambiguity predicts a *different measurement* off the rendered page (vMerge overflow distribution, `tblCellSpacing` at the table edges, vertical `inside`/`outside`). Built by `scripts/make_issue165_fixtures.py`, alongside `issue-165-cellspacing-scale.docx` — four otherwise-identical tables at four spacings, which asks the follow-up probe B turned up: whether Word's gap is the declared value or twice it, whether the spacing is carved out of `tblW` or added to it, and whether a row-level value supersedes the table-level one as §17.4.45 says. They answer nothing on their own — each needed a Word render to measure against, which is what #165 tracked. All three have now been measured; the readings and what they settled are recorded where the code acts on them (`table/grid.rs` for vMerge, `table/borders.rs` for cell spacing, `build/floating.rs`'s test module for the vertical mirror), and `issue-165-floatv.docx` is asserted end-to-end by `tests/floating_anchor_parity.rs` |
| `issue-159-minimal.docx` | The reporter's document from issue #159 — four `w:fldSimple` DATE fields whose `\@` pictures cover an escaped space, no escape, an escaped letter, and no picture. Each carries a deliberately wrong cached result (`CACHED`) so a renderer that fails to evaluate is obvious. Exercised by `tests/date_field_picture.rs` |
| `footer-path-wrap.docx` | A token UAX #14 cannot break (a Windows path) in a footer-table cell far narrower than it. Built by `scripts/make_footer_path_fixture.py`; exercised by `tests/footer_path_wrap.rs` |
| `hidden-text.docx` | §17.3.2 `w:vanish` in every position that resolves differently — a hidden run between two visible ones (with a control paragraph that has no hidden run, so the geometry is comparable without measuring a glyph), a character style that hides, a run un-hiding itself with `w:val="0"`, a paragraph whose every run is hidden, a hidden tab-and-break, and a hidden `w:sym` that **still draws** — the known limit, since the model drops a run's properties on its non-text children. Built by `scripts/make_hidden_text_fixture.py`; exercised by `tests/hidden_text.rs` |
| `duplicate-children.docx` | Every property bag (`pPr`, `rPr`, `tblPr`, `trPr`, `tcPr`) plus a VML `<v:roundrect>` repeating a child the schema allows once, each pair disagreeing so the fixture pins *which* wins, plus a style name whose byte 4 splits a codepoint. Built by `scripts/make_duplicate_children_fixture.py`; exercised by `tests/duplicate_children.rs` |

`tests/parse_test_files.rs` parses these and validates the resulting `Document`, so anything added here becomes part of the test suite. Add a new fixture when reproducing a bug — a committed fixture is what makes a fix verifiable by anyone.

`output/` holds generated PDFs. Scratch only, gitignored; never commit generated PDFs.

**Render-and-verify loop.** Rendering changes need visual confirmation, not just green tests:

```bash
cargo build --release
./target/release/dxpdf test-files/sample-docx-files-sample1.docx -o output/sample1.pdf

# Targeted before/after on a single page:
pdftoppm -png -r 150 -f 1 -l 1 output/sample1.pdf /tmp/after
magick compare -metric AE /tmp/before-1.png /tmp/after-1.png null:
```

`scripts/compare_pdfs.py` batch-diffs rendered output against `*_real.pdf` reference files (needs poppler + Pillow). It reads a local reference corpus that is not part of the repo, so it reports "No test pairs found" unless you have those references locally.

For any paint or subset change, pixel-diff before vs after — a passing test suite does not prove the output is unchanged.

**But a clean pixel diff does not prove it either, and one class is invisible to this loop.** `pdftoppm` composites two abutting fills cleanly; CoreGraphics — macOS Preview, Quick Look, Safari — does not, and leaves a pale hairline wherever the shared edge falls on a fractional device pixel. A seam reported in the `MediumShading2-Accent5` header of `sample-docx-files-sample1.docx` was present at every zoom in Preview and at *no* resolution in poppler. `tests/table_shading_seams.rs` audits the command stream for the pairs instead of looking for their pixels, which is the only check that can see them; when a rendering question is about compositing rather than geometry, rasterize with CoreGraphics (a 30-line `CGContextDrawPDFPage` program) rather than trusting poppler.

**Debian package** (issue #92) — `cargo deb` builds it from `[package.metadata.deb]` in `Cargo.toml`; `scripts/verify_deb.py` checks the result and `tests/packaging.rs` checks the inputs (man page in step with `--help`, `debian/changelog` in step with the version). CI builds amd64 on every PR and `deb.yml` builds both architectures per release.

It has to be built **inside a `debian:12` container**, and both halves of that matter. `depends = "$auto"` runs `dpkg-shlibdeps`, which resolves the binary against the packages of whatever distribution it runs on. And glibc is a floor, not a ceiling: built on `ubuntu-latest` the package requires glibc 2.39 and will not install on Debian 12 at all. Bookworm's 2.36 reaches Debian 12 and 13, Ubuntu 24.04+, and derivatives — moving that base later silently drops users who have already installed. Use `clang-19`, not bookworm's default clang 14, which cannot compile Skia m150's C++20 `<ranges>` against GCC 12's libstdc++.

```bash
docker run --rm -v "$PWD:/w" -w /w debian:12-slim bash -c '
  apt-get update && apt-get install -y --no-install-recommends \
    build-essential clang-19 libclang-19-dev ninja-build python3 curl \
    ca-certificates git pkg-config libfontconfig1-dev libfreetype-dev lintian
  curl -sSf https://sh.rustup.rs | sh -s -- -y --default-toolchain none
  . "$HOME/.cargo/env"; export CC=clang-19 CXX=clang++-19
  cargo install cargo-deb --locked && cargo deb
  python3 scripts/verify_deb.py target/debian/*.deb
  lintian --fail-on error,warning target/debian/*.deb'
```

Two things that are not obvious from the files. The `assets` list is explicit because cargo-deb's default set also picks up C-ABI dynamic libraries, and `crate-type = ["rlib", "cdylib"]` means a release build emits `libdxpdf.so` — the PyO3 extension body, which has no business in `/usr/lib`. And when testing an install in a Debian *container*, delete `/etc/dpkg/dpkg.cfg.d/docker` first: it sets `path-exclude /usr/share/man/*`, so dpkg drops the man page and the test appears to prove the package has none.

**How a change is built here.** Three conventions that the CI commands below do not enforce and that reviewers assume have happened:

- **Tests written from the spec, first, and watched fail.** Not written against current output, which only pins the bug. Where a change touches uncovered code, the characterization tests land first as their own commit.
- **A mutation check on every new test.** A test written alongside its implementation passes on the first run, which proves nothing. Break the implementation deliberately and confirm the right tests fail — if none does, the test is decoration.
- **Pixel-diff any change that moves geometry**, across `test-files/` + `test-cases/`, and *explain* every diff rather than merely observing it. A document that changes is either the fix working or a regression, and only reading the pixels says which.

**Before handing work back**, run what CI runs (`.github/workflows/ci.yml`):

```bash
cargo fmt --all -- --check
cargo clippy --all-targets -- -D warnings   # CI enforces zero warnings
cargo test --all
RUSTDOCFLAGS="-D warnings" cargo doc --no-deps   # CI enforces zero doc warnings
cargo build --no-default-features           # `subset-fonts` off still compiles
```

The doc check catches dangling `[`links`]`, links from public docs to private
items, and prose rustdoc reads as HTML (`Vec<Thing>` outside backticks). Link to
a private item with a plain code span, not `[`brackets`]`.

**Logging**: `RUST_LOG=debug` gives per-phase timing — parse/render/total from `convert`, then resolve, registry, layout, subset and paint from `render_with_font_mgr` — plus the font-resolution decision for every requested family; `RUST_LOG=warn` surfaces unsupported-feature warnings from parse and layout. Prefer these numbers to intuition. The registry build used to be the largest cost on an ordinary document — a fixed 78–95 ms on every render regardless of document size — because it indexed the whole host font system up front. It is now tiered and lazy: a document whose fonts are all present or embedded costs ~3 ms, and one that has to reach the metadata index costs ~105 ms, paid once. See `src/render/fonts/catalog.rs`'s module doc for the per-operation breakdown; `FontMgr::match_family` dominates, at 28 ms across a 210-family host.
