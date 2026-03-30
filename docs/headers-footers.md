# Headers and Footers — §17.10

## Resolution

§17.10.5: sections without explicit header/footer references inherit from the previous section.

Footer types: `default`, `even`, `first`. Currently only `default` is resolved (via `footerReference` → `RelId` → `footer{N}.xml`).

## Content Structure

Headers/footers can contain:
- Paragraphs (text, fields, hyperlinks)
- Tables (address blocks, multi-column layouts)
- Floating images (logos, decorative elements)
- VML shapes with absolute positioning

## Building (`build_header_footer_content`)

Produces `HeaderFooterContent`:

```rust
pub struct HeaderFooterContent {
    pub blocks: Vec<LayoutBlock>,           // paragraphs + tables
    pub absolute_position: Option<(Pt, Pt)>, // VML override
    pub floating_images: Vec<FloatingImage>,  // page-relative
}
```

- **Paragraphs**: built via `build_fragments` (same as body paragraphs)
- **Tables**: built via `build_table` (full table style cascade applies)
- **Floating images**: extracted separately — they use page-relative coordinates, not stack-relative
- **VML absolute positioning**: detected from `Pict` inlines; overrides the default margin-based position

### §17.10.1: Empty Paragraph Handling

Empty non-last paragraphs still occupy vertical space (line height from the paragraph mark's font). A `Fragment::LineBreak` is inserted. The last empty paragraph produces no height.

## Layout (`render_headers_footers`)

Uses `stack_blocks` (same engine as table cells) to vertically stack paragraphs and tables.

### Header Positioning

```
offset_x = margins.left  (or VML abs_x)
offset_y = header_margin  (or VML abs_y)
```

Commands from `stack_blocks` are shifted by `(offset_x, offset_y)` and prepended before body content.

### Footer Positioning

```
footer_y = page_height - footer_margin - content_height
```

Commands shifted by `(margins.left, footer_y)` and appended after body content.

### Floating Images in Headers/Footers

Handled separately from `stack_blocks` because they use page-absolute coordinates:
- `FloatingImageY::Absolute(y)` — used as-is (no shift)
- `FloatingImageY::RelativeToParagraph(offset)` — shifted by header/footer offset

## Per-Page Field Evaluation

Headers/footers are built **per-page** to evaluate PAGE and NUMPAGES fields correctly.

### Two-Phase Layout (lib.rs)

1. **Phase 1**: layout all sections → determine total page count
2. **Phase 2**: render headers/footers with correct `page_number` and `total_pages`

`SectionHfInfo` stores the page range and raw block references for each section. Phase 2 iterates these, setting `field_ctx_cell` per-page before building content.

```rust
ctx.field_ctx_cell.set(FieldContext {
    page_number: Some(page_base + page_idx + 1),
    num_pages: Some(total_pages),
});
let hf = build_header_footer_content(blocks, ctx);
```

The field context is reset to default after header/footer rendering to prevent leakage into body layout.
