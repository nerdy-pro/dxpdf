# Paragraph Spacing — §17.3.1.33

## Space Before / After

- `w:spacing w:before="N"` — N twips above the paragraph (1 twip = 1/20 pt)
- `w:spacing w:after="N"` — N twips below the paragraph
- `w:spacing w:beforeAutospacing="1"` — use 14pt instead of explicit value (§17.3.1.33)

### Page-Top Suppression

**§17.3.1.33**: space_before is suppressed for "the first paragraph in a body/text story that begins on a page." This means only the **structural first paragraph** of a section on its initial page:

- First paragraph of the document body
- First paragraph after a section break

Space_before is **NOT** suppressed for:

- Paragraphs that start a new page via `pageBreakBefore` (e.g., Heading1 with `spacing before="480"` retains its 24pt offset)
- Paragraphs pushed to a new page by overflow
- Paragraphs re-laid after `keepNext` forces a page break

Implementation: `cursor_y <= column_top && first_on_section_page` in `layout_section`. The `first_on_section_page` flag is true only until the first paragraph or table is processed.

### Spacing Collapse

When two consecutive paragraphs meet, their spacing collapses (§17.3.1.33 note):

```
effective_gap = max(prev.space_after, current.space_before) - min(prev.space_after, current.space_before)
```

Simplified: `collapse = min(prev.space_after, current.space_before)`, then `cursor_y -= collapse`.

### Contextual Spacing

§17.3.1.9 `contextualSpacing`: when adjacent paragraphs share the same `styleId`, both `space_after` and `space_before` between them are eliminated entirely.

## Empty Paragraphs in the Body

§17.3.1.29: every paragraph ends with a paragraph mark (¶). A paragraph with no runs still occupies one line — the mark's line height comes from `w:pPr/w:rPr` (captured as `mark_run_properties`), falling back to the paragraph style and doc defaults.

Implementation: `build_paragraph_block` injects a `Fragment::LineBreak { line_height }` when the collected fragments are empty. This is required because `section::layout_section` splits paragraph fragments at page breaks and drops empty chunks (§17.3.3.1) — without the injected break, an empty-from-the-start paragraph would split into one empty chunk and be skipped.

Table cells cover the §17.4.66 trailing-empty-after-table case *before* calling `build_paragraph_block` (`build_cell_blocks`), so genuinely structural cell terminators still produce zero height.

## Empty Paragraphs in Headers/Footers

§17.10.1: empty non-last paragraphs in headers/footers still occupy a line height derived from the paragraph mark's font size (`w:rPr` on `w:pPr`, or `mark_run_properties`). A `Fragment::LineBreak { line_height }` is inserted to produce this vertical offset.

The last empty paragraph produces no height (it's a structural terminator).
