# Paragraph Spacing — §17.3.1.33

## Space Before / After

- `w:spacing w:before="N"` — N twips above the paragraph (1 twip = 1/20 pt)
- `w:spacing w:after="N"` — N twips below the paragraph
- `w:spacing w:beforeAutospacing="1"` — use 14pt instead of explicit value (§17.3.1.33)

### Page-Top Suppression

**§17.3.1.33**: space_before is suppressed when a paragraph begins at the top of any page. This includes:

- First paragraph of the document
- Paragraphs after `pageBreakBefore` (Heading1 typically has this)
- Paragraphs pushed to a new page by overflow
- Paragraphs re-laid after `keepNext` forces a page break

Implementation: `cursor_y <= column_top` check in `layout_section`. No `first_on_section_page` gating — suppression applies universally at any page top.

### Spacing Collapse

When two consecutive paragraphs meet, their spacing collapses (§17.3.1.33 note):

```
effective_gap = max(prev.space_after, current.space_before) - min(prev.space_after, current.space_before)
```

Simplified: `collapse = min(prev.space_after, current.space_before)`, then `cursor_y -= collapse`.

### Contextual Spacing

§17.3.1.9 `contextualSpacing`: when adjacent paragraphs share the same `styleId`, both `space_after` and `space_before` between them are eliminated entirely.

## Empty Paragraphs in Headers/Footers

§17.10.1: empty non-last paragraphs in headers/footers still occupy a line height derived from the paragraph mark's font size (`w:rPr` on `w:pPr`, or `mark_run_properties`). A `Fragment::LineBreak { line_height }` is inserted to produce this vertical offset.

The last empty paragraph produces no height (it's a structural terminator).
