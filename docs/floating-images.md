# Floating Images — §20.4.2

## Anchor Positioning

Floating images use `wp:anchor` inside `w:drawing`:

```xml
<wp:anchor distT="0" distB="0" distL="114300" distR="114300"
           simplePos="0" relativeHeight="251658240"
           behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
  <wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH>
  <wp:positionV relativeFrom="margin"><wp:align>top</wp:align></wp:positionV>
  <wp:extent cx="975360" cy="975360"/>
  <wp:wrapSquare wrapText="bothSides"/>
</wp:anchor>
```

### Vertical Position Variants

Images resolve to one of two internal representations:

- `FloatingImageY::Absolute(y)` — page-absolute position. Used for `relativeFrom="margin"`, `"page"`, `"topMargin"`, `"bottomMargin"`.
- `FloatingImageY::RelativeToParagraph(offset)` — offset from the anchor paragraph's content area top. Used for `relativeFrom="paragraph"` and `"line"`.

### Alignment

Named alignments (`top`, `center`, `bottom` for vertical; `left`, `center`, `right` for horizontal) are resolved to absolute positions during building, based on the reference area.

## Text Wrapping

Wrapping mode determines how text flows around the image:

- `wrapSquare` / `wrapTight` — text wraps on both sides (registered as `ActiveFloat`)
- `wrapTopAndBottom` (§20.4.2.18) — image acts as a block spacer; cursor_y advances past it
- `wrapNone` — no text wrapping, image overlays text (behind or in front based on `behindDoc`)

## Forward-Scan for Absolute Floats

Word uses multi-pass layout where all floats on a page affect all text. Our single-pass renderer approximates this with forward-scanning:

1. When a new page starts (`abs_floats_dirty = true`), scan `blocks[page_start_block..]` for paragraphs with absolute-positioned floating images
2. Register these as `current_page_abs_floats`
3. When building `effective_floats` for each paragraph, merge `current_page_abs_floats` with `page_floats`
4. Dedup by coordinate proximity (< 0.1pt)

### Merge Cutoff

Forward-scanned floats are only included if their `page_y_start <= cursor_y + space_before`. This prevents floats positioned far below the current paragraph from constraining it.

## Float Constraint Zone

Each `ActiveFloat` defines a rectangular constraint zone:

```
page_x = image_x - dist_left
width  = image_width + dist_left + dist_right
```

The `float_adjustments` function computes left/right indentation for each text line based on y-overlap with active floats.

## Pruning

Floats are pruned when `cursor_y >= float.page_y_end` — the cursor has passed below the float's bottom edge.
