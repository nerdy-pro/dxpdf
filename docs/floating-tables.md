# Floating Tables — §17.4.58 / §17.4.59

## Table Positioning Properties

Floating tables use `w:tblpPr` on `w:tblPr`:

```xml
<w:tblpPr w:rightFromText="187" w:bottomFromText="72"
          w:vertAnchor="text" w:tblpY="1"/>
```

### Horizontal Positioning

- `tblpXSpec` (§17.4.58): named alignment — `left`, `center`, `right`
- `tblpX`: absolute X offset from `horzAnchor`
- `horzAnchor`: reference frame — `text` (content area), `margin`, `page`
- `leftFromText` / `rightFromText`: gap between table edge and surrounding text

### Vertical Positioning

- `tblpY` (§17.4.59): absolute Y offset from `vertAnchor`
- `vertAnchor` (§17.4.58): reference frame:
  - `text` — top of the nearest preceding paragraph (default)
  - `margin` — top margin edge (`margins.top`)
  - `page` — top of the page (y=0)

### Y Position Computation

```rust
let anchor_y = match vert_anchor {
    Text   => last_para_start_y + y_offset,
    Margin => margins.top + y_offset,
    Page   => y_offset,
};
// Table must not start before the current cursor
// (preceding content already occupies space above cursor_y).
let float_y_start = anchor_y.max(cursor_y);
```

The `max(cursor_y)` floor prevents the table from overlapping already-rendered paragraph content above it.

### `last_para_start_y`

Tracked in `layout_section` — set to `cursor_y` at the start of each paragraph's processing (before spacing adjustments). Used as the anchor reference for `vertAnchor="text"`.

## Float Registration

Floating tables are registered as `ActiveFloat` with `FloatSource::Table { owner_block_idx }`. The `owner_block_idx` identifies which paragraph should wrap around the table.

## `tblOverlap`

§17.4.47: `w:tblOverlap val="never"` prevents overlap with **other floating tables**, not with paragraph text. Tables can still visually overlap paragraph content.

## Data Flow

```
build_table() → TableFloatInfo { right_gap, bottom_gap, x_align, y_offset, vert_anchor }
    ↓
LayoutBlock::Table { float_info: Option<TableFloatInfo> }
    ↓
layout_section() → compute float_y_start, register ActiveFloat, emit draw commands
```
