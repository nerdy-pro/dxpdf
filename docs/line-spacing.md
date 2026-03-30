# Line Spacing — §17.3.1.33

## Spacing Element

```xml
<w:spacing w:line="276" w:lineRule="auto"/>
```

### lineRule Modes

| lineRule | Meaning | Computation |
|----------|---------|-------------|
| `auto` | Proportional to font size | `natural_height * (line / 240)` |
| `exact` | Fixed height, clips if content is taller | `line` value in twips → Pt |
| `atLeast` | Minimum height, grows for taller content | `max(natural_height, line_in_pt)` |

Default when absent: `auto` with `line=240` (single spacing, 1.0x multiplier).

### Auto Mode Multiplier

The `line` value for `auto` mode is in 240ths of a line:

```
multiplier = Pt::from(line_twips).raw() / 12.0
```

Common values:
- `240` → 1.0x (single)
- `276` → 1.15x (Word default for modern templates)
- `259` → ~1.08x (Aptos/Calibri default in newer documents)
- `360` → 1.5x
- `480` → 2.0x (double)

### Natural Height

The "natural" line height is the tallest fragment on the line:
- Text: `ascent + descent` from font metrics
- Images: image height
- LineBreak: specified `line_height`

### Implementation

```rust
fn resolve_line_height(natural: Pt, rule: &LineSpacingRule) -> Pt {
    match rule {
        LineSpacingRule::Auto(multiplier) => natural * *multiplier,
        LineSpacingRule::Exact(h) => *h,
        LineSpacingRule::AtLeast(min) => natural.max(*min),
    }
}
```

## Paragraph Height

Total paragraph height = `space_before + sum(line_heights) + space_after`.

This height is added to `cursor_y` after the paragraph is rendered. The `space_after` portion can collapse with the next paragraph's `space_before` (see [Paragraph Spacing](paragraph-spacing.md)).
