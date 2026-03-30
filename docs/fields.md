# Fields — §17.16.18

## Simple Fields

`Inline::Field` contains a parsed `FieldInstruction` and cached result `content`:

```xml
<w:fldSimple w:instr="PAGE \* Arabic">
  <w:r><w:t>3</w:t></w:r>
</w:fldSimple>
```

The `content` inlines contain the cached result from when the document was last saved.

## Complex Fields

Complex fields use a state machine of `FieldChar` markers:

```
Begin → InstrText... → Separate → result TextRuns → End
```

```xml
<w:r><w:rPr>...formatting...</w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:rPr>...formatting...</w:rPr><w:instrText>PAGE \* Arabic</w:instrText></w:r>
<w:r><w:rPr>...formatting...</w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:rPr>...formatting...</w:rPr><w:t>3</w:t></w:r>
<w:r><w:rPr>...formatting...</w:rPr><w:fldChar w:fldCharType="end"/></w:r>
```

### State Machine in `collect_fragments`

```
field_depth: i32    — tracks nesting (Begin increments, Separate decrements)
field_instr: String — accumulated InstrText between Begin and Separate
field_sub_pending: Option<String> — substitution text awaiting first result TextRun
field_sub_emitted: bool — substitution was applied, skip remaining result TextRuns
```

- **Begin**: `field_depth += 1`, clear state
- **InstrText**: append to `field_instr` (only when `field_depth > 0`)
- **Separate**: parse `field_instr` via `dxpdf_field::parse()`, evaluate for PAGE/NUMPAGES. Set `field_sub_pending` if substitution available. `field_depth -= 1`
- **TextRun** (between Separate and End):
  - If `field_sub_pending` is set: use substituted text with this TextRun's resolved font properties (§17.16.19 MERGEFORMAT), then set `field_sub_emitted = true`
  - If `field_sub_emitted`: skip
  - Otherwise: collect normally (cached result)
- **End**: if `field_sub_pending` still set (no result TextRun was present), emit with paragraph default font. Clear state.

## Dynamic Field Evaluation

`FieldContext` provides runtime values:

```rust
pub struct FieldContext {
    pub page_number: Option<usize>,  // 1-based
    pub num_pages: Option<usize>,
}
```

Currently evaluated fields:
- `PAGE` (§17.16.4.1) — current page number
- `NUMPAGES` (§17.16.4.1) — total document page count

When `FieldContext` has no value for a field (e.g., body text without per-page context), the cached result is used as fallback.

## §17.16.19 MERGEFORMAT

The `\* MERGEFORMAT` switch preserves the formatting of the first result run when the field is updated. Our substitution honors this: the first `TextRun` between Separate and End provides font family, size, bold, italic, color — the substituted text replaces only the content while preserving the style.

Fallback (no result TextRun present): paragraph default font properties are used via `make_field_text_fragment`.
