# Style Cascade — §17.7.2

## Priority Order

For a **body paragraph**:

1. Direct formatting (highest)
2. Paragraph style (e.g., Normal, Heading1)
3. Document defaults (lowest)

For a **table cell paragraph**:

1. Direct formatting (highest)
2. Paragraph style
3. Table conditional formatting (§17.7.6)
4. Table style paragraph properties
5. Document defaults (lowest)

## Document Defaults

Document defaults (`w:docDefaults/w:pPrDefault`) provide the lowest-priority fallback for all paragraph properties. Common values:

```xml
<w:pPrDefault>
  <w:pPr>
    <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
  </w:pPr>
</w:pPrDefault>
```

### Resolved Styles vs. Doc Defaults

**Resolved styles must NOT include document defaults for paragraph properties.** Doc defaults are merged by the caller (`resolve_paragraph_defaults` / `build_fragments`) at the correct cascade level.

Rationale: if the resolved Normal style includes doc defaults (e.g., `spacing after=160`), and a table style defines `spacing after=0`, the table style cannot override — `merge_paragraph_properties` only fills `None` fields. By keeping doc defaults out of resolved styles, the table style's properties merge before doc defaults.

Run defaults (`w:rPrDefault`) ARE merged during style resolution because they don't have the same cascade conflict — character formatting doesn't have a "table style" insertion point.

### Implementation

In `resolve_one` (styles.rs): only `merge_run_properties` is called, NOT `merge_paragraph_properties` with doc defaults.

In `resolve_paragraph_defaults` (build.rs):
- `defer_doc_defaults=false` (body paragraphs): doc defaults merged after paragraph style
- `defer_doc_defaults=true` (table cell paragraphs): doc defaults deferred, merged after table style + conditional formatting

## Property Merging

`merge_paragraph_properties` fills in `None` fields from the base — it never overwrites `Some` values. Spacing sub-fields (`before`, `after`, `line`) are merged individually (§17.3.1.33), not as an atomic `Option<Spacing>` block. This allows partial overrides:

```
// Direct: spacing.before = Some(240)
// Style:  spacing.before = None, spacing.line = Some(Auto(276))
// Result: spacing.before = Some(240), spacing.line = Some(Auto(276))
```
