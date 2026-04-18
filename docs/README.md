# OOXML Rendering Reference

Technical reference for DOCX-to-PDF rendering decisions in dxpdf, organized by OOXML spec section.

## Contents

- [Style Cascade](style-cascade.md) — §17.7.2 property resolution, doc defaults, table style interaction
- [Paragraph Spacing](paragraph-spacing.md) — §17.3.1.33 spacing, page-top suppression, collapse rules
- [Floating Tables](floating-tables.md) — §17.4.58/§17.4.59 tblpPr positioning, vertical anchors
- [Floating Images](floating-images.md) — §20.4.2 anchor positioning, text wrapping, forward-scan
- [Fields](fields.md) — §17.16.18 complex/simple fields, PAGE/NUMPAGES evaluation
- [Headers and Footers](headers-footers.md) — §17.10.1 rendering, table support, per-page fields
- [Line Spacing](line-spacing.md) — §17.3.1.33 line/lineRule, Auto/Exact/AtLeast modes
- [DrawingML Plan](drawingml-plan.md) — §20 shapes/fills/outlines/effects — tiered implementation plan, recommended cut-point, architectural decisions
- [DrawingML Tier 0](drawingml-tier-0.md) — foundational infrastructure: color ADT, fills/strokes/effects, geometry generator, FloatingShape — standalone blueprint
- [DrawingML Tier 0 — Phase 2](drawingml-tier-0-phase-2.md) — detailed plan for fill / stroke / effect ADTs + parsers
- [Serde Migration Plan](serde-migration-plan.md) — full plan to replace event-driven OOXML parsing with serde schemas; phases, risks, decisions
