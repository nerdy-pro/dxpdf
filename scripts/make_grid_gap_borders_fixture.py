#!/usr/bin/env python3
"""Build test-files/grid-gap-borders.docx — §17.4.36 `insideV` against §17.4.15
`gridBefore` and §17.4.14 `gridAfter`: what is painted on the vertical edge
where a row's cells stop short of the table's grid.

`gridBefore` leaves the leftmost grid columns of a row **blank** — no `<w:tc>`
covers them. The row's first cell therefore has a grid column to its left but no
*cell* to its left, and §17.4.36 defines `insideV` as the border on the table's
"interior vertical edges" without ever defining "interior".

**This fixture has now been rendered in Word, and it answers the question**: the
gap-facing edge takes the table's own `w:left` (or `w:right`). Word draws a 3pt
red line on the leading edge of `D`, `F` and `G`, and a 3pt green one on the
trailing edge of `E` and `F`. A row's first `<w:tc>` has no cell facing it, so
§17.4.66's "cell borders and outer table borders" leaves the table's border as
what it faces; `gridBefore` moves where that edge is, not what it is.

Two other readings were held and refuted, and the fixture is kept because it is
what refuted them: `insideV`, because grid columns exist to the left — ruled out
by `bidi-visual-table.docx`, whose `nil` outer border leaves that edge bare; and
*nothing at all*, because §17.4.35 places `w:left` "around the table" and this
edge is 50pt inside it — argued from the spec's wording, consistent with every
measurement available then, and ruled out by this render.

Table 2 is the same rows with the outer borders set to `nil`, which under the
answer above predicts a bare edge — the originally reported symptom, and now the
control that keeps one rule covering both documents.

---------------------------------------------------------------------------
How to read a Word render
---------------------------------------------------------------------------

Every vertical border is identifiable on sight, by colour *and* by weight:

    left     3pt red    (C00000)
    right    3pt green  (00B050)
    insideV  1pt blue   (0070C0)

Horizontals are 1pt grey so they cannot be confused with any of the three. Row 1
of each table spans the whole grid and is the **legend**: it shows all three at
once, so the reader never has to trust a colour name.

What Word draws, and what this engine now draws with it:

    D, F, G  leading edge   3pt red    (the table's `w:left`)
    E, F     trailing edge  3pt green  (the table's `w:right`)
    G|H                     1pt blue   (`insideV` — a cell faces it)

Row 3 is the §17.4.14 `gridAfter` mirror of row 2, and row 4 gaps cell `F` at
both ends — the case that exposes a fix reaching one end and not the other. Row 5
gaps the row's start but then has **two** cells, so `G|H` is genuinely interior
and keeps its blue: without it, giving every cell of a gapped row the outer
border would satisfy every other row here.

Table 2 is the same five rows with `w:left` and `w:right` set to `nil`, so the
same rule predicts a bare edge. That is the originally reported symptom, and
keeping both tables is what holds one rule to both documents.

---------------------------------------------------------------------------
Why the styles part
---------------------------------------------------------------------------

Same reason as `make_bidi_visual_fixture.py`, whose `STYLES` constant carries the
full argument: a package that declares no face, size or spacing is read one way
here and another by Word, because §17.7.2 leaves an absent `w:docDefaults`
application-defined. A fixture meant to be *measured against Word* has to state
them. Times New Roman because macOS and Word ship the same file.

Nothing about this fixture's question depends on the font — no assertion here
measures a glyph — but a reader comparing the two renders side by side should not
have to discount a row-height difference while looking for a border.

Run `scripts/verify_docx.py` on the result before committing. Regenerate and
commit if the content changes; the build is deterministic and needs no
third-party packages.

    scripts/make_grid_gap_borders_fixture.py
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

#: The `xmlns` is the relationship-*part* URI. The relationship-*type* URI here
#: instead yields a package Word rejects and this parser accepts — see AGENTS.md.
DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

#: §17.7.5 — see the module docstring, and `make_bidi_visual_fixture.py` for why
#: each value is stated rather than defaulted.
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"
                  w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
        <w:kern w:val="2"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>
"""

#: Unequal on purpose: 50 / 100 / 150 pt. With equal columns a reader cannot say
#: which column a gap consumed, and the two `gridSpan` cells below would be
#: indistinguishable from one another.
GRID = '<w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/>'

#: The three vertical borders, each identifiable by colour and weight alone.
LEFT = '<w:left w:val="single" w:sz="24" w:space="0" w:color="C00000"/>'
RIGHT = '<w:right w:val="single" w:sz="24" w:space="0" w:color="00B050"/>'
INSIDE_V = '<w:insideV w:val="single" w:sz="8" w:space="0" w:color="0070C0"/>'
#: Horizontals in a fourth colour, so no vertical reading can be confused by one.
HORIZONTAL = (
    '<w:top w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
    '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
    '<w:insideH w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
)

NIL_LEFT = '<w:left w:val="nil"/>'
NIL_RIGHT = '<w:right w:val="nil"/>'


def cell(label: str, fill: str, extra: str = "") -> str:
    """One shaded, labelled cell.

    Deliberately carries **no** `w:tcBorders`: the whole question is what the
    table-level borders resolve to at this cell's edges, and a cell-level border
    would answer it locally and hide the answer.
    """
    return (
        "<w:tc>"
        f'<w:tcPr><w:tcW w:w="0" w:type="auto"/>{extra}'
        f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/></w:tcPr>'
        f'<w:p><w:r><w:t xml:space="preserve">{label}</w:t></w:r></w:p>'
        "</w:tc>"
    )


def rows() -> str:
    """The four rows, identical in both tables.

    `w:trPr` children follow the `CT_TrPrBase` sequence (`gridBefore` then
    `gridAfter`), so nothing rests on a reader's tolerance.
    """
    span2 = '<w:gridSpan w:val="2"/>'
    return "".join(
        [
            # 1 — the legend: every border on the page at once.
            "<w:tr>"
            + cell("A", "F8CBAD")
            + cell("B", "C6E0B4")
            + cell("C", "BDD7EE")
            + "</w:tr>",
            # 2 — §17.4.15: one column blank at the row's start. The question is
            #     what sits at D's leading edge, 50pt in from the table's left.
            "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/></w:trPr>"
            + cell("D", "FFE699", span2)
            + "</w:tr>",
            # 3 — §17.4.14: the mirror image, at E's trailing edge, 150pt in
            #     from the table's right.
            "<w:tr><w:trPr><w:gridAfter w:val=\"1\"/></w:trPr>"
            + cell("E", "D9D2E9", span2)
            + "</w:tr>",
            # 4 — both at once. A fix that handles one end and not the other
            #     shows up here and nowhere else.
            "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/><w:gridAfter w:val=\"1\"/></w:trPr>"
            + cell("F", "F4CCCC")
            + "</w:tr>",
            # 5 — a gapped row that still has an *interior* boundary of its own,
            #     between G and H. This is the trap-detector: every other row
            #     here has a single cell, so a change that suppressed every
            #     vertical border in a row with a gap would satisfy all of them.
            #     The G|H edge has a cell on both sides and must stay painted
            #     whatever is decided about the gap-facing one.
            "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/></w:trPr>"
            + cell("G", "D9EAD3")
            + cell("H", "CFE2F3")
            + "</w:tr>",
        ]
    )


def table(nil_outer: bool) -> str:
    """One table. `nil_outer` reproduces the reported case; otherwise the outer
    borders are visible and the render discriminates all three readings."""
    left = NIL_LEFT if nil_outer else LEFT
    right = NIL_RIGHT if nil_outer else RIGHT
    return (
        "<w:tbl><w:tblPr>"
        '<w:tblW w:w="6000" w:type="dxa"/>'
        f"<w:tblBorders>{HORIZONTAL}{left}{right}{INSIDE_V}</w:tblBorders>"
        '<w:tblLayout w:type="fixed"/>'
        "<w:tblCellMar>"
        '<w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
        "</w:tblCellMar>"
        "</w:tblPr>"
        f"<w:tblGrid>{GRID}</w:tblGrid>{rows()}</w:tbl>"
    )


def heading(text: str) -> str:
    # `<` and `&` would become markup, which `verify_docx.py` cannot catch — it
    # checks that the package is a sound OPC container, not that the parts are
    # schema-valid. An earlier probe lost an afternoon to a heading that named an
    # element literally, so the guard is here rather than in a reviewer's memory.
    if "<" in text or "&" in text:
        raise ValueError(f"heading text would be parsed as markup: {text!r}")
    return f'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


DOCUMENT = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {heading("Table 1 - visible outer borders. Row 1 is the legend: 3pt red left, 1pt blue insideV, 3pt green right.")}
    {table(False)}
    <w:p/>
    {heading("At cell D's leading edge: 3pt red means the row's leading edge takes w:left; 1pt blue means insideV; nothing means the edge is never painted.")}
    <w:p/>
    {heading("Table 2 - the same rows with w:left and w:right set to nil. This is the reported case, without bidiVisual.")}
    {table(True)}
    <w:p/>
    {heading("Word draws nothing at either gap edge here. Any line at D's left or E's right is the defect.")}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"""


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    target = OUT / "grid-gap-borders.docx"
    # Fixed timestamps so regenerating an unchanged fixture produces identical
    # bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/_rels/document.xml.rels", DOC_RELS),
            ("word/styles.xml", STYLES),
            ("word/document.xml", DOCUMENT),
        ):
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
