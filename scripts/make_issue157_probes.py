#!/usr/bin/env python3
"""Build the two issue #157 probes — documents that *ask* a question ECMA-376
does not answer, rather than fixtures that pin an answer.

Neither changes behaviour on its own. Each is authored so that the candidate
readings predict a **different measurement off the rendered page**, so one Word
render settles it. That is the pattern `test-files/issue-165-*.docx` used for
vMerge overflow and cell spacing, and the reason those are now decided.

---------------------------------------------------------------------------
1. issue-157-tblprex-bidi.docx  — §17.4.1 `w:bidiVisual` on one row
---------------------------------------------------------------------------

`w:tblPrEx` may carry `bidiVisual`, flipping a single row against a table whose
other rows run left to right. The engine parses it (`TableRowPropertyExceptions::bidi_visual`)
and does not act on it, because the grid is shared and the two readings differ:

  A. the row's cells keep their **declared widths** and swap places, so the
     mirrored row's slot boundaries do not line up with the rows around it;
  B. the row's cells take the **widths of the slots they land in**, so every
     boundary stays aligned and only the contents move.

The grid is 1000/2000/3000 twips — deliberately unequal, because with three
equal columns the two readings render identically and the document would prove
nothing. Row 2 is the flipped one; rows 1 and 3 are the same cells unflipped.

**How to read a Word render.** Look at cell `B` in row 2. Under (A) it is 100pt
wide, as it is in rows 1 and 3. Under (B) it is 100pt only by coincidence of
being the middle column — so look at `A` instead: 50pt under (A), 150pt under
(B). The cells are labelled and separately shaded so this is a ruler
measurement, not a judgement.

---------------------------------------------------------------------------
2. issue-157-empty-row-edge.docx — §17.4.66 across an empty `<w:tr/>`
---------------------------------------------------------------------------

§17.4.66 resolves a shared horizontal edge between two rows by picking one
cell's border and clearing the facing one. A `<w:tr/>` with no `<w:tc>` has no
cell to face, so today neither neighbour is resolved away and **both** paint:
that boundary gets two 0.5pt lines where the same table without the empty row
gets one.

That may well be correct — the conflict genuinely has no subject — which is why
this is a probe and not a fix. The document puts the two tables side by side
with identical rows and borders, so a Word render answers it by measuring one
boundary against the other:

  * equal thickness  -> Word treats the empty row as transparent to §17.4.66,
                        and the engine should resolve across it;
  * double thickness -> today's behaviour is right and the difference is the
                        author's, not a defect.

`w:sz="24"` (3pt) rather than a hairline, so "one line or two" is visible at
100% zoom instead of needing a loupe.

---------------------------------------------------------------------------

Run `scripts/verify_docx.py` on the results before committing. Regenerate and
commit if the content changes; the build is deterministic and needs no
third-party packages.

    scripts/make_issue157_probes.py
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
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

SECT_PR = """<w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>"""


def heading(text: str) -> str:
    # `<` and `&` in a `<w:t>` would become markup, which `verify_docx.py`
    # cannot catch — it checks that the package is a sound OPC container, not
    # that the parts are schema-valid. This probe lost an afternoon to a heading
    # that named an element literally, so the guard is here rather than in a
    # reviewer's memory.
    if "<" in text or "&" in text:
        raise ValueError(f"heading text would be parsed as markup: {text!r}")
    return f'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def cell(label: str, fill: str, extra: str = "") -> str:
    return (
        "<w:tc>"
        f'<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>{extra}</w:tcPr>'
        f'<w:p><w:r><w:t xml:space="preserve">{label}</w:t></w:r></w:p>'
        "</w:tc>"
    )


def document(body: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body}
    {SECT_PR}
  </w:body>
</w:document>
"""


# ── probe 1: w:tblPrEx/w:bidiVisual ─────────────────────────────────────────

#: Unequal on purpose — see the module docstring.
GRID3 = '<w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/>'


def tblprex_probe() -> str:
    plain = "<w:tr>" + cell("A", "F8CBAD") + cell("B", "C6E0B4") + cell("C", "BDD7EE") + "</w:tr>"
    flipped = (
        "<w:tr><w:tblPrEx><w:bidiVisual/></w:tblPrEx>"
        + cell("A", "F8CBAD")
        + cell("B", "C6E0B4")
        + cell("C", "BDD7EE")
        + "</w:tr>"
    )
    table = (
        "<w:tbl><w:tblPr>"
        '<w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        '<w:left w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        '<w:right w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        '<w:insideH w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        '<w:insideV w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        "</w:tblBorders></w:tblPr>"
        f"<w:tblGrid>{GRID3}</w:tblGrid>{plain}{flipped}{plain}</w:tbl>"
    )
    return document(
        heading("Row 2 carries w:tblPrEx/w:bidiVisual. Rows 1 and 3 do not.")
        + table
        + heading("Measure cell A in row 2: 50pt keeps its declared width, 150pt takes the slot's.")
    )


# ── probe 2: §17.4.66 across an empty <w:tr/> ───────────────────────────────

GRID1 = '<w:gridCol w:w="4000"/>'


def empty_row_probe() -> str:
    def table(with_empty: bool) -> str:
        rows = "<w:tr>" + cell("upper", "D9EAD3") + "</w:tr>"
        if with_empty:
            rows += "<w:tr/>"
        rows += "<w:tr>" + cell("lower", "CFE2F3") + "</w:tr>"
        return (
            "<w:tbl><w:tblPr>"
            '<w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
            "<w:tblBorders>"
            '<w:top w:val="single" w:sz="24" w:space="0" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="24" w:space="0" w:color="000000"/>'
            '<w:left w:val="nil"/><w:right w:val="nil"/>'
            '<w:insideH w:val="single" w:sz="24" w:space="0" w:color="000000"/>'
            '<w:insideV w:val="nil"/>'
            "</w:tblBorders></w:tblPr>"
            f"<w:tblGrid>{GRID1}</w:tblGrid>{rows}</w:tbl>"
        )

    return document(
        heading("Control: two rows, one shared edge between them.")
        + table(False)
        + "<w:p/>"
        + heading("The same, with an empty w:tr element between the two rows.")
        + table(True)
        + heading("Compare the thickness of the middle line in each.")
    )


def write(name: str, body: str) -> None:
    target = OUT / name
    # Fixed timestamps so regenerating an unchanged fixture produces identical
    # bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for part, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/document.xml", body),
        ):
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    write("issue-157-tblprex-bidi.docx", tblprex_probe())
    write("issue-157-empty-row-edge.docx", empty_row_probe())


if __name__ == "__main__":
    main()
