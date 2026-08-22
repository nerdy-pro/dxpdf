#!/usr/bin/env python3
"""Build test-files/bidi-visual.docx — §17.4.1 w:bidiVisual (issue #157).

Three tables over one deliberately lopsided grid (1440/2880/4320 twips =
72/144/216 pt), so a mirrored column order changes every x-coordinate and a
test can tell the mirror from a mere reversal of the text:

1. *mirror*   — `bidiVisual`, one row AA|BB|CC. The logical first cell (AA)
   must come out rightmost, and the gaps between neighbours must be the
   *mirrored* column widths.
2. *merges*   — `bidiVisual`, a `gridSpan=2` cell and a `vMerge` pair, plus
   `tblBorders`, so spans, merges and border resolution all cross the flip.
3. *control*  — the same grid without `bidiVisual`; KK|LL|MM stays LTR and
   pins that the flag mirrors only the table that carries it.

Each cell holds one unique two-letter token in a single run, so every token
is exactly one text draw command and `tests/bidi_visual.rs` can assert on
positions token-by-token.
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

GRID = '<w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="2880"/><w:gridCol w:w="4320"/></w:tblGrid>'


def cell(text: str, tc_pr: str = "") -> str:
    body = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:tc><w:tcPr>{tc_pr}</w:tcPr><w:p>{body}</w:p></w:tc>"


def table(tbl_pr: str, rows: str) -> str:
    return f"<w:tbl><w:tblPr>{tbl_pr}</w:tblPr>{GRID}{rows}</w:tbl>"


def spacer() -> str:
    return "<w:p/>"


BORDERS = (
    "<w:tblBorders>"
    '<w:top w:val="single" w:sz="4" w:color="000000"/>'
    '<w:left w:val="single" w:sz="4" w:color="000000"/>'
    '<w:bottom w:val="single" w:sz="4" w:color="000000"/>'
    '<w:right w:val="single" w:sz="4" w:color="000000"/>'
    '<w:insideH w:val="single" w:sz="4" w:color="000000"/>'
    '<w:insideV w:val="single" w:sz="4" w:color="000000"/>'
    "</w:tblBorders>"
)

MIRROR = table(
    "<w:bidiVisual/>",
    "<w:tr>" + cell("AA") + cell("BB") + cell("CC") + "</w:tr>",
)

MERGES = table(
    "<w:bidiVisual/>" + BORDERS,
    "<w:tr>"
    + cell("DD", '<w:gridSpan w:val="2"/>')
    + cell("EE")
    + "</w:tr>"
    + "<w:tr>"
    + cell("FF", '<w:vMerge w:val="restart"/>')
    + cell("GG")
    + cell("HH")
    + "</w:tr>"
    + "<w:tr>"
    + cell("", "<w:vMerge/>")
    + cell("II")
    + cell("JJ")
    + "</w:tr>",
)

CONTROL = table(
    "",
    "<w:tr>" + cell("KK") + cell("LL") + cell("MM") + "</w:tr>",
)

DOCUMENT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body>"
    + MIRROR
    + spacer()
    + MERGES
    + spacer()
    + CONTROL
    + spacer()
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
    ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
    "</w:body></w:document>"
)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    target = OUT / "bidi-visual.docx"
    # Fixed timestamps so regenerating an unchanged fixture produces
    # identical bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/document.xml", DOCUMENT),
        ):
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
