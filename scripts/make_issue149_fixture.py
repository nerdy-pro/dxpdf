#!/usr/bin/env python3
"""Build test-files/shading-patterns.docx — §17.18.78 ST_Shd cell shading
patterns (issue #149).

One borderless 3x5 table, one `w:shd` value per cell. Every cell's colours are
unique across the fixture, so a test can identify each cell's output by colour
alone — no coordinates, no draw order:

- the percentage tints and `solid`/`clear` render as one flat colour each,
  identifiable as an exact blended RGB;
- each geometric family (horz/vert/diag/reverseDiag stripes, horz/diag cross,
  one thin variant) renders as its background fill plus stripe lines in its
  foreground colour, identifiable by that foreground;
- `nil` must paint nothing, and `solid` must paint its *pattern* colour, not
  its fill — the two colours no cell may produce.
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

#: (w:val, w:color, w:fill) per cell, reading order. Colours unique per cell.
CELLS = [
    # Flat family: tints blend color into fill; solid takes the colour whole.
    ("pct25", "0000FF", "FFFFFF"),
    ("pct50", "auto", "auto"),
    ("pct12", "000000", "FFFFFF"),
    ("solid", "CC0000", "00CC00"),
    ("clear", "auto", "FFCC00"),
    ("nil", "AA0000", "AA0000"),
    # Geometric families: fill behind, stripes in the pattern colour.
    ("horzStripe", "220022", "DDFFDD"),
    ("thinHorzStripe", "330033", "DDFFEE"),
    ("vertStripe", "440044", "EEFFDD"),
    ("diagStripe", "550055", "EEFFEE"),
    ("reverseDiagStripe", "660066", "FFEEDD"),
    ("horzCross", "770077", "FFEEEE"),
    ("diagCross", "880088", "FFDDEE"),
    # Padding to complete the 3x5 grid; one carries the document's only text
    # so the fixture parses as a document with body content.
    ("clear", "auto", "F0F0F1"),
    ("clear", "auto", "F0F0F2", "shading"),
]


def cell(val: str, color: str, fill: str, text: str = "") -> str:
    body = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return (
        "<w:tc><w:tcPr>"
        f'<w:shd w:val="{val}" w:color="{color}" w:fill="{fill}"/>'
        f"</w:tcPr><w:p>{body}</w:p></w:tc>"
    )


def document() -> str:
    rows = []
    for i in range(0, len(CELLS), 3):
        rows.append("<w:tr>" + "".join(cell(*c) for c in CELLS[i : i + 3]) + "</w:tr>")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body><w:tbl><w:tblPr/>"
        '<w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>'
        + "".join(rows)
        + "</w:tbl><w:p/>"
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
        "</w:body></w:document>"
    )


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    target = OUT / "shading-patterns.docx"
    # Fixed timestamps so regenerating an unchanged fixture produces
    # identical bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/document.xml", document()),
        ):
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
