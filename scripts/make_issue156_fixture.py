#!/usr/bin/env python3
"""Build test-files/bidi-tabs.docx — §17.3.1.37 tab stops under w:bidi
(issue #156).

Four paragraph pairs on one Letter page (text column 72..540 pt), each RTL
paragraph mirrored by an LTR control, every token unique so a test can find
each one's x without coordinates or draw order:

1. *stops*   — ``RA\tRB\tRC`` under ``w:bidi`` with **end** stops at 100 and
   200 pt, so the zones' *left* edges land metric-free at 540−100 and
   540−200; the control ``LA\tLB\tLC`` uses **start** stops at the same
   positions, landing at 72+100 and 72+200.
2. *numbered* — an RTL numbered item (start=7, hanging 18 pt inside a 36 pt
   indent): the "7." label must sit against the right margin, right of its
   own body text. The LTR control (start=3) pins the label at exactly
   72+36−18 = 90 pt.
3. *bar*     — an RTL paragraph with a ``bar`` stop at 150 pt: the rule must
   be drawn at 540−150, mirrored like every other stop position.
4. *grid*    — ``RE\tRF`` with no custom stops: the default 36 pt grid is
   walked from the right, so RF ends up left of RE.
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
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
"""

#: Two single-level decimal definitions, differing only in start (7 for the
#: RTL item, 3 for the LTR control) so the two labels are distinct strings.
NUMBERING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + "".join(
        f'<w:abstractNum w:abstractNumId="{aid}">'
        f'<w:lvl w:ilvl="0">'
        f'<w:start w:val="{start}"/>'
        f'<w:numFmt w:val="decimal"/>'
        f'<w:lvlText w:val="%1."/>'
        f'<w:lvlJc w:val="left"/>'
        f'<w:pPr><w:ind w:start="720" w:hanging="360"/></w:pPr>'
        f"</w:lvl>"
        f"</w:abstractNum>"
        for aid, start in ((0, 7), (1, 3))
    )
    + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    + '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
    + "</w:numbering>"
)


def para(p_pr: str, runs: str) -> str:
    return f"<w:p><w:pPr>{p_pr}</w:pPr>{runs}</w:p>"


def tabbed_runs(tokens: list[str]) -> str:
    out = []
    for i, t in enumerate(tokens):
        if i:
            out.append("<w:r><w:tab/></w:r>")
        out.append(f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>')
    return "".join(out)


def stops(entries: list[tuple[str, int]]) -> str:
    tabs = "".join(f'<w:tab w:val="{v}" w:pos="{p}"/>' for v, p in entries)
    return f"<w:tabs>{tabs}</w:tabs>"


BODY = (
    # 1. Custom stops. 2000/4000 twips = 100/200 pt.
    para("<w:bidi/>" + stops([("end", 2000), ("end", 4000)]), tabbed_runs(["RA", "RB", "RC"]))
    + para(stops([("start", 2000), ("start", 4000)]), tabbed_runs(["LA", "LB", "LC"]))
    # 2. Numbered items (the level's w:ind carries indent 720 / hanging 360).
    + para(
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:bidi/>',
        '<w:r><w:t xml:space="preserve">RNUM</w:t></w:r>',
    )
    + para(
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>',
        '<w:r><w:t xml:space="preserve">LNUM</w:t></w:r>',
    )
    # 3. A bar stop at 3000 twips = 150 pt.
    + para("<w:bidi/>" + stops([("bar", 3000)]), tabbed_runs(["RD"]))
    # 4. The default tab-stop grid.
    + para("<w:bidi/>", tabbed_runs(["RE", "RF"]))
)

DOCUMENT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body>" + BODY + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
    ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
    "</w:body></w:document>"
)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    target = OUT / "bidi-tabs.docx"
    # Fixed timestamps so regenerating an unchanged fixture produces
    # identical bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/_rels/document.xml.rels", DOC_RELS),
            ("word/document.xml", DOCUMENT),
            ("word/numbering.xml", NUMBERING),
        ):
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
