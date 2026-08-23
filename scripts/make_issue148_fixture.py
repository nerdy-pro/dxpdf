#!/usr/bin/env python3
"""Build test-files/text-effects.docx — the four legacy run text effects
(issue #148): `w:shadow` §17.3.2.31, `w:outline` §17.3.2.23, `w:emboss`
§17.3.2.13, `w:imprint` §17.3.2.18.

One paragraph per case, one unique token per run, so a test can find each
run's draw commands by text alone:

- PLAIN     — the control: one text command, no effect.
- SHDW      — shadow on default (black) text: the copy must be light gray.
- SHDWRED   — shadow on red text: the copy must be black (the shadow colour
              is keyed on the text's luminance, not fixed).
- OUTL      — outline: hollow glyphs, one stroked command.
- SHOUT     — shadow + outline, the one combination §17.3.2.31 permits.
- EMBS      — emboss: light-gray copy down-right, black text turns white.
- IMPR      — imprint: the mirror — copy up-left.
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

#: (token, rPr children)
RUNS = [
    ("PLAIN", ""),
    ("SHDW", "<w:shadow/>"),
    ("SHDWRED", '<w:shadow/><w:color w:val="FF0000"/>'),
    ("OUTL", "<w:outline/>"),
    ("SHOUT", "<w:shadow/><w:outline/>"),
    ("EMBS", "<w:emboss/>"),
    ("IMPR", "<w:imprint/>"),
]


def para(token: str, r_pr: str) -> str:
    return (
        f"<w:p><w:r><w:rPr>{r_pr}</w:rPr>"
        f'<w:t xml:space="preserve">{token}</w:t></w:r></w:p>'
    )


DOCUMENT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body>"
    + "".join(para(t, pr) for t, pr in RUNS)
    + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
    ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
    "</w:body></w:document>"
)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    target = OUT / "text-effects.docx"
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
