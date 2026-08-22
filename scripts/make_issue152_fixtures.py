#!/usr/bin/env python3
"""Build test-files/counting-*.docx — §17.18.59 counting systems (issue #152).

One fixture per language. Each `w:numFmt` under test gets two lists: one
starting at 1 (items 1, 2, 3) and one starting at 10 (items 10, 11, 12), so
every fixture crosses its language's tens boundary in-document — the place
counting systems first diverge from positional digits (十二, not 一二).

`w:lvlText` is a bare `%1` so the rendered label *is* the formatted counter,
which is what `tests/counting_systems.rs` asserts against, string-for-string.
"""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files"

#: fixture stem → the numFmt values it exercises.
FIXTURES = {
    "counting-zh-hans": ["chineseCounting", "chineseCountingThousand", "chineseLegalSimplified"],
    "counting-zh-hant": [
        "taiwaneseCounting",
        "taiwaneseCountingThousand",
        "taiwaneseDigital",
        "ideographLegalTraditional",
    ],
    "counting-ja": ["japaneseCounting", "japaneseLegal", "japaneseDigitalTenThousand"],
    "counting-ko": ["koreanCounting", "koreanLegal", "koreanDigital", "koreanDigital2"],
    "counting-vi": ["vietnameseCounting"],
    "counting-hi": ["hindiCounting"],
    "counting-th": ["thaiCounting", "bahtText"],
    "counting-en-dollar": ["dollarText"],
}

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


def numbering_xml(fmts: list[str]) -> str:
    """One abstractNum per (format, start) pair; numId = 1-based, in order."""
    abstracts = []
    nums = []
    num_id = 0
    for i, fmt in enumerate(fmts):
        for start in (1, 10):
            abstract_id = num_id
            num_id += 1
            abstracts.append(
                f'<w:abstractNum w:abstractNumId="{abstract_id}">'
                f'<w:lvl w:ilvl="0">'
                f'<w:start w:val="{start}"/>'
                f'<w:numFmt w:val="{fmt}"/>'
                f'<w:lvlText w:val="%1"/>'
                f'<w:lvlJc w:val="left"/>'
                f'<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
                f"</w:lvl>"
                f"</w:abstractNum>"
            )
            nums.append(
                f'<w:num w:numId="{num_id}"><w:abstractNumId w:val="{abstract_id}"/></w:num>'
            )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + "".join(abstracts)
        + "".join(nums)
        + "</w:numbering>"
    )


def document_xml(fmts: list[str]) -> str:
    paras = []
    num_id = 0
    for fmt in fmts:
        for start in (1, 10):
            num_id += 1
            for item in range(3):
                n = start + item
                paras.append(
                    "<w:p><w:pPr>"
                    f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>'
                    "</w:pPr>"
                    f'<w:r><w:t xml:space="preserve">{fmt} item {n}</w:t></w:r></w:p>'
                )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body>"
        + "".join(paras)
        + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
        "</w:body></w:document>"
    )


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for stem, fmts in FIXTURES.items():
        target = OUT / f"{stem}.docx"
        # Fixed timestamps so regenerating an unchanged fixture produces
        # identical bytes and does not show up as a diff.
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in (
                ("[Content_Types].xml", CONTENT_TYPES),
                ("_rels/.rels", ROOT_RELS),
                ("word/_rels/document.xml.rels", DOC_RELS),
                ("word/document.xml", document_xml(fmts)),
                ("word/numbering.xml", numbering_xml(fmts)),
            ):
                info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(info, data)
        print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
