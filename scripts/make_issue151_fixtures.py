#!/usr/bin/env python3
"""Build the issue #151 fixtures: page borders and an applied document grid.

Two documents, each two sections, each section its own probe:

``page-borders.docx``
    Section 1: ``<w:pgBorders w:offsetFrom="page">`` with four *distinct*
    sides — every side a different colour, width and space, so a test can
    pin each edge's band position independently and a transposed edge is
    caught by colour, not just coordinates. Section 2: the same four sides
    with ``offsetFrom="text"``, so the two sections differ only in the
    reference frame the offsets measure from.

``doc-grid.docx``
    Section 1: ``<w:docGrid w:type="lines" w:linePitch="360"/>`` (the
    Japanese-Word default, 18pt) over four paragraphs — a wrapping Latin
    paragraph, a CJK paragraph, a ``<w:snapToGrid w:val="0"/>`` opt-out, and
    a 20pt-font paragraph whose natural line is taller than one pitch (two
    slots). Section 2: the identical content under ``<w:docGrid
    w:linePitch="360"/>`` with **no type** — the grid every Word document
    carries and no renderer may apply — making section 2 the ungridded
    control the tests measure section 1 against.

Both are plain OPC packages (document + content types + rels), verified by
``scripts/verify_docx.py``.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
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


def document(body: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body>" + body + "</w:body></w:document>"
    )


def para(text: str, ppr: str = "") -> str:
    ppr_xml = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    return f'<w:p>{ppr_xml}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def write_docx(path: Path, document_xml: str) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/document.xml", document_xml),
        ):
            # Fixed date_time so the archive is reproducible byte-for-byte.
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    path.write_bytes(buf.getvalue())
    print(f"wrote {path.relative_to(ROOT)} ({path.stat().st_size} bytes)")


# US Letter, 1in margins throughout, so the tests can use round numbers.
# CT_SectPr is a strict xsd:sequence: pgSz, pgMar, then pgBorders, ...,
# then docGrid near the end — Word refuses packages that reorder it (the
# issue-165 lesson: this parser tolerates what Word will not open).
SECT_GEOMETRY = (
    '<w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
    ' w:header="720" w:footer="720" w:gutter="0"/>'
)


def make_page_borders() -> None:
    # Four distinct sides: colour, sz (eighths) and space (points) all differ.
    sides = (
        '<w:top w:val="single" w:sz="48" w:space="24" w:color="FF0000"/>'
        '<w:left w:val="single" w:sz="24" w:space="12" w:color="00FF00"/>'
        '<w:bottom w:val="double" w:sz="24" w:space="6" w:color="0000FF"/>'
        '<w:right w:val="single" w:sz="8" w:space="0" w:color="FF00FF"/>'
    )
    body = (
        para("Section one: borders offset from the page edge.")
        # A mid-body sectPr closes section 1 (its own paragraph).
        + f'<w:p><w:pPr><w:sectPr>{SECT_GEOMETRY}<w:pgBorders w:offsetFrom="page">{sides}</w:pgBorders>'
        + "</w:sectPr></w:pPr></w:p>"
        + para("Section two: the same borders offset from the text margins.")
        + f'<w:sectPr>{SECT_GEOMETRY}<w:pgBorders w:offsetFrom="text">{sides}</w:pgBorders></w:sectPr>'
    )
    write_docx(OUT / "page-borders.docx", document(body))


def make_doc_grid() -> None:
    latin = (
        "The quick brown fox jumps over the lazy dog and keeps running until "
        "this paragraph is comfortably longer than a single line of a page."
    )
    cjk = "文書グリッドは行送りを揃えます。" * 4
    content = (
        para(latin)
        + para(cjk)
        + para("This paragraph opts out of the grid.", '<w:snapToGrid w:val="0"/>')
        + para(
            "Tall text takes two grid slots.",
            "",
        ).replace(
            "<w:r>",
            '<w:r><w:rPr><w:sz w:val="40"/></w:rPr>',
        )
    )
    body = (
        content
        + f'<w:p><w:pPr><w:sectPr>{SECT_GEOMETRY}<w:docGrid w:type="lines" w:linePitch="360"/>'
        + "</w:sectPr></w:pPr></w:p>"
        + content
        + f'<w:sectPr>{SECT_GEOMETRY}<w:docGrid w:linePitch="360"/></w:sectPr>'
    )
    write_docx(OUT / "doc-grid.docx", document(body))


if __name__ == "__main__":
    make_page_borders()
    make_doc_grid()
