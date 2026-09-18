#!/usr/bin/env python3
"""Build test-files/comments.docx and comments-hidden.docx (issue #154).

One body, two views — like the tracked-changes pair:

- Paragraph 1 holds a range commented by Ann (comment 1, two paragraphs of
  body — the balloon must hold both).
- A range opened in paragraph 2 closes in paragraph 3: the shading stamp has
  to survive a paragraph boundary (comment 2, by Bob — second palette color;
  Bob also has no initials, pinning the optional attribute).
- Paragraph 4 is the control: no comment anywhere near it.

The section's right margin is 2160 twips (1.5in) so balloons have room; the
`comment-reference.docx` corpus fixture keeps the narrow-margin case.
comments-hidden.docx adds `<w:revisionView w:comments="0"/>`: no shading, no
anchors, no balloons.

Deterministic: fixed ZIP metadata; non-ASCII as NCRs. Regenerate and commit
if content changes; verify with `python3 scripts/verify_docx.py`.
"""

from pathlib import Path
import zipfile

OUT_DIR = Path(__file__).resolve().parent.parent / "test-files"

CT_BASE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>
"""

CT_SETTINGS = CT_BASE.replace(
    "</Types>",
    '  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n</Types>',
)

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOC_RELS_BASE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>
"""

DOC_RELS_SETTINGS = DOC_RELS_BASE.replace(
    "</Relationships>",
    '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\n</Relationships>',
)

SETTINGS_HIDDEN = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:revisionView w:comments="0"/>
</w:settings>
"""


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def ref(cid):
    return f'<w:r><w:commentReference w:id="{cid}"/></w:r>'


DOCUMENT = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>{run("Before ")}<w:commentRangeStart w:id="1"/>{run("first range")}<w:commentRangeEnd w:id="1"/>{ref(1)}{run(" after")}</w:p>
    <w:p>{run("Open ")}<w:commentRangeStart w:id="2"/>{run("spans down")}</w:p>
    <w:p>{run("and closes")}<w:commentRangeEnd w:id="2"/>{ref(2)}{run(" here")}</w:p>
    <w:p>{run("Control paragraph")}</w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="2160" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>
"""

COMMENTS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="1" w:author="Ann" w:initials="A" w:date="2026-01-01T00:00:00Z">
    <w:p><w:r><w:annotationRef/></w:r><w:r><w:t>First balloon line.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second balloon line.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="2" w:author="Bob" w:date="2026-01-02T00:00:00Z">
    <w:p><w:r><w:annotationRef/></w:r><w:r><w:t>Cross-paragraph note.</w:t></w:r></w:p>
  </w:comment>
</w:comments>
"""


def write(path, parts):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in parts:
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.external_attr = 0o600 << 16
            z.writestr(info, content)
    print(f"wrote {path} ({path.stat().st_size} bytes)")


def main():
    write(
        OUT_DIR / "comments.docx",
        [
            ("[Content_Types].xml", CT_BASE),
            ("_rels/.rels", ROOT_RELS),
            ("word/_rels/document.xml.rels", DOC_RELS_BASE),
            ("word/document.xml", DOCUMENT),
            ("word/comments.xml", COMMENTS),
        ],
    )
    write(
        OUT_DIR / "comments-hidden.docx",
        [
            ("[Content_Types].xml", CT_SETTINGS),
            ("_rels/.rels", ROOT_RELS),
            ("word/_rels/document.xml.rels", DOC_RELS_SETTINGS),
            ("word/document.xml", DOCUMENT),
            ("word/comments.xml", COMMENTS),
            ("word/settings.xml", SETTINGS_HIDDEN),
        ],
    )


if __name__ == "__main__":
    main()
