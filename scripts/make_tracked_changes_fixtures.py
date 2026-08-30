#!/usr/bin/env python3
"""Build test-files/tracked-changes.docx and tracked-changes-final.docx (issue #154).

Both share one body; they differ only in `w:settings/w:revisionView`, which is
the whole point — the display decision lives in the document:

1. "Alpha <ins Ann>beta </ins>gamma" — an unaccepted insertion mid-sentence.
2. "Delta <del Ann>epsilon </del>zeta" — an unaccepted deletion; the case
   with the correctness edge (its text must not paint in the final view).
3. "<ins Bob>eta</ins>" — a second author, pinning the second palette color.
4. "Theta" — the control: no revision anywhere near it, must never move.
5. "<strike>struck</strike> <dstrike>double</dstrike>" — plain §17.3.2.37
   formatting, NOT a revision: it must stay struck in BOTH views, which is
   what tells a revision mark apart from ordinary strike formatting.

tracked-changes.docx has no revisionView (spec default: markup shown) —
deletions render struck through, insertions underlined, both in per-author
colors. tracked-changes-final.docx adds `<w:revisionView w:insDel="0"/>` —
deletions suppressed, insertions plain.

Deterministic: fixed ZIP metadata. Regenerate and commit if content changes;
verify with `python3 scripts/verify_docx.py test-files/tracked-changes*.docx`.
"""

from pathlib import Path
import zipfile

OUT_DIR = Path(__file__).resolve().parent.parent / "test-files"

CONTENT_TYPES_BASE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

CONTENT_TYPES_SETTINGS = CONTENT_TYPES_BASE.replace(
    "</Types>",
    '  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n</Types>',
)

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOC_RELS_SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>
"""

SETTINGS_FINAL = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:revisionView w:insDel="0"/>
</w:settings>
"""


def run(text, rpr=""):
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def del_run(text):
    return f'<w:r><w:delText xml:space="preserve">{text}</w:delText></w:r>'


DOCUMENT = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>{run("Alpha ")}<w:ins w:id="1" w:author="Ann" w:date="2026-01-01T00:00:00Z">{run("beta ")}</w:ins>{run("gamma")}</w:p>
    <w:p>{run("Delta ")}<w:del w:id="2" w:author="Ann" w:date="2026-01-01T00:00:00Z">{del_run("epsilon ")}</w:del>{run("zeta")}</w:p>
    <w:p><w:ins w:id="3" w:author="Bob" w:date="2026-01-02T00:00:00Z">{run("eta")}</w:ins></w:p>
    <w:p>{run("Theta")}</w:p>
    <w:p>{run("struck", "<w:rPr><w:strike/></w:rPr>")}{run(" ")}{run("double", "<w:rPr><w:dstrike/></w:rPr>")}</w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>
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
        OUT_DIR / "tracked-changes.docx",
        [
            ("[Content_Types].xml", CONTENT_TYPES_BASE),
            ("_rels/.rels", ROOT_RELS),
            ("word/document.xml", DOCUMENT),
        ],
    )
    write(
        OUT_DIR / "tracked-changes-final.docx",
        [
            ("[Content_Types].xml", CONTENT_TYPES_SETTINGS),
            ("_rels/.rels", ROOT_RELS),
            ("word/_rels/document.xml.rels", DOC_RELS_SETTINGS),
            ("word/settings.xml", SETTINGS_FINAL),
            ("word/document.xml", DOCUMENT),
        ],
    )


if __name__ == "__main__":
    main()
