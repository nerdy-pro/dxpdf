#!/usr/bin/env python3
"""Build test-files/devanagari.docx — the issue #153 fixture.

Five paragraphs, each pinning one leg of complex-script shaping with the
shaped cluster as the spacing unit:

1. Plain Devanagari — "किताब क्षमता धर्म हिन्दी": a prebase matra (कि — the
   vowel is stored after क and drawn before it), a virama conjunct (क्ष), a
   reph (र्म), and a word mixing both (हिन्दी). Names no font anywhere, like
   issue-139-minimal.docx: the base face cannot draw Devanagari, so per-glyph
   fallback must pick a covering family first and shaping must then see the
   family fallback chose — the pass order build/block.rs enforces.
2. The same text with §17.3.2.35 `w:spacing w:val="40"` (2pt): spacing on a
   run whose unit is the shaped cluster, never the inside of a conjunct.
3. §17.3.1.13 `jc=distribute` over a short Devanagari line — the line's slack
   reaches the emitted command as distribution extra on a *shaped* run.
4. Latin control with the same `w:spacing w:val="40"`: must stay off the
   shaping path (shaped: none) with its spacing intact — the fast-path pin.
5. Arabic under `jc=distribute` — before #153 the distribution extra widened
   the decorations of a shaped run while the painter ignored it for the
   glyphs; this paragraph is the regression pin for that repair.

Deterministic: no timestamps, fixed ZIP metadata; non-ASCII is written as
numeric character references so this file stays pure ASCII. Regenerate and
commit the result if the content changes; verify with
`python3 scripts/verify_docx.py test-files/devanagari.docx`.
"""

from pathlib import Path
import zipfile

OUT = Path(__file__).resolve().parent.parent / "test-files" / "devanagari.docx"

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

# NCR spellings, so the source stays ASCII and the output cannot depend on
# the editor's encoding.
KITAAB = "&#x0915;&#x093F;&#x0924;&#x093E;&#x092C;"  # किताब — prebase matra
KSHAMATA = "&#x0915;&#x094D;&#x0937;&#x092E;&#x0924;&#x093E;"  # क्षमता — conjunct
DHARM = "&#x0927;&#x0930;&#x094D;&#x092E;"  # धर्म — reph
HINDI = "&#x0939;&#x093F;&#x0928;&#x094D;&#x0926;&#x0940;"  # हिन्दी
DEVA_SENTENCE = f"{KITAAB} {KSHAMATA} {DHARM} {HINDI}"
MAT_HAI = "&#x092E;&#x0924; &#x0939;&#x0948;"  # मत है — short distribute line
ARABIC = (
    "&#x0645;&#x0631;&#x062D;&#x0628;&#x0627; "
    "&#x0628;&#x0627;&#x0644;&#x0639;&#x0627;&#x0644;&#x0645;"
)  # مرحبا بالعالم

SPACING_2PT = '<w:rPr><w:spacing w:val="40"/></w:rPr>'


def run(text, rpr=""):
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(runs, ppr=""):
    return f"<w:p>{ppr}{runs}</w:p>"


DOCUMENT = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {para(run(DEVA_SENTENCE))}
    {para(run(DEVA_SENTENCE, SPACING_2PT))}
    {para(run(MAT_HAI), '<w:pPr><w:jc w:val="distribute"/></w:pPr>')}
    {para(run("Latin control", SPACING_2PT))}
    {para(run(ARABIC), '<w:pPr><w:jc w:val="distribute"/></w:pPr>')}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>
"""

PARTS = [
    ("[Content_Types].xml", CONTENT_TYPES),
    ("_rels/.rels", ROOT_RELS),
    ("word/document.xml", DOCUMENT),
]


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in PARTS:
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.external_attr = 0o600 << 16
            z.writestr(info, content)
    print(f"wrote {OUT} ({OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
