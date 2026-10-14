#!/usr/bin/env python3
"""Build the §17.15.1.10 automatic-hyphenation fixtures under test-files/
(issue #158).

Seven packages, one document in seven settings:

  auto-hyphenation-base.docx       `<w:autoHyphenation/>`, zone omitted (360
                                   twips per §17.15.1.53), no limit, capitals
                                   hyphenated — the spec defaults.
  auto-hyphenation-zone-wide.docx  + `<w:hyphenationZone w:val="7200"/>`: a
                                   zone wider than the column, so no gap can
                                   ever exceed it and nothing is hyphenated.
  auto-hyphenation-limit-1.docx    + zone 0 and `<w:consecutiveHyphenLimit
                                   w:val="1"/>`: no two consecutive lines may
                                   end in an inserted hyphen.
  auto-hyphenation-nocaps.docx     + zone 0 and `<w:doNotHyphenateCaps/>`: the
                                   ALL-CAPS compound stays whole while the
                                   lowercase one beside it is still cut.
  auto-hyphenation-off.docx        `<w:autoHyphenation w:val="0"/>`: the
                                   document says no; nothing is hyphenated.
  auto-hyphenation-compat15.docx   The base document with
                                   `compatibilityMode` 15 instead of 14 — a
                                   **probe**, not a fixture: Word's VBA
                                   documentation says native-mode Word
                                   ignores the hyphenation zone, which would
                                   make this render differently from the
                                   base in Word while this engine (which does
                                   not read `w:compat`) renders the two
                                   identically. Kept so the question survives
                                   between sessions; asserted by nothing.
  auto-hyphenation-columns.docx    Zone 0, one long paragraph flowing from a
                                   2 in column into a 4 in one. A word
                                   hyphenated for the narrow column may fit
                                   whole in the wide one, so the continuation
                                   re-fit must see the paragraph's words, not
                                   this placement's cut fragments.

The text column is 2 in wide (4680-twip margins either side of a Letter
page), justified, Times New Roman 12 pt, `de-DE` from `docDefaults`, so that
German compounds overflow often and hyphenation is visible. The compounds are
chosen so the TeX `de-1996` patterns and Word's dictionary agree on their
seams (Donau|dampf|schiff|fahrt, Kraft|fahr|zeug …) — this engine hyphenates
from patterns, Word from its dictionary, and the two differ on some words, so
"breaks where Word does" can only be asserted for words where both find the
same cut. No paragraph contains an authored hyphen or a digit, so every `-`
at a line end in the render is one the engine inserted.

Paragraphs, in order:

  1. the German prose (hyphenated under the document setting);
  2. the same prose with `<w:suppressAutoHyphens/>` — §17.3.1.34 exempts it;
  3. `KRAFTFAHRZEUG` and `Kraftfahrzeug` alternating behind prefixes of
     varying length, so each lands at several column positions and at least
     one of each is cut when the setting allows it;
  4. German prose carrying a run tagged `en-US` — a word in that run must be
     cut at an American-English pattern point, not a German one.

The columns document holds paragraph 1 six times over as one paragraph.

Same package skeleton as make_issue165_fixtures.py (whose docstring records
why the skeleton is as full as it is), with a `w:compat` block because the
question these documents ask Word is compat-mode dependent. Deterministic:
fixed ZIP dates and a fixed docProps timestamp. Regenerate rather than
hand-edit; run `scripts/verify_docx.py` on the result.

    python3 scripts/make_auto_hyphenation_fixtures.py
"""

import pathlib
import zipfile

OUT = pathlib.Path(__file__).resolve().parent.parent / "test-files"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"
MS_WORD = "http://schemas.microsoft.com/office/word"

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
STAMP = "2026-10-12T00:00:00Z"


def join(*parts):
    """Concatenate with no separator — `w:body` is element-only content."""
    return "".join(parts)


# ── package skeleton ─────────────────────────────────────────────────────────

STYLES = f"""{XML_DECL}<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr>\
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>\
<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="de-DE" w:eastAsia="de-DE" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>\
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>\
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal">\
<w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>"""

FONT_TABLE = f"""{XML_DECL}<w:fonts xmlns:w="{W}"><w:font w:name="Times New Roman">\
<w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font></w:fonts>"""

WEB_SETTINGS = f'{XML_DECL}<w:webSettings xmlns:w="{W}"><w:optimizeForBrowser/></w:webSettings>'

CORE = f"""{XML_DECL}<cp:coreProperties \
xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" \
xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" \
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\
<dc:title>dxpdf issue #158 fixture</dc:title><dc:creator>dxpdf</dc:creator>\
<cp:lastModifiedBy>dxpdf</cp:lastModifiedBy>\
<dcterms:created xsi:type="dcterms:W3CDTF">{STAMP}</dcterms:created>\
<dcterms:modified xsi:type="dcterms:W3CDTF">{STAMP}</dcterms:modified></cp:coreProperties>"""

APP = f"""{XML_DECL}<Properties \
xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" \
xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\
<Application>dxpdf</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>\
<SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>\
<AppVersion>16.0000</AppVersion></Properties>"""

CONTENT_TYPES = (
    f'{XML_DECL}<Types xmlns="{CT_NS}">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    + "".join(
        f'<Override PartName="/{p}" ContentType="{c}"/>'
        for p, c in [
            ("word/document.xml", f"{WML}.document.main+xml"),
            ("word/styles.xml", f"{WML}.styles+xml"),
            ("word/settings.xml", f"{WML}.settings+xml"),
            ("word/webSettings.xml", f"{WML}.webSettings+xml"),
            ("word/fontTable.xml", f"{WML}.fontTable+xml"),
            ("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
            (
                "docProps/app.xml",
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
            ),
        ]
    )
    + "</Types>"
)

ROOT_RELS = (
    f'{XML_DECL}<Relationships xmlns="{NS_PKG_REL}">'
    f'<Relationship Id="rId1" Type="{NS_R}/officeDocument" Target="word/document.xml"/>'
    f'<Relationship Id="rId2" Type="{NS_PKG_REL}/metadata/core-properties" '
    'Target="docProps/core.xml"/>'
    f'<Relationship Id="rId3" Type="{NS_R}/extended-properties" Target="docProps/app.xml"/>'
    "</Relationships>"
)

DOC_RELS = (
    f'{XML_DECL}<Relationships xmlns="{NS_PKG_REL}">'
    f'<Relationship Id="rId1" Type="{NS_R}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rId2" Type="{NS_R}/settings" Target="settings.xml"/>'
    f'<Relationship Id="rId3" Type="{NS_R}/webSettings" Target="webSettings.xml"/>'
    f'<Relationship Id="rId4" Type="{NS_R}/fontTable" Target="fontTable.xml"/>'
    "</Relationships>"
)


def settings_xml(body, compat_mode):
    compat = (
        "<w:compat>"
        f'<w:compatSetting w:name="compatibilityMode" w:uri="{MS_WORD}" w:val="{compat_mode}"/>'
        "</w:compat>"
    )
    return f'{XML_DECL}<w:settings xmlns:w="{W}">{body}{compat}</w:settings>'


def write(name, document, settings_body, compat_mode=14):
    parts = [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("docProps/core.xml", CORE),
        ("docProps/app.xml", APP),
        ("word/_rels/document.xml.rels", DOC_RELS),
        ("word/document.xml", document),
        ("word/styles.xml", STYLES),
        ("word/settings.xml", settings_xml(settings_body, compat_mode)),
        ("word/webSettings.xml", WEB_SETTINGS),
        ("word/fontTable.xml", FONT_TABLE),
    ]
    path = OUT / name
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for member, body in parts:
            info = zipfile.ZipInfo(member, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            zf.writestr(info, body)
    print(f"wrote {path.name} ({path.stat().st_size} bytes)")


def document(body):
    return f'{XML_DECL}<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>'


# ── content ──────────────────────────────────────────────────────────────────

# Long compounds with seams the de-1996 patterns find (checked against the
# vendored dictionary): Do|nau|dampf|schiff|fahrt, Ge|schwin|dig|keits|be|gren|zung, Bin|nen|schiff|fahrt, Ar|beits|ge|mein|schaft,
# Wirt|schafts|wis|sen|schaf|ten, Bun|des|ver|fas|sungs|ge|richt,
# Recht|schreib|re|form, Zei|tungs|re|dak|ti|on, Fahr|kar|ten|au|to|mat,
# Ei|sen|bahn|kno|ten|punkt, Um|welt|schutz|or|ga|ni|sa|ti|on,
# Hoch|ge|schwin|dig|keits|zug, Le|bens|ver|si|che|rung, Sprach|wis|sen|schaft.
# Every word is narrower than the 144 pt column, so none takes the
# per-cluster path `split_oversized_fragments` keeps for over-wide words:
# `Donaudampfschifffahrtsgesellschaft` (34 letters, ~190 pt) did, and came
# out as `…gesel|lschaft`, which is that path's pre-existing behaviour and
# not what this fixture is about.
PROSE = (
    "Die Donaudampfschifffahrt beschloss, die Geschwindigkeitsbegrenzung "
    "auf der Binnenschifffahrt gemeinsam mit einer Arbeitsgemeinschaft der "
    "Wirtschaftswissenschaften zu überprüfen. Das Bundesverfassungsgericht verhandelte "
    "die Rechtschreibreform, während die Zeitungsredaktion den Fahrkartenautomaten am "
    "Eisenbahnknotenpunkt beschrieb. Die Umweltschutzorganisation forderte vom "
    "Hochgeschwindigkeitszug eine Lebensversicherung für die Sprachwissenschaft."
)

# `Kraft|fahr|zeug`, 13 letters (about 100 pt in 12 pt Times): narrower than
# the 144 pt column, so it is never cut by `split_oversized_fragments` and
# only ever by hyphenation; wide enough to overflow often. The prefixes
# between the words step the column position so that at least one capital
# and one lowercase instance each land where a cut is both needed and fits.
CAPS = (
    "Ein KRAFTFAHRZEUG und ein Kraftfahrzeug; dann das KRAFTFAHRZEUG, dann wieder ein "
    "Kraftfahrzeug und noch ein KRAFTFAHRZEUG sowie das letzte Kraftfahrzeug; am Ende "
    "folgt das KRAFTFAHRZEUG und davor das Kraftfahrzeug, zuletzt ein KRAFTFAHRZEUG."
)

MIXED_BEFORE = "Die Textverarbeitung schreibt "
# en-US patterns: in|ter|na|tional un|der|stand|ing au|to|mat|i|cally
# in|for|ma|tion dic|tio|nary hy|phen|ation para|graph.
ENGLISH = (
    "the international understanding of automatically generated information, "
    "dictionary hyphenation and paragraph documentation"
)
MIXED_AFTER = " in die Silbentrennung der Sprachwissenschaft hinein."


def run(text, lang=None):
    rpr = f'<w:rPr><w:lang w:val="{lang}"/></w:rPr>' if lang else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(runs, suppress=False):
    ppr = '<w:pPr><w:jc w:val="both"/>' + ("<w:suppressAutoHyphens/>" if suppress else "") + "</w:pPr>"
    return f"<w:p>{ppr}{runs}</w:p>"


# Letter page, 4680-twip side margins → a 2880-twip (2 in, 144 pt) column.
SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="4680" w:bottom="1440" w:left="4680" '
    'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)

# Letter page, 1440-twip margins (9360 twips of text), split unequally:
# 2880 (2 in) + 720 gap + 5760 (4 in). §17.6.4 `w:equalWidth="0"`.
SECT_COLUMNS = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="720" w:footer="720" w:gutter="0"/>'
    '<w:cols w:num="2" w:space="720" w:equalWidth="0">'
    '<w:col w:w="2880" w:space="720"/><w:col w:w="5760"/></w:cols></w:sectPr>'
)

BODY = join(
    para(run(PROSE)),
    para(run(PROSE), suppress=True),
    para(run(CAPS)),
    para(join(run(MIXED_BEFORE), run(ENGLISH, lang="en-US"), run(MIXED_AFTER))),
    SECT,
)

ON = "<w:autoHyphenation/>"
ZONE_0 = '<w:hyphenationZone w:val="0"/>'


def main():
    doc = document(BODY)
    write("auto-hyphenation-base.docx", doc, ON)
    write("auto-hyphenation-zone-wide.docx", doc, ON + '<w:hyphenationZone w:val="7200"/>')
    write(
        "auto-hyphenation-limit-1.docx",
        doc,
        ON + ZONE_0 + '<w:consecutiveHyphenLimit w:val="1"/>',
    )
    write("auto-hyphenation-nocaps.docx", doc, ON + ZONE_0 + "<w:doNotHyphenateCaps/>")
    write("auto-hyphenation-off.docx", doc, '<w:autoHyphenation w:val="0"/>')
    write("auto-hyphenation-compat15.docx", doc, ON, compat_mode=15)

    # One paragraph long enough to fill the 2 in column and spill into the
    # 4 in one: paragraph 1 six times over.
    columns = document(join(para(run(" ".join([PROSE] * 6))), SECT_COLUMNS))
    write("auto-hyphenation-columns.docx", columns, ON + ZONE_0)


if __name__ == "__main__":
    main()
