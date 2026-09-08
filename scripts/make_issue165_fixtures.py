#!/usr/bin/env python3
"""Build the issue #165 probe fixtures — the documents that ask Word the three
questions ECMA-376 leaves open.

Each is authored so the answer is a *measurement off the rendered page*, and so
that every candidate reading predicts a visibly different number. The
predictions are recorded in the comment above each builder below and in
`plans/issue-165-word-reference-renders.md`; they are written before rendering
on purpose, because with a PDF in hand it is easy to decide after the fact which
reading the output "obviously" supports.

    python3 scripts/make_issue165_fixtures.py

# These have to open in *Word*, which is stricter than this engine

The first cut of these fixtures was rejected by Word with "unreadable content"
while dxpdf parsed them, `textutil` read them, and the test suite stayed green.

**The cause was one wrong namespace**, and it is the only one of the three
things fixed here that actually mattered: the `.rels` parts declared
`xmlns=".../officeDocument/2006/relationships"`, which is the URI relationship
*types* are built from. A relationships *part* must be
`.../package/2006/relationships`. Word could not read the relationship part, so
it could not find `document.xml`, so the file was unreadable. See `NS_PKG_REL`
below — the two constants are kept apart and commented for exactly this reason,
and `scripts/verify_docx.py` now checks it.

Two other things were wrong and are fixed too, though neither alone would have
stopped Word opening the file:

1. **The package was bare** — only `[Content_Types].xml`, `_rels/.rels` and
   `word/document.xml`. A real Word document also carries
   `word/_rels/document.xml.rels`, `styles.xml`, `settings.xml`, `fontTable.xml`,
   `webSettings.xml` and `docProps/`, so the full skeleton is built below.
2. **Whitespace text nodes.** `w:tbl`, `w:tr` and `w:body` are element-only
   content models, and the pretty-printed source put newlines *between* their
   children. `join()` below emits no inter-element whitespace at all.

None of the package skeleton comes from a real document: `test-cases/` holds
genuine client files, and copying their `docProps` or `theme` into a committed
fixture would publish someone's metadata. Everything here is written from
scratch.

Deterministic: fixed ZIP dates, a fixed docProps timestamp, and a hand-built
PNG, so re-running produces byte-identical archives. Regenerate rather than
hand-edit.
"""

import pathlib
import struct
import zlib
import zipfile

OUT = pathlib.Path(__file__).resolve().parent.parent / "test-files"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
# Two different URIs that are easy to conflate, and conflating them is why the
# first cut of these fixtures failed to open in Word at all: a `.rels` part
# lives in the *package* relationships namespace, while a relationship's `Type`
# (and `r:embed` in document.xml) uses the *officeDocument* one. Getting the
# `xmlns` wrong makes the whole part unreadable, so Word cannot even find
# document.xml and offers to recover the file.
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
# A fixed instant, so the archive is reproducible.
STAMP = "2026-08-12T00:00:00Z"


def join(*parts):
    """Concatenate XML fragments with **no** separator.

    Deliberately not `"\\n".join`. `w:body`, `w:tbl` and `w:tr` are element-only
    content models; a newline between their children is a text node, and Word
    refuses the document over it even though tolerant readers do not.
    """
    return "".join(parts)


# ── package skeleton ─────────────────────────────────────────────────────────

STYLES = f"""{XML_DECL}<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr>\
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>\
<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>\
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
<dc:title>dxpdf issue #165 probe</dc:title><dc:creator>dxpdf</dc:creator>\
<cp:lastModifiedBy>dxpdf</cp:lastModifiedBy>\
<dcterms:created xsi:type="dcterms:W3CDTF">{STAMP}</dcterms:created>\
<dcterms:modified xsi:type="dcterms:W3CDTF">{STAMP}</dcterms:modified></cp:coreProperties>"""

APP = f"""{XML_DECL}<Properties \
xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" \
xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\
<Application>dxpdf</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>\
<SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>\
<AppVersion>16.0000</AppVersion></Properties>"""


def content_types(image=False):
    overrides = "".join(
        f'<Override PartName="/{p}" ContentType="{c}"/>'
        for p, c in [
            ("word/document.xml", f"{WML}.document.main+xml"),
            ("word/styles.xml", f"{WML}.styles+xml"),
            ("word/settings.xml", f"{WML}.settings+xml"),
            ("word/webSettings.xml", f"{WML}.webSettings+xml"),
            ("word/fontTable.xml", f"{WML}.fontTable+xml"),
            (
                "docProps/core.xml",
                "application/vnd.openxmlformats-package.core-properties+xml",
            ),
            (
                "docProps/app.xml",
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
            ),
        ]
    )
    png = '<Default Extension="png" ContentType="image/png"/>' if image else ""
    return (
        f'{XML_DECL}<Types xmlns="{CT_NS}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        f'<Default Extension="xml" ContentType="application/xml"/>{png}{overrides}</Types>'
    )


ROOT_RELS = (
    f'{XML_DECL}<Relationships xmlns="{NS_PKG_REL}">'
    f'<Relationship Id="rId1" Type="{NS_R}/officeDocument" Target="word/document.xml"/>'
    # Core properties is the one relationship type that lives under the package
    # namespace rather than the officeDocument one.
    f'<Relationship Id="rId2" Type="{NS_PKG_REL}/metadata/core-properties" '
    'Target="docProps/core.xml"/>'
    f'<Relationship Id="rId3" Type="{NS_R}/extended-properties" Target="docProps/app.xml"/>'
    "</Relationships>"
)


def doc_rels(image=False):
    img = (
        f'<Relationship Id="rIdI" Type="{NS_R}/image" Target="media/dot.png"/>'
        if image
        else ""
    )
    return (
        f'{XML_DECL}<Relationships xmlns="{NS_PKG_REL}">'
        f'<Relationship Id="rId1" Type="{NS_R}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rId2" Type="{NS_R}/settings" Target="settings.xml"/>'
        f'<Relationship Id="rId3" Type="{NS_R}/webSettings" Target="webSettings.xml"/>'
        f'<Relationship Id="rId4" Type="{NS_R}/fontTable" Target="fontTable.xml"/>'
        f"{img}</Relationships>"
    )


def solid_png(width, height, rgb):
    """A minimal solid-colour PNG, from stdlib so the fixture needs no Pillow."""

    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c))

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + bytes(rgb) * width for _ in range(height))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(raw, 9))
        + chunk(b"IEND", b"")
    )


def write(name, document, settings_body="", image=None):
    settings = f'{XML_DECL}<w:settings xmlns:w="{W}">{settings_body}</w:settings>'
    parts = [
        ("[Content_Types].xml", content_types(image is not None)),
        ("_rels/.rels", ROOT_RELS),
        ("docProps/core.xml", CORE),
        ("docProps/app.xml", APP),
        ("word/_rels/document.xml.rels", doc_rels(image is not None)),
        ("word/document.xml", document),
        ("word/styles.xml", STYLES),
        ("word/settings.xml", settings),
        ("word/webSettings.xml", WEB_SETTINGS),
        ("word/fontTable.xml", FONT_TABLE),
    ]
    if image is not None:
        parts.append(("word/media/dot.png", image))

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


def para(text=None):
    inner = f"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>" if text else ""
    return f"<w:p>{inner}</w:p>"


SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)

BORDER_SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")
TBL_BORDERS = (
    "<w:tblBorders>"
    + "".join(
        f'<w:{s} w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        for s in BORDER_SIDES
    )
    + "</w:tblBorders>"
)
TC_BORDERS = (
    "<w:tcBorders>"
    + "".join(
        f'<w:{s} w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
        for s in ("top", "left", "bottom", "right")
    )
    + "</w:tcBorders>"
)


# ── A. vMerge overflow distribution ──────────────────────────────────────────
#
# Column 1 is a restart+continue pair holding ten lines; column 2 holds one
# short line per row, so the row boundary is a visible rule whose y can be
# measured. NO w:trHeight anywhere — the question is what the auto sizer does,
# and an authored height would answer a different one.
#
#   even distribution (dxpdf today) → boundary at H/2      (measured: 0.500)
#   last row absorbs                → boundary at h,   near the top  (~0.12)
#   restart row absorbs             → boundary at H-h, near the bottom (~0.88)
def build_vmerge():
    tall = join(*(para(f"merged line {i:02d}") for i in range(1, 11)))
    body = join(
        para("A: vMerge overflow distribution"),
        "<w:tbl>",
        f'<w:tblPr><w:tblW w:w="8640" w:type="dxa"/>{TBL_BORDERS}'
        '<w:tblLayout w:type="fixed"/></w:tblPr>',
        '<w:tblGrid><w:gridCol w:w="4320"/><w:gridCol w:w="4320"/></w:tblGrid>',
        "<w:tr>",
        f'<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/>'
        f'<w:vMerge w:val="restart"/>{TC_BORDERS}</w:tcPr>{tall}</w:tc>',
        f'<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/>{TC_BORDERS}</w:tcPr>'
        f"{para('R1')}</w:tc>",
        "</w:tr>",
        "<w:tr>",
        f'<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/><w:vMerge/>{TC_BORDERS}'
        f"</w:tcPr>{para()}</w:tc>",
        f'<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/>{TC_BORDERS}</w:tcPr>'
        f"{para('R2')}</w:tc>",
        "</w:tr>",
        "</w:tbl>",
        para(),
        SECT,
    )
    write("issue-165-vmerge.docx", document(body))


# ── B. tblCellSpacing at the table's own edges ───────────────────────────────
#
# 20pt spacing, every border drawn (per [MS-OI29500] §17.4.66 a non-zero
# spacing means all borders display, which is what makes the gaps measurable).
#
#   one full spacing everywhere (dxpdf today) → edge 20pt, inner 20pt
#                                               (measured: 20.4 / 21.1)
#   half at the edges                         → edge 10pt, inner 20pt
def build_cellspacing():
    cells = join(
        *(
            f'<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/>{TC_BORDERS}</w:tcPr>'
            f"{para(f'C{i}')}</w:tc>"
            for i in (1, 2, 3)
        )
    )
    body = join(
        para("B: tblCellSpacing at the table edges"),
        "<w:tbl>",
        f'<w:tblPr><w:tblW w:w="7200" w:type="dxa"/>'
        f'<w:tblCellSpacing w:w="400" w:type="dxa"/>{TBL_BORDERS}'
        '<w:tblLayout w:type="fixed"/></w:tblPr>',
        '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/>'
        '<w:gridCol w:w="2400"/></w:tblGrid>',
        '<w:tr><w:trPr><w:tblCellSpacing w:w="400" w:type="dxa"/></w:trPr>',
        cells,
        "</w:tr>",
        "</w:tbl>",
        para(),
        SECT,
    )
    write("issue-165-cellspacing.docx", document(body))


# ── D. tblCellSpacing magnitude, and carve-vs-add ────────────────────────────
#
# MEASURED IN WORD, 2026-08-18. This probe was built to ask three questions and
# the render answered all three; what follows records both the questions and the
# answers, because two of them refuted a written finding rather than filling a
# blank. `tests/table_cell_spacing.rs` asserts the answers and
# `build::table::resolve_cell_spacing` holds the reasoning.
#
# Probe B settled *where* the spacing goes (edge gap = inter-cell gap) and then
# turned up something it was not built to ask: Word's gaps came out about twice
# this engine's. Neither ECMA-376 §17.4.44 nor [MS-OI29500] states a factor —
# both say only "the minimum amount of space which shall be left between all
# cells in the table including the width of the table borders in the
# calculation" — so the factor had to be measured, not read.
#
# Four tables, identical but for their spacing, all `tblW=7200` (360pt) with
# three 2400-twip (120pt) columns and `tblLayout=fixed`. Stacked so their edges
# line up on the page, which is what makes the second question free to read.
#
# **Question 1 — the factor. ANSWER: the declared value is a half-gap.**
#
#   declared value is the gap (dxpdf then) → S200 gap 10pt, S400 gap 20pt
#   declared value is a half-gap  ← WORD   → S200 gap 20pt, S400 gap 40pt
#
# Two values rather than one so the answer is a *ratio* rather than a single
# reading: whatever the factor is, S400's gap must be exactly twice S200's. A
# constant offset — a border width folded in, say — would break that and is
# worth knowing about before anything is multiplied by anything. It did not:
# the ratio holds, and the gap is uniform at 2x across both tables.
#
# The reading survives being taken off a screenshot because it is scale-free.
# With `3w + 4g = 360pt`, table 3 shows 189px of cell against 120px of gap — a
# ratio of 1.58, forcing g ≈ 41pt. Had the declared 20pt been the gap, the cells
# would be 93.3pt and the ratio 4.67. Different picture, not different rounding.
#
# This is what killed the argument that there is no factor. It came from
# ONLYOFFICE, an independent implementation that both renders and targets Word
# compatibility, and it explained Word's doubling as Word *summing* the
# `tblPr` and `trPr` declarations `issue-165-cellspacing.docx` happens to carry
# in both places. Tables 2 and 3 here declare the spacing at table level ONLY.
# There is nothing to sum, and Word doubles them anyway.
#
# **Question 2 — carve or add. ANSWER: carved.**
#
#   carved out of the table  ← WORD → all four tables the same width,
#                                     cells shrink as spacing grows
#   added to the table              → each table wider than the one above
#
# Just look at whether the right edges line up. They do.
#
# **Question 3 — precedence. ANSWER: the row value supersedes.**
#
# §17.4.44's own text says the table-level value "shall be superseded by a
# table-level exception or the row cell spacing value in that order", and the
# doubt was never about the text — it was whether Word obeys it, since probe B
# declares the same value at both levels and cannot tell supersede from sum.
# The last table here declares 400 at table level and 800 on its row, and with
# the factor known the three readings predict three different pictures:
#
#   supersede  ← WORD → 80pt gaps, 13.3pt cells (Word draws 13.5)
#   table wins        → 40pt gaps, 66.7pt cells
#   sum               → 120pt gaps, which 360pt of table cannot hold
#
# The cells come out narrow enough that their labels wrap to one glyph per line
# — in Word as in dxpdf, which is the most distinctive number in the fixture and
# the reason this one is not a close call.
#
# Still unmeasured, and the reason `resolve_table_cell_spacing` keeps a warning:
# rows that disagree with *each other*. Every table here is one row, so nothing
# says whether such a row re-carves the grid from its own spacing, what the
# table's reported width then is, or whether border collapsing stays table-wide.
# A fifth table carrying the row value on one row only would ask it.
def build_cellspacing_scale():
    # Cell text is one unbroken alphanumeric token per cell — no space, no
    # hyphen. UAX #14 breaks at both, and the line fitter emits one draw command
    # per piece, so `S=400 C1` would reach the page as two commands and
    # `S400-C1` as two more. A test that identifies a table by its cell string
    # needs that string to survive as one command, and needs no other cell's
    # string to be a prefix of it.
    def table(label, tag, tbl_spacing, row_spacing=None):
        cells = join(
            *(
                f'<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/>{TC_BORDERS}</w:tcPr>'
                f"{para(f'{tag}C{i}')}</w:tc>"
                for i in (1, 2, 3)
            )
        )
        spacing = (
            f'<w:tblCellSpacing w:w="{tbl_spacing}" w:type="dxa"/>'
            if tbl_spacing is not None
            else ""
        )
        tr_pr = (
            f'<w:trPr><w:tblCellSpacing w:w="{row_spacing}" w:type="dxa"/></w:trPr>'
            if row_spacing is not None
            else ""
        )
        return join(
            para(label),
            "<w:tbl>",
            f'<w:tblPr><w:tblW w:w="7200" w:type="dxa"/>{spacing}{TBL_BORDERS}'
            '<w:tblLayout w:type="fixed"/></w:tblPr>',
            '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/>'
            '<w:gridCol w:w="2400"/></w:tblGrid>',
            f"<w:tr>{tr_pr}",
            cells,
            "</w:tr>",
            "</w:tbl>",
            para(),
        )

    body = join(
        para("D: tblCellSpacing magnitude — same table, four spacings"),
        # The zero row is the reference: it fixes the cell width and the table
        # width that the other three are read against, on the same page and in
        # the same face, so no external measurement is needed.
        # Short cell tags on purpose: doubling the spacing narrows the cells,
        # and a label that no longer fits gets split across draw commands, which
        # is exactly what a test identifying a table by its cell text cannot
        # survive. The readable description lives in the heading above each
        # table, where a human looks anyway.
        table("Table 1 — no spacing", "T1", None),
        table("Table 2 — tblCellSpacing 200", "T2", 200),
        table("Table 3 — tblCellSpacing 400", "T3", 400),
        table("Table 4 — tblCellSpacing 400, row 800", "T4", 400, row_spacing=800),
        SECT,
    )
    write("issue-165-cellspacing-scale.docx", document(body))


# ── C. vertical inside/outside for floats ────────────────────────────────────
#
# Mirrored margins with ASYMMETRIC top and bottom (1in / 2in) so a vertical
# mirror is visible at all, and one anchor per page so each y is unambiguous.
# Six pages: {margin+align, insideMargin+offset, margin+outside} x {odd, even}.
#
#   aligns to region top (dxpdf today) → identical y on odd and even
#                                        (measured: (72,72) on all six)
#   mirrors vertically                 → y differs between odd and even
#
# That w:mirrorMargins mirrors left and right, not top and bottom, is the
# question — not an objection to the probe.
def build_floatv():
    def anchor(idx, rel_from, pos_body):
        return (
            "<w:r><w:drawing>"
            '<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/'
            'wordprocessingDrawing" distT="0" distB="0" distL="114300" distR="114300" '
            f'simplePos="0" relativeHeight="{idx}" behindDoc="0" locked="0" '
            'layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="{rel_from}">{pos_body}</wp:positionV>'
            '<wp:extent cx="457200" cy="457200"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="{idx}" name="Probe{idx}"/>'
            "<wp:cNvGraphicFramePr/>"
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="{idx}" name="dot.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip xmlns:r="{NS_R}" r:embed="rIdI"/>'
            "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            "</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
        )

    cases = [
        ("margin", "<wp:align>inside</wp:align>", "margin+align=inside"),
        ("insideMargin", "<wp:posOffset>0</wp:posOffset>", "insideMargin+offset=0"),
        ("margin", "<wp:align>outside</wp:align>", "margin+align=outside"),
    ]
    pages, idx = [], 1
    for rel_from, pos_body, label in cases:
        for parity in ("odd", "even"):
            pages.append(
                f'<w:p><w:r><w:t xml:space="preserve">{label} / {parity}</w:t></w:r>'
                f"{anchor(idx, rel_from, pos_body)}</w:p>"
            )
            idx += 1
    brk = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

    # Asymmetric top/bottom is what makes a vertical mirror detectable.
    sect = (
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="2880" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
    )
    write(
        "issue-165-floatv.docx",
        document(join(brk.join(pages), sect)),
        settings_body="<w:mirrorMargins/>",
        image=solid_png(64, 64, (200, 30, 30)),
    )


if __name__ == "__main__":
    build_vmerge()
    build_cellspacing()
    build_cellspacing_scale()
    build_floatv()
