#!/usr/bin/env python3
"""Build the two issue #157 probes — documents that *ask* a question ECMA-376
does not answer, rather than fixtures that pin an answer.

Neither changes behaviour on its own. Each is authored so that the candidate
readings predict a **different measurement off the rendered page**, so one Word
render settles it. That is the pattern `test-files/issue-165-*.docx` used for
vMerge overflow and cell spacing, and the reason those are now decided.

---------------------------------------------------------------------------
1. issue-157-tblprex-bidi.docx  — §17.4.1 `w:bidiVisual` on one row
---------------------------------------------------------------------------

`w:tblPrEx` may carry `bidiVisual`, flipping a single row against a table whose
other rows run left to right. The engine parses it (`TableRowPropertyExceptions::bidi_visual`);
it is now acted on — see `RowBidiOverride` and the `bidi_override` checks in
`table/borders.rs`, `table/measure.rs`, and the row-swap in `build/table.rs`.

Two readings of the width question were live before a render:

  A. the row's cells keep their **declared widths** and swap places, so the
     mirrored row's slot boundaries do not line up with the rows around it;
  B. the row's cells take the **widths of the slots they land in**, so every
     boundary stays aligned and only the contents move.

The grid is 1000/2000/3000 twips — deliberately unequal, because with three
equal columns the two readings render identically and the document would prove
nothing. Row 2 is the flipped one; rows 1 and 3 are the same cells unflipped.

**How to read a Word render.** Look at cell `B` in row 2. Under (A) it is 100pt
wide, as it is in rows 1 and 3. Under (B) it is 100pt only by coincidence of
being the middle column — so look at `A` instead: 50pt under (A), 150pt under
(B). The cells are labelled and separately shaded so this is a ruler
measurement, not a judgement.

**Measured 2026-09-05** (pixel-counted off a fresh render, calibrated against
the page's own margins): cell `A` comes out 50pt — reading (A). The same render
answered two questions this docstring never posed: the flipped row sits with
its own right edge on the *page's* content width, as if it were a mini
right-to-left table of its own, and it paints **after** the row that would
otherwise follow it rather than between its neighbours (row 2 of 3 renders
third). Both are implemented alongside the width finding and are measured from
this one arrangement only — see `RowBidiOverride`'s doc for what remains open.

**`with_styles=True`, added the same day**, for the reason `write()` gives: the
first Word render used a package with no `styles.xml`, and Word's own template
defaults (Calibri 11pt, non-zero paragraph spacing) made every row's height
disagree with this engine's ECMA-default read (Times New Roman 10pt, zero
spacing) for a reason that had nothing to do with `bidiVisual` at all — the row
height mismatch was a fixture gap, not a rendering defect. Stating the defaults
explicitly, the same fix `bidi-visual-table.docx` and `empty_row_probe` already
use, does not change any width, position or paint-order finding above — none of
them depend on the font — only the row heights, which now read the same values
in both renderers instead of two different application defaults.

---------------------------------------------------------------------------
2. issue-157-empty-row-edge.docx — §17.4.66 across a row of no height
---------------------------------------------------------------------------

§17.4.66 resolves a shared horizontal edge between two rows by picking one
cell's border and clearing the facing one. A row of no height puts its top and
its bottom boundary at the same y, so there are two edges where a table without
it has one, and nothing in the section says whether they resolve into each other
or both stand.

**This probe used to ask that with a cell-less `<w:tr/>`, and it cannot.**
Measured 2026-08-19: **Word refuses to open any document containing one.**
Isolated against two variants differing in exactly one thing each — the same
package with the row deleted opens, and the same row in a fuller package
(`document.xml.rels` plus `styles.xml`, one relationship, no dangling target)
does not. `CT_Row` makes the cell group `minOccurs="0"`, so `<w:tr/>` is
schema-valid and Word's reader is simply stricter than the schema, the same
class of rejection the three `issue-165-*` fixtures hit over a `.rels`
namespace. `verify_docx.py` cannot catch it either: the package is sound, and it
is the content Word declines.

So the question is asked in the two spellings Word does accept, both of which a
real producer can emit and both of which put two boundaries within a hair of
each other:

  Table 2 — a row of `hRule="exact" w:val="0"`, holding one empty cell. Its two
            boundaries are at the same y exactly.
  Table 3 — a row of `hRule="exact" w:val="40"`, which is 2pt: **shorter than
            its own two 3pt borders**, so the rules must overlap whatever else
            happens.

Table 4 asks the question the first three cannot, and it is the one that matters
most: **what does `w:trHeight` measure?** Its middle row declares 40pt,
comfortably more than its own 6pt of borders, so no floor and no collapse is
involved and the two readings differ by a clean 6pt:

  * the yellow row measures 40pt <- WORD -> the declared height is the row's
                                   *content box*, its rules standing outside it;
  * the yellow row measures 34pt        -> it would include them.

MEASURED 2026-08-19: Word draws **40pt**. Table 5 declares the same as `atLeast`
and draws 40pt too, so one rule covers both — which matters more than its
position suggests, since `atLeast` is what an omitted `hRule` becomes at the
parse seam and so carries nearly every `<w:trHeight>` written.

That refuted a reading this engine briefly shipped: that the declared height
*contains* the rules, by analogy with the other axis, where
`border-content-charge.docx` settled that a shared border is charged half to each
cell's *content box*. The analogy is wrong here. What made it plausible was
tables 2 and 3 — rows of 0 and 2pt against 6pt of rule, where a 2pt cell and a
hairline are not distinguishable by eye. Table 4 is twenty times that size.

**Table 2 is where the real defect was.** It declares `hRule="exact"` with
`w:val="0"`, and Word draws a full row of cell — about one empty line. Zero is
not a height, it is the marker for *unconstrained*, which [MS-OI29500] §2.4.77(c)
corroborates from the producing side: Word requires `val="0"` whenever
`hRule="auto"`. A literal reading draws a flat row and loses the cell, which is
what this engine did until that render.

Read the black band at the middle boundary against table 1's, which is one 3pt
rule and the control for the other two:

  * 3pt in tables 2 and 3 -> Word resolves the coincident edges into one, and
                             the engine should collapse them too;
  * 6pt in table 2        -> both rules stand, and today's behaviour is right;
  * a white gap in table 3 -> Word honours the 2pt height between the rules
                             rather than letting them meet.

`w:sz="24"` (3pt) rather than a hairline, so the difference is visible at 100%
zoom instead of needing a loupe. The fixture carries its own `styles.xml` for
the reason `bidi-visual-table.docx` does: §17.7.2 leaves an absent
`w:docDefaults` application-defined, so a package declaring no face or spacing
is read one way here and another by Word, and the two renders have to be
comparable to be compared.

**What is not asked any more, and will not be.** Whether a *cell-less* row
separates its neighbours has no fidelity answer: a document containing one is
one Word itself calls corrupt, so dxpdf's tolerance there is a robustness
decision rather than a match. `table::borders` (which merges the coincident
boundaries) and `table::measure` (which pins the zero height) both say so at the
point where they decide it. Do not wait on a render that cannot be produced.

---------------------------------------------------------------------------

Run `scripts/verify_docx.py` on the results before committing. Regenerate and
commit if the content changes; the build is deterministic and needs no
third-party packages.

    scripts/make_issue157_probes.py
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

SECT_PR = """<w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>"""


def heading(text: str) -> str:
    # `<` and `&` in a `<w:t>` would become markup, which `verify_docx.py`
    # cannot catch — it checks that the package is a sound OPC container, not
    # that the parts are schema-valid. This probe lost an afternoon to a heading
    # that named an element literally, so the guard is here rather than in a
    # reviewer's memory.
    if "<" in text or "&" in text:
        raise ValueError(f"heading text would be parsed as markup: {text!r}")
    return f'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def cell(label: str, fill: str, extra: str = "") -> str:
    return (
        "<w:tc>"
        f'<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>{extra}</w:tcPr>'
        f'<w:p><w:r><w:t xml:space="preserve">{label}</w:t></w:r></w:p>'
        "</w:tc>"
    )


def document(body: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body}
    {SECT_PR}
  </w:body>
</w:document>
"""


# ── probe 1: w:tblPrEx/w:bidiVisual ─────────────────────────────────────────

#: Unequal on purpose — see the module docstring.
GRID3 = '<w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/>'


def tblprex_probe() -> str:
    plain = "<w:tr>" + cell("A", "F8CBAD") + cell("B", "C6E0B4") + cell("C", "BDD7EE") + "</w:tr>"
    flipped = (
        "<w:tr><w:tblPrEx><w:bidiVisual/></w:tblPrEx>"
        + cell("A", "F8CBAD")
        + cell("B", "C6E0B4")
        + cell("C", "BDD7EE")
        + "</w:tr>"
    )
    table = (
        "<w:tbl><w:tblPr>"
        '<w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        '<w:left w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        '<w:right w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        '<w:insideH w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        '<w:insideV w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        "</w:tblBorders></w:tblPr>"
        f"<w:tblGrid>{GRID3}</w:tblGrid>{plain}{flipped}{plain}</w:tbl>"
    )
    return document(
        heading("Row 2 carries w:tblPrEx/w:bidiVisual. Rows 1 and 3 do not.")
        + table
        + heading("Measure cell A in row 2: 50pt keeps its declared width, 150pt takes the slot's.")
    )


# ── probe 2: §17.4.66 across an empty <w:tr/> ───────────────────────────────

GRID1 = '<w:gridCol w:w="4000"/>'


def empty_row_probe() -> str:
    #: An empty cell rather than no cell: `<w:tr/>` is what Word will not open.
    EMPTY_CELL = '<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:tcPr><w:p/></w:tc>'

    def table(middle: str = "") -> str:
        rows = "<w:tr>" + cell("upper", "D9EAD3") + "</w:tr>"
        rows += middle
        rows += "<w:tr>" + cell("lower", "CFE2F3") + "</w:tr>"
        return (
            "<w:tbl><w:tblPr>"
            '<w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
            "<w:tblBorders>"
            '<w:top w:val="single" w:sz="24" w:space="0" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="24" w:space="0" w:color="000000"/>'
            '<w:left w:val="nil"/><w:right w:val="nil"/>'
            '<w:insideH w:val="single" w:sz="24" w:space="0" w:color="000000"/>'
            '<w:insideV w:val="nil"/>'
            "</w:tblBorders></w:tblPr>"
            f"<w:tblGrid>{GRID1}</w:tblGrid>{rows}</w:tbl>"
        )

    def short_row(twips: int, rule: str = "exact") -> str:
        return (
            "<w:tr><w:trPr>"
            f'<w:trHeight w:val="{twips}" w:hRule="{rule}"/>'
            "</w:trPr>" + EMPTY_CELL + "</w:tr>"
        )

    return document(
        heading("1. Control: two rows, one shared edge between them.")
        + table()
        + "<w:p/>"
        + heading("2. The same, with a row of exactly zero height between them.")
        + table(short_row(0))
        + "<w:p/>"
        + heading("3. The same, with a 2pt row between them - shorter than its own borders.")
        + table(short_row(40))
        + "<w:p/>"
        + heading("4. The same, with a 40pt row - taller than its borders, so nothing is floored.")
        + table(short_row(800))
        + "<w:p/>"
        + heading("5. The same 40pt, declared atLeast rather than exact.")
        + table(short_row(800, rule="atLeast"))
        + heading("Measure the yellow rows of tables 4 and 5: 40pt means trHeight is the content box, 34 means it includes the borders.")
    )


STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"
                  w:eastAsia="Times New Roman" w:cs="Times New Roman"/>
        <w:kern w:val="2"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>
"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

STYLES_OVERRIDE = (
    '<Override PartName="/word/styles.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
)


def write(name: str, body: str, with_styles: bool = False) -> None:
    """Write one probe. `with_styles` adds a `styles.xml` stating the defaults.

    §17.7.2 leaves an absent `w:docDefaults` application-defined, so a package
    declaring no face, size or spacing is read one way by this engine (ECMA's
    Times New Roman 10pt, zero spacing) and another by Word (its template's
    Calibri 11pt with `after=160`, `line=259`). A fixture meant to be *measured
    against a Word render* has to state them or the two renders are not
    comparable — the same reason `bidi-visual-table.docx` carries one.
    """
    target = OUT / name
    parts = [
        (
            "[Content_Types].xml",
            CONTENT_TYPES.replace("</Types>", STYLES_OVERRIDE + "</Types>")
            if with_styles
            else CONTENT_TYPES,
        ),
        ("_rels/.rels", ROOT_RELS),
        ("word/document.xml", body),
    ]
    if with_styles:
        parts.append(("word/_rels/document.xml.rels", DOC_RELS))
        parts.append(("word/styles.xml", STYLES))
    # Fixed timestamps so regenerating an unchanged fixture produces identical
    # bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for part, data in parts:
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    write("issue-157-tblprex-bidi.docx", tblprex_probe(), with_styles=True)
    write("issue-157-empty-row-edge.docx", empty_row_probe(), with_styles=True)


if __name__ == "__main__":
    main()
