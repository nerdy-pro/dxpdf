#!/usr/bin/env python3
"""Build the three §17.4.66 border-geometry **probes**.

    test-files/border-content-charge.docx
    test-files/border-outer-box.docx
    test-files/border-junction-colour.docx

None of them pins behaviour. Each asks a question ECMA-376 does not answer and
is authored so that the candidate readings predict a **different picture** — the
`issue-165-*` pattern. They answer nothing until someone renders them in Word
and reports what they see; until then the code says what it does and why, and
these say what would settle it.

The three exist because ECMA-376 specifies **no stroke geometry for table
borders at all**. [MS-OI29500] §17.4.66 adds a precedence order for *conflicting*
declarations on one edge and stops there — it never says where on its edge a
border sits, how much of a cell it takes, or what happens where two of them
cross. Every one of those is this engine's own convention, and the file that
holds the convention (`src/render/layout/table/borders.rs`) names each of these
three probes at the site that guesses.

---------------------------------------------------------------------------
1. border-content-charge.docx — how much of a shared border is inside a cell?
---------------------------------------------------------------------------

A collapsed border stands on an edge two cells share. How far does it push each
cell's text?

Three readings, and this engine currently holds **two of them at once**, which is
the whole reason for the probe. `resolve_table_cell_borders` charges the winner's
cell the full width and charges the facing cell nothing; `rasterize_border_grid`
paints the line straddling the edge, half in each. They disagree, and no
measurement has ever been taken.

The fixture is one table, two columns, **zero cell margins** — so nothing but the
border can inset the text — and four rows whose second cell declares `w:left` at
0.5, 3, 6 and 12 pt. What moves is the second column's glyph, and the reading is
by eye rather than by ruler because the 12pt row makes the three cases look
completely different:

    charged nothing   the glyph starts *on* the grid line, so the border's
                      inner half is painted over it — a 12pt bar through the X
    charged half      the glyph starts flush against the border's inner edge
    charged in full   a 6pt gap between the border and the glyph

The first is what this engine does today.

---------------------------------------------------------------------------
2. border-outer-box.docx — does the table's box contain its outer borders?
---------------------------------------------------------------------------

The table's own four edges are shared with nothing, so there is no second cell to
halve a border with. Two readings: the border goes *inside* the box, or the box
grows to hold a border that straddles its edge.

It matters beyond the border. `w:tblInd` (§17.4.50) measures to the table's left
edge, and the two readings put that edge half a border apart; `TableSlice::size`
is what the stacker fits to a page, so under the straddling reading it does not
contain what the slice draws.

This engine took the inside reading this session, and the evidence for it is
one-sided: straddling put 0.2pt of a full-width table off the *paper*, which
§17.4.63's auto-width guard is drawn at (`tests/table_auto_width.rs`). That rules
straddling-without-widening out. It does not rule out straddling with a box wide
enough to hold it, and nothing here distinguishes those two.

The fixture is a paragraph of text — whose left edge is the page margin, and the
reference — followed by two tables at `w:tblInd="0"` differing **only** in outer
border weight, 0.5pt and 12pt. The reading:

    inside      both tables' left borders begin at the paragraph's left edge,
                and both tables end at the same right edge
    straddling  the thick table's border begins 6pt into the margin, and the
                table is 12pt wider than the thin one

---------------------------------------------------------------------------
3. border-junction-colour.docx — who owns the square where two borders cross?
---------------------------------------------------------------------------

A junction is not a conflict: all four segments meeting at it are correct, and
all four want the square. §17.4.66's precedence list is about two declarations on
*one* edge and says nothing about this.

**MEASURED.** This one has been rendered in Word, and the answer refuted what
the engine did: it gave the square to the `border_precedence` winner among the
segments reaching it, breaking a tie toward the vertical.

Tables 1 and 2 are the discrimination. Both have 12pt `insideV` and `insideH` of
equal weight and style, so precedence falls through to §17.4.66's colour rule
(darker wins), and the two tables **swap which axis is darker**:

    table 1   vertical black, horizontal light grey
    table 2   vertical light grey, horizontal black

    precedence          table 1's crossing is black, table 2's is black
                        (i.e. the *darker* line, whichever axis it is on)
    vertical always     table 1 black, table 2 light grey
    horizontal always   table 1 light grey, table 2 black   <- Word

Word draws pale then dark: **the horizontal takes the square**, whichever axis
carries the darker line. Both of the engine's rules died in one render. (What it
does when the two differ in *weight* is table 5, below, and is still open.)

Table 3 was the known limit rather than a discrimination, and Word settled it
too. Both axes are 12pt `double`, so each is two 4pt rules with a 4pt gap; Word
draws the crossing as the 2 x 2 lattice of ink with **both gaps running through
it**, reported as "the borders are negative space, so it looks like every cell
has its own border" — which is what separated per-cell rectangles look like, and
what a continuous double-ruled grid never does. The engine drew two rungs.

So a crossing is the **product** of its two axes' §17.18.2 rules, coloured by the
horizontal, and a `single` contributing one full band is that same rule rather
than a case of its own.

Tables 4 and 5 are what those two answers newly put at stake, and both are
**open**.

Table 4 asks about the product: a 12pt `single` horizontal crossing a 12pt
`double` vertical. The product punches the double's 4pt gap through the solid
line; the rival reading is that a solid line runs through unbroken and only the
double is interrupted.

Table 5 asks how far "the horizontal wins" reaches. Tables 1 and 2 tie the two
axes on weight *on purpose* — that is what makes them a test of colour — so they
say nothing about a crossing whose two lines differ in weight. A 12pt black
vertical against a 3pt pale horizontal separates the two readings:

    horizontal always   the thick black vertical is interrupted at every
                        crossing by a 12pt wide, 3pt tall pale band
    heavier wins        the crossing is black and the pale horizontal is the
                        one interrupted

The engine takes the first, because the rule that explains tables 1 and 2 most
simply is a paint order — horizontals over verticals — and a paint order has no
weight term in it. That is inference, not measurement, and it is visible: it is
what changed 27 crossings in `grid-gap-borders.docx`, whose 3pt red and green
verticals now carry a 1pt grey square wherever a row boundary crosses them.

---------------------------------------------------------------------------

All three carry the same styles part as `grid-gap-borders.docx`, for the reason
its script sets out at length: §17.7.2 makes an absent `w:docDefaults`
application-defined, so a package that names no face or size is read one way here
and another by Word. A fixture meant to be *measured against Word* has to state
them.

Run `scripts/verify_docx.py` on the results before committing. The build is
deterministic and needs no third-party packages.

    scripts/make_border_geometry_probes.py
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
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

#: The `xmlns` is the relationship-*part* URI. The relationship-*type* URI here
#: instead yields a package Word rejects and this parser accepts — see AGENTS.md.
DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

#: §17.7.5 — see the module docstring, and `make_bidi_visual_fixture.py` for why
#: each value is stated rather than defaulted.
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

SECT_PR = """<w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>"""

#: §17.18.2 `ST_EighthPointMeasure` caps `w:sz` at 96, which is 12pt. Every probe
#: here uses that cap where it wants a difference visible without a ruler.
MAX_SZ = 96


def heading(text: str) -> str:
    # `<` and `&` would become markup, which `verify_docx.py` cannot catch — it
    # checks that the package is a sound OPC container, not that the parts are
    # schema-valid. An earlier probe lost an afternoon to a heading that named an
    # element literally, so the guard is here rather than in a reviewer's memory.
    if "<" in text or "&" in text:
        raise ValueError(f"heading text would be parsed as markup: {text!r}")
    return f'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def para(text: str) -> str:
    if "<" in text or "&" in text:
        raise ValueError(f"paragraph text would be parsed as markup: {text!r}")
    return f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def document(body: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        f"  <w:body>\n    {body}\n    {SECT_PR}\n  </w:body>\n"
        "</w:document>\n"
    )


#: Zero on all four sides. Without this the inset is `max(border, margin)` and a
#: thin border is masked by the margin, which is exactly the measurement these
#: probes are trying to take.
NO_CELL_MARGIN = (
    "<w:tblCellMar>"
    '<w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
    '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
    "</w:tblCellMar>"
)


def write(name: str, body: str) -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    target = OUT / name
    # Fixed timestamps so regenerating an unchanged fixture produces identical
    # bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for part, data in (
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("word/_rels/document.xml.rels", DOC_RELS),
            ("word/styles.xml", STYLES),
            ("word/document.xml", document(body)),
        ):
            info = zipfile.ZipInfo(part, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


# ── 1. how much of a shared border is inside a cell? ────────────────────────


def charge_probe() -> str:
    """Four rows, identical but for the weight of the vertical they share.

    The measured cell is the **second** one, because only an interior edge is
    shared — the table's own left is a different question and probe 2's.

    The border is declared on the second cell's `w:left` rather than as a
    table-level `insideV` so that the row can vary it, and because a per-cell
    declaration faces an undeclared neighbour: §17.4.66 resolves it to itself and
    the reader does not have to reason about which of two lines won.
    """
    rows = []
    for sz in (4, 24, 48, MAX_SZ):
        pt = sz / 8
        rows.append(
            "<w:tr>"
            # The first cell names the weight, so the render is self-describing.
            "<w:tc>"
            '<w:tcPr><w:tcW w:w="2000" w:type="dxa"/>'
            '<w:shd w:val="clear" w:color="auto" w:fill="EDEDED"/></w:tcPr>'
            f'<w:p><w:r><w:t xml:space="preserve">{pt:g}pt border</w:t></w:r></w:p>'
            "</w:tc>"
            # The measured cell. Its glyph is a single X so its left side is a
            # vertical stroke, which is what makes "covered / flush / clear of
            # the border" readable without a ruler.
            "<w:tc>"
            '<w:tcPr><w:tcW w:w="4000" w:type="dxa"/>'
            "<w:tcBorders>"
            f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="C00000"/>'
            "</w:tcBorders>"
            '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/></w:tcPr>'
            '<w:p><w:r><w:t xml:space="preserve">X</w:t></w:r></w:p>'
            "</w:tc>"
            "</w:tr>"
        )
    table = (
        "<w:tbl><w:tblPr>"
        '<w:tblW w:w="6000" w:type="dxa"/>'
        "<w:tblBorders>"
        '<w:top w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        '<w:insideH w:val="single" w:sz="8" w:space="0" w:color="808080"/>'
        "</w:tblBorders>"
        '<w:tblLayout w:type="fixed"/>'
        f"{NO_CELL_MARGIN}"
        "</w:tblPr>"
        '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid>'
        + "".join(rows)
        + "</w:tbl>"
    )
    return "\n    ".join(
        [
            heading(
                "PROBE - how much of a shared border is inside a cell? "
                "Nothing is asserted here; this asks a question ECMA-376 does not answer."
            ),
            para(
                "Cell margins are zero, so the red line is the only thing that can "
                "move the X. Read the bottom row, where the border is 12pt:"
            ),
            para(
                "  - the red bar is painted over the X          -> a cell is charged nothing for a border its neighbour won"
            ),
            para(
                "  - the X sits flush against the red bar       -> half the border is inside each cell"
            ),
            para(
                "  - a 6pt gap between the red bar and the X    -> the whole border is inside this cell"
            ),
            table,
            para(
                "The same three readings predict 0, 3 and 6pt on the third row, and "
                "0, 1.5 and 3pt on the second. A slope across the four rows is what "
                "separates a real answer from a one-off."
            ),
        ]
    )


# ── 2. does the table's box contain its outer borders? ──────────────────────


def outer_box_probe() -> str:
    """Two tables differing only in outer border weight, both at `tblInd` 0.

    The reference is the paragraph above them: its first glyph sits at the page
    margin, which is where `w:tblInd="0"` puts the table's left edge. Whether the
    thick table's *border* starts there or 6pt to the left of it is the question.
    """

    def table(sz: int) -> str:
        return (
            "<w:tbl><w:tblPr>"
            '<w:tblW w:w="6000" w:type="dxa"/>'
            '<w:tblInd w:w="0" w:type="dxa"/>'
            "<w:tblBorders>"
            f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="C00000"/>'
            f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="C00000"/>'
            f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="C00000"/>'
            f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="C00000"/>'
            '<w:insideV w:val="single" w:sz="8" w:space="0" w:color="0070C0"/>'
            "</w:tblBorders>"
            '<w:tblLayout w:type="fixed"/>'
            f"{NO_CELL_MARGIN}"
            "</w:tblPr>"
            '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
            "<w:tr>"
            '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:r><w:t xml:space="preserve">{sz / 8:g}pt outer border</w:t></w:r></w:p></w:tc>'
            '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>'
            '<w:p><w:r><w:t xml:space="preserve">right column</w:t></w:r></w:p></w:tc>'
            "</w:tr>"
            "</w:tbl>"
        )

    return "\n    ".join(
        [
            heading(
                "PROBE - does a table's box contain its outer borders, or straddle them?"
            ),
            para(
                "This paragraph starts at the page margin, and so does a table with "
                "w:tblInd of zero. Both tables below declare the same width and the "
                "same indent; only the outer border weight differs."
            ),
            para(
                "  - both red frames begin under this paragraph's first letter, and both "
                "tables end at the same right edge  -> the box contains its borders"
            ),
            para(
                "  - the thick frame begins 6pt into the left margin and the table is 12pt "
                "wider than the thin one            -> the box straddles them"
            ),
            table(4),
            para(""),
            table(MAX_SZ),
            para(
                "The blue 1pt line inside each table is the control: it is on a shared "
                "edge, so it must straddle under either reading, and it must not move "
                "between the two tables."
            ),
        ]
    )


# ── 3. who owns the square where two borders cross? ─────────────────────────


def junction_probe() -> str:
    """Three tables. The first two swap which axis carries the darker line; the
    third asks what a `double` crossing a `double` should look like."""

    def table(v_colour: str, h_colour: str, style: str, label: str) -> str:
        cells = "".join(
            '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:r><w:t xml:space="preserve">{c}</w:t></w:r></w:p></w:tc>'
            for c in label
        )
        row = f"<w:tr>{cells}</w:tr>"
        return (
            "<w:tbl><w:tblPr>"
            '<w:tblW w:w="6000" w:type="dxa"/>'
            "<w:tblBorders>"
            f'<w:insideV w:val="{style}" w:sz="{MAX_SZ}" w:space="0" w:color="{v_colour}"/>'
            f'<w:insideH w:val="{style}" w:sz="{MAX_SZ}" w:space="0" w:color="{h_colour}"/>'
            "</w:tblBorders>"
            '<w:tblLayout w:type="fixed"/>'
            "</w:tblPr>"
            '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
            + row * 3
            + "</w:tbl>"
        )

    def mixed_table(v_colour: str, h_colour: str, label: str) -> str:
        """`insideV` double against `insideH` single, both 12pt."""
        cells = "".join(
            '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:r><w:t xml:space="preserve">{c}</w:t></w:r></w:p></w:tc>'
            for c in label
        )
        row = f"<w:tr>{cells}</w:tr>"
        return (
            "<w:tbl><w:tblPr>"
            '<w:tblW w:w="6000" w:type="dxa"/>'
            "<w:tblBorders>"
            f'<w:insideV w:val="double" w:sz="{MAX_SZ}" w:space="0" w:color="{v_colour}"/>'
            f'<w:insideH w:val="single" w:sz="{MAX_SZ}" w:space="0" w:color="{h_colour}"/>'
            "</w:tblBorders>"
            '<w:tblLayout w:type="fixed"/>'
            "</w:tblPr>"
            '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
            + row * 3
            + "</w:tbl>"
        )

    def uneven_table(v_colour: str, h_colour: str, label: str) -> str:
        """A 12pt vertical against a 3pt horizontal, both `single`."""
        cells = "".join(
            '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:r><w:t xml:space="preserve">{c}</w:t></w:r></w:p></w:tc>'
            for c in label
        )
        row = f"<w:tr>{cells}</w:tr>"
        return (
            "<w:tbl><w:tblPr>"
            '<w:tblW w:w="6000" w:type="dxa"/>'
            "<w:tblBorders>"
            f'<w:insideV w:val="single" w:sz="{MAX_SZ}" w:space="0" w:color="{v_colour}"/>'
            f'<w:insideH w:val="single" w:sz="24" w:space="0" w:color="{h_colour}"/>'
            "</w:tblBorders>"
            '<w:tblLayout w:type="fixed"/>'
            "</w:tblPr>"
            '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>'
            + row * 3
            + "</w:tbl>"
        )

    black, grey = "000000", "BFBFBF"
    return "\n    ".join(
        [
            heading(
                "PROBE - who owns the square where a vertical border crosses a horizontal one?"
            ),
            para(
                "Both lines in tables 1 and 2 are 12pt single, so they tie on weight and "
                "on style; only their colour differs, and the two tables swap which axis "
                "carries the darker one. Read the colour of the four crossing squares in each."
            ),
            para(
                "  - dark in table 1 and dark in table 2   -> the square follows the "
                "MS-OI29500 17.4.66 order, darker wins"
            ),
            para(
                "  - dark in table 1, pale in table 2      -> the vertical always wins"
            ),
            para(
                "  - pale in table 1, dark in table 2      -> the horizontal always wins"
            ),
            para(
                "MEASURED: Word draws pale in table 1 and dark in table 2. The "
                "horizontal takes the square, whichever axis carries the darker line."
            ),
            heading("Table 1 - vertical dark, horizontal pale."),
            table(black, grey, "single", "abc"),
            para(""),
            heading("Table 2 - vertical pale, horizontal dark."),
            table(grey, black, "single", "def"),
            para(""),
            heading("Table 3 - both axes 12pt double. MEASURED."),
            para(
                "A double is two rules with a gap, so the crossing is ideally a "
                "two-by-two lattice of ink with both gaps running through it, and that "
                "is what Word draws: the borders read as negative space, so every cell "
                "looks separately enclosed. The engine drew the square along one axis "
                "instead, which gave two rungs."
            ),
            table(black, black, "double", "ghi"),
            para(""),
            heading(
                "Table 4 - 12pt single horizontal crossing a 12pt double vertical. OPEN."
            ),
            para(
                "Tables 1 to 3 settle a crossing into the product of its two axes: each "
                "one divides the square across its own short side, and the horizontal "
                "colours it. This is the shape that product makes a claim about with no "
                "measurement behind it. Read whether the solid horizontal runs through "
                "the crossing unbroken, or whether the vertical's 4pt gap is punched "
                "through it, leaving the horizontal in two pieces at every crossing."
            ),
            mixed_table(black, black, "efg"),
            para(""),
            heading(
                "Table 5 - a 12pt black vertical crossing a 3pt pale horizontal. OPEN."
            ),
            para(
                "Tables 1 and 2 tie the two axes on weight, which is what makes them a "
                "test of colour and also what keeps them silent about weight. Read the "
                "crossings: a pale 3pt band interrupting the thick black vertical means "
                "the horizontal wins whatever its weight, and a black crossing that "
                "interrupts the pale line instead means the heavier line wins."
            ),
            uneven_table(black, grey, "hij"),
        ]
    )


def main() -> None:
    write("border-content-charge.docx", charge_probe())
    write("border-outer-box.docx", outer_box_probe())
    write("border-junction-colour.docx", junction_probe())


if __name__ == "__main__":
    main()
