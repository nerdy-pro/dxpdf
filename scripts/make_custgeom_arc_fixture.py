#!/usr/bin/env python3
"""Build test-files/custgeom-arc.docx — §20.1.9.3 arcTo in a custom geometry.

Four anchored `wps:wsp` shapes, each a `<a:custGeom>` whose path bends through
`<a:arcTo>`. No corpus document carries a custGeom at all, so this fixture is
the only end-to-end witness for the arc path verb. The four shapes are chosen
so that every known way to misread §20.1.9.3 renders visibly differently:

1. **Quarter pie** — moveTo the right point of a circle, arcTo 0°+90°, two
   straight edges back through the center. Reading the pen as the ellipse
   *center* (instead of a point ON the ellipse at `stAng`) shifts the whole
   wedge one radius right and adds a spurious pen-to-arc-start line.
2. **Ellipse chord** — wR≠hR with stAng=45°, swAng=90°. The spec's angles are
   ray angles measured at the center; Skia/AWT arc angles are parametric.
   Skipping the atan2(wR·sinθ, hR·cosθ) skew moves both endpoints.
3. **Full disc** — swAng=21600000 (360°). Arc APIs treat sweeps modulo 360°,
   so an unsplit full swing draws nothing at all.
4. **Crescent** — two chained arcs. The second arc's center derives from the
   pen position the first arc ended on, so it doubles any endpoint error.

Reference render: LibreOffice (whose ARCANGLETO implements the same
construction as Apache POI's ArcToCommand — pen on the ellipse at stAng,
ray→parametric angle skew).

Deterministic: no timestamps, fixed ZIP metadata, so re-running produces a
byte-identical archive. Regenerate rather than hand-edit.

    python3 scripts/make_custgeom_arc_fixture.py
"""

import pathlib
import zipfile

OUT = pathlib.Path(__file__).resolve().parent.parent / "test-files" / "custgeom-arc.docx"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

EMU_PER_INCH = 914400


def shape(sid, x_emu, y_emu, cx, cy, path_w, path_h, commands, fill="4472C4"):
    """One page-anchored wps shape with a custGeom path."""
    return f"""<w:r><w:drawing>
<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{sid}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="page"><wp:posOffset>{x_emu}</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>{y_emu}</wp:posOffset></wp:positionV>
<wp:extent cx="{cx}" cy="{cy}"/>
<wp:wrapNone/>
<wp:docPr id="{sid}" name="arc{sid}"/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp>
<wps:cNvSpPr/>
<wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:custGeom>
<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>
<a:rect l="0" t="0" r="{path_w}" b="{path_h}"/>
<a:pathLst><a:path w="{path_w}" h="{path_h}">
{commands}
</a:path></a:pathLst>
</a:custGeom>
<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>
<a:ln w="12700"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:ln>
</wps:spPr>
</wps:wsp>
</a:graphicData></a:graphic>
</wp:anchor>
</w:drawing></w:r>"""


INCH = EMU_PER_INCH
SHAPES = [
    # 1. Quarter pie: pen on the circle at 0°, +90° swing, edges back
    #    through the center (50,50).
    shape(
        1, INCH, INCH, INCH, INCH, 100, 100,
        '<a:moveTo><a:pt x="100" y="50"/></a:moveTo>'
        '<a:arcTo wR="50" hR="50" stAng="0" swAng="5400000"/>'
        '<a:lnTo><a:pt x="50" y="50"/></a:lnTo>'
        '<a:close/>',
    ),
    # 2. Ellipse chord: wR=8000, hR=4000, center (10000,5000). The pen sits
    #    on the ellipse at ray angle 45°: 45° hits it at
    #    center + (3578, 3578); +90° swing exits at ray 135°.
    shape(
        2, 3 * INCH, INCH, 2 * INCH, INCH, 20000, 10000,
        '<a:moveTo><a:pt x="13578" y="8578"/></a:moveTo>'
        '<a:arcTo wR="8000" hR="4000" stAng="2700000" swAng="5400000"/>'
        '<a:close/>',
        fill="ED7D31",
    ),
    # 3. Full disc: a single 360° swing from the circle's right point.
    shape(
        3, INCH, 3 * INCH, INCH, INCH, 100, 100,
        '<a:moveTo><a:pt x="100" y="50"/></a:moveTo>'
        '<a:arcTo wR="50" hR="50" stAng="0" swAng="21600000"/>'
        '<a:close/>',
        fill="70AD47",
    ),
    # 4. Crescent: outer right semicircle (r=50) from the top point down,
    #    then a narrower half-ellipse (wR=30) swinging back up to the start.
    shape(
        4, 3 * INCH, 3 * INCH, INCH, INCH, 100, 100,
        '<a:moveTo><a:pt x="50" y="0"/></a:moveTo>'
        '<a:arcTo wR="50" hR="50" stAng="16200000" swAng="10800000"/>'
        '<a:arcTo wR="30" hR="50" stAng="5400000" swAng="-10800000"/>'
        '<a:close/>',
        fill="FFC000",
    ),
]

DOCUMENT = (
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<w:body>
<w:p>"""
    + "".join(SHAPES)
    + """<w:r><w:t xml:space="preserve">custGeom arcTo fixture: pie, ellipse chord, full disc, crescent</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body>
</w:document>"""
)

PARTS = [
    ("[Content_Types].xml", CONTENT_TYPES),
    ("_rels/.rels", RELS),
    ("word/document.xml", DOCUMENT),
]


def main() -> None:
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, body in PARTS:
            # Fixed date_time so the archive is reproducible.
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            zf.writestr(info, body)
    print(f"wrote {OUT} ({OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
