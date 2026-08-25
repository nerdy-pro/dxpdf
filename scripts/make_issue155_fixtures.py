#!/usr/bin/env python3
"""Build the issue #155 fixtures: SmartArt and charts.

``smartart.docx``
    One inline diagram wired the way Word wires it: ``dgm:relIds`` naming
    four parts through the document's rels, the *data* part carrying the
    ``dsp:dataModelExt`` extension whose ``@relId`` — resolved against the
    **document's** rels, the MS-ODRAWXML quirk — names the pre-laid-out
    ``dsp:`` drawing part. The drawing holds three ``roundRect`` process
    nodes (accent1/2/3 scheme fills, white 14pt labels One/Two/Three) and a
    ``rightArrow`` between the first two, all in absolute EMU on the
    drawing's 5486400×914400 canvas. A theme part carries the stock Office
    palette, so the scheme fills resolve to the familiar accent RGBs and the
    tests can pin them deterministically.

``charts.docx``
    Three inline drawings, one chart part each: a clustered column chart
    (two series × three categories, explicit series names, value axis with
    gridlines, bottom legend, a rich title), a pie (one series, four
    points), and a line chart (one series, five points, circle markers).
    Values live in the caches, as every real producer writes them.

Both are plain OPC packages verified by ``scripts/verify_docx.py``.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files"

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
NS_DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def write_docx(path: Path, parts: dict[str, str]) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    path.write_bytes(buf.getvalue())
    print(f"wrote {path.relative_to(ROOT)} ({path.stat().st_size} bytes)")


def content_types(overrides: dict[str, str]) -> str:
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    ]
    for part, ct in overrides.items():
        lines.append(f'<Override PartName="{part}" ContentType="{ct}"/>')
    lines.append("</Types>")
    return "\n".join(lines)


THEME = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2>
<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
</a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>
"""

THEME_CT = "application/vnd.openxmlformats-officedocument.theme+xml"
THEME_REL = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

SECT = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
    ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)


def document(body: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f' xmlns:wp="{NS_WP}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">'
        f"<w:body>{body}{SECT}</w:body></w:document>"
    )


def inline_drawing(cx: int, cy: int, graphic_data: str, name: str) -> str:
    return (
        "<w:p><w:r><w:drawing>"
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:docPr id="1" name="{name}"/>'
        f'<a:graphic><a:graphicData uri="{graphic_data.split("|")[0]}">'
        f'{graphic_data.split("|", 1)[1]}'
        "</a:graphicData></a:graphic></wp:inline>"
        "</w:drawing></w:r></w:p>"
    )


# ── SmartArt ───────────────────────────────────────────────────────────────


def dsp_shape(prst: str, x: int, y: int, cx: int, cy: int, accent: str, label: str | None) -> str:
    tx = ""
    if label is not None:
        tx = (
            "<dsp:txBody><a:bodyPr/>"
            '<a:p><a:pPr algn="ctr"/>'
            f'<a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:rPr>'
            f"<a:t>{label}</a:t></a:r></a:p></dsp:txBody>"
            f'<dsp:txXfrm><a:off x="{x + cx // 8}" y="{y + cy // 4}"/>'
            f'<a:ext cx="{cx * 3 // 4}" cy="{cy // 2}"/></dsp:txXfrm>'
        )
    return (
        f'<dsp:sp modelId="{{{label or prst}}}">'
        "<dsp:nvSpPr><dsp:cNvPr id=\"1\" name=\"\"/><dsp:cNvSpPr/></dsp:nvSpPr>"
        "<dsp:spPr>"
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:schemeClr val="{accent}"/></a:solidFill>'
        '<a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>'
        "</dsp:spPr>"
        "<dsp:style>"
        '<a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>'
        '<a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>'
        '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>'
        '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>'
        "</dsp:style>"
        f"{tx}</dsp:sp>"
    )


def make_smartart() -> None:
    # Canvas = wp:extent: 5486400 × 914400 EMU (432 × 72 pt).
    node = 1371600  # node width
    gap = 685800
    shapes = [
        dsp_shape("roundRect", 0, 0, node, 914400, "accent1", "One"),
        dsp_shape("rightArrow", node, 228600, gap, 457200, "accent2", None),
        dsp_shape("roundRect", node + gap, 0, node, 914400, "accent2", "Two"),
        dsp_shape("roundRect", 2 * (node + gap), 0, node, 914400, "accent3", "Three"),
    ]
    drawing1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<dsp:drawing xmlns:dsp="{NS_DSP}" xmlns:a="{NS_A}">'
        "<dsp:spTree>"
        "<dsp:nvGrpSpPr><dsp:cNvPr id=\"0\" name=\"\"/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>"
        "<dsp:grpSpPr/>"
        f'{"".join(shapes)}'
        "</dsp:spTree></dsp:drawing>"
    )
    data1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<dgm:dataModel xmlns:dgm="{NS_DGM}" xmlns:a="{NS_A}" xmlns:dsp="{NS_DSP}">'
        '<dgm:ptLst><dgm:pt modelId="{0}" type="doc"/></dgm:ptLst>'
        "<dgm:extLst>"
        f'<a:ext uri="{NS_DSP}">'
        f'<dsp:dataModelExt relId="rId11" minVer="{NS_DGM}"/>'
        "</a:ext></dgm:extLst></dgm:dataModel>"
    )
    layout1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<dgm:layoutDef xmlns:dgm="{NS_DGM}" uniqueId="urn:fixture/layout"/>'
    )
    quick_style1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<dgm:styleDef xmlns:dgm="{NS_DGM}" uniqueId="urn:fixture/qs"/>'
    )
    colors1 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<dgm:colorsDef xmlns:dgm="{NS_DGM}" uniqueId="urn:fixture/colors"/>'
    )
    doc_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{THEME_REL}
<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>
<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="diagrams/layout1.xml"/>
<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="diagrams/quickStyle1.xml"/>
<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="diagrams/colors1.xml"/>
<Relationship Id="rId11" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/>
</Relationships>
"""
    graphic = (
        f"{NS_DGM}|"
        f'<dgm:relIds xmlns:dgm="{NS_DGM}" xmlns:r="{NS_R}" '
        'r:dm="rId7" r:lo="rId8" r:qs="rId9" r:cs="rId10"/>'
    )
    body = (
        "<w:p><w:r><w:t>Before the diagram.</w:t></w:r></w:p>"
        + inline_drawing(5486400, 914400, graphic, "Process")
        + "<w:p><w:r><w:t>After the diagram.</w:t></w:r></w:p>"
    )
    write_docx(
        OUT / "smartart.docx",
        {
            "[Content_Types].xml": content_types(
                {
                    "/word/diagrams/data1.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
                    "/word/diagrams/layout1.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
                    "/word/diagrams/quickStyle1.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
                    "/word/diagrams/colors1.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
                    "/word/diagrams/drawing1.xml": "application/vnd.ms-office.drawingml.diagramDrawing+xml",
                    "/word/theme/theme1.xml": THEME_CT,
                }
            ),
            "_rels/.rels": ROOT_RELS,
            "word/_rels/document.xml.rels": doc_rels,
            "word/theme/theme1.xml": THEME,
            "word/document.xml": document(body),
            "word/diagrams/data1.xml": data1,
            "word/diagrams/layout1.xml": layout1,
            "word/diagrams/quickStyle1.xml": quick_style1,
            "word/diagrams/colors1.xml": colors1,
            "word/diagrams/drawing1.xml": drawing1,
        },
    )


# ── Charts ─────────────────────────────────────────────────────────────────


def str_cache(values: list[str]) -> str:
    pts = "".join(
        f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(values)
    )
    return f'<c:ptCount val="{len(values)}"/>{pts}'


def num_cache(values: list[float | None]) -> str:
    pts = "".join(
        f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
        for i, v in enumerate(values)
        if v is not None
    )
    return f'<c:formatCode>General</c:formatCode><c:ptCount val="{len(values)}"/>{pts}'


def chart_part(inner: str, title: str | None = None, legend: str | None = None) -> str:
    title_xml = (
        "<c:title><c:tx><c:rich><a:bodyPr/>"
        f'<a:p><a:pPr><a:defRPr sz="1400"/></a:pPr><a:r><a:t>{title}</a:t></a:r></a:p>'
        "</c:rich></c:tx><c:overlay val=\"0\"/></c:title><c:autoTitleDeleted val=\"0\"/>"
        if title
        else '<c:autoTitleDeleted val="1"/>'
    )
    legend_xml = (
        f'<c:legend><c:legendPos val="{legend}"/><c:overlay val="0"/></c:legend>'
        if legend
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<c:chartSpace xmlns:c="{NS_C}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">'
        f"<c:chart>{title_xml}<c:plotArea><c:layout/>{inner}</c:plotArea>"
        f'{legend_xml}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
    )


def bar_chart() -> str:
    def ser(idx: int, name: str, vals: list[float]) -> str:
        return (
            f"<c:ser><c:idx val=\"{idx}\"/><c:order val=\"{idx}\"/>"
            f"<c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache>{str_cache([name])}</c:strCache></c:strRef></c:tx>"
            f"<c:cat><c:strRef><c:f>S!$A$2:$A$4</c:f><c:strCache>{str_cache(['Q1', 'Q2', 'Q3'])}</c:strCache></c:strRef></c:cat>"
            f"<c:val><c:numRef><c:f>S!$B$2:$B$4</c:f><c:numCache>{num_cache(vals)}</c:numCache></c:numRef></c:val>"
            "</c:ser>"
        )

    inner = (
        '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
        + ser(0, "North", [4.0, 7.0, 5.0])
        + ser(1, "South", [3.0, 2.0, 6.0])
        + '<c:gapWidth val="150"/><c:overlap val="-27"/>'
        + '<c:axId val="111"/><c:axId val="222"/></c:barChart>'
        + '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        + '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>'
        + '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        + '<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="111"/></c:valAx>'
    )
    return chart_part(inner, title="Sales", legend="b")


def pie_chart() -> str:
    inner = (
        '<c:pieChart><c:varyColors val="1"/>'
        "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
        f"<c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache>{str_cache(['Share'])}</c:strCache></c:strRef></c:tx>"
        f"<c:cat><c:strRef><c:f>S!$A$2:$A$5</c:f><c:strCache>{str_cache(['A', 'B', 'C', 'D'])}</c:strCache></c:strRef></c:cat>"
        f"<c:val><c:numRef><c:f>S!$B$2:$B$5</c:f><c:numCache>{num_cache([40.0, 30.0, 20.0, 10.0])}</c:numCache></c:numRef></c:val>"
        "</c:ser><c:firstSliceAng val=\"0\"/></c:pieChart>"
    )
    return chart_part(inner, title=None, legend="r")


def line_chart() -> str:
    inner = (
        '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
        "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
        f"<c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache>{str_cache(['Trend'])}</c:strCache></c:strRef></c:tx>"
        '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>'
        f"<c:cat><c:strRef><c:f>S!$A$2:$A$6</c:f><c:strCache>{str_cache(['a', 'b', 'c', 'd', 'e'])}</c:strCache></c:strRef></c:cat>"
        f"<c:val><c:numRef><c:f>S!$B$2:$B$6</c:f><c:numCache>{num_cache([1.0, 3.0, 2.0, None, 5.0])}</c:numCache></c:numRef></c:val>"
        "<c:smooth val=\"0\"/></c:ser>"
        '<c:marker val="1"/>'
        '<c:axId val="111"/><c:axId val="222"/></c:lineChart>'
        + '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        + '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>'
        + '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        + '<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="111"/></c:valAx>'
    )
    return chart_part(inner, title=None, legend=None)


def make_charts() -> None:
    chart_ct = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    doc_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{THEME_REL}
<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
<Relationship Id="rId21" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart2.xml"/>
<Relationship Id="rId22" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart3.xml"/>
</Relationships>
"""

    def chart_ref(rid: str) -> str:
        return f'{NS_C}|<c:chart xmlns:c="{NS_C}" xmlns:r="{NS_R}" r:id="{rid}"/>'

    body = (
        "<w:p><w:r><w:t>Charts:</w:t></w:r></w:p>"
        + inline_drawing(4572000, 2743200, chart_ref("rId20"), "Bar")
        + inline_drawing(3657600, 2743200, chart_ref("rId21"), "Pie")
        + inline_drawing(4572000, 2286000, chart_ref("rId22"), "Line")
    )
    write_docx(
        OUT / "charts.docx",
        {
            "[Content_Types].xml": content_types(
                {
                    "/word/charts/chart1.xml": chart_ct,
                    "/word/charts/chart2.xml": chart_ct,
                    "/word/charts/chart3.xml": chart_ct,
                    "/word/theme/theme1.xml": THEME_CT,
                }
            ),
            "_rels/.rels": ROOT_RELS,
            "word/_rels/document.xml.rels": doc_rels,
            "word/theme/theme1.xml": THEME,
            "word/document.xml": document(body),
            "word/charts/chart1.xml": bar_chart(),
            "word/charts/chart2.xml": pie_chart(),
            "word/charts/chart3.xml": line_chart(),
        },
    )


if __name__ == "__main__":
    make_smartart()
    make_charts()
