#!/usr/bin/env python3
"""Build test-files/wmf-image.docx and test-files/svg-image.docx (issue #150).

*wmf-image.docx* holds one inline picture whose media part is a hand-built
placeable WMF ([MS-WMF]): META_PLACEABLE + META_HEADER + one
META_DIBSTRETCHBLT record wrapping a 2×2 24-bpp DIB with four distinct
quadrant colours — red/green over blue/white, stored bottom-up as DIBs are —
then META_EOF. The colours make the decode assertable pixel-by-pixel.

*svg-image.docx* holds two inline pictures over three media parts:

1. the Word shape — main `a:blip r:embed` → a 1×1 **blue** PNG fallback,
   plus the `{96DAC541-7B7A-43D3-8B79-37D633B846F1}` extension carrying
   `asvg:svgBlip r:embed` → a solid **red** SVG. Word always writes this
   pair ([MS-ODRAWXML] "Pictures"); a consumer that renders SVG reads the
   svgBlip, one that doesn't reads the blue fallback — so the drawn colour
   *is* the assertion.
2. the no-fallback shape (pandoc writes these when it cannot rasterize):
   the main blip points straight at a **green** SVG part.
"""

import pathlib
import struct
import zipfile
import zlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "test-files"

# ── WMF ──────────────────────────────────────────────────────────────────────

SRCCOPY = 0x00CC0020


def bitmapinfoheader(width: int, height: int, bpp: int) -> bytes:
    return struct.pack(
        "<IiiHHIIiiII",
        40,  # biSize
        width,
        height,  # positive = bottom-up
        1,  # biPlanes
        bpp,
        0,  # biCompression = BI_RGB
        0,  # biSizeImage
        0,
        0,  # biXPelsPerMeter, biYPelsPerMeter
        0,
        0,  # biClrUsed, biClrImportant
    )


def dib_24bpp_2x2() -> bytes:
    """2×2, rows bottom-up, BGR, each row padded to 4 bytes.

    Visual layout (top-down): red green / blue white.
    """
    bottom_row = bytes([0xFF, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0, 0])  # blue, white
    top_row = bytes([0x00, 0x00, 0xFF, 0x00, 0xFF, 0x00, 0, 0])  # red, green
    return bitmapinfoheader(2, 2, 24) + bottom_row + top_row


def wmf_record(function: int, payload: bytes) -> bytes:
    total = 6 + len(payload)
    assert total % 2 == 0
    return struct.pack("<IH", total // 2, function) + payload


def make_wmf() -> bytes:
    # META_DIBSTRETCHBLT ([MS-WMF] §2.3.1.3.1): RasterOperation u32, then
    # SrcHeight, SrcWidth, YSrc, XSrc, DestHeight, DestWidth, YDest, XDest
    # (all i16), then the packed DIB.
    blt = wmf_record(
        0x0B41,
        struct.pack("<Ihhhhhhhh", SRCCOPY, 2, 2, 0, 0, 2, 2, 0, 0) + dib_24bpp_2x2(),
    )
    eof = wmf_record(0x0000, b"")

    records = blt + eof
    size_words = (18 + len(records)) // 2
    max_record_words = len(blt) // 2
    header = struct.pack(
        "<HHHHHHIH",
        0x0002,  # DISKMETAFILE
        9,  # headerSize in words
        0x0300,  # METAVERSION300 (DIBs supported)
        size_words & 0xFFFF,
        size_words >> 16,
        0,  # numberOfObjects
        max_record_words,
        0,  # numberOfMembers
    )

    # META_PLACEABLE: 2×2 px at 96 dpi = 30 twips a side at 1440/inch.
    placeable = struct.pack("<IHhhhhHI", 0x9AC6CDD7, 0, 0, 0, 30, 30, 1440, 0)
    checksum = 0
    for (word,) in struct.iter_unpack("<H", placeable):
        checksum ^= word
    placeable += struct.pack("<H", checksum)

    return placeable + header + records


# ── PNG (1×1 blue, hand-built) ───────────────────────────────────────────────


def png_chunk(tag: bytes, data: bytes) -> bytes:
    body = tag + data
    return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))


def make_png_1x1_blue() -> bytes:
    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)  # 8-bit RGB
    idat = zlib.compress(b"\x00\x00\x00\xff")  # filter 0 + one blue pixel
    return (
        b"\x89PNG\r\n\x1a\n"
        + png_chunk(b"IHDR", ihdr)
        + png_chunk(b"IDAT", idat)
        + png_chunk(b"IEND", b"")
    )


# ── SVG ──────────────────────────────────────────────────────────────────────


def make_svg(fill: str) -> bytes:
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4" '
        f'viewBox="0 0 4 4"><rect width="4" height="4" fill="{fill}"/></svg>'
    ).encode()


# ── DOCX assembly ────────────────────────────────────────────────────────────

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def content_types(extensions: list[tuple[str, str]]) -> str:
    defaults = "".join(
        f'<Default Extension="{ext}" ContentType="{ct}"/>' for ext, ct in extensions
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + defaults
        + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )


def doc_rels(entries: list[tuple[str, str]]) -> str:
    rels = "".join(
        f'<Relationship Id="{rid}" Type="{IMAGE_REL_TYPE}" Target="media/{target}"/>'
        for rid, target in entries
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + rels
        + "</Relationships>"
    )


def inline_picture(doc_pr_id: int, name: str, blip: str) -> str:
    """One paragraph holding one inline drawing, 1×1 inch (914400 EMU)."""
    return (
        "<w:p><w:r><w:drawing>"
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="914400" cy="914400"/>'
        f'<wp:docPr id="{doc_pr_id}" name="{name}"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{doc_pr_id}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f"<pic:blipFill>{blip}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
    )


def document(body: str) -> str:
    # The trailing text run keeps the fixture a document with body content.
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<w:body>" + body + "<w:p><w:r><w:t>images</w:t></w:r></w:p>"
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
        ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
        "</w:body></w:document>"
    )


#: Word's shape: PNG fallback in the main blip, the SVG in the svgBlip ext —
#: alongside the unrelated useLocalDpi ext Word also writes, so the parser is
#: exercised against an extLst it must filter, not just read.
WORD_STYLE_BLIP = (
    '<a:blip r:embed="rIdP">'
    "<a:extLst>"
    '<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">'
    '<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>'
    "</a:ext>"
    '<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
    '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdS"/>'
    "</a:ext>"
    "</a:extLst>"
    "</a:blip>"
)

#: The no-fallback shape: the main blip references the SVG part directly.
DIRECT_SVG_BLIP = '<a:blip r:embed="rIdS2"/>'


def write_docx(target: pathlib.Path, parts: list[tuple[str, bytes]]) -> None:
    # Fixed timestamps so regenerating an unchanged fixture produces
    # identical bytes and does not show up as a diff.
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts:
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    print(f"wrote {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)

    write_docx(
        OUT / "wmf-image.docx",
        [
            ("[Content_Types].xml", content_types([("wmf", "image/x-wmf")]).encode()),
            ("_rels/.rels", ROOT_RELS.encode()),
            (
                "word/_rels/document.xml.rels",
                doc_rels([("rIdW", "image1.wmf")]).encode(),
            ),
            (
                "word/document.xml",
                document(
                    inline_picture(1, "wmf", '<a:blip r:embed="rIdW"/>')
                ).encode(),
            ),
            ("word/media/image1.wmf", make_wmf()),
        ],
    )

    write_docx(
        OUT / "svg-image.docx",
        [
            (
                "[Content_Types].xml",
                content_types(
                    [("png", "image/png"), ("svg", "image/svg+xml")]
                ).encode(),
            ),
            ("_rels/.rels", ROOT_RELS.encode()),
            (
                "word/_rels/document.xml.rels",
                doc_rels(
                    [
                        ("rIdP", "image1.png"),
                        ("rIdS", "image2.svg"),
                        ("rIdS2", "image3.svg"),
                    ]
                ).encode(),
            ),
            (
                "word/document.xml",
                document(
                    inline_picture(1, "word-style", WORD_STYLE_BLIP)
                    + inline_picture(2, "direct-svg", DIRECT_SVG_BLIP)
                ).encode(),
            ),
            ("word/media/image1.png", make_png_1x1_blue()),
            ("word/media/image2.svg", make_svg("#FF0000")),
            ("word/media/image3.svg", make_svg("#00FF00")),
        ],
    )


if __name__ == "__main__":
    main()
