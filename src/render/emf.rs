//! Minimal EMF (Enhanced Metafile) image extractor.
//!
//! Handles the common case of EMF files wrapping a single embedded bitmap
//! via `EMR_STRETCHDIBITS` or `EMR_BITBLT` — the pattern used by Word when
//! inserting raster images via Windows clipboard or older authoring tools.
//!
//! Returns a decoded Skia image on success. Returns `None` for complex EMF
//! files that require full GDI record replay (beziers, text, paths, etc.).
//!
//! ## References
//! - MS-EMF §2.3.1.7: EMR_STRETCHDIBITS record
//! - MS-EMF §2.3.1.2: EMR_BITBLT record
//! - MS-WMF §2.2.2.3: BITMAPINFOHEADER

use skia_safe::Image;

use super::dib::{self, read_u32, DibImage};

/// EMF record type identifiers (MS-EMF §2.3).
const EMR_HEADER: u32 = 0x00000001;
const EMR_EOF: u32 = 0x0000000E;
const EMR_BITBLT: u32 = 0x0000004C;
const EMR_STRETCHDIBITS: u32 = 0x00000051;

/// EMF file signature in the header (MS-EMF §2.2.1, dSignature field).
const EMF_SIGNATURE: u32 = 0x464D4520;

/// DIB colour usage — pixel data is RGB (not palette index).
const DIB_RGB_COLORS: u32 = 0;

/// Raster-operation code for a straight source copy (no blending).
const SRCCOPY: u32 = 0x00CC0020;

/// Try to extract an embedded raster bitmap from an EMF file and return a
/// decoded Skia image.
///
/// Scans the record list for `EMR_STRETCHDIBITS` / `EMR_BITBLT` containing a
/// Device-Independent Bitmap. Converts the DIB pixel data (24-bpp BGR or
/// 32-bpp BGRA, bottom-up row order) to a top-down RGBA raster, then creates
/// a Skia image from it.
///
/// Returns `None` if:
/// - the data is not a valid EMF file,
/// - no supported bitmap record is found,
/// - the DIB uses an unsupported bit-depth or compression.
pub fn decode_emf_bitmap(emf_data: &[u8]) -> Option<Image> {
    validate_emf_header(emf_data)?;
    dib::to_image(extract_bitmap(emf_data)?)
}

// ── Header validation ────────────────────────────────────────────────────────

/// Check the mandatory EMF header record (MS-EMF §2.2.1).
fn validate_emf_header(data: &[u8]) -> Option<()> {
    if data.len() < 88 {
        return None;
    }
    let record_type = read_u32(data, 0)?;
    if record_type != EMR_HEADER {
        return None;
    }
    // dSignature at byte 40 within the header record.
    let signature = read_u32(data, 40)?;
    if signature != EMF_SIGNATURE {
        return None;
    }
    Some(())
}

// ── Record scanning ──────────────────────────────────────────────────────────

fn extract_bitmap(data: &[u8]) -> Option<DibImage> {
    let mut offset: usize = 0;
    while offset + 8 <= data.len() {
        let record_type = read_u32(data, offset)?;
        let record_size = read_u32(data, offset + 4)? as usize;

        if record_size < 8 || offset + record_size > data.len() {
            break;
        }

        match record_type {
            EMR_STRETCHDIBITS => {
                if let Some(result) = parse_stretchdibits(data, offset, record_size) {
                    return Some(result);
                }
            }
            EMR_BITBLT => {
                if let Some(result) = parse_bitblt(data, offset, record_size) {
                    return Some(result);
                }
            }
            EMR_EOF => break,
            _ => {}
        }

        offset += record_size;
    }
    None
}

// ── EMR_STRETCHDIBITS parser (MS-EMF §2.3.1.7) ──────────────────────────────

/// Fixed-field layout of EMR_STRETCHDIBITS after the 8-byte record header:
///
/// ```text
/// Offset  Size  Field
///   8      16   Bounds (RECTL)
///  24       4   xDest
///  28       4   yDest
///  32       4   xSrc
///  36       4   ySrc
///  40       4   cxSrc
///  44       4   cySrc
///  48       4   offBmiSrc   — byte offset from record start to BITMAPINFOHEADER
///  52       4   cbBmiSrc
///  56       4   offBitsSrc  — byte offset from record start to pixel data
///  60       4   cbBitsSrc
///  64       4   iUsageSrc
///  68       4   dwRop
///  72       4   cxDest
///  76       4   cyDest
/// ```
fn parse_stretchdibits(data: &[u8], record_start: usize, record_size: usize) -> Option<DibImage> {
    if record_size < 80 {
        return None;
    }

    let off_bmi = read_u32(data, record_start + 48)? as usize;
    let cb_bmi = read_u32(data, record_start + 52)? as usize;
    let off_bits = read_u32(data, record_start + 56)? as usize;
    let cb_bits = read_u32(data, record_start + 60)? as usize;
    let usage = read_u32(data, record_start + 64)?;
    let rop = read_u32(data, record_start + 68)?;

    // Only handle RGB colour usage and straight source-copy ROP.
    if usage != DIB_RGB_COLORS || rop != SRCCOPY {
        return None;
    }
    if cb_bmi == 0 || cb_bits == 0 {
        return None;
    }

    let bmi_abs = record_start + off_bmi;
    let bits_abs = record_start + off_bits;

    if bmi_abs + cb_bmi > data.len() || bits_abs + cb_bits > data.len() {
        return None;
    }

    dib::decode_dib(
        &data[bmi_abs..bmi_abs + cb_bmi],
        &data[bits_abs..bits_abs + cb_bits],
    )
}

// ── EMR_BITBLT parser (MS-EMF §2.3.1.2) ─────────────────────────────────────

/// Fixed-field layout of EMR_BITBLT after the 8-byte record header:
///
/// ```text
/// Offset  Size  Field
///   8      16   Bounds
///  24       4   xDest
///  28       4   yDest
///  32       4   cxDest
///  36       4   cyDest
///  40       4   dwRop
///  44       4   xSrc
///  48       4   ySrc
///  52      16   xformSrc (XFORM — 6 floats)
///  68       4   crBkColorSrc
///  72       4   iUsageSrc
///  76       4   offBmiSrc
///  80       4   cbBmiSrc
///  84       4   offBitsSrc
///  88       4   cbBitsSrc
/// ```
fn parse_bitblt(data: &[u8], record_start: usize, record_size: usize) -> Option<DibImage> {
    if record_size < 92 {
        return None;
    }

    let rop = read_u32(data, record_start + 40)?;
    let usage = read_u32(data, record_start + 72)?;
    let off_bmi = read_u32(data, record_start + 76)? as usize;
    let cb_bmi = read_u32(data, record_start + 80)? as usize;
    let off_bits = read_u32(data, record_start + 84)? as usize;
    let cb_bits = read_u32(data, record_start + 88)? as usize;

    if usage != DIB_RGB_COLORS || rop != SRCCOPY {
        return None;
    }
    if cb_bmi == 0 || cb_bits == 0 {
        return None;
    }

    let bmi_abs = record_start + off_bmi;
    let bits_abs = record_start + off_bits;

    if bmi_abs + cb_bmi > data.len() || bits_abs + cb_bits > data.len() {
        return None;
    }

    dib::decode_dib(
        &data[bmi_abs..bmi_abs + cb_bmi],
        &data[bits_abs..bits_abs + cb_bits],
    )
}

// ── Tests ────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::dib::BI_RGB;

    /// Unwrap the raster arm — every fixture here decodes to raw pixels.
    fn raster(dib: DibImage) -> (u32, u32, Vec<u8>) {
        match dib {
            DibImage::Raster {
                width,
                height,
                rgba,
            } => (width, height, rgba),
            other => panic!("expected a raster, got {other:?}"),
        }
    }

    /// Parameters for a synthetic single-bitmap EMF.
    struct DibSpec {
        /// Negative height means a top-down DIB (rows already in visual order).
        height: i32,
        width: i32,
        bpp: u16,
        compression: u32,
        /// Raw DIB pixel bytes, including any 4-byte row padding.
        pixels: Vec<u8>,
    }

    /// Build a minimal valid EMF wrapping one bitmap record of `record_type`
    /// (`EMR_STRETCHDIBITS` or `EMR_BITBLT`). Both records carry the same
    /// BITMAPINFOHEADER and pixel payload — only the fixed-field layout in
    /// front of them differs, which is exactly what the offsets encode.
    fn make_emf(record_type: u32, dib: &DibSpec) -> Vec<u8> {
        let bmi: Vec<u8> = {
            let mut v = vec![0u8; 40];
            v[0..4].copy_from_slice(&40u32.to_le_bytes()); // biSize
            v[4..8].copy_from_slice(&dib.width.to_le_bytes());
            v[8..12].copy_from_slice(&dib.height.to_le_bytes());
            v[12..14].copy_from_slice(&1u16.to_le_bytes()); // biPlanes
            v[14..16].copy_from_slice(&dib.bpp.to_le_bytes());
            v[16..20].copy_from_slice(&dib.compression.to_le_bytes());
            v
        };

        // Fixed-field size of each record, per MS-EMF §2.3.1.7 / §2.3.1.2.
        let fixed: usize = if record_type == EMR_BITBLT { 92 } else { 80 };
        let off_bmi = fixed as u32;
        let off_bits = off_bmi + 40;
        let mut record = vec![0u8; fixed];
        record[0..4].copy_from_slice(&record_type.to_le_bytes());
        record[4..8].copy_from_slice(&(fixed as u32 + 40 + dib.pixels.len() as u32).to_le_bytes());
        if record_type == EMR_BITBLT {
            record[40..44].copy_from_slice(&SRCCOPY.to_le_bytes()); // dwRop
            record[72..76].copy_from_slice(&DIB_RGB_COLORS.to_le_bytes()); // iUsageSrc
            record[76..80].copy_from_slice(&off_bmi.to_le_bytes());
            record[80..84].copy_from_slice(&40u32.to_le_bytes());
            record[84..88].copy_from_slice(&off_bits.to_le_bytes());
            record[88..92].copy_from_slice(&(dib.pixels.len() as u32).to_le_bytes());
        } else {
            record[48..52].copy_from_slice(&off_bmi.to_le_bytes());
            record[52..56].copy_from_slice(&40u32.to_le_bytes());
            record[56..60].copy_from_slice(&off_bits.to_le_bytes());
            record[60..64].copy_from_slice(&(dib.pixels.len() as u32).to_le_bytes());
            record[64..68].copy_from_slice(&DIB_RGB_COLORS.to_le_bytes()); // iUsageSrc
            record[68..72].copy_from_slice(&SRCCOPY.to_le_bytes()); // dwRop
        }
        record.extend_from_slice(&bmi);
        record.extend_from_slice(&dib.pixels);

        // EMR_HEADER (88 bytes, signature at byte 40).
        let mut header_rec = vec![0u8; 88];
        header_rec[0..4].copy_from_slice(&EMR_HEADER.to_le_bytes());
        header_rec[4..8].copy_from_slice(&88u32.to_le_bytes());
        header_rec[40..44].copy_from_slice(&EMF_SIGNATURE.to_le_bytes());

        // EMR_EOF (20 bytes minimum).
        let mut eof_rec = vec![0u8; 20];
        eof_rec[0..4].copy_from_slice(&EMR_EOF.to_le_bytes());
        eof_rec[4..8].copy_from_slice(&20u32.to_le_bytes());

        let mut emf = Vec::new();
        emf.extend_from_slice(&header_rec);
        emf.extend_from_slice(&record);
        emf.extend_from_slice(&eof_rec);
        emf
    }

    /// 2×2 32-bpp bottom-up pixel payload.
    ///
    /// The 4th byte is `rgbReserved` and MS-WMF §2.2.2.3 requires it to be
    /// **zero**. It is written as zero here deliberately: the fixture used to
    /// carry 0xFF, which is what let a decoder that mistook it for alpha pass
    /// its own tests while rendering conformant files invisible.
    #[rustfmt::skip]
    fn pixels_32bpp_2x2() -> Vec<u8> {
        vec![
            // Physical row 0 = bottom of image (BGRX)
            0x10, 0x20, 0x30, 0x00,  // B=0x10 G=0x20 R=0x30 → RGBA 0x30,0x20,0x10,0xFF
            0x40, 0x50, 0x60, 0x00,
            // Physical row 1 = top of image (BGRX)
            0x70, 0x80, 0x90, 0x00,  // B=0x70 G=0x80 R=0x90 → RGBA 0x90,0x80,0x70,0xFF
            0xA0, 0xB0, 0xC0, 0x00,
        ]
    }

    /// Minimal valid EMF containing one EMR_STRETCHDIBITS with a 2×2 32-bpp
    /// bottom-up DIB.
    fn make_test_emf_32bpp() -> Vec<u8> {
        make_emf(
            EMR_STRETCHDIBITS,
            &DibSpec {
                width: 2,
                height: 2, // positive = bottom-up
                bpp: 32,
                compression: BI_RGB,
                pixels: pixels_32bpp_2x2(),
            },
        )
    }

    #[test]
    fn extracts_32bpp_bitmap() {
        let emf = make_test_emf_32bpp();
        let (w, h, rgba) = raster(extract_bitmap(&emf).expect("should extract bitmap"));
        assert_eq!(w, 2);
        assert_eq!(h, 2);
        assert_eq!(rgba.len(), 2 * 2 * 4);

        // After flip: output row 0 (top) = physical row 1; output row 1 (bottom) = physical row 0.
        assert_eq!(&rgba[0..4], &[0x90, 0x80, 0x70, 0xFF]); // top-left:    physical row 1 pixel 0
        assert_eq!(&rgba[4..8], &[0xC0, 0xB0, 0xA0, 0xFF]); // top-right:   physical row 1 pixel 1
        assert_eq!(&rgba[8..12], &[0x30, 0x20, 0x10, 0xFF]); // bottom-left: physical row 0 pixel 0
    }

    #[test]
    fn rejects_invalid_header() {
        let mut bad = make_test_emf_32bpp();
        bad[40..44].copy_from_slice(&0xDEADBEEFu32.to_le_bytes()); // corrupt signature
        assert!(validate_emf_header(&bad).is_none());
    }

    #[test]
    fn rejects_truncated_data() {
        assert!(extract_bitmap(&[0u8; 10]).is_none());
    }

    #[test]
    fn decode_emf_bitmap_returns_skia_image() {
        let emf = make_test_emf_32bpp();
        let image = decode_emf_bitmap(&emf);
        assert!(image.is_some(), "should produce a Skia image");
        let img = image.unwrap();
        assert_eq!(img.width(), 2);
        assert_eq!(img.height(), 2);
    }

    /// H1#1 regression. MS-WMF §2.2.2.3: for a 32-bpp `BI_RGB` DIB the 4th
    /// byte is `rgbReserved`, required to be zero and ignored — it is not an
    /// alpha channel, and no compression this decoder accepts declares one
    /// (§2.1.1.7). Copying it into alpha made every conformant bitmap fully
    /// transparent, i.e. silently invisible, while a producer that wrote 0xFF
    /// there — as both corpus EMFs and the old fixture did — looked fine.
    #[test]
    fn reserved_byte_is_forced_opaque_not_read_as_alpha() {
        let (_, _, rgba) = raster(extract_bitmap(&make_test_emf_32bpp()).expect("extract"));
        assert!(
            rgba.chunks(4).all(|p| p[3] == 0xFF),
            "every pixel must be opaque; the source reserved bytes are all 0x00"
        );
        // Colour must survive the change untouched.
        assert_eq!(&rgba[0..4], &[0x90, 0x80, 0x70, 0xFF]);
    }

    /// The same guarantee restated at the boundary: the `ImageInfo` handed to
    /// Skia must declare what the decoders actually produce.
    #[test]
    fn decoded_image_declares_opaque_alpha() {
        let img = decode_emf_bitmap(&make_test_emf_32bpp()).expect("image");
        assert_eq!(img.alpha_type(), skia_safe::AlphaType::Opaque);
    }

    /// A negative `biHeight` means the rows are already in visual order and
    /// must **not** be flipped (MS-WMF §2.2.2.3).
    #[test]
    fn top_down_dib_is_not_flipped() {
        let emf = make_emf(
            EMR_STRETCHDIBITS,
            &DibSpec {
                width: 2,
                height: -2, // negative = top-down
                bpp: 32,
                compression: BI_RGB,
                pixels: pixels_32bpp_2x2(),
            },
        );
        let (w, h, rgba) = raster(extract_bitmap(&emf).expect("extract"));
        assert_eq!((w, h), (2, 2), "height is the absolute value");
        // Physical row 0 stays the top row, the opposite of the bottom-up case.
        assert_eq!(&rgba[0..4], &[0x30, 0x20, 0x10, 0xFF]);
    }

    /// `parse_bitblt`'s six field offsets are transcribed constants that
    /// nothing checked. Feeding the identical DIB through both record types
    /// must yield the identical bitmap — any offset that is wrong by even four
    /// bytes reads a neighbouring field and the record is rejected.
    #[test]
    fn bitblt_yields_the_same_bitmap_as_stretchdibits() {
        let spec = || DibSpec {
            width: 2,
            height: 2,
            bpp: 32,
            compression: BI_RGB,
            pixels: pixels_32bpp_2x2(),
        };
        let via_stretch = extract_bitmap(&make_emf(EMR_STRETCHDIBITS, &spec()));
        let via_bitblt = extract_bitmap(&make_emf(EMR_BITBLT, &spec()));
        assert!(via_bitblt.is_some(), "EMR_BITBLT must be parsed at all");
        assert_eq!(via_stretch, via_bitblt);
    }

    /// 24-bpp rows are padded to a 4-byte boundary, so a width of 3 stores
    /// 9 bytes of pixels plus 3 bytes of padding per row. The padding must be
    /// skipped, not consumed as pixel data — otherwise every row after the
    /// first is shifted by its predecessors' padding.
    #[test]
    fn decodes_24bpp_skipping_row_padding() {
        #[rustfmt::skip]
        let pixels: Vec<u8> = vec![
            // Physical row 0 = bottom (BGR ×3, then 3 padding bytes)
            0x10, 0x20, 0x30,  0x40, 0x50, 0x60,  0x70, 0x80, 0x90,  0xEE, 0xEE, 0xEE,
            // Physical row 1 = top
            0x11, 0x22, 0x33,  0x44, 0x55, 0x66,  0x77, 0x88, 0x99,  0xEE, 0xEE, 0xEE,
        ];
        let emf = make_emf(
            EMR_STRETCHDIBITS,
            &DibSpec {
                width: 3,
                height: 2,
                bpp: 24,
                compression: BI_RGB,
                pixels,
            },
        );
        let (w, h, rgba) = raster(extract_bitmap(&emf).expect("extract 24bpp"));
        assert_eq!((w, h), (3, 2));
        assert_eq!(rgba.len(), 3 * 2 * 4);
        // Top row = physical row 1, BGR → RGBA, always opaque.
        assert_eq!(&rgba[0..4], &[0x33, 0x22, 0x11, 0xFF]);
        assert_eq!(&rgba[8..12], &[0x99, 0x88, 0x77, 0xFF]);
        // Bottom row = physical row 0. If the 0xEE padding were consumed this
        // would start at 0xEE rather than the real first pixel.
        assert_eq!(&rgba[12..16], &[0x30, 0x20, 0x10, 0xFF]);
    }

    /// A palette-indexed depth whose colour table is *missing* must decline
    /// rather than produce garbage: `cbBmiSrc` here covers the bare 40-byte
    /// header, so the 8-bpp indices have no table to resolve against. (With
    /// the table present the shared decoder handles it — `render::dib`'s
    /// tests pin that side.)
    #[test]
    fn a_paletted_depth_without_its_table_is_declined() {
        let emf = make_emf(
            EMR_STRETCHDIBITS,
            &DibSpec {
                width: 2,
                height: 2,
                bpp: 8,
                compression: BI_RGB,
                pixels: vec![0u8; 16],
            },
        );
        assert!(extract_bitmap(&emf).is_none());
    }
}
