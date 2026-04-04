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

use skia_safe::{images, AlphaType, ColorType, Data, Image, ImageInfo};

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

/// Fixed-size prefix of `BITMAPINFOHEADER` that we parse (40 bytes per spec).
const BITMAPINFOHEADER_SIZE: u32 = 40;

/// Compression: uncompressed RGB.
const BI_RGB: u32 = 0;
/// Compression: uncompressed BITFIELDS (masks stored after header).
const BI_BITFIELDS: u32 = 3;

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
    let (width, height, rgba) = extract_bitmap(emf_data)?;
    let info = ImageInfo::new(
        (width as i32, height as i32),
        ColorType::RGBA8888,
        AlphaType::Premul,
        None,
    );
    images::raster_from_data(&info, Data::new_copy(&rgba), width as usize * 4)
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

fn extract_bitmap(data: &[u8]) -> Option<(u32, u32, Vec<u8>)> {
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
fn parse_stretchdibits(
    data: &[u8],
    record_start: usize,
    record_size: usize,
) -> Option<(u32, u32, Vec<u8>)> {
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

    decode_dib(
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
fn parse_bitblt(
    data: &[u8],
    record_start: usize,
    record_size: usize,
) -> Option<(u32, u32, Vec<u8>)> {
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

    decode_dib(
        &data[bmi_abs..bmi_abs + cb_bmi],
        &data[bits_abs..bits_abs + cb_bits],
    )
}

// ── DIB decoding ─────────────────────────────────────────────────────────────

/// Decode a Device-Independent Bitmap to a top-down RGBA pixel buffer.
///
/// Supports 24-bpp (BGR) and 32-bpp (BGRA / BGRX) uncompressed DIBs.
/// DIB rows are stored bottom-up (positive `biHeight`) per the Windows spec;
/// we flip them so the output is top-down as expected by Skia.
fn decode_dib(bmi: &[u8], bits: &[u8]) -> Option<(u32, u32, Vec<u8>)> {
    if bmi.len() < BITMAPINFOHEADER_SIZE as usize {
        return None;
    }

    let bi_size = read_u32(bmi, 0)?;
    if bi_size < BITMAPINFOHEADER_SIZE {
        return None;
    }

    let bi_width = read_i32(bmi, 4)?;
    let bi_height = read_i32(bmi, 8)?; // positive = bottom-up
    let bi_bit_count = read_u16(bmi, 14)?;
    let bi_compression = read_u32(bmi, 16)?;

    if bi_width <= 0 {
        return None;
    }
    // biHeight may be negative for top-down DIBs; use absolute value for sizing.
    let width = bi_width as u32;
    let height = bi_height.unsigned_abs();
    let bottom_up = bi_height > 0;

    if height == 0 || width == 0 {
        return None;
    }

    match (bi_bit_count, bi_compression) {
        (32, BI_RGB | BI_BITFIELDS) => decode_32bpp(bits, width, height, bottom_up),
        (24, BI_RGB) => decode_24bpp(bits, width, height, bottom_up),
        _ => None,
    }
}

/// Decode a 32-bpp bottom-up-or-top-down DIB (BGRA or BGRX) to top-down RGBA.
fn decode_32bpp(
    bits: &[u8],
    width: u32,
    height: u32,
    bottom_up: bool,
) -> Option<(u32, u32, Vec<u8>)> {
    let row_bytes = width as usize * 4;
    let total = row_bytes * height as usize;
    if bits.len() < total {
        return None;
    }

    let mut rgba = vec![0u8; total];
    for y in 0..height as usize {
        let src_row = if bottom_up {
            height as usize - 1 - y
        } else {
            y
        };
        let src = &bits[src_row * row_bytes..(src_row + 1) * row_bytes];
        let dst = &mut rgba[y * row_bytes..(y + 1) * row_bytes];
        for x in 0..width as usize {
            // BGRA → RGBA
            dst[x * 4] = src[x * 4 + 2]; // R
            dst[x * 4 + 1] = src[x * 4 + 1]; // G
            dst[x * 4 + 2] = src[x * 4]; // B
            dst[x * 4 + 3] = src[x * 4 + 3]; // A (may be 0xFF for BGRX)
        }
    }
    Some((width, height, rgba))
}

/// Decode a 24-bpp bottom-up-or-top-down DIB (BGR, 4-byte row padding) to top-down RGBA.
fn decode_24bpp(
    bits: &[u8],
    width: u32,
    height: u32,
    bottom_up: bool,
) -> Option<(u32, u32, Vec<u8>)> {
    // DIB rows are padded to a 4-byte boundary.
    let src_row_bytes = ((width as usize * 3) + 3) & !3;
    let dst_row_bytes = width as usize * 4;
    let total_src = src_row_bytes * height as usize;
    let total_dst = dst_row_bytes * height as usize;
    if bits.len() < total_src {
        return None;
    }

    let mut rgba = vec![0u8; total_dst];
    for y in 0..height as usize {
        let src_row = if bottom_up {
            height as usize - 1 - y
        } else {
            y
        };
        let src = &bits[src_row * src_row_bytes..(src_row * src_row_bytes) + width as usize * 3];
        let dst = &mut rgba[y * dst_row_bytes..(y + 1) * dst_row_bytes];
        for x in 0..width as usize {
            // BGR → RGBA (fully opaque)
            dst[x * 4] = src[x * 3 + 2]; // R
            dst[x * 4 + 1] = src[x * 3 + 1]; // G
            dst[x * 4 + 2] = src[x * 3]; // B
            dst[x * 4 + 3] = 0xFF; // A
        }
    }
    Some((width, height, rgba))
}

// ── Byte reading helpers ─────────────────────────────────────────────────────

#[inline]
fn read_u32(data: &[u8], offset: usize) -> Option<u32> {
    data.get(offset..offset + 4)
        .map(|b| u32::from_le_bytes(b.try_into().unwrap()))
}

#[inline]
fn read_i32(data: &[u8], offset: usize) -> Option<i32> {
    data.get(offset..offset + 4)
        .map(|b| i32::from_le_bytes(b.try_into().unwrap()))
}

#[inline]
fn read_u16(data: &[u8], offset: usize) -> Option<u16> {
    data.get(offset..offset + 2)
        .map(|b| u16::from_le_bytes(b.try_into().unwrap()))
}

// ── Tests ────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a minimal valid EMF containing one EMR_STRETCHDIBITS with a 2×2
    /// 32-bpp bottom-up DIB.
    fn make_test_emf_32bpp() -> Vec<u8> {
        // Pixel data: 2×2, BGRA, bottom-up → rows ordered: row1, row0.
        // Row 1 (top in DIB order = bottom logical): [B0,G0,R0,A0, B1,G1,R1,A1]
        // Row 0 (bottom in DIB order = top logical): [B2,G2,R2,A2, B3,G3,R3,A3]
        // Bottom-up DIB: physical row 0 (stored first) = bottom of image;
        // physical row 1 (stored second) = top of image.
        #[rustfmt::skip]
        let pixels: Vec<u8> = vec![
            // Physical row 0 = bottom of image (BGRA)
            0x10, 0x20, 0x30, 0xFF,  // B=0x10 G=0x20 R=0x30 → RGBA 0x30,0x20,0x10,0xFF
            0x40, 0x50, 0x60, 0xFF,
            // Physical row 1 = top of image (BGRA)
            0x70, 0x80, 0x90, 0xFF,  // B=0x70 G=0x80 R=0x90 → RGBA 0x90,0x80,0x70,0xFF
            0xA0, 0xB0, 0xC0, 0xFF,
        ];

        let bmi: Vec<u8> = {
            let mut v = vec![0u8; 40];
            v[0..4].copy_from_slice(&40u32.to_le_bytes()); // biSize
            v[4..8].copy_from_slice(&2i32.to_le_bytes()); // biWidth
            v[8..12].copy_from_slice(&2i32.to_le_bytes()); // biHeight (positive = bottom-up)
            v[12..14].copy_from_slice(&1u16.to_le_bytes()); // biPlanes
            v[14..16].copy_from_slice(&32u16.to_le_bytes()); // biBitCount
                                                             // biCompression, biSizeImage, etc. = 0 (BI_RGB)
            v
        };

        // EMR_STRETCHDIBITS: 80-byte fixed header (offsets 0–79 per MS-EMF §2.3.1.7),
        // followed immediately by bmi at offset 80 and bits at offset 120.
        let mut record = vec![0u8; 80];
        record[0..4].copy_from_slice(&EMR_STRETCHDIBITS.to_le_bytes()); // type
        record[4..8].copy_from_slice(&(80u32 + 40 + 16).to_le_bytes()); // size
                                                                        // Bounds: 0,0,2,2
                                                                        // xDest=0, yDest=0, xSrc=0, ySrc=0, cxSrc=2, cySrc=2
        record[48..52].copy_from_slice(&80u32.to_le_bytes()); // offBmiSrc
        record[52..56].copy_from_slice(&40u32.to_le_bytes()); // cbBmiSrc
        record[56..60].copy_from_slice(&120u32.to_le_bytes()); // offBitsSrc
        record[60..64].copy_from_slice(&16u32.to_le_bytes()); // cbBitsSrc
        record[64..68].copy_from_slice(&DIB_RGB_COLORS.to_le_bytes()); // iUsageSrc
        record[68..72].copy_from_slice(&SRCCOPY.to_le_bytes()); // dwRop

        record.extend_from_slice(&bmi);
        record.extend_from_slice(&pixels);

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

    #[test]
    fn extracts_32bpp_bitmap() {
        let emf = make_test_emf_32bpp();
        let (w, h, rgba) = extract_bitmap(&emf).expect("should extract bitmap");
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
}
