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
pub(crate) fn decode_emf_bitmap_with_limits(
    emf_data: &[u8],
    limits: &crate::ResourceLimits,
) -> Result<Option<Image>, crate::ResourceLimitError> {
    if validate_emf_header(emf_data).is_none() {
        return Ok(None);
    }
    let bitmap = visit_bitmap(emf_data, &|bmi, bits| {
        if bmi.len() < BITMAPINFOHEADER_SIZE as usize || read_u32(bmi, 0)? < BITMAPINFOHEADER_SIZE {
            return None;
        }
        let width = read_i32(bmi, 4)?;
        let height = read_i32(bmi, 8)?.unsigned_abs();
        if width <= 0 || height == 0 {
            return None;
        }
        if let Err(error) = crate::limits::check_image_pixels(width as u32, height, limits) {
            return Some(Err(error));
        }
        if height > i32::MAX as u32 {
            return None;
        }
        decode_dib(bmi, bits).map(Ok)
    })
    .transpose()?;
    let Some((width, height, rgba)) = bitmap else {
        return Ok(None);
    };
    // `Opaque`, not `Premul`: every decoder below writes `0xFF` alpha, because
    // no DIB format this module accepts carries an alpha channel (see
    // `decode_32bpp`). Declaring `Premul` over straight DIB bytes was a lie
    // that happened to be harmless only while alpha was uniformly `0xFF`.
    // Adding a format that *does* carry alpha means revisiting this line.
    let info = ImageInfo::new(
        (width as i32, height as i32),
        ColorType::RGBA8888,
        AlphaType::Opaque,
        None,
    );
    Ok(images::raster_from_data(
        &info,
        Data::new_copy(&rgba),
        width as usize * 4,
    ))
}

#[cfg(test)]
fn decode_emf_bitmap(emf_data: &[u8]) -> Option<Image> {
    decode_emf_bitmap_with_limits(emf_data, &crate::ResourceLimits::default())
        .expect("test bitmap fits default resource limits")
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

#[cfg(test)]
fn extract_bitmap(data: &[u8]) -> Option<(u32, u32, Vec<u8>)> {
    visit_bitmap(data, &decode_dib)
}

fn visit_bitmap<T>(data: &[u8], decode: &impl Fn(&[u8], &[u8]) -> Option<T>) -> Option<T> {
    let mut offset: usize = 0;
    while offset + 8 <= data.len() {
        let record_type = read_u32(data, offset)?;
        let record_size = read_u32(data, offset + 4)? as usize;

        let record_end = offset.checked_add(record_size)?;
        if record_size < 8 || record_end > data.len() {
            break;
        }

        match record_type {
            EMR_STRETCHDIBITS => {
                if let Some(result) = parse_stretchdibits(data, offset, record_size, decode) {
                    return Some(result);
                }
            }
            EMR_BITBLT => {
                if let Some(result) = parse_bitblt(data, offset, record_size, decode) {
                    return Some(result);
                }
            }
            EMR_EOF => break,
            _ => {}
        }

        offset = record_end;
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
fn parse_stretchdibits<T>(
    data: &[u8],
    record_start: usize,
    record_size: usize,
    decode: &impl Fn(&[u8], &[u8]) -> Option<T>,
) -> Option<T> {
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

    let bmi_abs = record_start.checked_add(off_bmi)?;
    let bits_abs = record_start.checked_add(off_bits)?;

    decode(
        data.get(bmi_abs..bmi_abs.checked_add(cb_bmi)?)?,
        data.get(bits_abs..bits_abs.checked_add(cb_bits)?)?,
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
fn parse_bitblt<T>(
    data: &[u8],
    record_start: usize,
    record_size: usize,
    decode: &impl Fn(&[u8], &[u8]) -> Option<T>,
) -> Option<T> {
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

    let bmi_abs = record_start.checked_add(off_bmi)?;
    let bits_abs = record_start.checked_add(off_bits)?;

    decode(
        data.get(bmi_abs..bmi_abs.checked_add(cb_bmi)?)?,
        data.get(bits_abs..bits_abs.checked_add(cb_bits)?)?,
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

/// Decode a 32-bpp bottom-up-or-top-down DIB (BGRX) to top-down **opaque** RGBA.
///
/// The fourth byte of each pixel is *not* an alpha channel. For `BI_RGB` it is
/// `rgbReserved`, which MS-WMF §2.2.2.3 requires to be zero and ignored; for
/// `BI_BITFIELDS` the compression's masks cover red, green and blue only
/// (MS-WMF §2.1.1.7 — no member of the Compression enumeration this decoder
/// accepts declares an alpha mask). Copying that byte into alpha made every
/// conformant 32-bpp bitmap fully transparent, i.e. invisible.
fn decode_32bpp(
    bits: &[u8],
    width: u32,
    height: u32,
    bottom_up: bool,
) -> Option<(u32, u32, Vec<u8>)> {
    let row_bytes = (width as usize).checked_mul(4)?;
    let total = row_bytes.checked_mul(height as usize)?;
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
            dst[x * 4 + 3] = 0xFF; // reserved byte — not alpha; see above
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
    let src_row_bytes = (width as usize).checked_mul(3)?.checked_add(3)? & !3;
    let dst_row_bytes = (width as usize).checked_mul(4)?;
    let total_src = src_row_bytes.checked_mul(height as usize)?;
    let total_dst = dst_row_bytes.checked_mul(height as usize)?;
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
    fn resource_limits_cover_both_emf_records_and_bitmap_depths() {
        let limits = crate::ResourceLimits {
            max_image_pixels: 3,
            ..Default::default()
        };
        for record in [EMR_STRETCHDIBITS, EMR_BITBLT] {
            for bpp in [24, 32] {
                for height in [2, -2] {
                    let emf = make_emf(
                        record,
                        &DibSpec {
                            width: 2,
                            height,
                            bpp,
                            compression: BI_RGB,
                            pixels: vec![0; 16],
                        },
                    );
                    let error = decode_emf_bitmap_with_limits(&emf, &limits)
                        .expect_err("4-pixel DIB exceeds 3-pixel budget");
                    assert_eq!(error.resource, "image pixels");
                    assert_eq!((error.actual, error.limit), (4, 3));
                    let exact = crate::ResourceLimits {
                        max_image_pixels: 4,
                        ..limits
                    };
                    let image = decode_emf_bitmap_with_limits(&emf, &exact)
                        .unwrap()
                        .unwrap();
                    assert_eq!((image.width(), image.height()), (2, 2));
                }
            }
        }
    }

    #[test]
    fn resource_limits_reject_emf_dimensions_before_pixel_allocation() {
        let emf = make_emf(
            EMR_STRETCHDIBITS,
            &DibSpec {
                width: 100_000,
                height: -100_000,
                bpp: 32,
                compression: BI_RGB,
                pixels: vec![0; 4],
            },
        );
        let error = decode_emf_bitmap_with_limits(&emf, &crate::ResourceLimits::default())
            .expect_err("oversized header must be rejected without decoding its pixels");
        assert_eq!(error.actual, 10_000_000_000);
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

    /// H1#1 regression. MS-WMF §2.2.2.3: for a 32-bpp `BI_RGB` DIB the 4th
    /// byte is `rgbReserved`, required to be zero and ignored — it is not an
    /// alpha channel, and no compression this decoder accepts declares one
    /// (§2.1.1.7). Copying it into alpha made every conformant bitmap fully
    /// transparent, i.e. silently invisible, while a producer that wrote 0xFF
    /// there — as both corpus EMFs and the old fixture did — looked fine.
    #[test]
    fn reserved_byte_is_forced_opaque_not_read_as_alpha() {
        let (_, _, rgba) = extract_bitmap(&make_test_emf_32bpp()).expect("extract");
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
        assert_eq!(img.alpha_type(), AlphaType::Opaque);
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
        let (w, h, rgba) = extract_bitmap(&emf).expect("extract");
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
        let (w, h, rgba) = extract_bitmap(&emf).expect("extract 24bpp");
        assert_eq!((w, h), (3, 2));
        assert_eq!(rgba.len(), 3 * 2 * 4);
        // Top row = physical row 1, BGR → RGBA, always opaque.
        assert_eq!(&rgba[0..4], &[0x33, 0x22, 0x11, 0xFF]);
        assert_eq!(&rgba[8..12], &[0x99, 0x88, 0x77, 0xFF]);
        // Bottom row = physical row 0. If the 0xEE padding were consumed this
        // would start at 0xEE rather than the real first pixel.
        assert_eq!(&rgba[12..16], &[0x30, 0x20, 0x10, 0xFF]);
    }

    /// Bit depths the module does not handle must decline rather than produce
    /// garbage — 8-bpp is palette-indexed and needs the colour table.
    #[test]
    fn unsupported_bit_depth_is_declined() {
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
