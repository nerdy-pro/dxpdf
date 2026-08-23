//! Device-Independent Bitmap decoding, shared by the EMF and WMF extractors.
//!
//! Both metafile formats carry the same payload object: [MS-WMF] §2.2.2.9
//! `DeviceIndependentBitmap` — [MS-EMF] defines its bitmap records by
//! reference to these [MS-WMF] objects — so the byte-level decoding lives
//! here once. A DIB is a header ([MS-WMF] §2.2.2.3 `BITMAPINFOHEADER`, or
//! the older 12-byte §2.2.2.2 `BitmapCoreHeader`), an optional colour table,
//! and bottom-up pixel rows padded to 4 bytes.
//!
//! Decoded here: 24-bpp BGR, 32-bpp BGRX, the palette-indexed 1/4/8-bpp
//! depths (clip-art WMFs are typically paletted), and the `BI_JPEG` /
//! `BI_PNG` compressions, whose "pixel data" is a whole embedded JPEG/PNG
//! stream handed back for Skia's ordinary decoder. Declined: the RLE
//! compressions and 16-bpp — rare in Office-produced metafiles, and a
//! decline draws nothing rather than something wrong.

use skia_safe::{images, AlphaType, ColorType, Data, Image, ImageInfo};

/// Fixed-size prefix of `BITMAPINFOHEADER` that we parse (40 bytes per spec).
pub(crate) const BITMAPINFOHEADER_SIZE: u32 = 40;
/// [MS-WMF] §2.2.2.2: the 12-byte `BitmapCoreHeader`.
const BITMAPCOREHEADER_SIZE: u32 = 12;

/// [MS-WMF] §2.1.1.7 Compression: uncompressed RGB.
pub(crate) const BI_RGB: u32 = 0;
/// Compression: uncompressed BITFIELDS (masks stored after header).
pub(crate) const BI_BITFIELDS: u32 = 3;
/// Compression: the pixel data is one embedded JPEG stream.
const BI_JPEG: u32 = 4;
/// Compression: the pixel data is one embedded PNG stream.
const BI_PNG: u32 = 5;

/// A decoded DIB, one step short of a Skia image.
///
/// `Encoded` exists because `BI_JPEG`/`BI_PNG` make the DIB a thin wrapper
/// around a stream Skia already decodes better than this module ever could.
#[derive(Debug, PartialEq, Eq)]
pub(crate) enum DibImage {
    /// Top-down, fully opaque RGBA.
    Raster {
        width: u32,
        height: u32,
        rgba: Vec<u8>,
    },
    /// A complete JPEG or PNG stream, for `Image::from_encoded`.
    Encoded(Vec<u8>),
}

/// Finish a decoded DIB into a Skia image.
pub(crate) fn to_image(dib: DibImage) -> Option<Image> {
    match dib {
        DibImage::Raster {
            width,
            height,
            rgba,
        } => {
            // `Opaque`, not `Premul`: every raster decoder below writes `0xFF`
            // alpha, because no DIB format this module accepts carries an
            // alpha channel (see `decode_32bpp`). Declaring `Premul` over
            // straight DIB bytes was a lie that happened to be harmless only
            // while alpha was uniformly `0xFF`. Adding a format that *does*
            // carry alpha means revisiting this line.
            let info = ImageInfo::new(
                (width as i32, height as i32),
                ColorType::RGBA8888,
                AlphaType::Opaque,
                None,
            );
            images::raster_from_data(&info, Data::new_copy(&rgba), width as usize * 4)
        }
        DibImage::Encoded(stream) => Image::from_encoded(Data::new_copy(&stream)),
    }
}

/// Decode a DIB whose header-plus-palette and pixel data arrive as separate
/// slices — the EMF shape, where `offBmiSrc`/`cbBmiSrc` cover the header
/// *and* the colour table and `offBitsSrc`/`cbBitsSrc` the pixels.
pub(crate) fn decode_dib(bmi: &[u8], bits: &[u8]) -> Option<DibImage> {
    let header_size = read_u32(bmi, 0)?;
    if header_size == BITMAPCOREHEADER_SIZE {
        return decode_core(bmi, bits);
    }
    if header_size < BITMAPINFOHEADER_SIZE || bmi.len() < BITMAPINFOHEADER_SIZE as usize {
        return None;
    }

    let width = read_i32(bmi, 4)?;
    let height = read_i32(bmi, 8)?; // positive = bottom-up
    let bit_count = read_u16(bmi, 14)?;
    let compression = read_u32(bmi, 16)?;
    let clr_used = read_u32(bmi, 32)?;

    // §2.1.1.7: an embedded stream replaces the pixel rows wholesale; the
    // header's width/height merely restate what the stream itself declares.
    if compression == BI_JPEG || compression == BI_PNG {
        return Some(DibImage::Encoded(bits.to_vec()));
    }

    if width <= 0 {
        return None;
    }
    // biHeight may be negative for top-down DIBs; use absolute value for sizing.
    let width = width as u32;
    let height_abs = height.unsigned_abs();
    let bottom_up = height > 0;
    if height_abs == 0 || width == 0 {
        return None;
    }

    match (bit_count, compression) {
        (32, BI_RGB | BI_BITFIELDS) => decode_32bpp(bits, width, height_abs, bottom_up),
        (24, BI_RGB) => decode_24bpp(bits, width, height_abs, bottom_up),
        (1 | 4 | 8, BI_RGB) => {
            // §2.2.2.3: the colour table follows the header inside the "bmi"
            // extent; entries are 4-byte RGBQUADs (B, G, R, reserved), and
            // `biClrUsed` shortens the table when non-zero.
            let entries = if clr_used != 0 {
                clr_used as usize
            } else {
                1usize << bit_count
            };
            let palette = bmi.get(header_size as usize..header_size as usize + entries * 4)?;
            decode_paletted(bits, width, height_abs, bottom_up, bit_count, palette, 4)
        }
        _ => None,
    }
}

/// Decode a *packed* DIB — header, colour table and pixels contiguous, the
/// WMF shape ([MS-WMF] §2.2.2.9: the record ends where the DIB does, with no
/// separate offsets).
pub(crate) fn decode_packed_dib(dib: &[u8]) -> Option<DibImage> {
    let header_size = read_u32(dib, 0)? as usize;
    let palette_bytes = if header_size == BITMAPCOREHEADER_SIZE as usize {
        // §2.2.2.2: RGBTRIPLE palette, always 2^bpp entries when indexed.
        let bit_count = read_u16(dib, 10)?;
        match bit_count {
            1 | 4 | 8 => (1usize << bit_count) * 3,
            _ => 0,
        }
    } else if header_size >= BITMAPINFOHEADER_SIZE as usize {
        let bit_count = read_u16(dib, 14)?;
        let compression = read_u32(dib, 16)?;
        let clr_used = read_u32(dib, 32)? as usize;
        match bit_count {
            1 | 4 | 8 => {
                (if clr_used != 0 {
                    clr_used
                } else {
                    1 << bit_count
                }) * 4
            }
            // §2.2.2.3: a 40-byte header with BI_BITFIELDS stores its three
            // channel masks where the palette would sit; the V4/V5 headers
            // (§2.2.2.4/§2.2.2.5) carry them inside the header instead. An
            // over-8-bpp `biClrUsed` is a legal "optimization palette".
            _ => {
                let masks = if compression == BI_BITFIELDS
                    && header_size == BITMAPINFOHEADER_SIZE as usize
                {
                    12
                } else {
                    0
                };
                clr_used * 4 + masks
            }
        }
    } else {
        return None;
    };

    let bits_at = header_size + palette_bytes;
    if bits_at > dib.len() {
        return None;
    }
    decode_dib(&dib[..bits_at], &dib[bits_at..])
}

/// [MS-WMF] §2.2.2.2 `BitmapCoreHeader`: 16-bit dimensions, always
/// bottom-up, RGBTRIPLE palette.
fn decode_core(bmi: &[u8], bits: &[u8]) -> Option<DibImage> {
    let width = read_u16(bmi, 4)? as u32;
    let height = read_u16(bmi, 6)? as u32;
    let bit_count = read_u16(bmi, 10)?;
    if width == 0 || height == 0 {
        return None;
    }
    match bit_count {
        24 => decode_24bpp(bits, width, height, true),
        1 | 4 | 8 => {
            let entries = 1usize << bit_count;
            let palette = bmi.get(12..12 + entries * 3)?;
            decode_paletted(bits, width, height, true, bit_count, palette, 3)
        }
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
fn decode_32bpp(bits: &[u8], width: u32, height: u32, bottom_up: bool) -> Option<DibImage> {
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
            dst[x * 4 + 3] = 0xFF; // reserved byte — not alpha; see above
        }
    }
    Some(DibImage::Raster {
        width,
        height,
        rgba,
    })
}

/// Decode a 24-bpp bottom-up-or-top-down DIB (BGR, 4-byte row padding) to top-down RGBA.
fn decode_24bpp(bits: &[u8], width: u32, height: u32, bottom_up: bool) -> Option<DibImage> {
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
    Some(DibImage::Raster {
        width,
        height,
        rgba,
    })
}

/// Decode a palette-indexed 1/4/8-bpp DIB to top-down opaque RGBA.
///
/// `entry_size` is 4 for the `BITMAPINFOHEADER` RGBQUAD table and 3 for the
/// core header's RGBTRIPLE one; both store channels as B, G, R. An index
/// past the table decodes as black rather than declining the whole bitmap —
/// the row layout is still sound, and a producer that undersized its table
/// via `biClrUsed` meant the missing entries to be unused.
fn decode_paletted(
    bits: &[u8],
    width: u32,
    height: u32,
    bottom_up: bool,
    bit_count: u16,
    palette: &[u8],
    entry_size: usize,
) -> Option<DibImage> {
    // Rows are padded to 32-bit boundaries whatever the depth.
    let src_row_bytes = (width as usize * bit_count as usize).div_ceil(32) * 4;
    let dst_row_bytes = width as usize * 4;
    if bits.len() < src_row_bytes * height as usize {
        return None;
    }

    let lookup = |index: usize| -> [u8; 3] {
        match palette.get(index * entry_size..index * entry_size + 3) {
            Some(bgr) => [bgr[2], bgr[1], bgr[0]],
            None => [0, 0, 0],
        }
    };

    let mut rgba = vec![0u8; dst_row_bytes * height as usize];
    for y in 0..height as usize {
        let src_row = if bottom_up {
            height as usize - 1 - y
        } else {
            y
        };
        let src = &bits[src_row * src_row_bytes..(src_row + 1) * src_row_bytes];
        let dst = &mut rgba[y * dst_row_bytes..(y + 1) * dst_row_bytes];
        for x in 0..width as usize {
            let index = match bit_count {
                8 => src[x] as usize,
                4 => {
                    let byte = src[x / 2];
                    if x % 2 == 0 {
                        (byte >> 4) as usize
                    } else {
                        (byte & 0x0F) as usize
                    }
                }
                // 1-bpp: most significant bit first.
                _ => ((src[x / 8] >> (7 - (x % 8))) & 1) as usize,
            };
            let [r, g, b] = lookup(index);
            dst[x * 4] = r;
            dst[x * 4 + 1] = g;
            dst[x * 4 + 2] = b;
            dst[x * 4 + 3] = 0xFF;
        }
    }
    Some(DibImage::Raster {
        width,
        height,
        rgba,
    })
}

// ── Byte reading helpers ─────────────────────────────────────────────────────

#[inline]
pub(crate) fn read_u32(data: &[u8], offset: usize) -> Option<u32> {
    data.get(offset..offset + 4)
        .map(|b| u32::from_le_bytes(b.try_into().unwrap()))
}

#[inline]
pub(crate) fn read_i32(data: &[u8], offset: usize) -> Option<i32> {
    data.get(offset..offset + 4)
        .map(|b| i32::from_le_bytes(b.try_into().unwrap()))
}

#[inline]
pub(crate) fn read_u16(data: &[u8], offset: usize) -> Option<u16> {
    data.get(offset..offset + 2)
        .map(|b| u16::from_le_bytes(b.try_into().unwrap()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn info_header(width: i32, height: i32, bpp: u16, compression: u32, clr_used: u32) -> Vec<u8> {
        let mut v = vec![0u8; 40];
        v[0..4].copy_from_slice(&40u32.to_le_bytes());
        v[4..8].copy_from_slice(&width.to_le_bytes());
        v[8..12].copy_from_slice(&height.to_le_bytes());
        v[12..14].copy_from_slice(&1u16.to_le_bytes());
        v[14..16].copy_from_slice(&bpp.to_le_bytes());
        v[16..20].copy_from_slice(&compression.to_le_bytes());
        v[32..36].copy_from_slice(&clr_used.to_le_bytes());
        v
    }

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

    /// An 8-bpp DIB is its palette applied to its indices: a two-entry
    /// red/blue table over a 2×2 bottom-up index grid.
    #[test]
    fn decodes_8bpp_through_its_palette() {
        let mut bmi = info_header(2, 2, 8, BI_RGB, 2);
        // RGBQUADs, stored B,G,R,reserved: entry 0 = red, entry 1 = blue.
        bmi.extend_from_slice(&[0x00, 0x00, 0xFF, 0x00, 0xFF, 0x00, 0x00, 0x00]);
        // Bottom-up rows padded to 4: bottom = [1, 0], top = [0, 1].
        let bits = [1u8, 0, 0, 0, 0, 1, 0, 0];

        let (w, h, rgba) = raster(decode_dib(&bmi, &bits).expect("decode"));
        assert_eq!((w, h), (2, 2));
        assert_eq!(&rgba[0..4], &[0xFF, 0, 0, 0xFF], "top-left red (index 0)");
        assert_eq!(&rgba[4..8], &[0, 0, 0xFF, 0xFF], "top-right blue");
        assert_eq!(&rgba[8..12], &[0, 0, 0xFF, 0xFF], "bottom-left blue");
    }

    /// 4-bpp packs two indices per byte, high nibble first; 1-bpp eight per
    /// byte, most significant bit first — pinned on asymmetric patterns that
    /// a nibble/bit-order mistake would flip.
    #[test]
    fn decodes_4bpp_and_1bpp_index_packing() {
        let mut bmi = info_header(2, 1, 4, BI_RGB, 0);
        let mut palette = vec![0u8; 16 * 4];
        palette[0..3].copy_from_slice(&[0x11, 0x22, 0x33]); // index 0: B,G,R
        palette[5 * 4..5 * 4 + 3].copy_from_slice(&[0x44, 0x55, 0x66]); // index 5
        bmi.extend_from_slice(&palette);
        // One row: indices [5, 0] → byte 0x50, padded to 4 bytes.
        let (_, _, rgba) = raster(decode_dib(&bmi, &[0x50, 0, 0, 0]).expect("4bpp"));
        assert_eq!(&rgba[0..4], &[0x66, 0x55, 0x44, 0xFF], "high nibble first");
        assert_eq!(&rgba[4..8], &[0x33, 0x22, 0x11, 0xFF]);

        let mut bmi = info_header(8, 1, 1, BI_RGB, 0);
        bmi.extend_from_slice(&[0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0x00]);
        // 0b1000_0001: first and last pixels white, the rest black.
        let (_, _, rgba) = raster(decode_dib(&bmi, &[0b1000_0001, 0, 0, 0]).expect("1bpp"));
        assert_eq!(&rgba[0..4], &[0xFF, 0xFF, 0xFF, 0xFF], "MSB is pixel 0");
        assert_eq!(&rgba[4..8], &[0, 0, 0, 0xFF]);
        assert_eq!(&rgba[28..32], &[0xFF, 0xFF, 0xFF, 0xFF]);
    }

    /// §2.1.1.7 `BI_PNG`: the "pixel data" is one whole PNG stream, handed
    /// back for Skia's decoder rather than interpreted as rows.
    #[test]
    fn png_compression_is_passed_through_encoded() {
        let bmi = info_header(1, 1, 0, super::BI_PNG, 0);
        let stream = [0x89, b'P', b'N', b'G', 1, 2, 3];
        assert_eq!(
            decode_dib(&bmi, &stream),
            Some(DibImage::Encoded(stream.to_vec()))
        );
    }

    /// A packed DIB splits at header + colour table: the same 8-bpp bitmap
    /// as above, but arriving as one contiguous slice (the WMF shape).
    #[test]
    fn packed_dib_splits_header_palette_and_bits() {
        let mut dib = info_header(2, 2, 8, BI_RGB, 2);
        dib.extend_from_slice(&[0x00, 0x00, 0xFF, 0x00, 0xFF, 0x00, 0x00, 0x00]);
        dib.extend_from_slice(&[1u8, 0, 0, 0, 0, 1, 0, 0]);
        let (w, h, rgba) = raster(decode_packed_dib(&dib).expect("packed"));
        assert_eq!((w, h), (2, 2));
        assert_eq!(&rgba[0..4], &[0xFF, 0, 0, 0xFF]);
    }

    /// The 12-byte core header: 16-bit dimensions, RGBTRIPLE palette, always
    /// bottom-up.
    #[test]
    fn decodes_a_core_header_dib() {
        let mut dib = vec![0u8; 12];
        dib[0..4].copy_from_slice(&12u32.to_le_bytes());
        dib[4..6].copy_from_slice(&1u16.to_le_bytes()); // width
        dib[6..8].copy_from_slice(&1u16.to_le_bytes()); // height
        dib[8..10].copy_from_slice(&1u16.to_le_bytes()); // planes
        dib[10..12].copy_from_slice(&8u16.to_le_bytes()); // bpp
        let mut palette = vec![0u8; 256 * 3];
        palette[7 * 3..7 * 3 + 3].copy_from_slice(&[0x10, 0x20, 0x30]); // B,G,R
        dib.extend_from_slice(&palette);
        dib.extend_from_slice(&[7, 0, 0, 0]); // one indexed pixel, padded

        let (w, h, rgba) = raster(decode_packed_dib(&dib).expect("core"));
        assert_eq!((w, h), (1, 1));
        assert_eq!(&rgba[..], &[0x30, 0x20, 0x10, 0xFF]);
    }

    /// An index past an undersized `biClrUsed` table paints black instead of
    /// declining the bitmap — the layout is sound, only the entry is absent.
    #[test]
    fn an_out_of_table_index_is_black() {
        let mut bmi = info_header(1, 1, 8, BI_RGB, 1);
        bmi.extend_from_slice(&[0xFF, 0xFF, 0xFF, 0x00]); // one white entry
        let (_, _, rgba) = raster(decode_dib(&bmi, &[9, 0, 0, 0]).expect("decode"));
        assert_eq!(&rgba[..], &[0, 0, 0, 0xFF]);
    }

    /// RLE compressions decline: drawing nothing beats drawing garbage.
    #[test]
    fn rle_compression_is_declined() {
        let bmi = info_header(2, 2, 8, 1, 0); // BI_RLE8
        assert!(decode_dib(&bmi, &[0u8; 16]).is_none());
    }
}
