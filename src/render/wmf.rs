//! Minimal WMF (Windows Metafile) image extractor — the [MS-WMF] sibling of
//! [`super::emf`], and built to its pattern: handle the common case of a
//! metafile wrapping a single embedded bitmap, decline everything that would
//! need full GDI record replay.
//!
//! A WMF is an optional 22-byte placeable header ([MS-WMF] §2.3.2.3 — most
//! standalone `.wmf` files carry it; metafiles lifted from RTF or the
//! clipboard do not), an 18-byte `META_HEADER` (§2.3.2.2), then records of
//! `RecordSize` (u32, in 16-bit words, *including* itself) and
//! `RecordFunction` (u16). The bitmap carriers scanned for:
//!
//! - `META_DIBSTRETCHBLT` (§2.3.1.3) and `META_DIBBITBLT` (§2.3.1.2) — each
//!   has a second, DIB-less form that reads the device context instead; the
//!   spec's own discriminator (`RecordSize == (RecordFunction >> 8) + 3`)
//!   identifies and skips it. Per §2.1.1.1 the *high* byte of these two
//!   record types is variable, so they are matched on the low byte.
//! - `META_STRETCHDIB` (§2.3.1.6) and `META_SETDIBTODEV` (§2.3.1.4) —
//!   single-form, matched in full.
//!
//! The payload is a *packed* DIB — header, colour table and pixels
//! contiguous to the record's end — decoded by the shared [`super::dib`],
//! which is the same object family EMF uses (§2.2.2.9). The Bitmap16
//! carriers (`META_BITBLT`/`META_STRETCHBLT`) are METAVERSION100-era
//! device-dependent bitmaps, not DIBs, and are declined with the rest.

use skia_safe::Image;

use super::dib::{self, read_u16, read_u32, DibImage};

/// [MS-WMF] §2.3.2.3: the placeable header's magic, first four bytes on disk.
pub(crate) const PLACEABLE_KEY: u32 = 0x9AC6CDD7;
/// The placeable header's fixed size, including its checksum.
const PLACEABLE_SIZE: usize = 22;
/// [MS-WMF] §2.3.2.2: the standard header's size in bytes (9 words).
const HEADER_SIZE: usize = 18;

/// §2.1.1.1 record types (low byte where the high byte is variable).
const META_EOF: u16 = 0x0000;
const META_DIBBITBLT_LOW: u8 = 0x40;
const META_DIBSTRETCHBLT_LOW: u8 = 0x41;
const META_SETDIBTODEV: u16 = 0x0D33;
const META_STRETCHDIB: u16 = 0x0F43;

/// §2.1.1.31 ColorUsage: pixel data is RGB (not palette index).
const DIB_RGB_COLORS: u16 = 0;
/// Raster-operation code for a straight source copy (no blending).
const SRCCOPY: u32 = 0x00CC0020;

/// Try to extract an embedded raster bitmap from a WMF file and return a
/// decoded Skia image. Returns `None` for files that are not WMF, carry no
/// supported bitmap record, or whose DIB the shared decoder declines.
pub fn decode_wmf_bitmap(wmf_data: &[u8]) -> Option<Image> {
    let records_at = validate_headers(wmf_data)?;
    dib::to_image(extract_bitmap(&wmf_data[records_at..])?)
}

/// Check the optional placeable header and the mandatory `META_HEADER`;
/// return the offset the record list starts at.
fn validate_headers(data: &[u8]) -> Option<usize> {
    let header_at = if read_u32(data, 0)? == PLACEABLE_KEY {
        PLACEABLE_SIZE
    } else {
        0
    };
    let header = data.get(header_at..header_at + HEADER_SIZE)?;
    // §2.3.2.2: Type is 1 (memory) or 2 (disk); HeaderSize is always 9
    // words; Version is 0x0100 or 0x0300. The two fixed fields are what
    // reject a file that merely starts with plausible bytes.
    let metafile_type = read_u16(header, 0)?;
    let header_words = read_u16(header, 2)?;
    let version = read_u16(header, 4)?;
    if !matches!(metafile_type, 1 | 2) || header_words != 9 {
        return None;
    }
    if !matches!(version, 0x0100 | 0x0300) {
        return None;
    }
    Some(header_at + HEADER_SIZE)
}

/// Walk the record list; first supported DIB wins.
fn extract_bitmap(records: &[u8]) -> Option<DibImage> {
    let mut offset = 0usize;
    while offset + 6 <= records.len() {
        let size_words = read_u32(records, offset)? as usize;
        let function = read_u16(records, offset + 4)?;
        if function == META_EOF {
            break;
        }
        // §2.3.2.2: a record is at least its own two header fields (3 words).
        let record_bytes = size_words.checked_mul(2)?;
        if size_words < 3 || offset + record_bytes > records.len() {
            break;
        }
        let record = &records[offset..offset + record_bytes];

        // §2.3.1.2/§2.3.1.3: the DIB-less form's discriminator, verbatim
        // from the spec: `RecordSize == ((RecordFunction >> 8) + 3)`.
        let has_bitmap = size_words != (function >> 8) as usize + 3;

        let dib_slice = match (function, function.to_le_bytes()[0]) {
            (_, META_DIBSTRETCHBLT_LOW) if has_bitmap => {
                // §2.3.1.3.1: RasterOperation u32 at 6, eight i16 coordinate
                // fields, DIB at byte 26.
                (read_u32(record, 6)? == SRCCOPY)
                    .then(|| record.get(26..))
                    .flatten()
            }
            (_, META_DIBBITBLT_LOW) if has_bitmap => {
                // §2.3.1.2.1: RasterOperation u32 at 6, six i16 fields, DIB
                // at byte 22.
                (read_u32(record, 6)? == SRCCOPY)
                    .then(|| record.get(22..))
                    .flatten()
            }
            (META_STRETCHDIB, _) => {
                // §2.3.1.6: RasterOperation u32 at 6, ColorUsage u16 at 10,
                // eight i16 fields, DIB at byte 28.
                (read_u32(record, 6)? == SRCCOPY && read_u16(record, 10)? == DIB_RGB_COLORS)
                    .then(|| record.get(28..))
                    .flatten()
            }
            (META_SETDIBTODEV, _) => {
                // §2.3.1.4: ColorUsage u16 at 6, eight u16 fields, DIB at 24.
                (read_u16(record, 6)? == DIB_RGB_COLORS)
                    .then(|| record.get(24..))
                    .flatten()
            }
            _ => None,
        };
        if let Some(dib_bytes) = dib_slice {
            if let Some(decoded) = dib::decode_packed_dib(dib_bytes) {
                return Some(decoded);
            }
        }

        offset += record_bytes;
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::dib::BI_RGB;

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

    /// A packed 2×2 24-bpp DIB — bottom-up rows, BGR, padded to 4 bytes.
    /// Top-down reading order: red green / blue white.
    fn packed_dib_2x2() -> Vec<u8> {
        let mut v = vec![0u8; 40];
        v[0..4].copy_from_slice(&40u32.to_le_bytes());
        v[4..8].copy_from_slice(&2i32.to_le_bytes());
        v[8..12].copy_from_slice(&2i32.to_le_bytes());
        v[12..14].copy_from_slice(&1u16.to_le_bytes());
        v[14..16].copy_from_slice(&24u16.to_le_bytes());
        v[16..20].copy_from_slice(&BI_RGB.to_le_bytes());
        v.extend_from_slice(&[0xFF, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0, 0]); // blue, white
        v.extend_from_slice(&[0x00, 0x00, 0xFF, 0x00, 0xFF, 0x00, 0, 0]); // red, green
        v
    }

    fn record(function: u16, payload: &[u8]) -> Vec<u8> {
        let total = 6 + payload.len();
        assert_eq!(total % 2, 0);
        let mut v = ((total / 2) as u32).to_le_bytes().to_vec();
        v.extend_from_slice(&function.to_le_bytes());
        v.extend_from_slice(payload);
        v
    }

    fn header(records: &[u8]) -> Vec<u8> {
        let size_words = ((HEADER_SIZE + records.len() + 6) / 2) as u32;
        let mut v = Vec::new();
        v.extend_from_slice(&2u16.to_le_bytes()); // DISKMETAFILE
        v.extend_from_slice(&9u16.to_le_bytes());
        v.extend_from_slice(&0x0300u16.to_le_bytes());
        v.extend_from_slice(&size_words.to_le_bytes());
        v.extend_from_slice(&0u16.to_le_bytes()); // objects
        v.extend_from_slice(&0u32.to_le_bytes()); // maxRecord
        v.extend_from_slice(&0u16.to_le_bytes()); // members
        v
    }

    fn placeable() -> Vec<u8> {
        let mut v = Vec::new();
        v.extend_from_slice(&PLACEABLE_KEY.to_le_bytes());
        v.extend_from_slice(&[0u8; 14]); // hWmf, bbox, inch, reserved
        let checksum = v
            .chunks(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .fold(0u16, |a, w| a ^ w);
        // Reserved u32 sits before the checksum; already zeroed above.
        v.extend_from_slice(&0u16.to_le_bytes());
        v.extend_from_slice(&checksum.to_le_bytes());
        v.truncate(PLACEABLE_SIZE);
        v
    }

    /// §2.3.1.3.1 payload in front of the DIB: RasterOperation + 8 × i16.
    fn dibstretchblt(dib: &[u8]) -> Vec<u8> {
        let mut p = SRCCOPY.to_le_bytes().to_vec();
        p.extend_from_slice(&[0u8; 16]);
        p.extend_from_slice(dib);
        record(0x0B41, &p)
    }

    fn eof() -> Vec<u8> {
        record(META_EOF, &[])
    }

    fn wmf(placeable_header: bool, records: &[u8]) -> Vec<u8> {
        let mut v = if placeable_header {
            placeable()
        } else {
            Vec::new()
        };
        v.extend_from_slice(&header(records));
        v.extend_from_slice(records);
        v.extend_from_slice(&eof());
        v
    }

    #[test]
    fn extracts_the_dib_with_and_without_the_placeable_header() {
        for with_placeable in [true, false] {
            let file = wmf(with_placeable, &dibstretchblt(&packed_dib_2x2()));
            let at = validate_headers(&file).expect("valid headers");
            let (w, h, rgba) = raster(extract_bitmap(&file[at..]).expect("bitmap"));
            assert_eq!((w, h), (2, 2), "placeable={with_placeable}");
            assert_eq!(&rgba[0..4], &[0xFF, 0, 0, 0xFF], "top-left red");
            assert_eq!(&rgba[4..8], &[0, 0xFF, 0, 0xFF], "top-right green");
            assert_eq!(&rgba[8..12], &[0, 0, 0xFF, 0xFF], "bottom-left blue");
            assert_eq!(&rgba[12..16], &[0xFF, 0xFF, 0xFF, 0xFF]);
        }
    }

    /// §2.3.1.2.1: the same DIB through META_DIBBITBLT — its shorter
    /// coordinate block puts the DIB at 22, and a wrong offset shears every
    /// pixel, so equality with the stretch variant pins both transcriptions.
    #[test]
    fn dibbitblt_yields_the_same_bitmap() {
        let mut p = SRCCOPY.to_le_bytes().to_vec();
        p.extend_from_slice(&[0u8; 12]);
        p.extend_from_slice(&packed_dib_2x2());
        let via_bitblt = wmf(true, &record(0x0940, &p));
        let via_stretch = wmf(true, &dibstretchblt(&packed_dib_2x2()));

        let extract = |file: &[u8]| {
            let at = validate_headers(file).expect("headers");
            extract_bitmap(&file[at..])
        };
        assert_eq!(extract(&via_bitblt), extract(&via_stretch));
    }

    /// §2.3.1.6 META_STRETCHDIB: ColorUsage at 10, DIB at 28.
    #[test]
    fn stretchdib_is_parsed_at_its_own_offsets() {
        let mut p = SRCCOPY.to_le_bytes().to_vec();
        p.extend_from_slice(&0u16.to_le_bytes()); // ColorUsage = DIB_RGB_COLORS
        p.extend_from_slice(&[0u8; 16]);
        p.extend_from_slice(&packed_dib_2x2());
        let file = wmf(true, &record(META_STRETCHDIB, &p));
        let at = validate_headers(&file).expect("headers");
        let (w, h, _) = raster(extract_bitmap(&file[at..]).expect("bitmap"));
        assert_eq!((w, h), (2, 2));
    }

    /// §2.3.1.3: the spec's own discriminator for the DIB-less form —
    /// `RecordSize == (RecordFunction >> 8) + 3`, i.e. 14 words for
    /// DIBSTRETCHBLT — must skip the record rather than read a bitmap out
    /// of its Reserved field.
    #[test]
    fn the_dibless_form_is_skipped() {
        // 14 words total: 3 header + RasterOperation (2) + Reserved (1) +
        // 8 coordinates.
        let mut p = SRCCOPY.to_le_bytes().to_vec();
        p.extend_from_slice(&[0u8; 18]);
        let rec = record(0x0B41, &p);
        assert_eq!(rec.len() / 2, 14, "the no-DIB form is exactly 14 words");
        let file = wmf(true, &rec);
        let at = validate_headers(&file).expect("headers");
        assert!(extract_bitmap(&file[at..]).is_none());
    }

    /// A raster operation other than SRCCOPY needs source blending this
    /// module does not do — decline, as the EMF twin does.
    #[test]
    fn a_blending_raster_operation_is_declined() {
        let mut p = 0x00EE0086u32.to_le_bytes().to_vec(); // SRCPAINT
        p.extend_from_slice(&[0u8; 16]);
        p.extend_from_slice(&packed_dib_2x2());
        let file = wmf(true, &record(0x0B41, &p));
        let at = validate_headers(&file).expect("headers");
        assert!(extract_bitmap(&file[at..]).is_none());
    }

    #[test]
    fn rejects_non_wmf_data() {
        assert!(validate_headers(&[0u8; 40]).is_none());
        assert!(validate_headers(b"\x89PNG\r\n\x1a\n rest of a png").is_none());
        // A truncated placeable header must not panic.
        assert!(validate_headers(&PLACEABLE_KEY.to_le_bytes()).is_none());
    }

    #[test]
    fn decode_wmf_bitmap_returns_a_skia_image() {
        let file = wmf(true, &dibstretchblt(&packed_dib_2x2()));
        let image = decode_wmf_bitmap(&file).expect("skia image");
        assert_eq!((image.width(), image.height()), (2, 2));
    }
}
