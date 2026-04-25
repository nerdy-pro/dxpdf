//! Font-format detection from raw bytes.
//!
//! Single source of truth for "what flavor of bytes are we looking at?" — used
//! by both byte extraction (to decide whether to unwrap WOFF) and subsetting
//! (to decide whether `subsetter` can handle the input directly).
//!
//! The classification follows the OpenType / WOFF / TrueType Collection
//! specifications by inspecting the first 4 bytes (the *sfnt version* or
//! *signature*). For TTC we additionally read `numFonts` from the header
//! (offset 8, big-endian u32) per the OpenType spec.

/// Flavor of an SFNT (Spline Font / Scalable Font) container — the family
/// of formats that share the offset-table header layout.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SfntFlavor {
    /// Standard `00 01 00 00` TrueType offset-table version.
    TrueTypeStandard,
    /// Apple `true` (sfnt-v2) TrueType offset-table version. Same on-disk
    /// structure as `TrueTypeStandard` — only the version tag differs.
    TrueTypeApple,
    /// OpenType with CFF or CFF2 outlines (`OTTO` signature).
    OpenTypeCff,
}

/// WOFF wrapper version — distinguishes the two compression schemes
/// (zlib per-table for v1, brotli all-tables for v2).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum WoffVersion {
    V1,
    V2,
}

/// Detected on-disk font format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FontFormat {
    /// A direct SFNT — bytes are subsettable as-is.
    Sfnt(SfntFlavor),
    /// A TrueType Collection containing `face_count` faces. The caller
    /// must select a face by index before subsetting.
    Ttc { face_count: u32 },
    /// A WOFF wrapper — must be unwrapped to SFNT before subsetting.
    Woff(WoffVersion),
}

impl FontFormat {
    /// Detect the format from the leading bytes of a font file.
    pub fn detect(bytes: &[u8]) -> Result<FontFormat, FormatError> {
        if bytes.len() < 4 {
            return Err(FormatError::TooShort {
                actual: bytes.len(),
            });
        }
        let magic = [bytes[0], bytes[1], bytes[2], bytes[3]];
        match magic {
            [0x00, 0x01, 0x00, 0x00] => Ok(Self::Sfnt(SfntFlavor::TrueTypeStandard)),
            [b't', b'r', b'u', b'e'] => Ok(Self::Sfnt(SfntFlavor::TrueTypeApple)),
            [b'O', b'T', b'T', b'O'] => Ok(Self::Sfnt(SfntFlavor::OpenTypeCff)),
            [b't', b't', b'c', b'f'] => {
                // OpenType TTC header:
                //   offset 0: TTCTag        ('ttcf')
                //   offset 4: majorVersion  uint16
                //   offset 6: minorVersion  uint16
                //   offset 8: numFonts      uint32 (big-endian)
                if bytes.len() < 12 {
                    return Err(FormatError::TtcHeaderTruncated {
                        actual: bytes.len(),
                    });
                }
                let face_count = u32::from_be_bytes([bytes[8], bytes[9], bytes[10], bytes[11]]);
                if face_count == 0 {
                    return Err(FormatError::TtcEmpty);
                }
                Ok(Self::Ttc { face_count })
            }
            [b'w', b'O', b'F', b'F'] => Ok(Self::Woff(WoffVersion::V1)),
            [b'w', b'O', b'F', b'2'] => Ok(Self::Woff(WoffVersion::V2)),
            other => Err(FormatError::UnknownMagic(other)),
        }
    }

    /// Whether the bytes can be fed to `subsetter::subset` without an
    /// unwrap step. TTCs need a face slice; WOFFs need decompression.
    pub fn is_directly_subsettable(self) -> bool {
        matches!(self, Self::Sfnt(_))
    }
}

#[derive(thiserror::Error, Debug, PartialEq, Eq)]
pub enum FormatError {
    #[error("font data is shorter than 4 bytes (got {actual})")]
    TooShort { actual: usize },
    #[error("TTC header truncated: need at least 12 bytes, got {actual}")]
    TtcHeaderTruncated { actual: usize },
    #[error("TTC header advertises zero faces")]
    TtcEmpty,
    #[error("unknown font magic bytes: {0:02x?}")]
    UnknownMagic([u8; 4]),
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn detect_truetype_standard_magic() {
        let bytes = [0x00, 0x01, 0x00, 0x00, 0xde, 0xad, 0xbe, 0xef];
        assert_eq!(
            FontFormat::detect(&bytes).unwrap(),
            FontFormat::Sfnt(SfntFlavor::TrueTypeStandard)
        );
    }

    #[test]
    fn detect_apple_truetype_magic() {
        let mut bytes = b"true".to_vec();
        bytes.extend_from_slice(&[0u8; 16]);
        assert_eq!(
            FontFormat::detect(&bytes).unwrap(),
            FontFormat::Sfnt(SfntFlavor::TrueTypeApple)
        );
    }

    #[test]
    fn detect_opentype_cff_magic() {
        let mut bytes = b"OTTO".to_vec();
        bytes.extend_from_slice(&[0u8; 16]);
        assert_eq!(
            FontFormat::detect(&bytes).unwrap(),
            FontFormat::Sfnt(SfntFlavor::OpenTypeCff)
        );
    }

    #[test]
    fn detect_ttc_with_face_count() {
        let mut bytes = b"ttcf".to_vec();
        bytes.extend_from_slice(&[0x00, 0x01, 0x00, 0x00]); // version 1.0
        bytes.extend_from_slice(&[0x00, 0x00, 0x00, 0x03]); // numFonts = 3 (BE)
        assert_eq!(
            FontFormat::detect(&bytes).unwrap(),
            FontFormat::Ttc { face_count: 3 }
        );
    }

    #[test]
    fn detect_ttc_truncated_header() {
        let bytes = b"ttcf\x00\x01".to_vec(); // only 6 bytes — short of 12
        assert!(matches!(
            FontFormat::detect(&bytes).unwrap_err(),
            FormatError::TtcHeaderTruncated { actual: 6 }
        ));
    }

    #[test]
    fn detect_ttc_zero_faces_is_invalid() {
        let mut bytes = b"ttcf".to_vec();
        bytes.extend_from_slice(&[0x00, 0x01, 0x00, 0x00]);
        bytes.extend_from_slice(&[0x00, 0x00, 0x00, 0x00]); // numFonts = 0
        assert_eq!(
            FontFormat::detect(&bytes).unwrap_err(),
            FormatError::TtcEmpty
        );
    }

    #[test]
    fn detect_woff1() {
        let mut bytes = b"wOFF".to_vec();
        bytes.extend_from_slice(&[0u8; 16]);
        assert_eq!(
            FontFormat::detect(&bytes).unwrap(),
            FontFormat::Woff(WoffVersion::V1)
        );
    }

    #[test]
    fn detect_woff2() {
        let mut bytes = b"wOF2".to_vec();
        bytes.extend_from_slice(&[0u8; 16]);
        assert_eq!(
            FontFormat::detect(&bytes).unwrap(),
            FontFormat::Woff(WoffVersion::V2)
        );
    }

    #[test]
    fn detect_rejects_unknown_magic() {
        let bytes = [0xab, 0xcd, 0xef, 0x00, 1, 2, 3, 4];
        assert!(matches!(
            FontFormat::detect(&bytes).unwrap_err(),
            FormatError::UnknownMagic(m) if m == [0xab, 0xcd, 0xef, 0x00]
        ));
    }

    #[test]
    fn detect_rejects_truncated() {
        let bytes = [0x00, 0x01, 0x00];
        assert_eq!(
            FontFormat::detect(&bytes).unwrap_err(),
            FormatError::TooShort { actual: 3 }
        );
    }

    #[test]
    fn detect_rejects_empty() {
        assert_eq!(
            FontFormat::detect(&[]).unwrap_err(),
            FormatError::TooShort { actual: 0 }
        );
    }

    #[test]
    fn is_directly_subsettable_only_for_sfnt() {
        assert!(FontFormat::Sfnt(SfntFlavor::TrueTypeStandard).is_directly_subsettable());
        assert!(FontFormat::Sfnt(SfntFlavor::TrueTypeApple).is_directly_subsettable());
        assert!(FontFormat::Sfnt(SfntFlavor::OpenTypeCff).is_directly_subsettable());
        assert!(!FontFormat::Ttc { face_count: 1 }.is_directly_subsettable());
        assert!(!FontFormat::Woff(WoffVersion::V1).is_directly_subsettable());
        assert!(!FontFormat::Woff(WoffVersion::V2).is_directly_subsettable());
    }
}
