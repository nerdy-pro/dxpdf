//! ZIP extraction and OOXML package-path resolution.

use std::collections::HashMap;

use crate::docx::error::{ParseError, Result};
use crate::limits::{check_limit, read_limited};

/// The contents of a DOCX package, extracted from the ZIP archive.
pub struct PackageContents {
    /// All files in the ZIP, keyed by normalized path (no leading slash).
    pub parts: HashMap<String, Vec<u8>>,
}

impl PackageContents {
    /// Extract all parts from a DOCX ZIP archive.
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(data, &crate::ResourceLimits::default())
    }

    /// Extract parts using caller-supplied resource budgets.
    pub fn from_bytes_with_limits(data: &[u8], limits: &crate::ResourceLimits) -> Result<Self> {
        check_limit("input bytes", data.len() as u64, limits.max_input_bytes)?;
        let cursor = std::io::Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor)?;
        check_limit("ZIP entries", archive.len() as u64, limits.max_zip_entries)?;
        let mut declared_total = 0;
        for i in 0..archive.len() {
            let file = archive.by_index_raw(i)?;
            check_limit("ZIP entry bytes", file.size(), limits.max_zip_entry_bytes)?;
            check_total_bytes(declared_total, file.size(), limits.max_zip_total_bytes)?;
            declared_total += file.size();
        }
        let mut parts = HashMap::with_capacity(archive.len());
        let mut total = 0;

        for i in 0..archive.len() {
            let file = archive.by_index(i)?;
            let name = normalize_path(file.name());
            let buf = read_limited(file, |size| {
                check_limit("ZIP entry bytes", size, limits.max_zip_entry_bytes)?;
                check_total_bytes(total, size, limits.max_zip_total_bytes)
            })?;
            total += buf.len() as u64;
            parts.insert(name, buf);
        }

        Ok(Self { parts })
    }

    /// Get the bytes for a part, case-insensitively.
    pub fn get_part(&self, path: &str) -> Option<&[u8]> {
        let normalized = normalize_path(path);
        self.parts.get(&normalized).map(|v| v.as_slice())
    }

    /// Get part bytes, or return MissingPart error.
    pub fn require_part(&self, path: &str) -> Result<&[u8]> {
        self.get_part(path)
            .ok_or_else(|| ParseError::MissingPart(path.to_string()))
    }

    /// Remove and return the owned bytes for a part. Avoids cloning.
    pub fn take_part(&mut self, path: &str) -> Option<Vec<u8>> {
        let normalized = normalize_path(path);
        self.parts.remove(&normalized)
    }
}

fn check_total_bytes(
    total: u64,
    size: u64,
    limit: u64,
) -> std::result::Result<(), crate::ResourceLimitError> {
    // Compare against the remaining budget before adding, including when a
    // caller has explicitly selected u64::MAX.
    if size > limit - total {
        Err(crate::ResourceLimitError {
            resource: "ZIP total bytes",
            actual: total.saturating_add(size),
            limit,
        })
    } else {
        Ok(())
    }
}

fn normalize_path(path: &str) -> String {
    path.trim_start_matches('/').to_lowercase()
}

/// Resolve a relationship target to an absolute path within the package.
/// base_dir is the directory containing the source part (e.g., "word" for "word/document.xml").
pub fn resolve_target(base_dir: &str, target: &str) -> String {
    if target.starts_with('/') {
        // Absolute path within the package
        normalize_path(target)
    } else {
        // Relative to base_dir
        let mut path = if base_dir.is_empty() {
            target.to_string()
        } else {
            format!("{}/{}", base_dir, target)
        };
        // Drop a leading "./" and collapse interior "/./" no-op segments so a
        // target like "./media/x.png" resolves the same as "media/x.png".
        if let Some(stripped) = path.strip_prefix("./") {
            path = stripped.to_string();
        }
        while let Some(pos) = path.find("/./") {
            path = format!("{}{}", &path[..pos], &path[pos + 2..]);
        }
        // Simplify "../" sequences
        while let Some(pos) = path.find("/../") {
            if let Some(parent_start) = path[..pos].rfind('/') {
                path = format!("{}{}", &path[..parent_start], &path[pos + 3..]);
            } else {
                path = path[pos + 4..].to_string();
            }
        }
        normalize_path(&path)
    }
}

/// Get the .rels path for a given part path.
/// e.g., "word/document.xml" → "word/_rels/document.xml.rels"
pub fn rels_path_for(part_path: &str) -> String {
    let normalized = normalize_path(part_path);
    if let Some(slash_pos) = normalized.rfind('/') {
        format!(
            "{}/_rels/{}.rels",
            &normalized[..slash_pos],
            &normalized[slash_pos + 1..]
        )
    } else {
        format!("_rels/{}.rels", normalized)
    }
}

/// Get the directory portion of a part path.
pub fn part_directory(part_path: &str) -> &str {
    match part_path.rfind('/') {
        Some(pos) => &part_path[..pos],
        None => "",
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn archive(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        for (name, data) in entries {
            writer
                .start_file(
                    *name,
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .unwrap();
            writer.write_all(data).unwrap();
        }
        writer.finish().unwrap().into_inner()
    }

    fn assert_limit(data: &[u8], limits: crate::ResourceLimits, resource: &str) {
        match PackageContents::from_bytes_with_limits(data, &limits) {
            Err(ParseError::ResourceLimit(error)) => assert_eq!(error.resource, resource),
            Err(error) => panic!("expected resource limit, got {error}"),
            Ok(_) => panic!("expected {resource} limit to reject archive"),
        }
    }

    #[test]
    fn resource_limits_reject_input_before_zip_parsing() {
        assert_limit(
            b"not a zip",
            crate::ResourceLimits {
                max_input_bytes: 8,
                ..Default::default()
            },
            "input bytes",
        );
    }

    #[test]
    fn resource_limits_reject_excess_zip_entries() {
        assert_limit(
            &archive(&[("one", b""), ("two/", b"")]),
            crate::ResourceLimits {
                max_zip_entries: 1,
                ..Default::default()
            },
            "ZIP entries",
        );
    }

    #[test]
    fn resource_limits_reject_expanded_entry_not_just_compressed_size() {
        let data = archive(&[("compressed", &[0; 8192])]);
        assert!(data.len() < 1024);
        assert_limit(
            &data,
            crate::ResourceLimits {
                max_zip_entry_bytes: 8191,
                ..Default::default()
            },
            "ZIP entry bytes",
        );
    }

    #[test]
    fn resource_limits_enforce_actual_bytes_when_zip_sizes_lie() {
        let mut data = archive(&[("compressed", &[0; 8192])]);
        let central = data
            .windows(4)
            .position(|bytes| bytes == b"PK\x01\x02")
            .expect("central directory");
        data[central + 24..central + 28].copy_from_slice(&1u32.to_le_bytes());
        data[22..26].copy_from_slice(&1u32.to_le_bytes());
        assert_limit(
            &data,
            crate::ResourceLimits {
                max_zip_entry_bytes: 8191,
                ..Default::default()
            },
            "ZIP entry bytes",
        );
        assert_limit(
            &data,
            crate::ResourceLimits {
                max_zip_total_bytes: 8191,
                ..Default::default()
            },
            "ZIP total bytes",
        );
    }

    #[test]
    fn resource_limits_count_colliding_paths_toward_total() {
        assert_limit(
            &archive(&[("ONE", b"123"), ("one", b"456")]),
            crate::ResourceLimits {
                max_zip_total_bytes: 5,
                ..Default::default()
            },
            "ZIP total bytes",
        );
    }

    #[test]
    fn resource_limits_accept_exact_boundaries() {
        let data = archive(&[("one", b"123"), ("two", b"456")]);
        let limits = crate::ResourceLimits {
            max_input_bytes: data.len() as u64,
            max_zip_entries: 2,
            max_zip_entry_bytes: 3,
            max_zip_total_bytes: 6,
            ..Default::default()
        };
        let package = PackageContents::from_bytes_with_limits(&data, &limits).unwrap();
        assert_eq!(package.get_part("one"), Some(b"123".as_slice()));
        assert_eq!(package.get_part("two"), Some(b"456".as_slice()));
    }

    #[test]
    fn resource_limits_total_arithmetic_is_inclusive_and_overflow_safe() {
        assert!(check_total_bytes(3, 3, 6).is_ok());
        assert!(check_total_bytes(3, 4, 6).is_err());
        assert!(check_total_bytes(u64::MAX - 1, 1, u64::MAX).is_ok());
        assert!(check_total_bytes(u64::MAX - 1, 2, u64::MAX).is_err());
    }

    #[test]
    fn normalize_strips_leading_slash_and_lowercases() {
        assert_eq!(normalize_path("/Word/Document.XML"), "word/document.xml");
        assert_eq!(
            normalize_path("word/media/Image1.PNG"),
            "word/media/image1.png"
        );
    }

    #[test]
    fn resolve_relative_target_joins_base_dir() {
        assert_eq!(
            resolve_target("word", "media/image1.png"),
            "word/media/image1.png"
        );
    }

    #[test]
    fn resolve_absolute_target_ignores_base_dir() {
        assert_eq!(
            resolve_target("word", "/word/media/image1.png"),
            "word/media/image1.png"
        );
    }

    #[test]
    fn resolve_empty_base_dir_uses_target_as_is() {
        assert_eq!(resolve_target("", "document.xml"), "document.xml");
    }

    #[test]
    fn resolve_collapses_parent_sequences() {
        // A theme rel from word/ pointing at word/theme/theme1.xml via "../".
        assert_eq!(
            resolve_target("word/theme", "../media/image1.png"),
            "word/media/image1.png"
        );
        // Multiple "../" segments.
        assert_eq!(resolve_target("a/b/c", "../../x.xml"), "a/x.xml");
    }

    #[test]
    fn resolve_parent_past_root_leaves_residual_dotdot() {
        // Traversing above the package root is malformed. One "../" is consumed
        // against "word"; the second has no parent left, so a residual "../"
        // remains — a path that matches no part, which is the safe outcome for
        // malformed input (the image is simply not found).
        assert_eq!(resolve_target("word", "../../x.xml"), "../x.xml");
    }

    #[test]
    fn resolve_collapses_current_dir_segments() {
        // Leading "./" and interior "/./" are no-ops and must normalize away,
        // otherwise the part lookup (an exact map key) would miss.
        assert_eq!(resolve_target("word", "./media/x.png"), "word/media/x.png");
        assert_eq!(resolve_target("", "./document.xml"), "document.xml");
        assert_eq!(resolve_target("word", "media/./x.png"), "word/media/x.png");
    }

    #[test]
    fn rels_path_for_builds_sibling_rels_file() {
        assert_eq!(
            rels_path_for("word/document.xml"),
            "word/_rels/document.xml.rels"
        );
        // A root-level part has no directory prefix.
        assert_eq!(
            rels_path_for("[Content_Types].xml"),
            "_rels/[content_types].xml.rels"
        );
    }

    #[test]
    fn part_directory_returns_dir_or_empty() {
        assert_eq!(part_directory("word/media/image1.png"), "word/media");
        assert_eq!(part_directory("document.xml"), "");
    }

    // KNOWN GAP: percent-encoded relationship targets (e.g. spaces as "%20")
    // are not decoded, so a target like "media/my%20image.png" resolves to a
    // key that won't match the ZIP part "media/my image.png". Word does not
    // percent-encode internal media names in practice, so this is left as a
    // documented limitation rather than pulling in a percent-decoder.
}
