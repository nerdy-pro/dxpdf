//! ZIP extraction and OOXML package-path resolution.

use std::collections::HashMap;
use std::io::Read;

use crate::docx::error::{ParseError, Result};
use crate::docx::whitespace_workaround::substitute_whitespace_only_runs;

/// The contents of a DOCX package, extracted from the ZIP archive.
pub struct PackageContents {
    /// All files in the ZIP, keyed by normalized path (no leading slash).
    pub parts: HashMap<String, Vec<u8>>,
}

impl PackageContents {
    /// Extract all parts from a DOCX ZIP archive.
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        let cursor = std::io::Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor)?;
        let mut parts = HashMap::with_capacity(archive.len());

        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let name = normalize_path(file.name());
            let mut buf = Vec::with_capacity(file.size() as usize);
            file.read_to_end(&mut buf)?;
            // Apply the whitespace workaround only to XML parts. Binary parts
            // (images, fonts, embedded OLE) must not be touched.
            // See `whitespace_workaround` module docs for the rationale.
            if name.ends_with(".xml") || name.ends_with(".rels") {
                buf = substitute_whitespace_only_runs(&buf);
            }
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
