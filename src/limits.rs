//! Per-conversion input budgets, not a substitute for process isolation.

/// Resource budgets applied by default to parsing and raster-image decoding.
///
/// Limits are inclusive. Zero permits no bytes, entries, or pixels for that
/// resource; it does not disable a limit. Raising a limit increases the memory
/// a document can consume. These budgets do not cap total RSS, CPU, XML nesting,
/// font processing, or output size. Use a resource-limited worker for untrusted
/// documents.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[cfg_attr(
    feature = "python",
    pyo3::pyclass(frozen, get_all, skip_from_py_object, module = "dxpdf")
)]
pub struct ResourceLimits {
    /// Maximum compressed DOCX input size (default 32 MiB).
    pub max_input_bytes: u64,
    /// Maximum ZIP entry count, including directories (default 4096).
    pub max_zip_entries: u64,
    /// Maximum uncompressed bytes in any ZIP entry (default 64 MiB).
    pub max_zip_entry_bytes: u64,
    /// Maximum total uncompressed ZIP bytes, including duplicate paths
    /// (default 256 MiB).
    pub max_zip_total_bytes: u64,
    /// Maximum width times height of each source raster image, before
    /// downsampling (default 25 million pixels).
    pub max_image_pixels: u64,
}

impl ResourceLimits {
    /// Finite default budgets, also used by the CLI and C ABI.
    pub const DEFAULT: Self = Self {
        max_input_bytes: 32 * 1024 * 1024,
        max_zip_entries: 4096,
        max_zip_entry_bytes: 64 * 1024 * 1024,
        max_zip_total_bytes: 256 * 1024 * 1024,
        max_image_pixels: 25_000_000,
    };

    /// Read a DOCX input without buffering more than `max_input_bytes`.
    ///
    /// The reader is checked incrementally, even when its size is not known
    /// or changes while reading.
    pub fn read_input(&self, reader: impl std::io::Read) -> crate::docx::error::Result<Vec<u8>> {
        read_limited(reader, |size| {
            check_limit("input bytes", size, self.max_input_bytes)
        })
    }
}

impl Default for ResourceLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

#[cfg(feature = "python")]
#[pyo3::pymethods]
impl ResourceLimits {
    #[new]
    #[pyo3(signature = (*,
        max_input_bytes = Self::DEFAULT.max_input_bytes,
        max_zip_entries = Self::DEFAULT.max_zip_entries,
        max_zip_entry_bytes = Self::DEFAULT.max_zip_entry_bytes,
        max_zip_total_bytes = Self::DEFAULT.max_zip_total_bytes,
        max_image_pixels = Self::DEFAULT.max_image_pixels
    ))]
    fn new(
        max_input_bytes: u64,
        max_zip_entries: u64,
        max_zip_entry_bytes: u64,
        max_zip_total_bytes: u64,
        max_image_pixels: u64,
    ) -> Self {
        Self {
            max_input_bytes,
            max_zip_entries,
            max_zip_entry_bytes,
            max_zip_total_bytes,
            max_image_pixels,
        }
    }
}

/// A document exceeded a configured resource budget.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[error("resource limit exceeded: {resource} is {actual}, limit is {limit}")]
pub struct ResourceLimitError {
    pub resource: &'static str,
    pub actual: u64,
    pub limit: u64,
}

pub(crate) fn check_limit(
    resource: &'static str,
    actual: u64,
    limit: u64,
) -> Result<(), ResourceLimitError> {
    if actual > limit {
        Err(ResourceLimitError {
            resource,
            actual,
            limit,
        })
    } else {
        Ok(())
    }
}

pub(crate) fn check_image_pixels(
    width: u32,
    height: u32,
    limits: &ResourceLimits,
) -> Result<(), ResourceLimitError> {
    check_limit(
        "image pixels",
        u64::from(width) * u64::from(height),
        limits.max_image_pixels,
    )
}

pub(crate) fn read_limited(
    mut reader: impl std::io::Read,
    check: impl Fn(u64) -> Result<(), ResourceLimitError>,
) -> crate::docx::error::Result<Vec<u8>> {
    let mut bytes = Vec::new();
    let mut chunk = [0; 8192];
    loop {
        let count = match reader.read(&mut chunk) {
            Err(error) if error.kind() == std::io::ErrorKind::Interrupted => continue,
            result => result?,
        };
        if count == 0 {
            return Ok(bytes);
        }
        check(bytes.len() as u64 + count as u64)?;
        bytes.extend_from_slice(&chunk[..count]);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Read};

    #[test]
    fn resource_limits_bound_actual_stream_bytes() {
        let limits = ResourceLimits {
            max_input_bytes: 8,
            ..Default::default()
        };
        let mut source = Cursor::new([0; 16_384]);
        let error = limits.read_input(&mut source).unwrap_err();
        assert!(matches!(
            error,
            crate::docx::error::ParseError::ResourceLimit(ResourceLimitError {
                resource: "input bytes",
                limit: 8,
                ..
            })
        ));
        assert_eq!(source.position(), 8192, "must stop reading after rejection");
        assert_eq!(
            limits.read_input(Cursor::new(b"12345678")).unwrap(),
            b"12345678"
        );
    }

    #[test]
    fn resource_limits_zero_allows_only_empty_input() {
        let limits = ResourceLimits {
            max_input_bytes: 0,
            ..Default::default()
        };
        assert!(limits.read_input(Cursor::new(b"")).unwrap().is_empty());
        assert!(matches!(
            limits.read_input(Cursor::new(b"x")),
            Err(crate::docx::error::ParseError::ResourceLimit(_))
        ));
    }

    #[test]
    fn resource_limits_retry_interrupted_reads_but_surface_io_errors() {
        struct InterruptedThenData(bool, Cursor<&'static [u8]>);
        impl Read for InterruptedThenData {
            fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
                if !self.0 {
                    self.0 = true;
                    Err(std::io::ErrorKind::Interrupted.into())
                } else {
                    self.1.read(buffer)
                }
            }
        }
        let limits = ResourceLimits {
            max_input_bytes: 3,
            ..Default::default()
        };
        assert_eq!(
            limits
                .read_input(InterruptedThenData(false, Cursor::new(b"abc")))
                .unwrap(),
            b"abc"
        );
        struct Broken;
        impl Read for Broken {
            fn read(&mut self, _: &mut [u8]) -> std::io::Result<usize> {
                Err(std::io::ErrorKind::PermissionDenied.into())
            }
        }
        assert!(matches!(
            limits.read_input(Broken),
            Err(crate::docx::error::ParseError::Io(error))
                if error.kind() == std::io::ErrorKind::PermissionDenied
        ));
    }
}
