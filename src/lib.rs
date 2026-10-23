#[cfg(feature = "capi")]
mod capi;
pub mod docx;
pub mod error;
pub mod field;
pub mod i18n;
pub mod limits;
pub mod model;
pub mod render;

pub use error::Error;
pub use limits::{ResourceLimitError, ResourceLimits};
pub use render::{RenderOptions, DEFAULT_IMAGE_DPI, MIN_IMAGE_DPI};

/// Convert raw DOCX bytes into PDF bytes using default [`RenderOptions`].
pub fn convert(docx_bytes: &[u8]) -> Result<Vec<u8>, Error> {
    convert_with_options(docx_bytes, &RenderOptions::default())
}

/// Convert raw DOCX bytes into PDF bytes with caller-supplied [`RenderOptions`]
/// (e.g. a non-default embedded-image DPI).
pub fn convert_with_options(docx_bytes: &[u8], options: &RenderOptions) -> Result<Vec<u8>, Error> {
    use std::time::Instant;

    let t0 = Instant::now();
    let document = crate::docx::parse_with_limits(docx_bytes, options.resource_limits())?;
    log::debug!("Parse:  {:?}", t0.elapsed());

    let t1 = Instant::now();
    let pdf_bytes = crate::render::render(document, options)?;
    log::debug!("Render: {:?}", t1.elapsed());

    log::debug!("Total:  {:?}", t0.elapsed());
    Ok(pdf_bytes)
}

// --- Python bindings (enabled with `python` feature) ---

#[cfg(feature = "python")]
mod python {
    use pyo3::exceptions::PyRuntimeError;
    use pyo3::prelude::*;

    /// Convert DOCX bytes to PDF bytes.
    ///
    /// `image_dpi` sets the target resolution (pixels per inch) embedded raster
    /// images are downsampled to; defaults to 220. Releases the GIL for the
    /// duration of the conversion. For untrusted inputs, use a resource-limited
    /// subprocess; releasing the GIL does not isolate crashes or exhaustion.
    #[pyfunction]
    #[pyo3(signature = (docx_bytes, image_dpi = crate::DEFAULT_IMAGE_DPI, *, limits = None))]
    fn convert(
        py: Python<'_>,
        docx_bytes: &[u8],
        image_dpi: f32,
        limits: Option<PyRef<'_, crate::ResourceLimits>>,
    ) -> PyResult<Vec<u8>> {
        let limits = limits.map(|limits| *limits).unwrap_or_default();
        crate::limits::check_limit(
            "input bytes",
            docx_bytes.len() as u64,
            limits.max_input_bytes,
        )
        .map_err(|error| PyRuntimeError::new_err(error.to_string()))?;
        // Copied out while the GIL is held: `docx_bytes` zero-copy-borrows a
        // Python-owned buffer, and releasing the GIL below lets another thread
        // mutate or free it (e.g. a `bytearray`) out from under a borrow.
        let docx_bytes = docx_bytes.to_vec();
        py.detach(|| {
            let options = crate::RenderOptions::default()
                .with_image_dpi(image_dpi)
                .with_resource_limits(limits);
            crate::convert_with_options(&docx_bytes, &options)
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))
        })
    }

    /// Convert a DOCX file to a PDF file.
    ///
    /// `image_dpi` sets the target resolution (pixels per inch) embedded raster
    /// images are downsampled to; defaults to 220. Releases the GIL for the
    /// duration of the read, conversion and write. Untrusted inputs still need
    /// a resource-limited subprocess.
    #[pyfunction]
    #[pyo3(signature = (input, output, image_dpi = crate::DEFAULT_IMAGE_DPI, *, limits = None))]
    fn convert_file(
        py: Python<'_>,
        input: &str,
        output: &str,
        image_dpi: f32,
        limits: Option<PyRef<'_, crate::ResourceLimits>>,
    ) -> PyResult<()> {
        let limits = limits.map(|limits| *limits).unwrap_or_default();
        // Copied out while the GIL is held, same reasoning as in `convert`.
        let input = input.to_owned();
        let output = output.to_owned();
        py.detach(|| {
            let file = std::fs::File::open(&input)
                .map_err(|e| PyRuntimeError::new_err(format!("Failed to read {input}: {e}")))?;
            let docx_bytes = limits
                .read_input(file)
                .map_err(|e| PyRuntimeError::new_err(format!("Failed to read {input}: {e}")))?;
            let options = crate::RenderOptions::default()
                .with_image_dpi(image_dpi)
                .with_resource_limits(limits);
            let pdf_bytes = crate::convert_with_options(&docx_bytes, &options)
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
            std::fs::write(&output, &pdf_bytes)
                .map_err(|e| PyRuntimeError::new_err(format!("Failed to write {output}: {e}")))
        })
    }

    /// A fast DOCX-to-PDF converter powered by Skia.
    #[pymodule]
    fn dxpdf(m: &Bound<'_, PyModule>) -> PyResult<()> {
        m.add_class::<crate::ResourceLimits>()?;
        m.add_function(wrap_pyfunction!(convert, m)?)?;
        m.add_function(wrap_pyfunction!(convert_file, m)?)?;
        Ok(())
    }
}
