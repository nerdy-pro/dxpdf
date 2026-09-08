#[cfg(feature = "capi")]
mod capi;
pub mod docx;
pub mod error;
pub mod field;
pub mod i18n;
pub mod model;
pub mod render;

pub use error::Error;
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
    let document = crate::docx::parse(docx_bytes)?;
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
    /// images are downsampled to; defaults to 220.
    #[pyfunction]
    #[pyo3(signature = (docx_bytes, image_dpi = crate::DEFAULT_IMAGE_DPI))]
    fn convert(docx_bytes: &[u8], image_dpi: f32) -> PyResult<Vec<u8>> {
        let options = crate::RenderOptions::default().with_image_dpi(image_dpi);
        crate::convert_with_options(docx_bytes, &options)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    /// Convert a DOCX file to a PDF file.
    ///
    /// `image_dpi` sets the target resolution (pixels per inch) embedded raster
    /// images are downsampled to; defaults to 220.
    #[pyfunction]
    #[pyo3(signature = (input, output, image_dpi = crate::DEFAULT_IMAGE_DPI))]
    fn convert_file(input: &str, output: &str, image_dpi: f32) -> PyResult<()> {
        let docx_bytes = std::fs::read(input)
            .map_err(|e| PyRuntimeError::new_err(format!("Failed to read {input}: {e}")))?;
        let options = crate::RenderOptions::default().with_image_dpi(image_dpi);
        let pdf_bytes = crate::convert_with_options(&docx_bytes, &options)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        std::fs::write(output, &pdf_bytes)
            .map_err(|e| PyRuntimeError::new_err(format!("Failed to write {output}: {e}")))?;
        Ok(())
    }

    /// A fast DOCX-to-PDF converter powered by Skia.
    #[pymodule]
    fn dxpdf(m: &Bound<'_, PyModule>) -> PyResult<()> {
        m.add_function(wrap_pyfunction!(convert, m)?)?;
        m.add_function(wrap_pyfunction!(convert_file, m)?)?;
        Ok(())
    }
}
