#![allow(clippy::collapsible_if, clippy::collapsible_match)]

pub mod docx;
pub mod error;
pub mod field;
pub mod model;
pub mod render;

pub use error::Error;

/// Convert raw DOCX bytes into PDF bytes.
pub fn convert(docx_bytes: &[u8]) -> Result<Vec<u8>, Error> {
    use std::time::Instant;

    let t0 = Instant::now();
    let document = crate::docx::parse(docx_bytes)?;
    log::debug!("Parse:  {:?}", t0.elapsed());

    let t1 = Instant::now();
    let pdf_bytes = crate::render::render(&document)?;
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
    #[pyfunction]
    fn convert(docx_bytes: &[u8]) -> PyResult<Vec<u8>> {
        crate::convert(docx_bytes).map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    /// Convert a DOCX file to a PDF file.
    #[pyfunction]
    fn convert_file(input: &str, output: &str) -> PyResult<()> {
        let docx_bytes = std::fs::read(input)
            .map_err(|e| PyRuntimeError::new_err(format!("Failed to read {input}: {e}")))?;
        let pdf_bytes =
            crate::convert(&docx_bytes).map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
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
