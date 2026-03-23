#![allow(
    clippy::too_many_arguments,
    clippy::collapsible_if,
    clippy::collapsible_match
)]

pub mod dimension;
pub mod error;
pub mod geometry;
pub mod model;
pub mod parse;
pub mod render;
pub mod units;

pub use error::Error;

/// Convert raw DOCX bytes into PDF bytes.
pub fn convert(docx_bytes: &[u8]) -> Result<Vec<u8>, Error> {
    let document = parse::parse(docx_bytes)?;
    convert_document(&document)
}

/// Convert a parsed `Document` into PDF bytes.
pub fn convert_document(document: &model::Document) -> Result<Vec<u8>, Error> {
    let font_mgr = skia_safe::FontMgr::new();
    render::fonts::preload_fonts(&font_mgr, &document.font_families());
    let measured = render::layout::measure::measure(document, &font_mgr);
    let pages = render::layout::layout(&measured, &font_mgr);
    render::painter::render_to_pdf_with_font_mgr(&pages, &font_mgr)
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
