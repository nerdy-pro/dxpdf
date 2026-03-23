#![allow(clippy::collapsible_if, clippy::collapsible_match)]

pub(crate) mod dimension;
pub(crate) mod error;
pub(crate) mod geometry;
pub(crate) mod model;
pub(crate) mod parse;
pub(crate) mod render;

pub use error::Error;

/// Convert raw DOCX bytes into PDF bytes.
pub fn convert(docx_bytes: &[u8]) -> Result<Vec<u8>, Error> {
    use std::time::Instant;

    let t0 = Instant::now();
    let mut document = parse::parse(docx_bytes)?;
    parse::resolve(&mut document);
    log::debug!("Parse:  {:?}", t0.elapsed());

    let t1 = Instant::now();
    let font_mgr = skia_safe::FontMgr::new();
    render::fonts::preload_fonts(&font_mgr, &document.font_families());
    log::debug!("Fonts:  {:?}", t1.elapsed());

    let t2 = Instant::now();
    let measured = render::layout::measure::measure(&document, &font_mgr);
    let pages = render::layout::layout(&measured, &font_mgr);
    log::debug!("Layout: {:?}", t2.elapsed());

    let t3 = Instant::now();
    let pdf_bytes = render::painter::render_to_pdf_with_font_mgr(&pages, &font_mgr)?;
    log::debug!("Paint:  {:?}", t3.elapsed());

    log::debug!("Total:  {:?}", t0.elapsed());
    log::info!("Pages: {} | Size: {} bytes", pages.len(), pdf_bytes.len());

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
