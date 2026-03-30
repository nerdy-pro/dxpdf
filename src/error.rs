use thiserror::Error;

#[derive(Debug, Error)]
pub enum Error {
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Parse error: {0}")]
    Parse(#[from] dxpdf_docx::error::ParseError),

    #[error("Render error: {0}")]
    Render(#[from] dxpdf_renderer::error::RenderError),
}
