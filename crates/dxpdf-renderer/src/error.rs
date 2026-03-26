//! Renderer error types.

/// Errors that can occur during rendering.
#[derive(Debug)]
pub enum RenderError {
    /// No sections found in the document.
    EmptyDocument,
}

impl std::fmt::Display for RenderError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            RenderError::EmptyDocument => write!(f, "document has no content to render"),
        }
    }
}

impl std::error::Error for RenderError {}
