//! Renderer error types.

/// Errors that can occur during rendering.
///
/// Emptiness is deliberately *not* an error: a document with
/// no content produces a single blank page (`render::layout_document`), which
/// is what Word does — so the `EmptyDocument` variant this type used to carry
/// was not merely unconstructed, it contradicted the behaviour the pipeline
/// had chosen.
#[derive(Debug)]
pub enum RenderError {
    /// A raster image exceeds the caller's resource budget.
    ResourceLimit(crate::ResourceLimitError),
    /// The host font system exposes no typeface at all, so there is nothing to
    /// fall back to when a requested family cannot be resolved. Seen on
    /// container images built without any fonts installed, and with a
    /// deliberately empty `FontMgr`.
    NoFontsAvailable,
}

impl std::fmt::Display for RenderError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            RenderError::ResourceLimit(error) => error.fmt(f),
            RenderError::NoFontsAvailable => write!(
                f,
                "no fonts available on this system — install at least one font, \
                 or check that fontconfig is configured"
            ),
        }
    }
}

impl std::error::Error for RenderError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::ResourceLimit(error) => Some(error),
            Self::NoFontsAvailable => None,
        }
    }
}

impl From<crate::ResourceLimitError> for RenderError {
    fn from(error: crate::ResourceLimitError) -> Self {
        Self::ResourceLimit(error)
    }
}
