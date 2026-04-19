use thiserror::Error;

/// All errors that can occur during DOCX parsing.
#[derive(Debug, Error)]
pub enum ParseError {
    #[error("failed to read ZIP archive: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("failed to deserialize XML: {0}")]
    XmlDeserialize(#[from] quick_xml::DeError),

    #[error("invalid UTF-8 in XML content: {0}")]
    Utf8(#[from] std::str::Utf8Error),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("missing required part: {0}")]
    MissingPart(String),

    #[error("invalid attribute value '{value}' for '{attr}': {reason}")]
    InvalidAttributeValue {
        attr: String,
        value: String,
        reason: String,
    },

    #[error("invalid integer: {0}")]
    ParseInt(#[from] std::num::ParseIntError),
}

pub type Result<T> = std::result::Result<T, ParseError>;
