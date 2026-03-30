use thiserror::Error;

/// All errors that can occur during DOCX parsing.
#[derive(Debug, Error)]
pub enum ParseError {
    #[error("failed to read ZIP archive: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("failed to parse XML: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("invalid XML attribute: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),

    #[error("invalid UTF-8 in XML content: {0}")]
    Utf8(#[from] std::str::Utf8Error),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("missing required part: {0}")]
    MissingPart(String),

    #[error("missing required attribute '{attr}' on element '{element}'")]
    MissingAttribute { element: String, attr: String },

    #[error("missing required child element '{child}' in '{parent}'")]
    MissingElement { parent: String, child: String },

    #[error("invalid attribute value '{value}' for '{attr}': {reason}")]
    InvalidAttributeValue {
        attr: String,
        value: String,
        reason: String,
    },

    #[error("invalid integer: {0}")]
    ParseInt(#[from] std::num::ParseIntError),

    #[error("unresolved style reference: {0}")]
    UnresolvedStyle(String),

    #[error("circular style inheritance involving: {0}")]
    CircularStyleInheritance(String),

    #[error("unexpected end of XML inside <{0}>")]
    UnexpectedEof(String),
}

pub type Result<T> = std::result::Result<T, ParseError>;
