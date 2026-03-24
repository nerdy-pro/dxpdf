use std::collections::HashMap;

/// Runtime context for evaluating field instructions.
///
/// Consumers provide whichever fields they have available — layout provides
/// page numbers, mail merge provides data fields, etc. Missing context
/// causes the evaluator to return `FieldValue::Unevaluable`.
#[derive(Clone, Debug, Default)]
pub struct FieldContext {
    /// Current page number (1-based).
    pub page_number: Option<u32>,
    /// Total number of pages in the document.
    pub total_pages: Option<u32>,
    /// Current section number (1-based).
    pub section_number: Option<u32>,
    /// Total pages in the current section.
    pub section_pages: Option<u32>,
    /// Current date for DATE fields.
    pub date: Option<Date>,
    /// Current time for TIME fields.
    pub time: Option<Time>,
    /// Document properties (author, title, subject, etc.).
    pub document_properties: HashMap<String, String>,
    /// Bookmark locations for REF/PAGEREF/NOTEREF.
    pub bookmarks: HashMap<String, BookmarkLocation>,
    /// Mail merge data for MERGEFIELD.
    pub merge_data: Option<HashMap<String, String>>,
    /// Sequence counters for SEQ fields.
    pub sequences: HashMap<String, u32>,
    /// File name for FILENAME field.
    pub file_name: Option<String>,
    /// Full file path for FILENAME \p.
    pub file_path: Option<String>,
}

/// A simple date representation (no external dependency required).
#[derive(Clone, Copy, Debug)]
pub struct Date {
    pub year: i32,
    pub month: u32,
    pub day: u32,
}

/// A simple time representation.
#[derive(Clone, Copy, Debug)]
pub struct Time {
    pub hour: u32,
    pub minute: u32,
    pub second: u32,
}

/// Location information for a bookmark.
#[derive(Clone, Debug)]
pub struct BookmarkLocation {
    /// The page number (1-based) where the bookmark appears.
    pub page: Option<u32>,
    /// The text content at the bookmark.
    pub text: Option<String>,
    /// Paragraph number, if applicable.
    pub paragraph_number: Option<String>,
    /// Footnote/endnote reference number, if the bookmark is in a note.
    pub note_ref: Option<String>,
}
