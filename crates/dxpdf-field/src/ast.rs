/// A parsed OOXML field instruction (§17.16.5).
#[derive(Clone, Debug, PartialEq)]
pub enum FieldInstruction {
    /// PAGE — current page number (§17.16.5.52).
    Page {
        switches: CommonSwitches,
    },
    /// NUMPAGES — total page count (§17.16.5.51).
    NumPages {
        switches: CommonSwitches,
    },
    /// SECTION — current section number (§17.16.5.66).
    Section {
        switches: CommonSwitches,
    },
    /// SECTIONPAGES — pages in current section (§17.16.5.67).
    SectionPages {
        switches: CommonSwitches,
    },
    /// DATE — current date (§17.16.5.13).
    Date {
        format: Option<String>,
        switches: CommonSwitches,
    },
    /// TIME — current time (§17.16.5.76).
    Time {
        format: Option<String>,
        switches: CommonSwitches,
    },
    /// AUTHOR — document author (§17.16.5.4).
    Author {
        format: Option<String>,
        switches: CommonSwitches,
    },
    /// TITLE — document title (§17.16.5.77).
    Title {
        switches: CommonSwitches,
    },
    /// SUBJECT — document subject (§17.16.5.72).
    Subject {
        switches: CommonSwitches,
    },
    /// FILENAME — document file name (§17.16.5.18).
    FileName {
        /// `\p` — include full path.
        path: bool,
        switches: CommonSwitches,
    },
    /// REF — cross-reference to a bookmark (§17.16.5.60).
    Ref {
        bookmark: String,
        /// `\f` — footnote/endnote reference number.
        footnote_ref: bool,
        /// `\n` — paragraph number only.
        paragraph_number: bool,
        /// `\h` — create hyperlink.
        hyperlink: bool,
        switches: CommonSwitches,
    },
    /// PAGEREF — page number of a bookmark (§17.16.5.53).
    PageRef {
        bookmark: String,
        /// `\h` — create hyperlink.
        hyperlink: bool,
        switches: CommonSwitches,
    },
    /// NOTEREF — footnote/endnote reference number (§17.16.5.50).
    NoteRef {
        bookmark: String,
        /// `\f` — superscript formatting.
        superscript: bool,
        /// `\h` — create hyperlink.
        hyperlink: bool,
        switches: CommonSwitches,
    },
    /// HYPERLINK — clickable link (§17.16.5.25).
    Hyperlink {
        target: String,
        /// `\l` — bookmark/anchor within target.
        anchor: Option<String>,
        /// `\t` — target frame.
        target_frame: Option<String>,
        /// `\m` — image map coordinates.
        image_map: bool,
        switches: CommonSwitches,
    },
    /// TOC — table of contents (§17.16.5.75).
    Toc {
        /// `\o` — outline levels, e.g. "1-3".
        outline_levels: Option<String>,
        /// `\h` — make entries hyperlinks.
        hyperlinks: bool,
        /// `\n` — omit page numbers.
        no_page_numbers: bool,
        /// `\t` — custom style/level pairs.
        custom_styles: Option<String>,
        /// `\f` — TC field identifier.
        tc_identifier: Option<String>,
        switches: CommonSwitches,
    },
    /// SEQ — sequence/autonumber (§17.16.5.68).
    Seq {
        identifier: String,
        /// `\r` — reset to value.
        reset_to: Option<String>,
        /// `\c` — repeat closest preceding number.
        repeat: bool,
        /// `\n` — next sequence number (default).
        next: bool,
        switches: CommonSwitches,
    },
    /// IF — conditional (§17.16.5.26).
    If {
        left: String,
        operator: ComparisonOp,
        right: String,
        then_text: String,
        else_text: String,
        switches: CommonSwitches,
    },
    /// MERGEFIELD — mail merge field (§17.16.5.44).
    MergeField {
        name: String,
        /// `\b` — text before if non-blank.
        text_before: Option<String>,
        /// `\f` — text after if non-blank.
        text_after: Option<String>,
        switches: CommonSwitches,
    },
    /// DOCPROPERTY — document property value (§17.16.5.15).
    DocProperty {
        name: String,
        switches: CommonSwitches,
    },
    /// SYMBOL — insert a character by code (§17.16.5.73).
    Symbol {
        char_code: u32,
        /// `\f` — font name.
        font: Option<String>,
        switches: CommonSwitches,
    },
    /// INCLUDEPICTURE — embedded picture (§17.16.5.28).
    IncludePicture {
        path: String,
        /// `\d` — don't store data with document.
        no_store: bool,
        switches: CommonSwitches,
    },
    /// EQ — equation (§17.16.5.16).
    Eq {
        equation: String,
        switches: CommonSwitches,
    },
    /// An unrecognized but syntactically valid field type.
    Unknown {
        field_type: String,
        raw: String,
    },
}

/// Common formatting switches shared by most field types (§17.16.4.1).
#[derive(Clone, Debug, Default, PartialEq)]
pub struct CommonSwitches {
    /// `\*` — general formatting switch (e.g., `\* MERGEFORMAT`, `\* Upper`).
    pub format: Option<String>,
    /// `\#` — numeric formatting switch (e.g., `\# "0.00"`).
    pub numeric_format: Option<String>,
    /// `\@` — date/time formatting switch (e.g., `\@ "dd/MM/yyyy"`).
    pub date_format: Option<String>,
}

/// Comparison operators for IF fields.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ComparisonOp {
    Equal,
    NotEqual,
    LessThan,
    LessEqual,
    GreaterThan,
    GreaterEqual,
}
