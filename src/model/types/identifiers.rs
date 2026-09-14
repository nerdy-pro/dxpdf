//! Typed identifiers and revision tracking.

/// A relationship ID (e.g., "rId1") — opaque, interned from the .rels files.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct RelId(String);

impl RelId {
    pub fn new(id: impl Into<String>) -> Self {
        Self(id.into())
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Footnote or endnote numeric ID.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct NoteId(i64);

impl NoteId {
    pub fn new(id: i64) -> Self {
        Self(id)
    }

    pub fn value(self) -> i64 {
        self.0
    }
}

/// Identifier of a comment (`w:comment/@w:id`), the key that ties the three
/// body anchors (`commentRangeStart`/`End`, `commentReference`) to the
/// comment's content in the comments part.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct CommentId(i64);

impl CommentId {
    pub fn new(v: i64) -> Self {
        Self(v)
    }
    pub fn value(&self) -> i64 {
        self.0
    }
}

/// A bookmark ID.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct BookmarkId(i64);

impl BookmarkId {
    pub fn new(id: i64) -> Self {
        Self(id)
    }

    pub fn value(self) -> i64 {
        self.0
    }
}

/// A style ID (e.g., "Heading1", "Normal") — reference into `Document.styles`.
/// Per §17.7.4.17, this is the `w:styleId` attribute value.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct StyleId(String);

impl StyleId {
    pub fn new(id: impl Into<String>) -> Self {
        Self(id.into())
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// §17.9.1: abstract numbering definition ID.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct AbstractNumId(i64);

impl AbstractNumId {
    pub fn new(id: i64) -> Self {
        Self(id)
    }

    pub fn value(self) -> i64 {
        self.0
    }
}

/// §17.9.19: numbering instance ID.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct NumId(i64);

impl NumId {
    pub fn new(id: i64) -> Self {
        Self(id)
    }

    pub fn value(self) -> i64 {
        self.0
    }
}

/// §17.9.21: picture bullet ID. References between `w:numPicBullet` and `w:lvlPicBulletId`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct NumPicBulletId(i64);

impl NumPicBulletId {
    pub fn new(id: i64) -> Self {
        Self(id)
    }

    pub fn value(self) -> i64 {
        self.0
    }
}

/// VML shape identifier (e.g., "_x0000_t202"). Used as a `v:shapetype` `id`
/// and referenced by `v:shape` `type` (with a leading `#` prefix).
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct VmlShapeId(String);

impl VmlShapeId {
    pub fn new(id: impl Into<String>) -> Self {
        Self(id.into())
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Revision Save ID — identifies which editing session produced a change.
/// Stored as a 32-bit value parsed from an 8-digit hex string (e.g., "00A2B3C4").
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct RevisionSaveId(u32);

impl RevisionSaveId {
    pub fn value(self) -> u32 {
        self.0
    }

    /// Parse from an OOXML hex string. Returns None if invalid.
    pub fn from_hex(s: &str) -> Option<Self> {
        u32::from_str_radix(s, 16).ok().map(Self)
    }
}

/// Revision tracking IDs attached to an element.
/// Each field records which editing session performed that type of change.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct RevisionIds {
    /// Session that added this element.
    pub r: Option<RevisionSaveId>,
    /// Session that last modified this element's properties.
    pub r_pr: Option<RevisionSaveId>,
    /// Session that deleted this element (for tracked deletions).
    pub del: Option<RevisionSaveId>,
}

/// Revision tracking IDs specific to paragraphs.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct ParagraphRevisionIds {
    /// Session that added this paragraph.
    pub r: Option<RevisionSaveId>,
    /// Session that added the default run content.
    pub r_default: Option<RevisionSaveId>,
    /// Session that last modified paragraph properties.
    pub p: Option<RevisionSaveId>,
    /// Session that last modified run properties on the paragraph mark.
    pub r_pr: Option<RevisionSaveId>,
    /// Session that deleted this paragraph.
    pub del: Option<RevisionSaveId>,
}

/// Revision tracking IDs specific to table rows.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct TableRowRevisionIds {
    /// Session that added this row.
    pub r: Option<RevisionSaveId>,
    /// Session that last modified row properties.
    pub r_pr: Option<RevisionSaveId>,
    /// Session that deleted this row.
    pub del: Option<RevisionSaveId>,
    /// Session that last modified this row's table-level formatting.
    pub tr: Option<RevisionSaveId>,
}

/// Revision tracking IDs specific to section properties.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct SectionRevisionIds {
    /// Session that added this section.
    pub r: Option<RevisionSaveId>,
    /// Session that last modified section run properties.
    pub r_pr: Option<RevisionSaveId>,
    /// Session that last modified section properties.
    pub sect: Option<RevisionSaveId>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn revision_save_id_parses_eight_digit_hex() {
        assert_eq!(
            RevisionSaveId::from_hex("00A2B3C4").unwrap().value(),
            0x00A2_B3C4
        );
        assert_eq!(
            RevisionSaveId::from_hex("FFFFFFFF").unwrap().value(),
            0xFFFF_FFFF
        );
    }

    #[test]
    fn revision_save_id_rejects_non_hex() {
        assert!(RevisionSaveId::from_hex("XYZ").is_none());
        assert!(RevisionSaveId::from_hex("").is_none());
        assert!(RevisionSaveId::from_hex("-1").is_none());
    }

    #[test]
    fn revision_save_id_rejects_overflow() {
        // Nine hex digits exceed u32::MAX.
        assert!(RevisionSaveId::from_hex("1FFFFFFFF").is_none());
    }

    #[test]
    fn typed_id_accessors_round_trip() {
        assert_eq!(RelId::new("rId7").as_str(), "rId7");
        assert_eq!(NumId::new(3).value(), 3);
        assert_eq!(StyleId::new("Heading1").as_str(), "Heading1");
        assert_eq!(VmlShapeId::new("_x0000_t202").as_str(), "_x0000_t202");
    }
}
