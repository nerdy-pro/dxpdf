//! Header / footer set ADT.
//!
//! Per ECMA-376 §17.10.5 (`headerReference`/`footerReference`'s
//! `@type` attribute), each section can supply up to three header XML
//! parts and three footer XML parts, distinguished by purpose:
//!
//! - `default` — used on most pages
//! - `first` — used on page 1 of a section when `<w:titlePg/>` is set
//!   (§17.10.6); falling back to a *blank* header (not `default`) when
//!   `first` is missing
//! - `even` — used on even-numbered logical pages when the document
//!   setting `<w:evenAndOddHeaders/>` is on (§17.10.1); falling back to
//!   *blank* (not `default`) when `even` is missing
//!
//! The data structure is generic over the slot value: parse layer holds
//! relationship ids, resolve layer holds the loaded block content, and
//! tests can use any owned/borrowed value type. Selection lives in
//! `crate::render::layout::header_footer`.

/// Three slots per ECMA-376 §17.10.5. Slots are independent; an absent
/// slot is `None` and is *not* substituted with `default` at selection
/// time — that matches Word's observed behavior and the spec literal.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct HeaderFooterSet<T> {
    /// `<w:headerReference w:type="default" .../>` — the everyday slot.
    pub default: Option<T>,
    /// `<w:headerReference w:type="first" .../>` — paired with
    /// `<w:titlePg/>` on the section.
    pub first: Option<T>,
    /// `<w:headerReference w:type="even" .../>` — paired with
    /// `<w:evenAndOddHeaders/>` on document settings.
    pub even: Option<T>,
}

/// The three reference kinds OOXML defines. Useful as a key when
/// iterating slots and as the result of a per-page selection.
#[derive(Copy, Clone, Debug, Eq, Hash, PartialEq)]
pub enum HeaderFooterKind {
    Default,
    First,
    Even,
}

impl<T> HeaderFooterSet<T> {
    /// True when no slot is populated. Allows callers to skip
    /// header/footer rendering entirely for sections that have neither.
    pub fn is_empty(&self) -> bool {
        self.default.is_none() && self.first.is_none() && self.even.is_none()
    }

    /// Borrow a specific slot.
    pub fn get(&self, kind: HeaderFooterKind) -> Option<&T> {
        match kind {
            HeaderFooterKind::Default => self.default.as_ref(),
            HeaderFooterKind::First => self.first.as_ref(),
            HeaderFooterKind::Even => self.even.as_ref(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_set_reports_empty() {
        let s: HeaderFooterSet<&str> = HeaderFooterSet::default();
        assert!(s.is_empty());
    }

    #[test]
    fn populated_set_is_not_empty() {
        let s: HeaderFooterSet<&str> = HeaderFooterSet {
            default: Some("d"),
            first: None,
            even: None,
        };
        assert!(!s.is_empty());
    }

    #[test]
    fn get_returns_each_slot() {
        let s = HeaderFooterSet {
            default: Some("D"),
            first: Some("F"),
            even: Some("E"),
        };
        assert_eq!(s.get(HeaderFooterKind::Default), Some(&"D"));
        assert_eq!(s.get(HeaderFooterKind::First), Some(&"F"));
        assert_eq!(s.get(HeaderFooterKind::Even), Some(&"E"));
    }

    #[test]
    fn get_returns_none_for_absent_slot() {
        let s: HeaderFooterSet<&str> = HeaderFooterSet {
            default: Some("D"),
            first: None,
            even: None,
        };
        assert_eq!(s.get(HeaderFooterKind::First), None);
        assert_eq!(s.get(HeaderFooterKind::Even), None);
    }
}
