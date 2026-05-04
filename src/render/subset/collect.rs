//! Codepoint-usage collection — walk every `DrawCommand::Text` in the laid-out
//! pages and record which Unicode codepoints each typeface actually paints.
//!
//! Keyed by [`TypefaceId`] (Skia `Typeface::unique_id`) so that two distinct
//! request keys (e.g. Calibri → Carlito substitution + a direct Carlito
//! request) that resolve to the same underlying typeface have their usage
//! merged — exactly what the subsetter needs.
//!
//! Codepoints (not glyph ids) are the right currency for two reasons:
//! 1. The subsetter (`fontcull`) is codepoint-driven; it walks the font's own
//!    cmap to derive the glyph closure, then keeps `GSUB` substitutions
//!    reachable from those glyphs (ligatures, contextual alternates), so any
//!    glyph the runtime might shape into is preserved.
//! 2. Codepoints are the source of truth in DOCX text — independent of the
//!    cmap of the specific font that happens to render them.

use std::collections::{BTreeMap, BTreeSet};

use skia_safe::FontStyle;

use crate::render::fonts::{FontRegistry, TypefaceId};
use crate::render::layout::draw_command::{DrawCommand, LayoutedPage};

/// Newtype around a Unicode scalar value (`char`'s underlying `u32`).
/// Prevents mixing with glyph ids (which are `u16`).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Codepoint(pub u32);

impl From<char> for Codepoint {
    fn from(c: char) -> Self {
        Self(c as u32)
    }
}

/// Single source of truth for "which codepoints does each typeface need?"
#[derive(Debug, Default, Clone)]
pub struct CodepointUsage {
    pub per_typeface: BTreeMap<TypefaceId, BTreeSet<Codepoint>>,
}

impl CodepointUsage {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn is_empty(&self) -> bool {
        self.per_typeface.is_empty()
    }

    pub fn typeface_count(&self) -> usize {
        self.per_typeface.len()
    }

    pub fn codepoints(&self, id: TypefaceId) -> Option<&BTreeSet<Codepoint>> {
        self.per_typeface.get(&id)
    }

    fn insert(&mut self, id: TypefaceId, chars: impl IntoIterator<Item = char>) {
        let entry = self.per_typeface.entry(id).or_default();
        for c in chars {
            entry.insert(Codepoint::from(c));
        }
    }
}

fn font_style(bold: bool, italic: bool) -> FontStyle {
    match (bold, italic) {
        (true, true) => FontStyle::bold_italic(),
        (true, false) => FontStyle::bold(),
        (false, true) => FontStyle::italic(),
        (false, false) => FontStyle::normal(),
    }
}

/// Walk every `DrawCommand::Text` in `pages` and accumulate codepoint usage
/// per resolved typeface. The resulting [`CodepointUsage`] is the input to
/// the subsetting pass.
pub fn collect(pages: &[LayoutedPage], registry: &FontRegistry) -> CodepointUsage {
    let mut usage = CodepointUsage::new();

    for page in pages {
        for cmd in &page.commands {
            if let DrawCommand::Text {
                text,
                font_family,
                bold,
                italic,
                ..
            } = cmd
            {
                if text.is_empty() {
                    continue;
                }
                let entry = registry.resolve(font_family, font_style(*bold, *italic));
                let typeface_id = TypefaceId::from(&entry.typeface);
                usage.insert(typeface_id, text.chars());
            }
        }
    }

    usage
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::dimension::Pt;
    use crate::render::geometry::{PtOffset, PtSize};
    use crate::render::resolve::color::RgbColor;
    use skia_safe::FontMgr;
    use std::rc::Rc;

    fn page_with_text(text: &str, family: &str, font_size: Pt, bold: bool) -> LayoutedPage {
        LayoutedPage {
            commands: vec![DrawCommand::Text {
                position: PtOffset::new(Pt::new(72.0), Pt::new(100.0)),
                text: Rc::from(text),
                font_family: Rc::from(family),
                char_spacing: Pt::ZERO,
                font_size,
                bold,
                italic: false,
                color: RgbColor::BLACK,
                text_scale: 1.0,
            }],
            page_size: PtSize::new(Pt::new(612.0), Pt::new(792.0)),
        }
    }

    fn registry() -> FontRegistry {
        FontRegistry::new(FontMgr::new())
    }

    #[test]
    fn collect_empty_pages_returns_empty_usage() {
        let r = registry();
        let usage = collect(&[], &r);
        assert!(usage.is_empty());
        assert_eq!(usage.typeface_count(), 0);
    }

    #[test]
    fn collect_skips_empty_text_commands() {
        let r = registry();
        let pages = vec![page_with_text("", "AnyFamily", Pt::new(12.0), false)];
        let usage = collect(&pages, &r);
        assert!(usage.is_empty());
    }

    #[test]
    fn collect_aggregates_same_typeface_across_sizes() {
        let r = registry();
        let pages = vec![
            page_with_text("hello", "AggregateProbe", Pt::new(10.0), false),
            page_with_text("world", "AggregateProbe", Pt::new(12.0), false),
        ];
        let usage = collect(&pages, &r);
        assert_eq!(
            usage.typeface_count(),
            1,
            "different font sizes share one typeface and merge into one usage entry"
        );
        let id = *usage.per_typeface.keys().next().unwrap();
        let cps = usage.codepoints(id).unwrap();
        // 'hello' ∪ 'world' = {h, e, l, o, w, r, d} → 7 unique chars
        assert_eq!(cps.len(), 7);
    }

    #[test]
    fn collect_keys_by_resolved_typeface_id() {
        // Two distinct unknown family names both fall back to the system
        // default — same Skia typeface, same TypefaceId, one usage entry.
        let r = registry();
        let pages = vec![
            page_with_text("alpha", "FallbackFamilyA", Pt::new(12.0), false),
            page_with_text("beta", "FallbackFamilyB", Pt::new(12.0), false),
        ];
        let usage = collect(&pages, &r);
        assert_eq!(
            usage.typeface_count(),
            1,
            "two requests resolving to one underlying typeface must share usage"
        );
    }

    #[test]
    fn collect_separates_bold_and_regular() {
        let r = registry();
        let pages = vec![
            page_with_text("regular", "WeightProbe", Pt::new(12.0), false),
            page_with_text("bold", "WeightProbe", Pt::new(12.0), true),
        ];
        let usage = collect(&pages, &r);
        assert_eq!(
            usage.typeface_count(),
            2,
            "regular and bold resolve to different typefaces — separate usage entries"
        );
    }

    #[test]
    fn collect_records_codepoints_for_simple_text() {
        let r = registry();
        let pages = vec![page_with_text("abc", "CpProbe", Pt::new(12.0), false)];
        let usage = collect(&pages, &r);
        let cps = usage.per_typeface.values().next().unwrap();
        let expected: BTreeSet<Codepoint> =
            ['a', 'b', 'c'].into_iter().map(Codepoint::from).collect();
        assert_eq!(cps, &expected);
    }

    #[test]
    fn collect_handles_unicode_supplementary_plane() {
        // CJK + emoji exercise multibyte UTF-8 and supplementary plane codepoints
        // (emoji are above U+FFFF). Collection must round-trip them as `u32`.
        let r = registry();
        let pages = vec![page_with_text(
            "日本語🎉",
            "UnicodeProbe",
            Pt::new(12.0),
            false,
        )];
        let usage = collect(&pages, &r);
        let cps = usage.per_typeface.values().next().unwrap();
        assert!(cps.contains(&Codepoint('🎉' as u32)));
        assert!(cps.contains(&Codepoint('日' as u32)));
        assert!(cps.contains(&Codepoint('本' as u32)));
        assert!(cps.contains(&Codepoint('語' as u32)));
    }

    #[test]
    fn collect_dedups_repeated_chars() {
        let r = registry();
        let pages = vec![
            page_with_text("aaaaaaaa", "DedupProbe", Pt::new(12.0), false),
            page_with_text("aaaa", "DedupProbe", Pt::new(14.0), false),
        ];
        let usage = collect(&pages, &r);
        let cps = usage.per_typeface.values().next().unwrap();
        assert_eq!(
            cps.len(),
            1,
            "a single repeated character must yield one codepoint in the set"
        );
        assert_eq!(cps.iter().next(), Some(&Codepoint::from('a')));
    }
}
