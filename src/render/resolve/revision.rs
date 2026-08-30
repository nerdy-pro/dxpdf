//! Issue #154: per-author colors for tracked-change marks.
//!
//! Neither ECMA-376 nor the document says what color a revision mark is —
//! Word assigns one per author, by order of appearance, from an
//! application-defined palette ("by author" is the default of its display
//! options). This module does the same, deterministically from the document
//! itself: authors are numbered by the first appearance of any run they
//! changed, in document order, and cycle through [`REVISION_PALETTE`]. The
//! hues are this engine's own; matching Word's exact colors is neither
//! possible (they are user-configurable) nor the point — one color per
//! author, stable across renders, is.

use std::collections::HashMap;

use crate::model::{Block, Inline};
use crate::render::resolve::color::RgbColor;

/// One color per author, cycling. Starts with the classic revision red.
pub const REVISION_PALETTE: [RgbColor; 6] = [
    RgbColor { r: 192, g: 0, b: 0 },
    RgbColor {
        r: 0,
        g: 112,
        b: 192,
    },
    RgbColor {
        r: 0,
        g: 146,
        b: 70,
    },
    RgbColor {
        r: 112,
        g: 48,
        b: 160,
    },
    RgbColor {
        r: 227,
        g: 108,
        b: 10,
    },
    RgbColor {
        r: 0,
        g: 134,
        b: 139,
    },
];

/// Issue #154: the fill behind a commented range — one fixed pale tone
/// rather than an author tint, so overlapping ranges and multi-author
/// documents stay readable. Word's print view brackets the range instead of
/// shading it; the shading is this engine's choice, made because a bracket
/// needs glyph-height geometry the fragment layer doesn't expose while
/// `Fragment::Text.shading` is already plumbed end to end.
pub const COMMENT_RANGE_SHADING: RgbColor = RgbColor {
    r: 255,
    g: 235,
    b: 156,
};

/// The color for an author no walk registered — a revision in a part the
/// palette walk does not cover (a header, say) still gets a mark.
pub const REVISION_FALLBACK_COLOR: RgbColor = REVISION_PALETTE[0];

/// Number the revision authors of `blocks` in first-appearance order and
/// color each from [`REVISION_PALETTE`].
///
/// The walk covers the document body, including table content; `register`
/// is public so the comments part can join the same numbering — a person who
/// both edited and commented keeps one color.
pub fn collect_revision_colors(blocks: &[Block]) -> HashMap<String, RgbColor> {
    let mut colors = HashMap::new();
    walk_blocks(blocks, &mut colors);
    colors
}

/// Assign `author` the next palette color unless already registered.
pub fn register(author: &str, colors: &mut HashMap<String, RgbColor>) {
    if !colors.contains_key(author) {
        let color = REVISION_PALETTE[colors.len() % REVISION_PALETTE.len()];
        colors.insert(author.to_string(), color);
    }
}

fn walk_blocks(blocks: &[Block], colors: &mut HashMap<String, RgbColor>) {
    for block in blocks {
        match block {
            Block::Paragraph(p) => walk_inlines(&p.content, colors),
            Block::Table(t) => {
                for row in &t.rows {
                    for cell in &row.cells {
                        walk_blocks(&cell.content, colors);
                    }
                }
            }
            Block::SectionBreak(_) => {}
        }
    }
}

fn walk_inlines(inlines: &[Inline], colors: &mut HashMap<String, RgbColor>) {
    for inline in inlines {
        match inline {
            Inline::TextRun(tr) => {
                if let Some(rev) = &tr.revision {
                    register(&rev.author, colors);
                }
            }
            Inline::Hyperlink(h) => walk_inlines(&h.content, colors),
            Inline::Field(f) => walk_inlines(&f.content, colors),
            _ => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{RevisionKind, RunRevision};

    fn revised_run(author: &str) -> Inline {
        Inline::TextRun(Box::new(crate::model::TextRun {
            style_id: None,
            properties: crate::model::RunProperties::default(),
            content: vec![crate::model::RunElement::Text("x".into())],
            rsids: crate::model::RevisionIds::default(),
            revision: Some(RunRevision {
                kind: RevisionKind::Inserted,
                author: author.into(),
            }),
            comment: None,
        }))
    }

    fn para(inlines: Vec<Inline>) -> Block {
        Block::Paragraph(Box::new(crate::model::Paragraph {
            style_id: None,
            properties: Default::default(),
            mark_run_properties: None,
            content: inlines,
            rsids: Default::default(),
        }))
    }

    /// Authors are numbered by first appearance in document order, so the
    /// same document always colors the same author the same way.
    #[test]
    fn authors_are_colored_by_first_appearance() {
        let blocks = vec![
            para(vec![revised_run("B"), revised_run("A")]),
            para(vec![revised_run("A")]),
        ];
        let colors = collect_revision_colors(&blocks);
        assert_eq!(colors.len(), 2);
        assert_eq!(colors["B"], REVISION_PALETTE[0], "first author seen");
        assert_eq!(colors["A"], REVISION_PALETTE[1]);
    }

    /// More authors than palette entries cycle rather than panic or collide
    /// with `HashMap` iteration order.
    #[test]
    fn the_palette_cycles() {
        let runs: Vec<Inline> = (0..8).map(|i| revised_run(&format!("a{i}"))).collect();
        let colors = collect_revision_colors(&[para(runs)]);
        assert_eq!(colors["a6"], REVISION_PALETTE[0]);
        assert_eq!(colors["a7"], REVISION_PALETTE[1]);
    }
}
