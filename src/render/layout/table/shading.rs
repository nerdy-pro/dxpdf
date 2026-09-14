//! §17.18.78 geometric shading patterns as draw commands (issue #149).
//!
//! A patterned cell is painted as its fill rect (when it has one) plus
//! stripe **lines** in the pattern colour, clipped to the cell box here
//! rather than by a clip state: every stripe reaches the page as an
//! ordinary [`DrawCommand::Line`], so the painter needs no shader machinery
//! and a test can assert the geometry command-by-command — the same
//! reasoning that keeps borders as lines.
//!
//! # Tile geometry — Word 97's 8×8 tiles, in points
//!
//! Word's patterns are its legacy 8×8 1-bit tiles, drawn at screen
//! resolution; the 5th-edition spec's own per-value swatches (Word-rendered
//! bitmaps, corroborated tile-for-tile by the legacy MS Office pattern set
//! preserved in GNOME goffice's `go-pattern.c`) measure out to a period of
//! **4 px in every family**, a thick ("dark") stripe covering 2 px of it
//! and a thin ("light") one 1 px. At 96 dpi one pixel is 0.75 pt, so the
//! constants below: 3 pt pitch, 1.5 / 0.75 pt stripes. A diagonal stripe's
//! 2 px are measured *along a pixel row*, so its stroked (perpendicular)
//! width is that over √2, and its 4-px period likewise lies along the edge
//! — which is what makes the diagonals visually denser than the
//! orthogonals, exactly as the tiles have it.
//!
//! The crosses are emitted as the union of their two constituent stripe
//! families. For three of the four the spec's tiles *decompose* into
//! exactly that union (phase aside, which a repeating fill cannot show);
//! the dark `horzCross` swatch is the recorded anomaly — see
//! [`PatternFamily::HorzCross`].
//!
//! Stripes are emitted only where they fit whole: a stripe whose width
//! would cross the cell edge is dropped rather than half-painted, since a
//! `Line` is stroked symmetrically about its spine and cannot be clipped
//! lengthwise. At most one stripe is lost at each edge, well under a tile.

use crate::render::dimension::Pt;
use crate::render::geometry::{PtLineSegment, PtOffset, PtRect};
use crate::render::layout::draw_command::DrawCommand;
use crate::render::resolve::shading::{PatternFamily, PatternGeometry, ResolvedShading};

/// One Word pattern period: 4 px at 96 dpi.
const TILE: f32 = 3.0;
/// A thick ("dark") stripe is 2 of the period's 4 px.
const THICK: f32 = 1.5;
/// A thin ("light") stripe is 1 px.
const THIN: f32 = 0.75;
/// A diagonal stripe's width is measured along a pixel row, so its stroked
/// width is the orthogonal one over √2.
const SQRT2: f32 = std::f32::consts::SQRT_2;

/// Emit one cell's shading: a flat colour is one rect; a pattern is its
/// background rect (absent for an `auto` fill) plus its stripes.
pub(super) fn emit_cell_shading(
    commands: &mut Vec<DrawCommand>,
    rect: PtRect,
    shading: &ResolvedShading,
) {
    match *shading {
        ResolvedShading::Flat(color) => commands.push(DrawCommand::Rect { rect, color }),
        ResolvedShading::Pattern {
            geometry,
            foreground,
            background,
        } => {
            if let Some(color) = background {
                commands.push(DrawCommand::Rect { rect, color });
            }
            for (line, width) in stripe_lines(rect, geometry) {
                commands.push(DrawCommand::Line {
                    line,
                    color: foreground,
                    width,
                });
            }
        }
    }
}

/// The stripes of one pattern over one box, as `(segment, stroke width)` —
/// every segment lies inside `rect`, pre-clipped.
pub(super) fn stripe_lines(rect: PtRect, geometry: PatternGeometry) -> Vec<(PtLineSegment, Pt)> {
    let width = if geometry.thin { THIN } else { THICK };
    let mut out = Vec::new();
    match geometry.family {
        PatternFamily::Horz => horizontal(rect, width, &mut out),
        PatternFamily::Vert => vertical(rect, width, &mut out),
        PatternFamily::Diag => diagonal(rect, width / SQRT2, false, &mut out),
        PatternFamily::ReverseDiag => diagonal(rect, width / SQRT2, true, &mut out),
        PatternFamily::HorzCross => {
            horizontal(rect, width, &mut out);
            vertical(rect, width, &mut out);
        }
        PatternFamily::DiagCross => {
            diagonal(rect, width / SQRT2, false, &mut out);
            diagonal(rect, width / SQRT2, true, &mut out);
        }
    }
    out
}

fn push(out: &mut Vec<(PtLineSegment, Pt)>, x0: f32, y0: f32, x1: f32, y1: f32, width: f32) {
    out.push((
        PtLineSegment {
            start: PtOffset {
                x: Pt::new(x0),
                y: Pt::new(y0),
            },
            end: PtOffset {
                x: Pt::new(x1),
                y: Pt::new(y1),
            },
        },
        Pt::new(width),
    ));
}

/// Horizontal stripes: spines every [`TILE`] down from the top edge, the
/// first flush against it (its spine half a stripe in).
fn horizontal(rect: PtRect, width: f32, out: &mut Vec<(PtLineSegment, Pt)>) {
    let (left, top) = (rect.origin.x.raw(), rect.origin.y.raw());
    let (w, h) = (rect.size.width.raw(), rect.size.height.raw());
    let mut y = width / 2.0;
    while y + width / 2.0 <= h {
        push(out, left, top + y, left + w, top + y, width);
        y += TILE;
    }
}

/// Vertical stripes: [`horizontal`] with the axes swapped.
fn vertical(rect: PtRect, width: f32, out: &mut Vec<(PtLineSegment, Pt)>) {
    let (left, top) = (rect.origin.x.raw(), rect.origin.y.raw());
    let (w, h) = (rect.size.width.raw(), rect.size.height.raw());
    let mut x = width / 2.0;
    while x + width / 2.0 <= w {
        push(out, left + x, top, left + x, top + h, width);
        x += TILE;
    }
}

/// 45° stripes, one spine per [`TILE`] measured along the top edge.
/// `falling` selects `reverseDiagStripe`'s `\` (y grows with x); otherwise
/// the stripes rise to the right — `diagStripe`'s `/`, emitted as runs from
/// the upper right toward the lower left. See [`PatternFamily`] for how
/// the two names map onto the slopes.
fn diagonal(rect: PtRect, width: f32, falling: bool, out: &mut Vec<(PtLineSegment, Pt)>) {
    let (left, top) = (rect.origin.x.raw(), rect.origin.y.raw());
    let (w, h) = (rect.size.width.raw(), rect.size.height.raw());
    // Inset so a stripe's stroked width stays inside the box (a 45° stroke
    // reaches width/(2√2) into each axis; width/2 is the safe bound). The
    // spine walks the family of lines y = ±x + c with c stepping one tile.
    let inset = width / 2.0;
    let mut c = -(((h - inset) / TILE).floor()) * TILE;
    while c < w - inset {
        let x0 = c.max(inset);
        let y0 = (-c).max(inset);
        let run = (w - inset - x0).min(h - inset - y0);
        if run > 0.0 {
            if falling {
                push(
                    out,
                    left + x0,
                    top + y0,
                    left + x0 + run,
                    top + y0 + run,
                    width,
                );
            } else {
                push(
                    out,
                    left + w - x0,
                    top + y0,
                    left + w - x0 - run,
                    top + y0 + run,
                    width,
                );
            }
        }
        c += TILE;
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::geometry::PtSize;
    use crate::render::resolve::color::RgbColor;

    fn geometry(family: PatternFamily, thin: bool) -> PatternGeometry {
        PatternGeometry { family, thin }
    }

    fn rect(w: f32, h: f32) -> PtRect {
        PtRect {
            origin: PtOffset {
                x: Pt::new(100.0),
                y: Pt::new(50.0),
            },
            size: PtSize {
                width: Pt::new(w),
                height: Pt::new(h),
            },
        }
    }

    /// Every stripe of every family stays inside the box it shades,
    /// **including its stroked width** — the pre-clipping this module
    /// exists for, over a box that cuts both diagonal runs short. A stroke
    /// extends half its width to each side of the spine, perpendicular to
    /// the run.
    #[test]
    fn every_stripe_lies_inside_the_rect() {
        use PatternFamily::*;
        let r = rect(40.0, 13.0);
        for family in [Horz, Vert, Diag, ReverseDiag, HorzCross, DiagCross] {
            for thin in [false, true] {
                for (line, width) in stripe_lines(r, geometry(family, thin)) {
                    // A diagonal stroke reaches width/(2√2) into each axis;
                    // an orthogonal one width/2 along its own. Half in both
                    // axes is the safe bound for every family at once.
                    let half = width.raw() / 2.0;
                    let horizontal = line.start.y == line.end.y;
                    let vertical = line.start.x == line.end.x;
                    let (hx, hy) = match (horizontal, vertical) {
                        (true, false) => (0.0, half),
                        (false, true) => (half, 0.0),
                        _ => (half, half),
                    };
                    for p in [line.start, line.end] {
                        assert!(
                            p.x.raw() - hx >= 100.0 - 1e-4 && p.x.raw() + hx <= 140.0 + 1e-4,
                            "{family:?} thin={thin}: x of {p:?} (width {width:?}) leaves the box"
                        );
                        assert!(
                            p.y.raw() - hy >= 50.0 - 1e-4 && p.y.raw() + hy <= 63.0 + 1e-4,
                            "{family:?} thin={thin}: y of {p:?} (width {width:?}) leaves the box"
                        );
                    }
                }
            }
        }
    }

    /// Horizontal spines sit one [`TILE`] apart, the first flush with the
    /// top edge: thick (1.5 pt) spines in a 13 pt box at 0.75 + n·3 while
    /// the stripe still fits — four of them — and the 0.75 pt thin stripe
    /// fits a fifth, which is the thick/thin distinction doing observable
    /// work beyond the widths themselves.
    #[test]
    fn horizontal_stripes_step_one_tile() {
        let thick = stripe_lines(rect(40.0, 13.0), geometry(PatternFamily::Horz, false));
        let ys: Vec<f32> = thick.iter().map(|(l, _)| l.start.y.raw()).collect();
        assert_eq!(ys, [50.75, 53.75, 56.75, 59.75], "top + width/2 + n·TILE");
        assert!(thick.iter().all(|(l, _)| l.start.y == l.end.y));
        assert!(
            thick
                .iter()
                .all(|(l, _)| (l.end.x - l.start.x).raw() == 40.0),
            "full-width runs"
        );
        assert!(thick.iter().all(|(_, w)| w.raw() == 1.5), "2 px stripes");

        let thin = stripe_lines(rect(40.0, 13.0), geometry(PatternFamily::Horz, true));
        assert_eq!(thin.len(), 5, "a thinner stripe fits once more");
        assert!(thin.iter().all(|(_, w)| w.raw() == 0.75), "1 px stripes");
    }

    /// The two diagonal families mirror each other — same spine count, 45°
    /// both, opposite slopes: `diagStripe` rises to the right (emitted
    /// upper-right → lower-left, so Δx < 0 with Δy > 0),
    /// `reverseDiagStripe` falls (Δx and Δy both positive). Their stroke is
    /// the orthogonal width over √2, the tile's 2 px measured along a row.
    #[test]
    fn diagonals_mirror() {
        let diag = stripe_lines(rect(40.0, 13.0), geometry(PatternFamily::Diag, false));
        let reverse = stripe_lines(
            rect(40.0, 13.0),
            geometry(PatternFamily::ReverseDiag, false),
        );
        assert_eq!(diag.len(), reverse.len());
        assert!(!diag.is_empty());
        for (l, w) in &diag {
            let (dx, dy) = ((l.end.x - l.start.x).raw(), (l.end.y - l.start.y).raw());
            assert!(dx < 0.0 && dy > 0.0, "diagStripe rises to the right: {l:?}");
            assert!((dx + dy).abs() < 1e-4, "…at 45°");
            assert!((w.raw() - 1.5 / SQRT2).abs() < 1e-5);
        }
        for (l, _) in &reverse {
            let (dx, dy) = ((l.end.x - l.start.x).raw(), (l.end.y - l.start.y).raw());
            assert!(
                dx > 0.0 && dy > 0.0,
                "reverseDiagStripe falls to the right: {l:?}"
            );
            assert!((dx - dy).abs() < 1e-4, "…at 45°");
        }
    }

    /// The crosses are exactly their two constituent families' unions —
    /// the reading the spec's own tiles prove for three of the four and
    /// [`PatternFamily::HorzCross`] records for the fourth.
    #[test]
    fn crosses_are_the_union_of_their_families() {
        let r = rect(40.0, 13.0);
        let cross = stripe_lines(r, geometry(PatternFamily::HorzCross, false));
        let horz = stripe_lines(r, geometry(PatternFamily::Horz, false));
        let vert = stripe_lines(r, geometry(PatternFamily::Vert, false));
        assert_eq!(cross.len(), horz.len() + vert.len());

        let dcross = stripe_lines(r, geometry(PatternFamily::DiagCross, false));
        let diag = stripe_lines(r, geometry(PatternFamily::Diag, false));
        assert_eq!(dcross.len(), 2 * diag.len());
    }

    /// A pattern is its background rect first, then only stripes — and no
    /// rect at all over an `auto` fill, which shades like `clear`: stripes
    /// over nothing.
    #[test]
    fn emit_paints_background_then_stripes() {
        let fg = RgbColor { r: 1, g: 2, b: 3 };
        let bg = RgbColor {
            r: 250,
            g: 250,
            b: 250,
        };
        let mut commands = Vec::new();
        emit_cell_shading(
            &mut commands,
            rect(40.0, 13.0),
            &ResolvedShading::Pattern {
                geometry: geometry(PatternFamily::Horz, false),
                foreground: fg,
                background: Some(bg),
            },
        );
        assert!(
            matches!(&commands[0], DrawCommand::Rect { color, .. } if *color == bg),
            "background rect first"
        );
        assert!(commands.len() > 1, "then stripes");
        assert!(commands[1..]
            .iter()
            .all(|c| matches!(c, DrawCommand::Line { color, .. } if *color == fg)));

        let mut transparent = Vec::new();
        emit_cell_shading(
            &mut transparent,
            rect(40.0, 13.0),
            &ResolvedShading::Pattern {
                geometry: geometry(PatternFamily::Horz, false),
                foreground: fg,
                background: None,
            },
        );
        assert!(
            transparent
                .iter()
                .all(|c| matches!(c, DrawCommand::Line { .. })),
            "an auto fill draws stripes over nothing"
        );
    }
}
