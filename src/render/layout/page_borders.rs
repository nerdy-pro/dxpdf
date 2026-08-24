//! §17.6.10 `<w:pgBorders>` — borders around each page of a section.
//!
//! Runs in the second layout pass, beside header/footer rendering: that is
//! the only point where a finished [`LayoutedPage`] and its section's
//! [`PageConfig`] coexist, and both halves of the geometry need them — the
//! per-page command list for `@w:zOrder`, the margins for
//! `@w:offsetFrom="text"`.
//!
//! What is honoured: the four line edges (through the same
//! `convert_model_border` approximation the tables use, so `double` draws its
//! §17.4.38 two-thirds band and the other 24 line styles come out solid, warn
//! once), `@w:offsetFrom`, `@w:display`, `@w:zOrder`, per-edge `@w:space` and
//! colours. What is declined: **art borders** draw nothing — the same reading
//! LibreOffice takes — and are reported once per render; `@w:shadow` and
//! `@w:frame` are accepted by the parser and not drawn;
//! `w:bordersDoNotSurroundHeader`/`Footer` (§17.15.1.7/.8) are not read, so a
//! `text`-offset border always takes the configured text margins as its
//! reference. The margins are the *configured* ones deliberately: a tall
//! header pushes the body's content down per page, but Word measures the
//! border from the page's declared margin, not from where the pushed content
//! landed.
//!
//! A page shared across a §17.6.22 continuous break takes the **owning**
//! (succeeding) section's borders — the same ownership rule the header
//! selection already applies to that page, and for the same reason: the page
//! belongs to two sections and nothing in §17.6 resolves which frame it
//! gets. The corollary is that a section whose every page is shared away
//! contributes no borders at all. **Word reference render**: a render
//! showing the preceding section's frame (or both) winning would move the
//! ownership rule in `render/mod.rs`, not this module.
//!
//! Geometry, from the ECMA-376 §17.6.2–§17.6.10 prose and the Word↔Writer
//! interop notes (`editeng::BorderDistanceFromWord`): `space` is in points and
//! measures to the border edge **nearest the reference** — from the page edge
//! to the border's outer edge (`offsetFrom="page"`, line growing inward), or
//! from the text margin to the border's inner edge (`offsetFrom="text"`, line
//! growing outward). Horizontal edges own the corner squares and span between
//! the vertical edges' outer x (or the reference frame where a vertical edge
//! is absent); vertical edges run between the horizontal bands, the same
//! junction convention the table painter uses.

use crate::model::{self, PageBorderDisplay, PageBorderEdge, PageBorderOffset, PageBorderZOrder};
use crate::render::dimension::Pt;
use crate::render::geometry::PtRect;

use super::build::convert::convert_model_border;
use super::build::BuildState;
use super::draw_command::LayoutedPage;
use super::page::PageConfig;
use super::table::{emit_border_rect, TableBorderLine};

/// Draw a section's page borders onto its pages.
///
/// `pages` is the section's own slice; `first_page_index` is the section
/// page index of its first element — 0 for the section's own range, and the
/// section's page count for the endnote page appended after it, which
/// §17.11.2 makes a *continuation* of the last section's flow (so
/// `display="firstPage"` must not re-fire there, and `"notFirstPage"` must).
pub fn render_page_borders(
    pages: &mut [LayoutedPage],
    first_page_index: usize,
    config: &PageConfig,
    borders: &model::PageBorders,
    state: &mut BuildState,
) {
    let display = borders.display.unwrap_or(PageBorderDisplay::AllPages);
    let frame = match resolve_frame(config, borders, state) {
        Some(frame) => frame,
        None => return,
    };

    for (page_idx, page) in pages.iter_mut().enumerate() {
        let page_idx = first_page_index + page_idx;
        let shown = match display {
            PageBorderDisplay::AllPages => true,
            PageBorderDisplay::FirstPage => page_idx == 0,
            PageBorderDisplay::NotFirstPage => page_idx > 0,
        };
        if !shown {
            continue;
        }

        let mut commands = Vec::new();
        if let Some((line, rect)) = &frame.top {
            emit_border_rect(&mut commands, line, *rect, true);
        }
        if let Some((line, rect)) = &frame.bottom {
            emit_border_rect(&mut commands, line, *rect, true);
        }
        if let Some((line, rect)) = &frame.left {
            emit_border_rect(&mut commands, line, *rect, false);
        }
        if let Some((line, rect)) = &frame.right {
            emit_border_rect(&mut commands, line, *rect, false);
        }

        // §17.18.67: painting order *is* z-order — the painter replays the
        // command list front to back. `front` goes after everything on the
        // page (headers, body, footers — all merged by the time this runs);
        // `back` before it.
        match borders.z_order.unwrap_or(PageBorderZOrder::Front) {
            PageBorderZOrder::Front => page.commands.extend(commands),
            PageBorderZOrder::Back => {
                commands.append(&mut page.commands);
                page.commands = commands;
            }
        }
    }
}

/// The four resolved edges with their band rectangles, shared by every page
/// of the section that displays them.
struct BorderFrame {
    top: Option<(TableBorderLine, PtRect)>,
    bottom: Option<(TableBorderLine, PtRect)>,
    left: Option<(TableBorderLine, PtRect)>,
    right: Option<(TableBorderLine, PtRect)>,
}

/// One edge's resolved line and its position along its axis (`near` = the
/// coordinate nearer the page's top-left, `far` = near + width).
struct EdgeBand {
    line: TableBorderLine,
    near: Pt,
    far: Pt,
}

fn resolve_frame(
    config: &PageConfig,
    borders: &model::PageBorders,
    state: &mut BuildState,
) -> Option<BorderFrame> {
    // [MS-OE376] §2.6.10: Word defaults an absent offsetFrom to `text`.
    let offset = borders.offset_from.unwrap_or(PageBorderOffset::Text);
    let page_w = config.page_size.width;
    let page_h = config.page_size.height;

    // Each edge's band along its own axis. `at_far_end` picks which end of
    // the axis the edge hugs and which direction the line grows.
    let mut band = |edge: &Option<PageBorderEdge>, extent: Pt, margin: Pt, at_far_end: bool| {
        let border = resolve_edge(edge, state)?;
        let width = border.width;
        if width <= Pt::ZERO {
            return None;
        }
        let space = edge_space(edge);
        let near = match (offset, at_far_end) {
            // From the page edge, growing inward.
            (PageBorderOffset::Page, false) => space,
            (PageBorderOffset::Page, true) => extent - space - width,
            // From the text margin, growing outward (toward the page edge).
            (PageBorderOffset::Text, false) => margin - space - width,
            (PageBorderOffset::Text, true) => extent - margin + space,
        };
        Some(EdgeBand {
            line: border,
            near,
            far: near + width,
        })
    };

    let top = band(&borders.top, page_h, config.margins.top, false);
    let bottom = band(&borders.bottom, page_h, config.margins.bottom, true);
    let left = band(&borders.left, page_w, config.margins.left, false);
    let right = band(&borders.right, page_w, config.margins.right, true);
    if top.is_none() && bottom.is_none() && left.is_none() && right.is_none() {
        return None;
    }

    // §17.6.2: where a perpendicular edge exists, an edge runs to meet it;
    // where it does not, `page` offsets run to the page edge and `text`
    // offsets to the text extent. Horizontal edges take the corner squares;
    // vertical edges run between the horizontal bands (the table painter's
    // junction convention, so a `double` corner keeps its gap).
    let (default_x0, default_x1, default_y0, default_y1) = match offset {
        PageBorderOffset::Page => (Pt::ZERO, page_w, Pt::ZERO, page_h),
        PageBorderOffset::Text => (
            config.margins.left,
            page_w - config.margins.right,
            config.margins.top,
            page_h - config.margins.bottom,
        ),
    };
    let x0 = left.as_ref().map_or(default_x0, |b| b.near);
    let x1 = right.as_ref().map_or(default_x1, |b| b.far);
    let y0 = top.as_ref().map_or(default_y0, |b| b.far);
    let y1 = bottom.as_ref().map_or(default_y1, |b| b.near);

    let clamp = |rect: PtRect| -> Option<(PtRect, ())> {
        (rect.size.width > Pt::ZERO && rect.size.height > Pt::ZERO).then_some((rect, ()))
    };
    let horizontal = |b: EdgeBand| {
        let (rect, ()) = clamp(PtRect::from_xywh(x0, b.near, x1 - x0, b.far - b.near))?;
        Some((b.line, rect))
    };
    let vertical = |b: EdgeBand| {
        let (rect, ()) = clamp(PtRect::from_xywh(b.near, y0, b.far - b.near, y1 - y0))?;
        Some((b.line, rect))
    };

    Some(BorderFrame {
        top: top.and_then(horizontal),
        bottom: bottom.and_then(horizontal),
        left: left.and_then(vertical),
        right: right.and_then(vertical),
    })
}

/// A line edge resolves through the table machinery (style approximation,
/// `auto` colour, warn-once); a `nil`/`none` edge and an absent one draw
/// nothing; an art edge is declined, once per render.
fn resolve_edge(edge: &Option<PageBorderEdge>, state: &mut BuildState) -> Option<TableBorderLine> {
    match edge {
        None => None,
        Some(PageBorderEdge::Line(border)) => {
            if border.style.draws_nothing() {
                return None;
            }
            Some(convert_model_border(border, state))
        }
        Some(PageBorderEdge::Art { name, .. }) => {
            if !state.warned_art_page_border {
                state.warned_art_page_border = true;
                log::warn!(
                    "[layout] §17.6.10 art page border {name:?} is not drawn \
                     (art borders are declined, matching LibreOffice)"
                );
            }
            None
        }
    }
}

fn edge_space(edge: &Option<PageBorderEdge>) -> Pt {
    match edge {
        Some(PageBorderEdge::Line(border)) => Pt::from(border.space),
        Some(PageBorderEdge::Art { space, .. }) => space.map(Pt::from).unwrap_or(Pt::ZERO),
        None => Pt::ZERO,
    }
}
