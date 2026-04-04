//! Private helper functions for section layout.

use super::super::draw_command::{DrawCommand, LayoutedPage};
use super::super::fragment::Fragment;
use super::super::page::PageConfig;
use super::super::paragraph::{layout_paragraph, ParagraphStyle};
use super::{FOOTNOTE_SEPARATOR_GAP, FOOTNOTE_SEPARATOR_LINE_WIDTH, FOOTNOTE_SEPARATOR_RATIO};
use crate::render::dimension::Pt;

/// Split a fragment slice at `Fragment::ColumnBreak` markers.
/// Returns a vec of slices; the column break fragments themselves are excluded.
pub(super) fn split_at_column_breaks(fragments: &[Fragment]) -> Vec<&[Fragment]> {
    let has_break = fragments.iter().any(|f| matches!(f, Fragment::ColumnBreak));
    if !has_break {
        return vec![fragments];
    }
    let mut chunks = Vec::new();
    let mut start = 0;
    for (i, frag) in fragments.iter().enumerate() {
        if matches!(frag, Fragment::ColumnBreak) {
            chunks.push(&fragments[start..i]);
            start = i + 1;
        }
    }
    chunks.push(&fragments[start..]);
    chunks
}

pub(super) fn render_page_footnotes(
    page: &mut LayoutedPage,
    config: &PageConfig,
    footnotes: &[(&[Fragment], &ParagraphStyle)],
    default_line_height: Pt,
    measure_text: super::super::paragraph::MeasureTextFn<'_>,
    separator_indent: Pt,
) {
    let content_width = config.content_width();
    let constraints = super::super::BoxConstraints::tight_width(content_width, Pt::INFINITY);
    let page_bottom = config.page_size.height - config.margins.bottom;

    // Layout all footnotes to compute total height.
    let mut footnote_layouts = Vec::new();
    let mut total_height = FOOTNOTE_SEPARATOR_GAP; // separator line + gap above first note
    for (frags, style) in footnotes {
        let para = layout_paragraph(
            frags,
            &constraints,
            style,
            default_line_height,
            measure_text,
        );
        total_height += para.size.height;
        footnote_layouts.push(para);
    }

    let footnote_top = page_bottom - total_height;

    // §17.11.23: separator line positioned per default paragraph indent.
    let sep_x = config.margins.left + separator_indent;
    let sep_width = content_width * FOOTNOTE_SEPARATOR_RATIO;
    page.commands.push(DrawCommand::Line {
        line: crate::render::geometry::PtLineSegment::new(
            crate::render::geometry::PtOffset::new(sep_x, footnote_top),
            crate::render::geometry::PtOffset::new(sep_x + sep_width, footnote_top),
        ),
        color: crate::render::resolve::color::RgbColor::BLACK,
        width: FOOTNOTE_SEPARATOR_LINE_WIDTH,
    });

    // Render footnote paragraphs.
    let mut cursor_y = footnote_top + FOOTNOTE_SEPARATOR_GAP;
    for para in footnote_layouts {
        for mut cmd in para.commands {
            cmd.shift_y(cursor_y);
            cmd.shift_x(config.margins.left);
            page.commands.push(cmd);
        }
        cursor_y += para.size.height;
    }
}

/// §17.4.28: compute the table's x offset based on alignment and indent.
pub(super) fn table_x_offset(
    alignment: Option<crate::model::Alignment>,
    indent: Pt,
    table_width: Pt,
    content_width: Pt,
    margin_left: Pt,
) -> Pt {
    use crate::model::Alignment;
    match alignment {
        Some(Alignment::Center) => margin_left + (content_width - table_width) * 0.5,
        Some(Alignment::End) => margin_left + content_width - table_width,
        _ => margin_left + indent,
    }
}
