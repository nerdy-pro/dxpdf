//! Cell layout — narrows constraints by cell margins, lays out child blocks.
//!
//! Uses the shared `stack_blocks` function from `section.rs` so that table
//! cells get the same features as body content (floating images, spacing
//! collapse, contextual spacing, etc.).

use crate::dimension::Pt;
use crate::geometry::PtEdgeInsets;

use super::section::{LayoutBlock, stack_blocks};

/// Result of laying out a cell.
#[derive(Debug)]
pub struct CellLayout {
    /// Draw commands relative to the cell's top-left origin.
    pub commands: Vec<super::draw_command::DrawCommand>,
    /// Content height (without margins).
    pub content_height: Pt,
}

/// Lay out blocks inside a table cell.
///
/// Receives the full cell width (from column sizing), deflates by margins,
/// lays out each block sequentially using `stack_blocks`, returns total
/// content height.
pub fn layout_cell(
    blocks: &[LayoutBlock],
    cell_width: Pt,
    margins: &PtEdgeInsets,
    default_line_height: Pt,
    measure_text: super::paragraph::MeasureTextFn<'_>,
) -> CellLayout {
    let content_width = (cell_width - margins.horizontal()).max(Pt::ZERO);

    let result = stack_blocks(blocks, content_width, default_line_height, measure_text);

    // Shift all commands by cell margins.
    let commands = result.commands.into_iter().map(|mut cmd| {
        cmd.shift(margins.left, margins.top);
        cmd
    }).collect();

    CellLayout {
        commands,
        content_height: result.height,
    }
}


#[cfg(test)]
mod tests {
    use super::*;
    use crate::layout::draw_command::DrawCommand;
    use crate::layout::fragment::{FontProps, Fragment};
    use crate::layout::paragraph::ParagraphStyle;
    use crate::resolve::color::RgbColor;
    use std::rc::Rc;

    fn text_frag(text: &str, width: f32) -> Fragment {
        Fragment::Text {
            text: text.into(),
            font: FontProps {
                family: Rc::from("Test"),
                size: Pt::new(12.0),
                bold: false,
                italic: false,
                underline: false,
                char_spacing: Pt::ZERO, underline_position: Pt::ZERO, underline_thickness: Pt::ZERO,
            },
            color: RgbColor::BLACK,
            width: Pt::new(width), trimmed_width: Pt::new(width),
            height: Pt::new(14.0),
            ascent: Pt::new(10.0),
            hyperlink_url: None,
            shading: None, border: None, baseline_offset: Pt::ZERO, text_offset: Pt::ZERO,
        }
    }

    fn simple_block(text: &str, width: f32) -> LayoutBlock {
        LayoutBlock::Paragraph {
            fragments: vec![text_frag(text, width)],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
        }
    }

    #[test]
    fn empty_cell_zero_height() {
        let result = layout_cell(
            &[],
            Pt::new(200.0),
            &PtEdgeInsets::ZERO,
            Pt::new(14.0),
            None,
        );
        assert_eq!(result.content_height.raw(), 0.0);
        assert!(result.commands.is_empty());
    }

    #[test]
    fn single_paragraph_in_cell() {
        let blocks = vec![simple_block("hello", 30.0)];
        let result = layout_cell(
            &blocks,
            Pt::new(200.0),
            &PtEdgeInsets::ZERO,
            Pt::new(14.0),
            None,
        );
        assert_eq!(result.content_height.raw(), 14.0);
        assert!(!result.commands.is_empty());
    }

    #[test]
    fn margins_offset_content() {
        let blocks = vec![simple_block("text", 30.0)];
        let margins = PtEdgeInsets::new(
            Pt::new(5.0),   // top
            Pt::new(10.0),  // right
            Pt::new(5.0),   // bottom
            Pt::new(10.0),  // left
        );
        let result = layout_cell(&blocks, Pt::new(200.0), &margins, Pt::new(14.0), None);

        // Text should be shifted right by left margin
        if let Some(DrawCommand::Text { position, .. }) = result.commands.first() {
            assert_eq!(position.x.raw(), 10.0, "left margin applied");
            assert!(position.y.raw() >= 5.0, "top margin applied");
        } else {
            panic!("expected Text command");
        }
    }

    #[test]
    fn margins_narrow_available_width() {
        // Cell is 100 wide, margins eat 60 (left=30, right=30), leaving 40 for content
        // Two fragments of 30 each = 60 > 40, so they should wrap
        let blocks = vec![LayoutBlock::Paragraph {
            fragments: vec![text_frag("aa ", 30.0), text_frag("bb", 30.0)],
            style: ParagraphStyle::default(),
            page_break_before: false,
            footnotes: vec![],
            floating_images: vec![],
        }];
        let margins = PtEdgeInsets::new(Pt::ZERO, Pt::new(30.0), Pt::ZERO, Pt::new(30.0));
        let result = layout_cell(&blocks, Pt::new(100.0), &margins, Pt::new(14.0), None);

        // Should wrap to 2 lines → height = 28
        assert_eq!(result.content_height.raw(), 28.0);
    }

    #[test]
    fn two_paragraphs_stack_vertically() {
        let blocks = vec![
            simple_block("first", 30.0),
            simple_block("second", 40.0),
        ];
        let result = layout_cell(
            &blocks,
            Pt::new(200.0),
            &PtEdgeInsets::ZERO,
            Pt::new(14.0),
            None,
        );
        assert_eq!(result.content_height.raw(), 28.0, "14 + 14");

        let text_cmds: Vec<_> = result
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCommand::Text { position, text, .. } => Some((text.clone(), position.y)),
                _ => None,
            })
            .collect();
        assert_eq!(text_cmds.len(), 2);
        assert!(
            text_cmds[1].1 > text_cmds[0].1,
            "second paragraph should be below first"
        );
    }
}
