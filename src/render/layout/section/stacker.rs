//! Shared block stacker — used by both page-level and cell-level layout.

use super::super::draw_command::DrawCommand;
use super::super::float;
use super::super::paragraph::layout_paragraph;
use super::super::table::layout_table;
use super::helpers::table_x_offset;
use super::types::{FloatingImageY, LayoutBlock};
use crate::render::dimension::Pt;
use crate::render::geometry::PtRect;

/// Result of stacking blocks vertically.
pub struct StackResult {
    /// Draw commands positioned relative to the stacking origin (0,0).
    pub commands: Vec<DrawCommand>,
    /// Total height consumed by all blocks.
    pub height: Pt,
}

/// Stack blocks vertically within a fixed-width area.
///
/// This is the shared core used by both page-level layout (`layout_section`)
/// and cell-level layout. It handles:
/// - Paragraph layout with spacing collapse and space_before suppression
/// - Table layout
/// - Floating image registration and text wrapping
///
/// It does NOT handle page breaks, column breaks, or footnote collection —
/// those are page-level concerns managed by `layout_section`.
pub fn stack_blocks(
    blocks: &[LayoutBlock],
    content_width: Pt,
    default_line_height: Pt,
    measure_text: super::super::paragraph::MeasureTextFn<'_>,
) -> StackResult {
    let constraints = super::super::BoxConstraints::tight_width(content_width, Pt::INFINITY);
    let mut commands = Vec::new();
    let mut cursor_y = Pt::ZERO;
    let mut prev_space_after = Pt::ZERO;
    let mut prev_style_id: Option<crate::model::StyleId> = None;
    let mut page_floats: Vec<float::ActiveFloat> = Vec::new();

    for block in blocks {
        match block {
            LayoutBlock::Paragraph {
                fragments,
                style,
                floating_images,
                floating_shapes,
                ..
            } => {
                let mut effective_style = style.clone_for_layout();

                // Spacing collapse.
                if effective_style.contextual_spacing
                    && effective_style.style_id.is_some()
                    && effective_style.style_id == prev_style_id
                {
                    cursor_y -= prev_space_after + effective_style.space_before;
                } else {
                    let collapse = prev_space_after.min(effective_style.space_before);
                    cursor_y -= collapse;
                }

                // Register floating images.
                let content_top = cursor_y + effective_style.space_before;
                for fi in floating_images.iter() {
                    let (y_start, y_end) = match fi.y {
                        FloatingImageY::RelativeToParagraph(offset) => {
                            (content_top + offset, content_top + offset + fi.size.height)
                        }
                        FloatingImageY::Absolute(img_y) => (img_y, img_y + fi.size.height),
                    };
                    if fi.is_wrap_top_and_bottom() {
                        let img_y = match fi.y {
                            FloatingImageY::Absolute(y) => y,
                            FloatingImageY::RelativeToParagraph(offset) => content_top + offset,
                        };
                        commands.push(DrawCommand::Image {
                            rect: PtRect::from_xywh(fi.x, img_y, fi.size.width, fi.size.height),
                            image_data: fi.image_data.clone(),
                        });
                        if y_end > cursor_y {
                            cursor_y = y_end;
                        }
                    } else {
                        page_floats.push(float::ActiveFloat {
                            page_x: fi.x - fi.dist_left,
                            page_y_start: y_start,
                            page_y_end: y_end,
                            width: fi.size.width + fi.dist_left + fi.dist_right,
                            source: float::FloatSource::Image,
                            wrap_text: fi.wrap_mode.wrap_text().into(),
                        });
                    }
                }

                // §20.4.2: register floating shapes (DrawingML). Mirrors the
                // image branch above: `TopAndBottom` emits now and advances
                // the cursor; wrap-enabled modes (Square/Tight/Through) are
                // registered as active floats so subsequent lines narrow
                // around them. `None` shapes emit after the paragraph.
                for fs in floating_shapes.iter() {
                    use crate::render::layout::section::WrapMode;
                    if matches!(fs.wrap_mode, WrapMode::None) {
                        continue;
                    }
                    let (y_start, y_end) = match fs.y {
                        FloatingImageY::RelativeToParagraph(offset) => {
                            (content_top + offset, content_top + offset + fs.size.height)
                        }
                        FloatingImageY::Absolute(y) => (y, y + fs.size.height),
                    };
                    if fs.is_wrap_top_and_bottom() {
                        let shape_y = match fs.y {
                            FloatingImageY::Absolute(y) => y,
                            FloatingImageY::RelativeToParagraph(offset) => content_top + offset,
                        };
                        commands.push(DrawCommand::Path {
                            origin: crate::render::geometry::PtOffset::new(fs.x, shape_y),
                            rotation: fs.rotation,
                            flip_h: fs.flip_h,
                            flip_v: fs.flip_v,
                            extent: fs.size,
                            paths: fs.paths.clone(),
                            fill: fs.fill.clone(),
                            stroke: fs.stroke.clone(),
                            effects: fs.effects.clone(),
                        });
                        if y_end > cursor_y {
                            cursor_y = y_end;
                        }
                    } else {
                        page_floats.push(float::ActiveFloat {
                            page_x: fs.x - fs.dist_left,
                            page_y_start: y_start,
                            page_y_end: y_end,
                            width: fs.size.width + fs.dist_left + fs.dist_right,
                            source: float::FloatSource::Shape,
                            wrap_text: fs.wrap_mode.wrap_text().into(),
                        });
                    }
                }

                float::prune_floats(&mut page_floats, cursor_y);

                effective_style.page_floats = page_floats.clone();
                effective_style.page_y = cursor_y;
                effective_style.page_x = Pt::ZERO;
                effective_style.page_content_width = content_width;

                let para = layout_paragraph(
                    fragments,
                    &constraints,
                    &effective_style,
                    default_line_height,
                    measure_text,
                );

                for mut cmd in para.commands {
                    cmd.shift_y(cursor_y);
                    commands.push(cmd);
                }

                cursor_y += para.size.height;

                // Emit non-wrapTopAndBottom floating images.
                let para_content_top = cursor_y - para.size.height + effective_style.space_before;
                for fi in floating_images {
                    if fi.is_wrap_top_and_bottom() {
                        continue;
                    }
                    let img_y = match fi.y {
                        FloatingImageY::Absolute(y) => y,
                        FloatingImageY::RelativeToParagraph(offset) => para_content_top + offset,
                    };
                    commands.push(DrawCommand::Image {
                        rect: PtRect::from_xywh(fi.x, img_y, fi.size.width, fi.size.height),
                        image_data: fi.image_data.clone(),
                    });
                    // Extend cursor to encompass the image so table cells
                    // expand to contain floating images.
                    let img_bottom = img_y + fi.size.height;
                    if img_bottom > cursor_y {
                        cursor_y = img_bottom;
                    }
                }

                // Emit floating shapes (DrawingML). `TopAndBottom` shapes
                // were emitted pre-layout along with the `cursor_y` advance;
                // skip them here. `None` + Square/Tight/Through emit now at
                // their resolved anchor position (the shape's bounding rect
                // is already registered as an active float for wrap modes).
                for fs in floating_shapes {
                    if fs.is_wrap_top_and_bottom() {
                        continue;
                    }
                    let shape_y = match fs.y {
                        FloatingImageY::Absolute(y) => y,
                        FloatingImageY::RelativeToParagraph(offset) => para_content_top + offset,
                    };
                    commands.push(DrawCommand::Path {
                        origin: crate::render::geometry::PtOffset::new(fs.x, shape_y),
                        rotation: fs.rotation,
                        flip_h: fs.flip_h,
                        flip_v: fs.flip_v,
                        extent: fs.size,
                        paths: fs.paths.clone(),
                        fill: fs.fill.clone(),
                        stroke: fs.stroke.clone(),
                        effects: fs.effects.clone(),
                    });
                    // §17.17.1 / §20.1.2.1.1: shape's text-box content paints
                    // *over* the shape's fill — emit after the path. Each
                    // command is in shape-local coords; shift by the shape's
                    // resolved page origin.
                    for mut cmd in fs.text_commands.iter().cloned() {
                        cmd.shift(fs.x, shape_y);
                        commands.push(cmd);
                    }
                }

                prev_space_after = effective_style.space_after;
                prev_style_id = effective_style.style_id.clone();
            }
            LayoutBlock::Table {
                rows,
                col_widths,
                border_config,
                indent,
                alignment,
                ..
            } => {
                // stack_blocks is used for table cells and header/footer —
                // no adjacent table collapse in these contexts.
                let table = layout_table(
                    rows,
                    col_widths,
                    &constraints,
                    default_line_height,
                    border_config.as_ref(),
                    measure_text,
                    false,
                );

                let table_x = table_x_offset(
                    *alignment,
                    *indent,
                    table.size.width,
                    content_width,
                    Pt::ZERO,
                );

                for mut cmd in table.commands {
                    cmd.shift_y(cursor_y);
                    cmd.shift_x(table_x);
                    commands.push(cmd);
                }

                cursor_y += table.size.height;
                prev_space_after = Pt::ZERO;
                prev_style_id = None;
            }
        }
    }

    StackResult {
        commands,
        height: cursor_y,
    }
}
