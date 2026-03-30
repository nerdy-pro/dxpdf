//! Draw commands — the output of layout, consumed by the painter.

use std::rc::Rc;

use crate::dimension::Pt;
use crate::geometry::{PtLineSegment, PtOffset, PtRect, PtSize};
use crate::resolve::color::RgbColor;

/// A positioned drawing command — absolute page coordinates.
#[derive(Debug, Clone)]
pub enum DrawCommand {
    Text {
        position: PtOffset,
        text: String,
        font_family: Rc<str>,
        char_spacing: Pt,
        font_size: Pt,
        bold: bool,
        italic: bool,
        color: RgbColor,
    },
    Underline {
        line: PtLineSegment,
        color: RgbColor,
        width: Pt,
    },
    Line {
        line: PtLineSegment,
        color: RgbColor,
        width: Pt,
    },
    Image {
        rect: PtRect,
        image_data: Rc<[u8]>,
    },
    Rect {
        rect: PtRect,
        color: RgbColor,
    },
    LinkAnnotation {
        rect: PtRect,
        url: String,
    },
    /// Internal link to a named destination (bookmark).
    InternalLink {
        rect: PtRect,
        destination: String,
    },
    /// Named destination marker (bookmark target).
    NamedDestination {
        position: PtOffset,
        name: String,
    },
}

impl DrawCommand {
    /// Shift all coordinates by `(dx, dy)`.
    pub fn shift(&mut self, dx: Pt, dy: Pt) {
        match self {
            DrawCommand::Text { position, .. } => {
                position.x += dx;
                position.y += dy;
            }
            DrawCommand::Underline { line, .. } | DrawCommand::Line { line, .. } => {
                line.start.x += dx;
                line.start.y += dy;
                line.end.x += dx;
                line.end.y += dy;
            }
            DrawCommand::Image { rect, .. }
            | DrawCommand::Rect { rect, .. }
            | DrawCommand::LinkAnnotation { rect, .. }
            | DrawCommand::InternalLink { rect, .. } => {
                rect.origin.x += dx;
                rect.origin.y += dy;
            }
            DrawCommand::NamedDestination { position, .. } => {
                position.x += dx;
                position.y += dy;
            }
        }
    }

    /// Shift all y-coordinates by `dy`.
    pub fn shift_y(&mut self, dy: Pt) {
        self.shift(Pt::ZERO, dy);
    }

    /// Shift all x-coordinates by `dx`.
    pub fn shift_x(&mut self, dx: Pt) {
        self.shift(dx, Pt::ZERO);
    }
}

/// A fully laid-out page — ready for painting.
#[derive(Debug, Clone)]
pub struct LayoutedPage {
    pub commands: Vec<DrawCommand>,
    pub page_size: PtSize,
}

impl LayoutedPage {
    pub fn new(page_size: PtSize) -> Self {
        Self {
            commands: Vec::new(),
            page_size,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shift_y_moves_text() {
        let mut cmd = DrawCommand::Text {
            position: PtOffset::new(Pt::new(10.0), Pt::new(20.0)),
            text: "hi".into(),
            font_family: Rc::from("Arial"),
            char_spacing: Pt::ZERO,
            font_size: Pt::new(12.0),
            bold: false,
            italic: false,
            color: RgbColor::BLACK,
        };
        cmd.shift_y(Pt::new(5.0));
        if let DrawCommand::Text { position, .. } = cmd {
            assert_eq!(position.y.raw(), 25.0);
            assert_eq!(position.x.raw(), 10.0); // x unchanged
        }
    }

    #[test]
    fn shift_y_moves_line() {
        let mut cmd = DrawCommand::Line {
            line: PtLineSegment::new(
                PtOffset::new(Pt::new(0.0), Pt::new(10.0)),
                PtOffset::new(Pt::new(100.0), Pt::new(10.0)),
            ),
            color: RgbColor::BLACK,
            width: Pt::new(1.0),
        };
        cmd.shift_y(Pt::new(50.0));
        if let DrawCommand::Line { line, .. } = cmd {
            assert_eq!(line.start.y.raw(), 60.0, "Line y shifted");
        }
    }

    #[test]
    fn shift_y_moves_rect() {
        let mut cmd = DrawCommand::Rect {
            rect: PtRect::from_xywh(Pt::new(0.0), Pt::new(10.0), Pt::new(50.0), Pt::new(20.0)),
            color: RgbColor::BLACK,
        };
        cmd.shift_y(Pt::new(100.0));
        if let DrawCommand::Rect { rect, .. } = cmd {
            assert_eq!(rect.origin.y.raw(), 110.0);
        }
    }

    #[test]
    fn layouted_page_new() {
        let page = LayoutedPage::new(PtSize::new(Pt::new(612.0), Pt::new(792.0)));
        assert!(page.commands.is_empty());
        assert_eq!(page.page_size.width.raw(), 612.0);
    }
}
