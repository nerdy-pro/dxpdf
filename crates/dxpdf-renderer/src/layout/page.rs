//! Page configuration — convert section properties to layout-ready config.

use dxpdf_docx_model::model::{Columns, SectionProperties};

use crate::dimension::Pt;
use crate::geometry::{PtEdgeInsets, PtSize};

/// §17.6.13: default page width when w:pgSz is absent. Word uses US Letter (8.5 inches = 12240 twips).
const SPEC_DEFAULT_PAGE_WIDTH: Pt = Pt::new(612.0);
/// §17.6.13: default page height when w:pgSz is absent. Word uses US Letter (11 inches = 15840 twips).
const SPEC_DEFAULT_PAGE_HEIGHT: Pt = Pt::new(792.0);
/// §17.6.11: default page margin when w:pgMar is absent (1 inch = 1440 twips).
const SPEC_DEFAULT_MARGIN: Pt = Pt::new(72.0);

/// A single column's layout geometry.
#[derive(Debug, Clone, Copy)]
pub struct ColumnGeometry {
    /// X offset of this column relative to the left page margin.
    pub x_offset: Pt,
    /// Available text width within this column.
    pub width: Pt,
}

/// Page layout configuration in points.
#[derive(Debug, Clone)]
pub struct PageConfig {
    pub page_size: PtSize,
    pub margins: PtEdgeInsets,
    pub header_margin: Pt,
    pub footer_margin: Pt,
    /// §17.6.4: column layout. Single-element vec for normal single-column.
    pub columns: Vec<ColumnGeometry>,
}

impl Default for PageConfig {
    fn default() -> Self {
        let content_width = SPEC_DEFAULT_PAGE_WIDTH - SPEC_DEFAULT_MARGIN - SPEC_DEFAULT_MARGIN;
        Self {
            page_size: PtSize::new(SPEC_DEFAULT_PAGE_WIDTH, SPEC_DEFAULT_PAGE_HEIGHT),
            margins: PtEdgeInsets::new(SPEC_DEFAULT_MARGIN, SPEC_DEFAULT_MARGIN, SPEC_DEFAULT_MARGIN, SPEC_DEFAULT_MARGIN),
            header_margin: SPEC_DEFAULT_MARGIN / 2.0,
            footer_margin: SPEC_DEFAULT_MARGIN / 2.0,
            columns: vec![ColumnGeometry { x_offset: Pt::ZERO, width: content_width }],
        }
    }
}

impl PageConfig {
    /// Build from section properties, falling back to US Letter defaults.
    pub fn from_section(sect: &SectionProperties) -> Self {
        let mut cfg = Self::default();

        if let Some(ref ps) = sect.page_size {
            if let Some(w) = ps.width {
                cfg.page_size.width = Pt::from(w);
            }
            if let Some(h) = ps.height {
                cfg.page_size.height = Pt::from(h);
            }
        }

        if let Some(ref pm) = sect.page_margins {
            if let Some(t) = pm.top {
                cfg.margins.top = Pt::from(t);
            }
            if let Some(r) = pm.right {
                cfg.margins.right = Pt::from(r);
            }
            if let Some(b) = pm.bottom {
                cfg.margins.bottom = Pt::from(b);
            }
            if let Some(l) = pm.left {
                cfg.margins.left = Pt::from(l);
            }
            if let Some(h) = pm.header {
                cfg.header_margin = Pt::from(h);
            }
            if let Some(f) = pm.footer {
                cfg.footer_margin = Pt::from(f);
            }
        }

        // §17.6.4: compute column geometry.
        let content_width = cfg.page_size.width - cfg.margins.left - cfg.margins.right;
        cfg.columns = compute_columns(content_width, &sect.columns);

        cfg
    }

    /// Available width for body content (page width minus left and right margins).
    pub fn content_width(&self) -> Pt {
        self.page_size.width - self.margins.left - self.margins.right
    }

    /// Number of columns in this section.
    pub fn num_columns(&self) -> usize {
        self.columns.len()
    }

    /// Available height for body content (page height minus top and bottom margins).
    pub fn content_height(&self) -> Pt {
        self.page_size.height - self.margins.top - self.margins.bottom
    }
}

/// §17.6.4: compute column geometry from section column properties.
fn compute_columns(content_width: Pt, columns: &Option<Columns>) -> Vec<ColumnGeometry> {
    let cols = match columns {
        Some(c) if c.count.unwrap_or(1) > 1 => c,
        _ => return vec![ColumnGeometry { x_offset: Pt::ZERO, width: content_width }],
    };

    let num = cols.count.unwrap_or(1) as usize;
    let default_space = cols.space.map(Pt::from).unwrap_or(Pt::new(36.0)); // 720tw = 0.5in

    // Use individual column definitions if provided and not equal_width.
    if !cols.columns.is_empty() && cols.equal_width != Some(true) {
        let mut result = Vec::with_capacity(cols.columns.len());
        let mut x = Pt::ZERO;
        for (i, col_def) in cols.columns.iter().enumerate() {
            let w = col_def.width.map(Pt::from).unwrap_or(content_width / num as f32);
            result.push(ColumnGeometry { x_offset: x, width: w });
            if i < cols.columns.len() - 1 {
                let gap = col_def.space.map(Pt::from).unwrap_or(default_space);
                x += w + gap;
            }
        }
        return result;
    }

    // Equal-width columns.
    let total_gap = default_space * (num as f32 - 1.0);
    let col_width = (content_width - total_gap) / num as f32;
    let mut result = Vec::with_capacity(num);
    for i in 0..num {
        let x = (col_width + default_space) * i as f32;
        result.push(ColumnGeometry { x_offset: x, width: col_width });
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use dxpdf_docx_model::dimension::{Dimension, Twips};
    use dxpdf_docx_model::model::{PageMargins, PageSize};

    #[test]
    fn default_is_us_letter() {
        let cfg = PageConfig::default();
        assert_eq!(cfg.page_size.width.raw(), 612.0);
        assert_eq!(cfg.page_size.height.raw(), 792.0);
        assert_eq!(cfg.margins.top.raw(), 72.0);
    }

    #[test]
    fn content_dimensions() {
        let cfg = PageConfig::default();
        assert_eq!(cfg.content_width().raw(), 468.0);  // 612 - 72 - 72
        assert_eq!(cfg.content_height().raw(), 648.0);  // 792 - 72 - 72
    }

    #[test]
    fn from_section_with_page_size() {
        let sect = SectionProperties {
            page_size: Some(PageSize {
                width: Some(Dimension::<Twips>::new(12240)),  // 8.5in = 612pt
                height: Some(Dimension::<Twips>::new(15840)), // 11in = 792pt
                orientation: None,
            }),
            ..Default::default()
        };
        let cfg = PageConfig::from_section(&sect);
        assert_eq!(cfg.page_size.width.raw(), 612.0);
        assert_eq!(cfg.page_size.height.raw(), 792.0);
    }

    #[test]
    fn from_section_with_margins() {
        let sect = SectionProperties {
            page_margins: Some(PageMargins {
                top: Some(Dimension::<Twips>::new(1440)),    // 1in = 72pt
                right: Some(Dimension::<Twips>::new(1440)),
                bottom: Some(Dimension::<Twips>::new(1440)),
                left: Some(Dimension::<Twips>::new(1440)),
                header: Some(Dimension::<Twips>::new(720)),  // 0.5in = 36pt
                footer: Some(Dimension::<Twips>::new(720)),
                gutter: None,
            }),
            ..Default::default()
        };
        let cfg = PageConfig::from_section(&sect);
        assert_eq!(cfg.margins.top.raw(), 72.0);
        assert_eq!(cfg.header_margin.raw(), 36.0);
    }

    #[test]
    fn from_section_partial_uses_defaults() {
        let sect = SectionProperties {
            page_margins: Some(PageMargins {
                top: Some(Dimension::<Twips>::new(2880)), // 2in = 144pt
                right: None,
                bottom: None,
                left: None,
                header: None,
                footer: None,
                gutter: None,
            }),
            ..Default::default()
        };
        let cfg = PageConfig::from_section(&sect);
        assert_eq!(cfg.margins.top.raw(), 144.0, "custom top margin");
        assert_eq!(cfg.margins.right.raw(), 72.0, "default right margin");
    }
}
