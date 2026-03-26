//! Page configuration — convert section properties to layout-ready config.

use dxpdf_docx_model::model::SectionProperties;

use crate::dimension::Pt;
use crate::geometry::{PtEdgeInsets, PtSize};

/// §17.6.13: default page width when w:pgSz is absent. Word uses US Letter (8.5 inches = 12240 twips).
const SPEC_DEFAULT_PAGE_WIDTH: Pt = Pt::new(612.0);
/// §17.6.13: default page height when w:pgSz is absent. Word uses US Letter (11 inches = 15840 twips).
const SPEC_DEFAULT_PAGE_HEIGHT: Pt = Pt::new(792.0);
/// §17.6.11: default page margin when w:pgMar is absent (1 inch = 1440 twips).
const SPEC_DEFAULT_MARGIN: Pt = Pt::new(72.0);

/// Page layout configuration in points.
#[derive(Debug, Clone, Copy)]
pub struct PageConfig {
    pub page_size: PtSize,
    pub margins: PtEdgeInsets,
    pub header_margin: Pt,
    pub footer_margin: Pt,
}

impl Default for PageConfig {
    fn default() -> Self {
        Self {
            page_size: PtSize::new(SPEC_DEFAULT_PAGE_WIDTH, SPEC_DEFAULT_PAGE_HEIGHT),
            margins: PtEdgeInsets::new(SPEC_DEFAULT_MARGIN, SPEC_DEFAULT_MARGIN, SPEC_DEFAULT_MARGIN, SPEC_DEFAULT_MARGIN),
            header_margin: SPEC_DEFAULT_MARGIN / 2.0,
            footer_margin: SPEC_DEFAULT_MARGIN / 2.0,
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

        cfg
    }

    /// Available width for body content (page width minus left and right margins).
    pub fn content_width(&self) -> Pt {
        self.page_size.width - self.margins.left - self.margins.right
    }

    /// Available height for body content (page height minus top and bottom margins).
    pub fn content_height(&self) -> Pt {
        self.page_size.height - self.margins.top - self.margins.bottom
    }
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
