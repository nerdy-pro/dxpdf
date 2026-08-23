//! Image extraction — navigate DrawingML hierarchy to extract image RelIds.

use std::sync::Arc;

use crate::model::dimension::{Dimension, ThousandthPercent};
use crate::model::{GraphicContent, Image, ImageFormat, RelId, RelativeRect};
use crate::render::dimension::Pt;
use crate::render::geometry::PtRect;

/// Resolved image entry — shared bytes with detected format.
///
/// All references to the same image share one `Arc<[u8]>` allocation — the very
/// one the parser read out of the package, since resolve moves the handle
/// rather than copying the bytes. That identity is load-bearing: the painter's
/// bitmap cache keys on the pointer, so two placements of one image decode once.
#[derive(Clone, Debug)]
pub struct MediaEntry {
    pub data: Arc<[u8]>,
    pub format: ImageFormat,
}

/// Extract the embedded image relationship ID from a DrawingML Image.
/// Navigates: Image → graphic → Picture → blip_fill → blip.
///
/// The `svgBlip` extension outranks the main embed (issue #150):
/// [MS-ODRAWXML] "Pictures" defines the main blip's raster as a copy kept
/// "for backward compatibility" with consumers that cannot draw SVG — this
/// renderer can, so the SVG part is the picture. LibreOffice 24.2 reads the
/// pair the same way. Word always writes the SVG part it references, so a
/// missing part means a package broken in a way no fallback rule saves.
pub fn extract_image_rel_id(image: &Image) -> Option<&RelId> {
    match image.graphic.as_ref()? {
        GraphicContent::Picture(pic) => {
            let blip = pic.blip_fill.blip.as_ref()?;
            blip.svg_embed.as_ref().or(blip.embed.as_ref())
        }
        GraphicContent::WordProcessingShape(_) => None,
    }
}

/// §20.1.10.48 CT_RelativeRect → the visible source region as a fraction rect
/// in `[0, 1]` (origin = top-left crop offset, size = visible fraction).
///
/// The single home for the `1 - l - r` crop math: both the picture blip
/// (`extract_src_rect`) and the shape-fill blip (`ResolvedBlip::src_rect`)
/// paths call it, so cropped pictures and cropped fills can never diverge.
///
/// Returns `None` when there is no crop (all edges zero → draw the whole image)
/// or when the insets overlap so the visible region collapses (`l + r ≥ 1` or
/// `t + b ≥ 1`, malformed → fall back to the whole image rather than a blank
/// `Strict`-sampled frame). Negative insets are preserved as an out-of-`[0, 1]`
/// rect — the painter resolves those into letterbox/pillarbox padding.
pub(crate) fn relative_rect_to_fraction(rel: &RelativeRect) -> Option<PtRect> {
    // CT_RelativeRect edges are in thousandths of a percent (100% = 100000).
    let frac = |d: Option<Dimension<ThousandthPercent>>| d.map_or(0.0, |v| v.to_fraction());
    let (left, top, right, bottom) = (
        frac(rel.left),
        frac(rel.top),
        frac(rel.right),
        frac(rel.bottom),
    );
    if left == 0.0 && top == 0.0 && right == 0.0 && bottom == 0.0 {
        return None;
    }
    let (visible_w, visible_h) = (1.0 - left - right, 1.0 - top - bottom);
    if visible_w <= 0.0 || visible_h <= 0.0 {
        return None;
    }
    Some(PtRect::from_xywh(
        Pt::new(left),
        Pt::new(top),
        Pt::new(visible_w),
        Pt::new(visible_h),
    ))
}

/// §20.1.10.48 `a:srcRect` — the picture's source crop, as a fraction rect in
/// `[0, 1]` relative to the image's natural extent. Returns `None` when there
/// is no crop, so the whole image is drawn. Word crops the source *before*
/// stretching it into the display frame; ignoring this squashes cropped logos
/// (the visible aspect ratio no longer matches the frame). The conversion
/// itself lives in `relative_rect_to_fraction`, shared with the shape-fill
/// blip path (`ResolvedBlip::src_rect`).
pub fn extract_src_rect(image: &Image) -> Option<PtRect> {
    let GraphicContent::Picture(pic) = image.graphic.as_ref()? else {
        return None;
    };
    relative_rect_to_fraction(pic.blip_fill.src_rect.as_ref()?)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::dimension::Dimension;
    use crate::model::geometry::{EdgeInsets, Size};
    use crate::model::*;

    fn make_image_with_blip(rel_id: &str) -> Image {
        Image {
            extent: Size::new(Dimension::new(0), Dimension::new(0)),
            effect_extent: None,
            doc_properties: DocProperties {
                id: 1,
                name: "img".into(),
                description: None,
                hidden: None,
                title: None,
            },
            graphic_frame_locks: None,
            graphic: Some(GraphicContent::Picture(Picture {
                nv_pic_pr: NvPicProperties {
                    cnv_pr: DocProperties {
                        id: 1,
                        name: "pic".into(),
                        description: None,
                        hidden: None,
                        title: None,
                    },
                    cnv_pic_pr: None,
                },
                blip_fill: BlipFill {
                    rotate_with_shape: None,
                    dpi: None,
                    blip: Some(Blip {
                        svg_embed: None,
                        embed: Some(RelId::new(rel_id)),
                        link: None,
                        compression: None,
                    }),
                    src_rect: None,
                    fill_kind: BlipFillKind::Unspecified,
                },
                shape_properties: None,
            })),
            placement: ImagePlacement::Inline {
                distance: EdgeInsets::new(
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                ),
            },
        }
    }

    #[test]
    fn extracts_rel_id_from_picture_blip() {
        let img = make_image_with_blip("rId5");
        let rel_id = extract_image_rel_id(&img);
        assert_eq!(rel_id.map(|r| r.as_str()), Some("rId5"));
    }

    #[test]
    fn no_graphic_returns_none() {
        let img = Image {
            extent: Size::new(Dimension::new(0), Dimension::new(0)),
            effect_extent: None,
            doc_properties: DocProperties {
                id: 1,
                name: "img".into(),
                description: None,
                hidden: None,
                title: None,
            },
            graphic_frame_locks: None,
            graphic: None,
            placement: ImagePlacement::Inline {
                distance: EdgeInsets::new(
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                ),
            },
        };
        assert!(extract_image_rel_id(&img).is_none());
    }

    #[test]
    fn no_blip_returns_none() {
        let img = Image {
            extent: Size::new(Dimension::new(0), Dimension::new(0)),
            effect_extent: None,
            doc_properties: DocProperties {
                id: 1,
                name: "img".into(),
                description: None,
                hidden: None,
                title: None,
            },
            graphic_frame_locks: None,
            graphic: Some(GraphicContent::Picture(Picture {
                nv_pic_pr: NvPicProperties {
                    cnv_pr: DocProperties {
                        id: 1,
                        name: "pic".into(),
                        description: None,
                        hidden: None,
                        title: None,
                    },
                    cnv_pic_pr: None,
                },
                blip_fill: BlipFill {
                    rotate_with_shape: None,
                    dpi: None,
                    blip: None,
                    src_rect: None,
                    fill_kind: BlipFillKind::Unspecified,
                },
                shape_properties: None,
            })),
            placement: ImagePlacement::Inline {
                distance: EdgeInsets::new(
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                ),
            },
        };
        assert!(extract_image_rel_id(&img).is_none());
    }

    #[test]
    fn word_processing_shape_returns_none() {
        let img = Image {
            extent: Size::new(Dimension::new(0), Dimension::new(0)),
            effect_extent: None,
            doc_properties: DocProperties {
                id: 1,
                name: "img".into(),
                description: None,
                hidden: None,
                title: None,
            },
            graphic_frame_locks: None,
            graphic: Some(GraphicContent::WordProcessingShape(WordProcessingShape {
                cnv_pr: None,
                shape_properties: None,
                style_line_ref: None,
                style_effect_ref: None,
                style_fill_ref: None,
                style_font_ref: None,
                body_pr: None,
                txbx_content: vec![],
            })),
            placement: ImagePlacement::Inline {
                distance: EdgeInsets::new(
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                    Dimension::new(0),
                ),
            },
        };
        assert!(extract_image_rel_id(&img).is_none());
    }

    #[test]
    fn src_rect_none_when_no_crop() {
        // `make_image_with_blip` leaves `blip_fill.src_rect = None`.
        assert!(extract_src_rect(&make_image_with_blip("rId1")).is_none());
    }

    #[test]
    fn src_rect_converts_thousandth_percent_crop() {
        // Real values from the IP-05 header logo: top/right/bottom crop, no
        // left. Thousandths of a percent → fractions of the natural extent.
        let mut img = make_image_with_blip("rId1");
        if let Some(GraphicContent::Picture(pic)) = img.graphic.as_mut() {
            pic.blip_fill.src_rect = Some(RelativeRect {
                left: None,
                top: Some(Dimension::new(15883)),
                right: Some(Dimension::new(14520)),
                bottom: Some(Dimension::new(17647)),
            });
        }
        let r = extract_src_rect(&img).expect("crop present");
        assert!((r.origin.x.raw() - 0.0).abs() < 1e-4, "left crop = 0");
        assert!((r.origin.y.raw() - 0.15883).abs() < 1e-4, "top crop");
        assert!((r.size.width.raw() - 0.85480).abs() < 1e-4, "visible width");
        assert!(
            (r.size.height.raw() - 0.66470).abs() < 1e-4,
            "visible height"
        );
    }

    #[test]
    fn src_rect_none_when_horizontal_crop_collapses_region() {
        // l + r ≥ 1: the insets overlap, so no visible width remains. Malformed
        // input must fall back to the whole image (None), not a blank frame.
        let mut img = make_image_with_blip("rId1");
        if let Some(GraphicContent::Picture(pic)) = img.graphic.as_mut() {
            pic.blip_fill.src_rect = Some(RelativeRect {
                left: Some(Dimension::new(60000)),
                top: None,
                right: Some(Dimension::new(60000)),
                bottom: None,
            });
        }
        assert!(extract_src_rect(&img).is_none());
    }

    #[test]
    fn src_rect_none_when_vertical_crop_collapses_region() {
        // t + b ≥ 1: no visible height remains.
        let mut img = make_image_with_blip("rId1");
        if let Some(GraphicContent::Picture(pic)) = img.graphic.as_mut() {
            pic.blip_fill.src_rect = Some(RelativeRect {
                left: None,
                top: Some(Dimension::new(100000)),
                right: None,
                bottom: Some(Dimension::new(5000)),
            });
        }
        assert!(extract_src_rect(&img).is_none());
    }

    #[test]
    fn relative_rect_to_fraction_is_the_shared_converter() {
        // Direct test of the seam both blip paths call: 25% left / 10% bottom
        // crop → origin (0.25, 0) and visible (0.65, 0.90).
        let rel = RelativeRect {
            left: Some(Dimension::new(25000)),
            top: None,
            right: Some(Dimension::new(10000)),
            bottom: Some(Dimension::new(10000)),
        };
        let r = relative_rect_to_fraction(&rel).expect("crop present");
        assert!((r.origin.x.raw() - 0.25).abs() < 1e-5);
        assert!((r.size.width.raw() - 0.65).abs() < 1e-5);
        assert!((r.size.height.raw() - 0.90).abs() < 1e-5);
    }

    #[test]
    fn relative_rect_to_fraction_preserves_negative_insets_for_padding() {
        // Negative insets (letterbox padding) stay as an out-of-[0,1] rect; the
        // painter turns that into padding rather than the converter clamping it.
        let rel = RelativeRect {
            left: Some(Dimension::new(-20000)),
            top: None,
            right: Some(Dimension::new(-20000)),
            bottom: None,
        };
        let r = relative_rect_to_fraction(&rel).expect("crop present");
        assert!(
            (r.origin.x.raw() + 0.2).abs() < 1e-5,
            "negative origin kept"
        );
        assert!((r.size.width.raw() - 1.4).abs() < 1e-5, "width > 1 kept");
    }
}
