//! Image extraction — navigate DrawingML hierarchy to extract image RelIds.

use dxpdf_docx_model::model::{GraphicContent, Image, RelId};

/// Extract the embedded image relationship ID from a DrawingML Image.
/// Navigates: Image → graphic → Picture → blip_fill → blip → embed.
pub fn extract_image_rel_id(image: &Image) -> Option<&RelId> {
    match image.graphic.as_ref()? {
        GraphicContent::Picture(pic) => pic.blip_fill.blip.as_ref()?.embed.as_ref(),
        GraphicContent::WordProcessingShape(_) => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use dxpdf_docx_model::dimension::Dimension;
    use dxpdf_docx_model::geometry::{EdgeInsets, Size};
    use dxpdf_docx_model::model::*;

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
                        embed: Some(RelId::new(rel_id)),
                        link: None,
                        compression: None,
                    }),
                    src_rect: None,
                    stretch: None,
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
                    stretch: None,
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
}
