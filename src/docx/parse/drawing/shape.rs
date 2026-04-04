use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::dimension::Dimension;
use crate::docx::error::{ParseError, Result};
use crate::docx::geometry::Offset;
use crate::docx::model::*;
use crate::docx::xml;

use super::{parse_cnv_pr, parse_cnv_pr_attrs, parse_positive_size_2d};

// ── wps:wsp (§14.5) ────────────────────────────────────────────────────────

/// Parse `wps:wsp` — Word Processing Shape.
pub(super) fn parse_word_processing_shape(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<WordProcessingShape> {
    let mut cnv_pr = None;
    let mut shape_properties = None;
    let mut body_pr = None;
    let mut txbx_content = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"cNvPr" => {
                        cnv_pr = Some(parse_cnv_pr(e, reader, buf)?);
                    }
                    b"spPr" => {
                        shape_properties = Some(parse_shape_properties(e, reader, buf)?);
                    }
                    b"bodyPr" => {
                        body_pr = Some(parse_body_properties(e, reader, buf)?);
                    }
                    b"txbx" => {
                        txbx_content = parse_wsp_txbx(reader, buf)?;
                    }
                    _ => {
                        xml::warn_unsupported_element("wsp", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"cNvPr" => {
                        cnv_pr = Some(parse_cnv_pr_attrs(e)?);
                    }
                    b"cNvSpPr" | b"cNvCnPr" => {} // non-visual shape props — attrs only
                    b"bodyPr" => {
                        body_pr = Some(parse_body_properties_attrs(e)?);
                    }
                    _ => xml::warn_unsupported_element("wsp", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"wsp" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"wsp")),
            _ => {}
        }
    }

    Ok(WordProcessingShape {
        cnv_pr,
        shape_properties,
        body_pr,
        txbx_content,
    })
}

/// Parse `wps:txbx` — contains `w:txbxContent`.
fn parse_wsp_txbx(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Vec<Block>> {
    let mut blocks = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"txbxContent" => {
                        let (content, _) = crate::docx::parse::body::parse_block_content(
                            reader,
                            buf,
                            b"txbxContent",
                        )?;
                        blocks = content;
                    }
                    _ => {
                        xml::warn_unsupported_element("txbx", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"txbx" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"txbx")),
            _ => {}
        }
    }

    Ok(blocks)
}

/// §20.1.2.1.1: parse `a:bodyPr` attributes only.
fn parse_body_properties_attrs(e: &BytesStart<'_>) -> Result<BodyProperties> {
    Ok(BodyProperties {
        rotation: xml::optional_attr_i64(e, b"rot")?.map(Dimension::new),
        vert: xml::optional_attr(e, b"vert")?
            .map(|s| parse_text_vertical_type(&s))
            .transpose()?,
        wrap: xml::optional_attr(e, b"wrap")?
            .map(|s| parse_text_wrapping_type(&s))
            .transpose()?,
        left_inset: xml::optional_attr_i64(e, b"lIns")?.map(Dimension::new),
        top_inset: xml::optional_attr_i64(e, b"tIns")?.map(Dimension::new),
        right_inset: xml::optional_attr_i64(e, b"rIns")?.map(Dimension::new),
        bottom_inset: xml::optional_attr_i64(e, b"bIns")?.map(Dimension::new),
        anchor: xml::optional_attr(e, b"anchor")?
            .map(|s| parse_text_anchoring_type(&s))
            .transpose()?,
        auto_fit: None,
    })
}

/// §20.1.10.82 ST_TextVerticalType.
fn parse_text_vertical_type(val: &str) -> Result<TextVerticalType> {
    match val {
        "horz" => Ok(TextVerticalType::Horz),
        "vert" => Ok(TextVerticalType::Vert),
        "vert270" => Ok(TextVerticalType::Vert270),
        "wordArtVert" => Ok(TextVerticalType::WordArtVert),
        "eaVert" => Ok(TextVerticalType::EaVert),
        "mongolianVert" => Ok(TextVerticalType::MongolianVert),
        "wordArtVertRtl" => Ok(TextVerticalType::WordArtVertRtl),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "vert".into(),
            value: other.into(),
            reason: "expected value per §20.1.10.82 ST_TextVerticalType".into(),
        }),
    }
}

/// §20.1.10.85 ST_TextWrappingType.
fn parse_text_wrapping_type(val: &str) -> Result<TextWrappingType> {
    match val {
        "none" => Ok(TextWrappingType::None),
        "square" => Ok(TextWrappingType::Square),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "wrap".into(),
            value: other.into(),
            reason: "expected none or square per §20.1.10.85 ST_TextWrappingType".into(),
        }),
    }
}

/// §20.1.10.59 ST_TextAnchoringType.
fn parse_text_anchoring_type(val: &str) -> Result<TextAnchoringType> {
    match val {
        "t" => Ok(TextAnchoringType::Top),
        "ctr" => Ok(TextAnchoringType::Center),
        "b" => Ok(TextAnchoringType::Bottom),
        "just" => Ok(TextAnchoringType::Justified),
        "dist" => Ok(TextAnchoringType::Distributed),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "anchor".into(),
            value: other.into(),
            reason: "expected value per §20.1.10.59 ST_TextAnchoringType".into(),
        }),
    }
}

/// §20.1.2.1.1: parse `a:bodyPr` with children.
fn parse_body_properties(
    e: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<BodyProperties> {
    let mut bp = parse_body_properties_attrs(e)?;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"noAutofit" => bp.auto_fit = Some(TextAutoFit::NoAutoFit),
                    b"normAutofit" => bp.auto_fit = Some(TextAutoFit::NormalAutoFit),
                    b"spAutoFit" => bp.auto_fit = Some(TextAutoFit::SpAutoFit),
                    _ => {} // skip prstTxWarp, scene3d, etc.
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"bodyPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"bodyPr")),
            _ => {}
        }
    }

    Ok(bp)
}

// ── pic:spPr (§20.1.2.2.35) ────────────────────────────────────────────────

pub(super) fn parse_shape_properties(
    start: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<ShapeProperties> {
    let bw_mode = xml::optional_attr(start, b"bwMode")?
        .map(|s| parse_black_white_mode(&s))
        .transpose()?;
    let mut transform = None;
    let mut preset_geometry = None;
    let mut fill = None;
    let mut outline = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"xfrm" => {
                        transform = Some(parse_transform_2d(e, reader, buf)?);
                    }
                    b"prstGeom" => {
                        preset_geometry = Some(parse_preset_geometry(e, reader, buf)?);
                    }
                    b"ln" => {
                        outline = Some(parse_outline(e, reader, buf)?);
                    }
                    b"custGeom" => {
                        xml::warn_unsupported_element("spPr", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                    // Fill group — complex, skip children for now.
                    b"solidFill" | b"gradFill" | b"blipFill" | b"pattFill" | b"grpFill" => {
                        xml::warn_unsupported_element("spPr", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                    b"effectLst" | b"effectDag" | b"scene3d" | b"sp3d" | b"extLst" => {
                        xml::skip_to_end(reader, buf, local)?;
                    }
                    _ => {
                        xml::warn_unsupported_element("spPr", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"noFill" => {
                        fill = Some(DrawingFill::NoFill);
                    }
                    _ => xml::warn_unsupported_element("spPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"spPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"spPr")),
            _ => {}
        }
    }

    Ok(ShapeProperties {
        bw_mode,
        transform,
        preset_geometry,
        fill,
        outline,
    })
}

fn parse_black_white_mode(val: &str) -> Result<BlackWhiteMode> {
    match val {
        "auto" => Ok(BlackWhiteMode::Auto),
        "black" => Ok(BlackWhiteMode::Black),
        "blackGray" => Ok(BlackWhiteMode::BlackGray),
        "blackWhite" => Ok(BlackWhiteMode::BlackWhite),
        "clr" => Ok(BlackWhiteMode::Clr),
        "gray" => Ok(BlackWhiteMode::Gray),
        "grayWhite" => Ok(BlackWhiteMode::GrayWhite),
        "hidden" => Ok(BlackWhiteMode::Hidden),
        "invGray" => Ok(BlackWhiteMode::InvGray),
        "ltGray" => Ok(BlackWhiteMode::LtGray),
        "white" => Ok(BlackWhiteMode::White),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "bwMode".into(),
            value: other.into(),
            reason: "expected value per §20.1.10.10 ST_BlackWhiteMode".into(),
        }),
    }
}

// ── a:xfrm (§20.1.7.6) ────────────────────────────────────────────────────

fn parse_transform_2d(
    start: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Transform2D> {
    let rotation = xml::optional_attr_i64(start, b"rot")?.map(Dimension::new);
    let flip_h = xml::optional_attr_bool(start, b"flipH")?;
    let flip_v = xml::optional_attr_bool(start, b"flipV")?;
    let mut offset = None;
    let mut extent = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"off" => {
                        let x = xml::optional_attr_i64(e, b"x")?.ok_or_else(|| {
                            ParseError::MissingAttribute {
                                element: "a:off".into(),
                                attr: "x".into(),
                            }
                        })?;
                        let y = xml::optional_attr_i64(e, b"y")?.ok_or_else(|| {
                            ParseError::MissingAttribute {
                                element: "a:off".into(),
                                attr: "y".into(),
                            }
                        })?;
                        offset = Some(Offset::new(Dimension::new(x), Dimension::new(y)));
                    }
                    b"ext" => {
                        extent = Some(parse_positive_size_2d(e)?);
                    }
                    _ => xml::warn_unsupported_element("xfrm", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"xfrm" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"xfrm")),
            _ => {}
        }
    }

    Ok(Transform2D {
        rotation,
        flip_h,
        flip_v,
        offset,
        extent,
    })
}

// ── a:prstGeom (§20.1.9.18) ────────────────────────────────────────────────

fn parse_preset_geometry(
    start: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<PresetGeometryDef> {
    let prst = xml::optional_attr(start, b"prst")?.ok_or_else(|| ParseError::MissingAttribute {
        element: "a:prstGeom".into(),
        attr: "prst".into(),
    })?;
    let preset = parse_preset_shape_type(&prst);
    let mut adjust_values = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"avLst" => {
                        adjust_values = parse_geom_guide_list(reader, buf)?;
                    }
                    _ => {
                        xml::warn_unsupported_element("prstGeom", local);
                        xml::skip_to_end(reader, buf, local)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"avLst" => {} // empty adjustment list
                    _ => xml::warn_unsupported_element("prstGeom", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"prstGeom" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"prstGeom")),
            _ => {}
        }
    }

    Ok(PresetGeometryDef {
        preset,
        adjust_values,
    })
}

fn parse_geom_guide_list(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Vec<GeomGuide>> {
    let mut guides = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) if xml::local_name(e.name().as_ref()) == b"gd" => {
                let name = xml::optional_attr(e, b"name")?.unwrap_or_default();
                let formula = xml::optional_attr(e, b"fmla")?.unwrap_or_default();
                guides.push(GeomGuide { name, formula });
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"avLst" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"avLst")),
            _ => {}
        }
    }

    Ok(guides)
}

fn parse_preset_shape_type(val: &str) -> PresetShapeType {
    match val {
        "rect" => PresetShapeType::Rect,
        "roundRect" => PresetShapeType::RoundRect,
        "ellipse" => PresetShapeType::Ellipse,
        "triangle" => PresetShapeType::Triangle,
        "rtTriangle" => PresetShapeType::RtTriangle,
        "diamond" => PresetShapeType::Diamond,
        "parallelogram" => PresetShapeType::Parallelogram,
        "trapezoid" => PresetShapeType::Trapezoid,
        "pentagon" => PresetShapeType::Pentagon,
        "hexagon" => PresetShapeType::Hexagon,
        "octagon" => PresetShapeType::Octagon,
        "star4" => PresetShapeType::Star4,
        "star5" => PresetShapeType::Star5,
        "star6" => PresetShapeType::Star6,
        "star8" => PresetShapeType::Star8,
        "star10" => PresetShapeType::Star10,
        "star12" => PresetShapeType::Star12,
        "star16" => PresetShapeType::Star16,
        "star24" => PresetShapeType::Star24,
        "star32" => PresetShapeType::Star32,
        "line" => PresetShapeType::Line,
        "plus" => PresetShapeType::Plus,
        "can" => PresetShapeType::Can,
        "cube" => PresetShapeType::Cube,
        "donut" => PresetShapeType::Donut,
        "noSmoking" => PresetShapeType::NoSmoking,
        "blockArc" => PresetShapeType::BlockArc,
        "heart" => PresetShapeType::Heart,
        "sun" => PresetShapeType::Sun,
        "moon" => PresetShapeType::Moon,
        "smileyFace" => PresetShapeType::SmileyFace,
        "lightningBolt" => PresetShapeType::LightningBolt,
        "cloud" => PresetShapeType::Cloud,
        "arc" => PresetShapeType::Arc,
        "plaque" => PresetShapeType::Plaque,
        "frame" => PresetShapeType::Frame,
        "bevel" => PresetShapeType::Bevel,
        "foldedCorner" => PresetShapeType::FoldedCorner,
        "chevron" => PresetShapeType::Chevron,
        "homePlate" => PresetShapeType::HomePlate,
        "ribbon" => PresetShapeType::Ribbon,
        "ribbon2" => PresetShapeType::Ribbon2,
        "pie" => PresetShapeType::Pie,
        "pieWedge" => PresetShapeType::PieWedge,
        "chord" => PresetShapeType::Chord,
        "teardrop" => PresetShapeType::Teardrop,
        "arrow" => PresetShapeType::Arrow,
        "leftArrow" => PresetShapeType::LeftArrow,
        "rightArrow" => PresetShapeType::RightArrow,
        "upArrow" => PresetShapeType::UpArrow,
        "downArrow" => PresetShapeType::DownArrow,
        "leftRightArrow" => PresetShapeType::LeftRightArrow,
        "upDownArrow" => PresetShapeType::UpDownArrow,
        "quadArrow" => PresetShapeType::QuadArrow,
        "bentArrow" => PresetShapeType::BentArrow,
        "uturnArrow" => PresetShapeType::UturnArrow,
        "circularArrow" => PresetShapeType::CircularArrow,
        "curvedRightArrow" => PresetShapeType::CurvedRightArrow,
        "curvedLeftArrow" => PresetShapeType::CurvedLeftArrow,
        "curvedUpArrow" => PresetShapeType::CurvedUpArrow,
        "curvedDownArrow" => PresetShapeType::CurvedDownArrow,
        "stripedRightArrow" => PresetShapeType::StripedRightArrow,
        "notchedRightArrow" => PresetShapeType::NotchedRightArrow,
        "bentUpArrow" => PresetShapeType::BentUpArrow,
        "leftUpArrow" => PresetShapeType::LeftUpArrow,
        "leftRightUpArrow" => PresetShapeType::LeftRightUpArrow,
        "leftArrowCallout" => PresetShapeType::LeftArrowCallout,
        "rightArrowCallout" => PresetShapeType::RightArrowCallout,
        "upArrowCallout" => PresetShapeType::UpArrowCallout,
        "downArrowCallout" => PresetShapeType::DownArrowCallout,
        "leftRightArrowCallout" => PresetShapeType::LeftRightArrowCallout,
        "upDownArrowCallout" => PresetShapeType::UpDownArrowCallout,
        "quadArrowCallout" => PresetShapeType::QuadArrowCallout,
        "swooshArrow" => PresetShapeType::SwooshArrow,
        "leftCircularArrow" => PresetShapeType::LeftCircularArrow,
        "leftRightCircularArrow" => PresetShapeType::LeftRightCircularArrow,
        "callout1" => PresetShapeType::Callout1,
        "callout2" => PresetShapeType::Callout2,
        "callout3" => PresetShapeType::Callout3,
        "accentCallout1" => PresetShapeType::AccentCallout1,
        "accentCallout2" => PresetShapeType::AccentCallout2,
        "accentCallout3" => PresetShapeType::AccentCallout3,
        "borderCallout1" => PresetShapeType::BorderCallout1,
        "borderCallout2" => PresetShapeType::BorderCallout2,
        "borderCallout3" => PresetShapeType::BorderCallout3,
        "accentBorderCallout1" => PresetShapeType::AccentBorderCallout1,
        "accentBorderCallout2" => PresetShapeType::AccentBorderCallout2,
        "accentBorderCallout3" => PresetShapeType::AccentBorderCallout3,
        "wedgeRectCallout" => PresetShapeType::WedgeRectCallout,
        "wedgeRoundRectCallout" => PresetShapeType::WedgeRoundRectCallout,
        "wedgeEllipseCallout" => PresetShapeType::WedgeEllipseCallout,
        "cloudCallout" => PresetShapeType::CloudCallout,
        "leftBracket" => PresetShapeType::LeftBracket,
        "rightBracket" => PresetShapeType::RightBracket,
        "leftBrace" => PresetShapeType::LeftBrace,
        "rightBrace" => PresetShapeType::RightBrace,
        "bracketPair" => PresetShapeType::BracketPair,
        "bracePair" => PresetShapeType::BracePair,
        "straightConnector1" => PresetShapeType::StraightConnector1,
        "bentConnector2" => PresetShapeType::BentConnector2,
        "bentConnector3" => PresetShapeType::BentConnector3,
        "bentConnector4" => PresetShapeType::BentConnector4,
        "bentConnector5" => PresetShapeType::BentConnector5,
        "curvedConnector2" => PresetShapeType::CurvedConnector2,
        "curvedConnector3" => PresetShapeType::CurvedConnector3,
        "curvedConnector4" => PresetShapeType::CurvedConnector4,
        "curvedConnector5" => PresetShapeType::CurvedConnector5,
        "flowChartProcess" => PresetShapeType::FlowChartProcess,
        "flowChartDecision" => PresetShapeType::FlowChartDecision,
        "flowChartInputOutput" => PresetShapeType::FlowChartInputOutput,
        "flowChartPredefinedProcess" => PresetShapeType::FlowChartPredefinedProcess,
        "flowChartInternalStorage" => PresetShapeType::FlowChartInternalStorage,
        "flowChartDocument" => PresetShapeType::FlowChartDocument,
        "flowChartMultidocument" => PresetShapeType::FlowChartMultidocument,
        "flowChartTerminator" => PresetShapeType::FlowChartTerminator,
        "flowChartPreparation" => PresetShapeType::FlowChartPreparation,
        "flowChartManualInput" => PresetShapeType::FlowChartManualInput,
        "flowChartManualOperation" => PresetShapeType::FlowChartManualOperation,
        "flowChartConnector" => PresetShapeType::FlowChartConnector,
        "flowChartPunchedCard" => PresetShapeType::FlowChartPunchedCard,
        "flowChartPunchedTape" => PresetShapeType::FlowChartPunchedTape,
        "flowChartSummingJunction" => PresetShapeType::FlowChartSummingJunction,
        "flowChartOr" => PresetShapeType::FlowChartOr,
        "flowChartCollate" => PresetShapeType::FlowChartCollate,
        "flowChartSort" => PresetShapeType::FlowChartSort,
        "flowChartExtract" => PresetShapeType::FlowChartExtract,
        "flowChartMerge" => PresetShapeType::FlowChartMerge,
        "flowChartOfflineStorage" => PresetShapeType::FlowChartOfflineStorage,
        "flowChartOnlineStorage" => PresetShapeType::FlowChartOnlineStorage,
        "flowChartMagneticTape" => PresetShapeType::FlowChartMagneticTape,
        "flowChartMagneticDisk" => PresetShapeType::FlowChartMagneticDisk,
        "flowChartMagneticDrum" => PresetShapeType::FlowChartMagneticDrum,
        "flowChartDisplay" => PresetShapeType::FlowChartDisplay,
        "flowChartDelay" => PresetShapeType::FlowChartDelay,
        "flowChartAlternateProcess" => PresetShapeType::FlowChartAlternateProcess,
        "flowChartOffpageConnector" => PresetShapeType::FlowChartOffpageConnector,
        "actionButtonBlank" => PresetShapeType::ActionButtonBlank,
        "actionButtonHome" => PresetShapeType::ActionButtonHome,
        "actionButtonHelp" => PresetShapeType::ActionButtonHelp,
        "actionButtonInformation" => PresetShapeType::ActionButtonInformation,
        "actionButtonForwardNext" => PresetShapeType::ActionButtonForwardNext,
        "actionButtonBackPrevious" => PresetShapeType::ActionButtonBackPrevious,
        "actionButtonEnd" => PresetShapeType::ActionButtonEnd,
        "actionButtonBeginning" => PresetShapeType::ActionButtonBeginning,
        "actionButtonReturn" => PresetShapeType::ActionButtonReturn,
        "actionButtonDocument" => PresetShapeType::ActionButtonDocument,
        "actionButtonSound" => PresetShapeType::ActionButtonSound,
        "actionButtonMovie" => PresetShapeType::ActionButtonMovie,
        "irregularSeal1" => PresetShapeType::IrregularSeal1,
        "irregularSeal2" => PresetShapeType::IrregularSeal2,
        "wave" => PresetShapeType::Wave,
        "doubleWave" => PresetShapeType::DoubleWave,
        "ellipseRibbon" => PresetShapeType::EllipseRibbon,
        "ellipseRibbon2" => PresetShapeType::EllipseRibbon2,
        "verticalScroll" => PresetShapeType::VerticalScroll,
        "horizontalScroll" => PresetShapeType::HorizontalScroll,
        "leftRightRibbon" => PresetShapeType::LeftRightRibbon,
        "gear6" => PresetShapeType::Gear6,
        "gear9" => PresetShapeType::Gear9,
        "funnel" => PresetShapeType::Funnel,
        "mathPlus" => PresetShapeType::MathPlus,
        "mathMinus" => PresetShapeType::MathMinus,
        "mathMultiply" => PresetShapeType::MathMultiply,
        "mathDivide" => PresetShapeType::MathDivide,
        "mathEqual" => PresetShapeType::MathEqual,
        "mathNotEqual" => PresetShapeType::MathNotEqual,
        "cornerTabs" => PresetShapeType::CornerTabs,
        "squareTabs" => PresetShapeType::SquareTabs,
        "plaqueTabs" => PresetShapeType::PlaqueTabs,
        "chartX" => PresetShapeType::ChartX,
        "chartStar" => PresetShapeType::ChartStar,
        "chartPlus" => PresetShapeType::ChartPlus,
        "halfFrame" => PresetShapeType::HalfFrame,
        "corner" => PresetShapeType::Corner,
        "diagStripe" => PresetShapeType::DiagStripe,
        "nonIsoscelesTrapezoid" => PresetShapeType::NonIsoscelesTrapezoid,
        "heptagon" => PresetShapeType::Heptagon,
        "decagon" => PresetShapeType::Decagon,
        "dodecagon" => PresetShapeType::Dodecagon,
        "round1Rect" => PresetShapeType::Round1Rect,
        "round2SameRect" => PresetShapeType::Round2SameRect,
        "round2DiagRect" => PresetShapeType::Round2DiagRect,
        "snipRoundRect" => PresetShapeType::SnipRoundRect,
        "snip1Rect" => PresetShapeType::Snip1Rect,
        "snip2SameRect" => PresetShapeType::Snip2SameRect,
        "snip2DiagRect" => PresetShapeType::Snip2DiagRect,
        other => {
            log::warn!("prstGeom: unrecognized shape type {:?}", other);
            PresetShapeType::Other(other.to_owned())
        }
    }
}

// ── a:ln (§20.1.2.2.24) ────────────────────────────────────────────────────

fn parse_outline(
    start: &BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Outline> {
    let width = xml::optional_attr_i64(start, b"w")?.map(Dimension::new);
    let cap = xml::optional_attr(start, b"cap")?
        .map(|s| parse_line_cap(&s))
        .transpose()?;
    let compound = xml::optional_attr(start, b"cmpd")?
        .map(|s| parse_compound_line(&s))
        .transpose()?;
    let alignment = xml::optional_attr(start, b"algn")?
        .map(|s| parse_pen_alignment(&s))
        .transpose()?;
    let mut fill = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                // Skip complex children (dash, join, head/tail end, fills).
                xml::skip_to_end(reader, buf, local)?;
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                if local == b"noFill" {
                    fill = Some(DrawingFill::NoFill);
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"ln" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"ln")),
            _ => {}
        }
    }

    Ok(Outline {
        width,
        cap,
        compound,
        alignment,
        fill,
    })
}

fn parse_line_cap(val: &str) -> Result<LineCap> {
    match val {
        "flat" => Ok(LineCap::Flat),
        "rnd" => Ok(LineCap::Round),
        "sq" => Ok(LineCap::Square),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "cap".into(),
            value: other.into(),
            reason: "expected flat, rnd, or sq per §20.1.10.31 ST_LineCap".into(),
        }),
    }
}

fn parse_compound_line(val: &str) -> Result<CompoundLine> {
    match val {
        "sng" => Ok(CompoundLine::Single),
        "dbl" => Ok(CompoundLine::Double),
        "thickThin" => Ok(CompoundLine::ThickThin),
        "thinThick" => Ok(CompoundLine::ThinThick),
        "tri" => Ok(CompoundLine::Triple),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "cmpd".into(),
            value: other.into(),
            reason: "expected value per §20.1.10.15 ST_CompoundLine".into(),
        }),
    }
}

fn parse_pen_alignment(val: &str) -> Result<PenAlignment> {
    match val {
        "ctr" => Ok(PenAlignment::Center),
        "in" => Ok(PenAlignment::Inset),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "algn".into(),
            value: other.into(),
            reason: "expected ctr or in per §20.1.10.39 ST_PenAlignment".into(),
        }),
    }
}
