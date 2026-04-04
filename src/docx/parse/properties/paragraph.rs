use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::dimension::Dimension;
use crate::docx::error::Result;
use crate::docx::model::*;
use crate::docx::xml;

use super::run::parse_run_properties;
use super::section::{parse_section_properties, parse_section_rsids};
use super::{
    invalid_value, opt_val, parse_alignment, parse_border, parse_cnf_style, parse_shading,
    toggle_attr,
};

/// Parsed result of a `w:pPr` element.
pub struct ParsedParagraphProperties {
    pub properties: ParagraphProperties,
    pub style_id: Option<StyleId>,
    pub run_properties: Option<RunProperties>,
    /// Per §17.6.18, a `w:sectPr` child of `w:pPr` means this paragraph
    /// is the last paragraph of a section. The section break occurs after
    /// the paragraph, and these are the properties for that section.
    pub section_properties: Option<SectionProperties>,
}

/// Parse `w:pPr` element. Reader must have just read the Start event for `pPr`.
pub fn parse_paragraph_properties(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<ParsedParagraphProperties> {
    let mut props = ParagraphProperties::default();
    let mut style_id: Option<StyleId> = None;
    let mut run_props: Option<RunProperties> = None;
    let mut sect_props: Option<SectionProperties> = None;

    loop {
        let event = xml::next_event(reader, buf)?;
        let is_start = matches!(event, Event::Start(_));
        match event {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    // attrs-only — valid as Start or Empty
                    b"pStyle" => {
                        style_id = xml::optional_attr(e, b"val")?.map(StyleId::new);
                    }
                    b"ind" => {
                        props.indentation = Some(parse_indentation(e)?);
                    }
                    b"spacing" => {
                        props.spacing = Some(parse_paragraph_spacing(e)?);
                    }
                    b"jc" => {
                        props.alignment = opt_val(e, parse_alignment)?;
                    }
                    b"shd" => {
                        props.shading = Some(parse_shading(e)?);
                    }
                    b"outlineLvl" => {
                        if let Some(val) = xml::optional_attr_u32(e, b"val")? {
                            props.outline_level = OutlineLevel::from_ooxml(val as u8);
                        }
                    }
                    // Start-only: have child elements
                    b"numPr" if is_start => {
                        props.numbering = Some(parse_numbering_pr(reader, buf)?);
                    }
                    b"tabs" if is_start => {
                        props.tabs = parse_tabs(reader, buf)?;
                    }
                    b"pBdr" if is_start => {
                        props.borders = Some(parse_paragraph_borders(reader, buf)?);
                    }
                    b"rPr" if is_start => {
                        let (rp, _) = parse_run_properties(reader, buf)?;
                        run_props = Some(rp);
                    }
                    // §17.6.18: sectPr inside pPr defines the section that ends
                    // with this paragraph. Contains pgSz, pgMar, cols, headerReference,
                    // footerReference, titlePg, type, pgNumType, docGrid, etc.
                    b"sectPr" if is_start => {
                        let rsids = parse_section_rsids(e)?;
                        let mut sp = parse_section_properties(reader, buf)?;
                        sp.rsids = rsids;
                        sect_props = Some(sp);
                    }
                    // Empty-only: no children
                    b"keepNext" => {
                        props.keep_next = toggle_attr(e)?;
                    }
                    b"keepLines" => {
                        props.keep_lines = toggle_attr(e)?;
                    }
                    b"widowControl" => {
                        props.widow_control = toggle_attr(e)?;
                    }
                    b"pageBreakBefore" => {
                        props.page_break_before = toggle_attr(e)?;
                    }
                    b"suppressAutoHyphens" => {
                        props.suppress_auto_hyphens = toggle_attr(e)?;
                    }
                    b"contextualSpacing" => {
                        props.contextual_spacing = toggle_attr(e)?;
                    }
                    b"bidi" => {
                        props.bidi = toggle_attr(e)?;
                    }
                    b"wordWrap" => {
                        props.word_wrap = toggle_attr(e)?;
                    }
                    b"textAlignment" => {
                        props.text_alignment = opt_val(e, parse_text_alignment)?;
                    }
                    b"cnfStyle" => {
                        props.cnf_style = Some(parse_cnf_style(e)?);
                    }
                    b"framePr" => {
                        props.frame_properties = Some(parse_frame_kind(e)?);
                    }
                    b"autoSpaceDE" => {
                        props.auto_space_de = toggle_attr(e)?;
                    }
                    b"autoSpaceDN" => {
                        props.auto_space_dn = toggle_attr(e)?;
                    }
                    // §17.6.18: an empty sectPr is valid (inherits all defaults).
                    b"sectPr" => {
                        let rsids = parse_section_rsids(e)?;
                        sect_props = Some(SectionProperties {
                            rsids,
                            ..SectionProperties::default()
                        });
                    }
                    _ => xml::warn_unsupported_element("pPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"pPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"pPr")),
            _ => {}
        }
    }

    Ok(ParsedParagraphProperties {
        properties: props,
        style_id,
        run_properties: run_props,
        section_properties: sect_props,
    })
}

fn parse_indentation(e: &BytesStart<'_>) -> Result<Indentation> {
    let start = xml::optional_attr_i64(e, b"left")?
        .or(xml::optional_attr_i64(e, b"start")?)
        .map(Dimension::new);
    let end = xml::optional_attr_i64(e, b"right")?
        .or(xml::optional_attr_i64(e, b"end")?)
        .map(Dimension::new);
    let first_line = if let Some(val) = xml::optional_attr_i64(e, b"hanging")? {
        Some(FirstLineIndent::Hanging(Dimension::new(val)))
    } else {
        xml::optional_attr_i64(e, b"firstLine")?
            .map(|val| FirstLineIndent::FirstLine(Dimension::new(val)))
    };
    let mirror = xml::optional_attr_bool(e, b"mirrorIndents")?;

    Ok(Indentation {
        start,
        end,
        first_line,
        mirror,
    })
}

fn parse_paragraph_spacing(e: &BytesStart<'_>) -> Result<ParagraphSpacing> {
    let before = xml::optional_attr_i64(e, b"before")?.map(Dimension::new);
    let after = xml::optional_attr_i64(e, b"after")?.map(Dimension::new);
    let line_val = xml::optional_attr_i64(e, b"line")?;
    let line_rule = xml::optional_attr(e, b"lineRule")?;

    let line = line_val.map(|val| match line_rule.as_deref() {
        Some("exact") => LineSpacing::Exact(Dimension::new(val)),
        Some("atLeast") => LineSpacing::AtLeast(Dimension::new(val)),
        _ => LineSpacing::Auto(Dimension::new(val)),
    });

    Ok(ParagraphSpacing {
        before,
        after,
        line,
        before_auto_spacing: xml::optional_attr_bool(e, b"beforeAutospacing")?,
        after_auto_spacing: xml::optional_attr_bool(e, b"afterAutospacing")?,
    })
}

fn parse_numbering_pr(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<NumberingReference> {
    let mut level = 0u8;
    let mut num_id = 0i64;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"ilvl" => {
                        if let Some(val) = xml::optional_attr_u32(e, b"val")? {
                            level = val as u8;
                        }
                    }
                    b"numId" => {
                        if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                            num_id = val;
                        }
                    }
                    _ => xml::warn_unsupported_element("numPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"numPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"numPr")),
            _ => {}
        }
    }

    Ok(NumberingReference { num_id, level })
}

fn parse_tabs(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Vec<TabStop>> {
    let mut tabs = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e)
                if xml::local_name(e.name().as_ref()) == b"tab" =>
            {
                let pos = xml::optional_attr_i64(e, b"pos")?.unwrap_or(0);
                let alignment = match xml::optional_attr(e, b"val")?.as_deref() {
                    Some("left") | None => TabAlignment::Left,
                    Some("center") => TabAlignment::Center,
                    Some("right") => TabAlignment::Right,
                    Some("decimal") => TabAlignment::Decimal,
                    Some("bar") => TabAlignment::Bar,
                    Some("clear") => TabAlignment::Clear,
                    Some("num") => TabAlignment::Left, // §17.18.81: num is legacy, treat as left
                    Some(other) => return Err(invalid_value("tab/val", other)),
                };
                let leader = match xml::optional_attr(e, b"leader")?.as_deref() {
                    Some("none") | None => TabLeader::None,
                    Some("dot") => TabLeader::Dot,
                    Some("hyphen") => TabLeader::Hyphen,
                    Some("underscore") => TabLeader::Underscore,
                    Some("heavy") => TabLeader::Heavy,
                    Some("middleDot") => TabLeader::MiddleDot,
                    Some(other) => return Err(invalid_value("tab/leader", other)),
                };
                tabs.push(TabStop {
                    position: Dimension::new(pos),
                    alignment,
                    leader,
                });
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tabs" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tabs")),
            _ => {}
        }
    }

    Ok(tabs)
}

fn parse_paragraph_borders(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<ParagraphBorders> {
    let mut borders = ParagraphBorders {
        top: None,
        bottom: None,
        left: None,
        right: None,
        between: None,
    };

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                let border = parse_border(e)?;
                match local {
                    b"top" => borders.top = Some(border),
                    b"bottom" => borders.bottom = Some(border),
                    b"left" | b"start" => borders.left = Some(border),
                    b"right" | b"end" => borders.right = Some(border),
                    b"between" => borders.between = Some(border),
                    _ => xml::warn_unsupported_element("pBdr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"pBdr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"pBdr")),
            _ => {}
        }
    }

    Ok(borders)
}

/// §17.18.91 ST_TextAlignment
fn parse_text_alignment(val: &str) -> Result<TextAlignment> {
    match val {
        "auto" => Ok(TextAlignment::Auto),
        "top" => Ok(TextAlignment::Top),
        "center" => Ok(TextAlignment::Center),
        "baseline" => Ok(TextAlignment::Baseline),
        "bottom" => Ok(TextAlignment::Bottom),
        other => Err(invalid_value("textAlignment/val", other)),
    }
}

/// §17.3.1.11: parse `w:framePr` attributes into a [`FrameKind`].
///
/// When `w:dropCap` is `"drop"` or `"margin"` the element represents a drop cap;
/// otherwise it is a floating text-box frame.
fn parse_frame_kind(e: &BytesStart<'_>) -> Result<FrameKind> {
    let drop_cap = match xml::optional_attr(e, b"dropCap")?.as_deref() {
        Some("drop") => Some(DropCap::Drop),
        Some("margin") => Some(DropCap::Margin),
        _ => None, // "none" or absent — not a drop cap
    };

    if let Some(style) = drop_cap {
        return Ok(FrameKind::DropCap {
            style,
            lines: xml::optional_attr_u32(e, b"lines")?.unwrap_or(3),
            h_space: xml::optional_attr_i64(e, b"hSpace")?.map(Dimension::new),
        });
    }

    Ok(FrameKind::TextBox(TextBoxPositioning {
        width: xml::optional_attr_i64(e, b"w")?.map(Dimension::new),
        height: xml::optional_attr_i64(e, b"h")?.map(Dimension::new),
        height_rule: match xml::optional_attr(e, b"hRule")?.as_deref() {
            Some("auto") => Some(HeightRule::Auto),
            Some("exact") => Some(HeightRule::Exact),
            Some("atLeast") => Some(HeightRule::AtLeast),
            Some(other) => return Err(invalid_value("framePr/hRule", other)),
            None => None,
        },
        h_space: xml::optional_attr_i64(e, b"hSpace")?.map(Dimension::new),
        v_space: xml::optional_attr_i64(e, b"vSpace")?.map(Dimension::new),
        wrap: match xml::optional_attr(e, b"wrap")?.as_deref() {
            Some("auto") => Some(FrameWrap::Auto),
            Some("notBeside") => Some(FrameWrap::NotBeside),
            Some("around") => Some(FrameWrap::Around),
            Some("tight") => Some(FrameWrap::Tight),
            Some("through") => Some(FrameWrap::Through),
            Some("none") => Some(FrameWrap::None),
            Some(other) => return Err(invalid_value("framePr/wrap", other)),
            None => None,
        },
        h_anchor: match xml::optional_attr(e, b"hAnchor")?.as_deref() {
            Some("text") => Some(TableAnchor::Text),
            Some("margin") => Some(TableAnchor::Margin),
            Some("page") => Some(TableAnchor::Page),
            Some(other) => return Err(invalid_value("framePr/hAnchor", other)),
            None => None,
        },
        v_anchor: match xml::optional_attr(e, b"vAnchor")?.as_deref() {
            Some("text") => Some(TableAnchor::Text),
            Some("margin") => Some(TableAnchor::Margin),
            Some("page") => Some(TableAnchor::Page),
            Some(other) => return Err(invalid_value("framePr/vAnchor", other)),
            None => None,
        },
        x: xml::optional_attr_i64(e, b"x")?.map(Dimension::new),
        x_align: match xml::optional_attr(e, b"xAlign")?.as_deref() {
            Some("left") => Some(TableXAlign::Left),
            Some("center") => Some(TableXAlign::Center),
            Some("right") => Some(TableXAlign::Right),
            Some("inside") => Some(TableXAlign::Inside),
            Some("outside") => Some(TableXAlign::Outside),
            Some(other) => return Err(invalid_value("framePr/xAlign", other)),
            None => None,
        },
        y: xml::optional_attr_i64(e, b"y")?.map(Dimension::new),
        y_align: match xml::optional_attr(e, b"yAlign")?.as_deref() {
            Some("top") => Some(TableYAlign::Top),
            Some("center") => Some(TableYAlign::Center),
            Some("bottom") => Some(TableYAlign::Bottom),
            Some("inside") => Some(TableYAlign::Inside),
            Some("outside") => Some(TableYAlign::Outside),
            Some("inline") => Some(TableYAlign::Inline),
            Some(other) => return Err(invalid_value("framePr/yAlign", other)),
            None => None,
        },
    }))
}
