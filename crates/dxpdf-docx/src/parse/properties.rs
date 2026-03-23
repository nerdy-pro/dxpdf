//! Parsers for OOXML property elements: pPr, rPr, tblPr, trPr, tcPr, sectPr.
//!
//! Each parser consumes events from the reader until the corresponding End event,
//! returning a fully-populated properties struct.

use log::warn;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::dimension::Dimension;
use crate::error::Result;
use crate::geometry::EdgeInsets;
use crate::model::*;
use crate::xml;

// ── Paragraph Properties ─────────────────────────────────────────────────────

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
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"pStyle" => {
                        style_id = xml::optional_attr(e, b"val")?.map(StyleId);
                    }
                    b"ind" => {
                        props.indentation = Some(parse_indentation(e)?);
                    }
                    b"spacing" => {
                        props.spacing = Some(parse_paragraph_spacing(e)?);
                    }
                    b"jc" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.alignment = Some(parse_alignment(&val));
                        }
                    }
                    b"numPr" => {
                        props.numbering = Some(parse_numbering_pr(reader, buf)?);
                    }
                    b"tabs" => {
                        props.tabs = parse_tabs(reader, buf)?;
                    }
                    b"pBdr" => {
                        props.borders = Some(parse_paragraph_borders(reader, buf)?);
                    }
                    b"shd" => {
                        props.shading = Some(parse_shading(e)?);
                    }
                    b"rPr" => {
                        let (rp, _) = parse_run_properties(reader, buf)?;
                        run_props = Some(rp);
                    }
                    b"outlineLvl" => {
                        if let Some(val) = xml::optional_attr_u32(e, b"val")? {
                            props.outline_level = OutlineLevel::from_ooxml(val as u8);
                        }
                    }
                    // §17.6.18: sectPr inside pPr defines the section that ends
                    // with this paragraph. Contains pgSz, pgMar, cols, headerReference,
                    // footerReference, titlePg, type, pgNumType, docGrid, etc.
                    b"sectPr" => {
                        let rsids = parse_section_rsids(e)?;
                        let mut sp = parse_section_properties(reader, buf)?;
                        sp.rsids = rsids;
                        sect_props = Some(sp);
                    }
                    _ => xml::warn_unsupported_element("pPr", &local),
                }
            }
            Event::Empty(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"pStyle" => {
                        style_id = xml::optional_attr(e, b"val")?.map(StyleId);
                    }
                    b"ind" => {
                        props.indentation = Some(parse_indentation(e)?);
                    }
                    b"spacing" => {
                        props.spacing = Some(parse_paragraph_spacing(e)?);
                    }
                    b"jc" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.alignment = Some(parse_alignment(&val));
                        }
                    }
                    b"shd" => {
                        props.shading = Some(parse_shading(e)?);
                    }
                    b"keepNext" => {
                        props.keep_next = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"keepLines" => {
                        props.keep_lines =
                            Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"widowControl" => {
                        props.widow_control =
                            Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"pageBreakBefore" => {
                        props.page_break_before =
                            Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"suppressAutoHyphens" => {
                        props.suppress_auto_hyphens =
                            Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"bidi" => {
                        props.bidi = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"outlineLvl" => {
                        if let Some(val) = xml::optional_attr_u32(e, b"val")? {
                            props.outline_level = OutlineLevel::from_ooxml(val as u8);
                        }
                    }
                    // §17.6.18: an empty sectPr is valid (inherits all defaults).
                    b"sectPr" => {
                        let rsids = parse_section_rsids(e)?;
                        sect_props = Some(SectionProperties {
                            rsids,
                            ..SectionProperties::default()
                        });
                    }
                    _ => xml::warn_unsupported_element("pPr", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"pPr" => break,
            Event::Eof => break,
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

// ── Run Properties ───────────────────────────────────────────────────────────

/// Parse `w:rPr` element. Returns (properties, optional style ID).
pub fn parse_run_properties(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<(RunProperties, Option<StyleId>)> {
    let mut props = RunProperties::default();
    let mut style_id: Option<StyleId> = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"rStyle" => {
                        style_id = xml::optional_attr(e, b"val")?.map(StyleId);
                    }
                    b"rFonts" => {
                        props.fonts = parse_font_set(e)?;
                    }
                    b"sz" => {
                        if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                            props.font_size = Some(Dimension::new(val));
                        }
                    }
                    b"szCs" => {}
                    b"b" => {
                        props.bold = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"bCs" => {}
                    b"i" => {
                        props.italic = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"iCs" => {}
                    b"u" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.underline = Some(parse_underline_style(&val));
                        } else {
                            props.underline = Some(UnderlineStyle::Single);
                        }
                    }
                    b"strike" => {
                        let on = xml::optional_attr_bool(e, b"val")?.unwrap_or(true);
                        if on {
                            props.strike = Some(StrikeStyle::Single);
                        }
                    }
                    b"dstrike" => {
                        let on = xml::optional_attr_bool(e, b"val")?.unwrap_or(true);
                        if on {
                            props.strike = Some(StrikeStyle::Double);
                        }
                    }
                    b"color" => {
                        props.color = Some(parse_color_attr(e)?);
                    }
                    b"highlight" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.highlight = parse_highlight_color(&val);
                        }
                    }
                    b"shd" => {
                        props.shading = Some(parse_shading(e)?);
                    }
                    b"vertAlign" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.vertical_align = Some(parse_vertical_align(&val));
                        }
                    }
                    b"spacing" => {
                        if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                            props.spacing = Some(Dimension::new(val));
                        }
                    }
                    b"kern" => {
                        if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                            props.kerning = Some(Dimension::new(val));
                        }
                    }
                    b"caps" => {
                        props.all_caps = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"smallCaps" => {
                        props.small_caps =
                            Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"vanish" => {
                        props.vanish = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"noProof" => {
                        props.no_proof = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"rtl" => {
                        props.rtl = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"emboss" => {
                        props.emboss = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"imprint" => {
                        props.imprint = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"outline" => {
                        props.outline = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"shadow" => {
                        props.shadow = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    _ => xml::warn_unsupported_element("rPr", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"rPr" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok((props, style_id))
}

// ── Table Properties ─────────────────────────────────────────────────────────

/// Parse `w:tblPr`. Returns (properties, optional style ID).
pub fn parse_table_properties(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<(TableProperties, Option<StyleId>)> {
    let mut props = TableProperties::default();
    let mut style_id: Option<StyleId> = None;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"tblStyle" => {
                        style_id = xml::optional_attr(e, b"val")?.map(StyleId);
                    }
                    b"tblBorders" => {
                        props.borders = Some(parse_table_borders(reader, buf)?);
                    }
                    b"tblCellMar" => {
                        props.cell_margins =
                            Some(parse_edge_insets_twips(reader, buf, b"tblCellMar")?);
                    }
                    b"tblLook" => {
                        props.look = Some(parse_table_look(e)?);
                    }
                    _ => xml::warn_unsupported_element("tblPr", &local),
                }
            }
            Event::Empty(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"tblStyle" => {
                        style_id = xml::optional_attr(e, b"val")?.map(StyleId);
                    }
                    b"jc" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.alignment = Some(parse_alignment(&val));
                        }
                    }
                    b"tblW" => {
                        props.width = Some(parse_table_measure(e)?);
                    }
                    b"tblLayout" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.layout = Some(match val.as_str() {
                                "fixed" => TableLayout::Fixed,
                                "autofit" | "auto" => TableLayout::Auto,
                                other => {
                                    warn!("unknown table layout: {other}");
                                    TableLayout::Auto
                                }
                            });
                        }
                    }
                    b"tblInd" => {
                        props.indent = Some(parse_table_measure(e)?);
                    }
                    b"tblCellSpacing" => {
                        props.cell_spacing = Some(parse_table_measure(e)?);
                    }
                    b"tblLook" => {
                        props.look = Some(parse_table_look(e)?);
                    }
                    _ => xml::warn_unsupported_element("tblPr", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tblPr" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok((props, style_id))
}

/// Parse `w:trPr`.
pub fn parse_table_row_properties(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<TableRowProperties> {
    let mut props = TableRowProperties::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"trHeight" => {
                        if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                            let rule = match xml::optional_attr(e, b"hRule")?.as_deref() {
                                Some("exact") => HeightRule::Exact,
                                Some("atLeast") => HeightRule::AtLeast,
                                _ => HeightRule::Auto,
                            };
                            props.height = Some(TableRowHeight {
                                value: Dimension::new(val),
                                rule,
                            });
                        }
                    }
                    b"tblHeader" => {
                        props.is_header = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"cantSplit" => {
                        props.cant_split =
                            Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    _ => xml::warn_unsupported_element("trPr", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"trPr" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(props)
}

/// Parse `w:tcPr`.
pub fn parse_table_cell_properties(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<TableCellProperties> {
    let mut props = TableCellProperties::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"tcBorders" => {
                        props.borders = Some(parse_table_cell_borders(reader, buf)?);
                    }
                    b"tcMar" => {
                        props.margins = Some(parse_edge_insets_twips(reader, buf, b"tcMar")?);
                    }
                    _ => xml::warn_unsupported_element("tcPr", &local),
                }
            }
            Event::Empty(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"tcW" => {
                        props.width = Some(parse_table_measure(e)?);
                    }
                    b"shd" => {
                        props.shading = Some(parse_shading(e)?);
                    }
                    b"vAlign" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.vertical_align = Some(parse_cell_vertical_align(&val));
                        }
                    }
                    b"vMerge" => {
                        let is_restart = xml::optional_attr(e, b"val")?
                            .map(|v| v == "restart")
                            .unwrap_or(false);
                        props.vertical_merge = Some(if is_restart {
                            VerticalMerge::Restart
                        } else {
                            VerticalMerge::Continue
                        });
                    }
                    b"gridSpan" => {
                        props.grid_span = xml::optional_attr_u32(e, b"val")?;
                    }
                    b"textDirection" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.text_direction = Some(parse_text_direction(&val));
                        }
                    }
                    b"noWrap" => {
                        props.no_wrap = Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    _ => xml::warn_unsupported_element("tcPr", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tcPr" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(props)
}

// ── Section Properties ───────────────────────────────────────────────────────

/// Parse `w:sectPr`. Reader should have just read the Start event.
/// Extract rsid attributes from a `w:sectPr` Start event.
pub fn parse_section_rsids(e: &BytesStart<'_>) -> Result<SectionRevisionIds> {
    Ok(SectionRevisionIds {
        r: xml::optional_rsid(e, b"rsidR")?,
        r_pr: xml::optional_rsid(e, b"rsidRPr")?,
        sect: xml::optional_rsid(e, b"rsidSect")?,
    })
}

pub fn parse_section_properties(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<SectionProperties> {
    let mut props = SectionProperties::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"pgSz" => {
                        props.page_size = Some(PageSize {
                            width: xml::optional_attr_i64(e, b"w")?.map(Dimension::new),
                            height: xml::optional_attr_i64(e, b"h")?.map(Dimension::new),
                            orientation: xml::optional_attr(e, b"orient")?.map(|s| {
                                match s.as_str() {
                                    "landscape" => PageOrientation::Landscape,
                                    "portrait" => PageOrientation::Portrait,
                                    other => {
                                        warn!("unknown page orientation: {other}");
                                        PageOrientation::Portrait
                                    }
                                }
                            }),
                        });
                    }
                    b"pgMar" => {
                        props.page_margins = Some(PageMargins {
                            top: xml::optional_attr_i64(e, b"top")?.map(Dimension::new),
                            right: xml::optional_attr_i64(e, b"right")?.map(Dimension::new),
                            bottom: xml::optional_attr_i64(e, b"bottom")?.map(Dimension::new),
                            left: xml::optional_attr_i64(e, b"left")?.map(Dimension::new),
                            header: xml::optional_attr_i64(e, b"header")?.map(Dimension::new),
                            footer: xml::optional_attr_i64(e, b"footer")?.map(Dimension::new),
                            gutter: xml::optional_attr_i64(e, b"gutter")?.map(Dimension::new),
                        });
                    }
                    b"cols" => {
                        props.columns = Some(Columns {
                            count: xml::optional_attr_u32(e, b"num")?,
                            space: xml::optional_attr_i64(e, b"space")?.map(Dimension::new),
                            equal_width: xml::optional_attr_bool(e, b"equalWidth")?,
                        });
                    }
                    b"headerReference" => {
                        if let Some(r_id) = xml::optional_attr(e, b"id")? {
                            let hf_type = xml::optional_attr(e, b"type")?;
                            let rel = RelId(r_id);
                            match hf_type.as_deref() {
                                Some("first") => props.header_refs.first = Some(rel),
                                Some("even") => props.header_refs.even = Some(rel),
                                _ => props.header_refs.default = Some(rel),
                            }
                        }
                    }
                    b"footerReference" => {
                        if let Some(r_id) = xml::optional_attr(e, b"id")? {
                            let hf_type = xml::optional_attr(e, b"type")?;
                            let rel = RelId(r_id);
                            match hf_type.as_deref() {
                                Some("first") => props.footer_refs.first = Some(rel),
                                Some("even") => props.footer_refs.even = Some(rel),
                                _ => props.footer_refs.default = Some(rel),
                            }
                        }
                    }
                    b"titlePg" => {
                        props.title_page =
                            Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true));
                    }
                    b"type" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.section_type = Some(parse_section_type(&val));
                        }
                    }
                    _ => xml::warn_unsupported_element("sectPr", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"sectPr" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(props)
}

// ── Shared parsing helpers ───────────────────────────────────────────────────

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

pub fn parse_alignment(val: &str) -> Alignment {
    match val {
        "start" | "left" => Alignment::Start,
        "center" => Alignment::Center,
        "end" | "right" => Alignment::End,
        "both" | "justify" => Alignment::Both,
        "distribute" => Alignment::Distribute,
        "thaiDistribute" => Alignment::Thai,
        other => {
            warn!("unknown alignment value: {other}");
            Alignment::Start
        }
    }
}

fn parse_numbering_pr(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<NumberingReference> {
    let mut level = 0u8;
    let mut num_id = 0i64;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
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
                    _ => xml::warn_unsupported_element("numPr", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"numPr" => break,
            Event::Eof => break,
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
                    Some("center") => TabAlignment::Center,
                    Some("right") => TabAlignment::Right,
                    Some("decimal") => TabAlignment::Decimal,
                    Some("bar") => TabAlignment::Bar,
                    Some("clear") => TabAlignment::Clear,
                    _ => TabAlignment::Left,
                };
                let leader = match xml::optional_attr(e, b"leader")?.as_deref() {
                    Some("dot") => TabLeader::Dot,
                    Some("hyphen") => TabLeader::Hyphen,
                    Some("underscore") => TabLeader::Underscore,
                    Some("heavy") => TabLeader::Heavy,
                    Some("middleDot") => TabLeader::MiddleDot,
                    _ => TabLeader::None,
                };
                tabs.push(TabStop {
                    position: Dimension::new(pos),
                    alignment,
                    leader,
                });
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tabs" => break,
            Event::Eof => break,
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
                let local = xml::local_name(e.name().as_ref()).to_vec();
                let border = parse_border(e)?;
                match local.as_slice() {
                    b"top" => borders.top = Some(border),
                    b"bottom" => borders.bottom = Some(border),
                    b"left" | b"start" => borders.left = Some(border),
                    b"right" | b"end" => borders.right = Some(border),
                    b"between" => borders.between = Some(border),
                    _ => xml::warn_unsupported_element("pBdr", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"pBdr" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(borders)
}

pub fn parse_border(e: &BytesStart<'_>) -> Result<Border> {
    let style = match xml::optional_attr(e, b"val")?.as_deref() {
        Some("single") => BorderStyle::Single,
        Some("thick") => BorderStyle::Thick,
        Some("double") => BorderStyle::Double,
        Some("dotted") => BorderStyle::Dotted,
        Some("dashed") => BorderStyle::Dashed,
        Some("dotDash") => BorderStyle::DotDash,
        Some("dotDotDash") => BorderStyle::DotDotDash,
        Some("triple") => BorderStyle::Triple,
        Some("thinThickSmallGap") => BorderStyle::ThinThickSmallGap,
        Some("thickThinSmallGap") => BorderStyle::ThickThinSmallGap,
        Some("thinThickThinSmallGap") => BorderStyle::ThinThickThinSmallGap,
        Some("thinThickMediumGap") => BorderStyle::ThinThickMediumGap,
        Some("thickThinMediumGap") => BorderStyle::ThickThinMediumGap,
        Some("thinThickThinMediumGap") => BorderStyle::ThinThickThinMediumGap,
        Some("thinThickLargeGap") => BorderStyle::ThinThickLargeGap,
        Some("thickThinLargeGap") => BorderStyle::ThickThinLargeGap,
        Some("thinThickThinLargeGap") => BorderStyle::ThinThickThinLargeGap,
        Some("wave") => BorderStyle::Wave,
        Some("doubleWave") => BorderStyle::DoubleWave,
        Some("dashSmallGap") => BorderStyle::DashSmallGap,
        Some("dashDotStroked") => BorderStyle::DashDotStroked,
        Some("threeDEmboss") => BorderStyle::ThreeDEmboss,
        Some("threeDEngrave") => BorderStyle::ThreeDEngrave,
        Some("outset") => BorderStyle::Outset,
        Some("inset") => BorderStyle::Inset,
        Some("none") | Some("nil") | None => BorderStyle::None,
        Some(other) => {
            warn!("unknown border style: {other}");
            BorderStyle::None
        }
    };

    let sz = xml::optional_attr_i64(e, b"sz")?.unwrap_or(0);
    let space = xml::optional_attr_i64(e, b"space")?.unwrap_or(0);
    let color = parse_color_from_attr(e)?;

    Ok(Border {
        style,
        width: Dimension::new(sz),
        space: Dimension::new(space),
        color,
    })
}

pub fn parse_shading(e: &BytesStart<'_>) -> Result<Shading> {
    let fill = match xml::optional_attr(e, b"fill")? {
        Some(ref s) if s.eq_ignore_ascii_case("auto") => Color::Auto,
        Some(ref s) => xml::parse_hex_color(s)
            .map(Color::Rgb)
            .unwrap_or(Color::Auto),
        None => Color::Auto,
    };

    let color = match xml::optional_attr(e, b"color")? {
        Some(ref s) if s.eq_ignore_ascii_case("auto") => Color::Auto,
        Some(ref s) => xml::parse_hex_color(s)
            .map(Color::Rgb)
            .unwrap_or(Color::Auto),
        None => Color::Auto,
    };

    let pattern = match xml::optional_attr(e, b"val")?.as_deref() {
        Some("clear") | None => ShadingPattern::Clear,
        Some("solid") => ShadingPattern::Solid,
        Some("horzStripe") => ShadingPattern::HorzStripe,
        Some("vertStripe") => ShadingPattern::VertStripe,
        Some("reverseDiagStripe") => ShadingPattern::ReverseDiagStripe,
        Some("diagStripe") => ShadingPattern::DiagStripe,
        Some("horzCross") => ShadingPattern::HorzCross,
        Some("diagCross") => ShadingPattern::DiagCross,
        Some("thinHorzStripe") => ShadingPattern::ThinHorzStripe,
        Some("thinVertStripe") => ShadingPattern::ThinVertStripe,
        Some("thinReverseDiagStripe") => ShadingPattern::ThinReverseDiagStripe,
        Some("thinDiagStripe") => ShadingPattern::ThinDiagStripe,
        Some("thinHorzCross") => ShadingPattern::ThinHorzCross,
        Some("thinDiagCross") => ShadingPattern::ThinDiagCross,
        Some("pct5") => ShadingPattern::Pct5,
        Some("pct10") => ShadingPattern::Pct10,
        Some("pct12") => ShadingPattern::Pct12,
        Some("pct15") => ShadingPattern::Pct15,
        Some("pct20") => ShadingPattern::Pct20,
        Some("pct25") => ShadingPattern::Pct25,
        Some("pct30") => ShadingPattern::Pct30,
        Some("pct35") => ShadingPattern::Pct35,
        Some("pct37") => ShadingPattern::Pct37,
        Some("pct40") => ShadingPattern::Pct40,
        Some("pct45") => ShadingPattern::Pct45,
        Some("pct50") => ShadingPattern::Pct50,
        Some("pct55") => ShadingPattern::Pct55,
        Some("pct60") => ShadingPattern::Pct60,
        Some("pct62") => ShadingPattern::Pct62,
        Some("pct65") => ShadingPattern::Pct65,
        Some("pct70") => ShadingPattern::Pct70,
        Some("pct75") => ShadingPattern::Pct75,
        Some("pct80") => ShadingPattern::Pct80,
        Some("pct85") => ShadingPattern::Pct85,
        Some("pct87") => ShadingPattern::Pct87,
        Some("pct90") => ShadingPattern::Pct90,
        Some("pct95") => ShadingPattern::Pct95,
        Some(other) => {
            warn!("unknown shading pattern: {other}");
            ShadingPattern::Clear
        }
    };

    Ok(Shading {
        fill,
        pattern,
        color,
    })
}

fn parse_color_attr(e: &BytesStart<'_>) -> Result<Color> {
    match xml::optional_attr(e, b"val")? {
        Some(ref s) if s.eq_ignore_ascii_case("auto") => Ok(Color::Auto),
        Some(ref s) => Ok(xml::parse_hex_color(s)
            .map(Color::Rgb)
            .unwrap_or(Color::Auto)),
        None => Ok(Color::Auto),
    }
}

fn parse_color_from_attr(e: &BytesStart<'_>) -> Result<Color> {
    match xml::optional_attr(e, b"color")? {
        Some(ref s) if s.eq_ignore_ascii_case("auto") => Ok(Color::Auto),
        Some(ref s) => Ok(xml::parse_hex_color(s)
            .map(Color::Rgb)
            .unwrap_or(Color::Auto)),
        None => Ok(Color::Auto),
    }
}

fn parse_font_set(e: &BytesStart<'_>) -> Result<FontSet> {
    Ok(FontSet {
        ascii: xml::optional_attr(e, b"ascii")?,
        high_ansi: xml::optional_attr(e, b"hAnsi")?,
        east_asian: xml::optional_attr(e, b"eastAsia")?,
        complex_script: xml::optional_attr(e, b"cs")?,
    })
}

fn parse_underline_style(val: &str) -> UnderlineStyle {
    match val {
        "single" => UnderlineStyle::Single,
        "words" => UnderlineStyle::Words,
        "double" => UnderlineStyle::Double,
        "thick" => UnderlineStyle::Thick,
        "dotted" => UnderlineStyle::Dotted,
        "dottedHeavy" => UnderlineStyle::DottedHeavy,
        "dash" => UnderlineStyle::Dash,
        "dashedHeavy" => UnderlineStyle::DashedHeavy,
        "dashLong" => UnderlineStyle::DashLong,
        "dashLongHeavy" => UnderlineStyle::DashLongHeavy,
        "dotDash" => UnderlineStyle::DotDash,
        "dashDotHeavy" => UnderlineStyle::DashDotHeavy,
        "dotDotDash" => UnderlineStyle::DotDotDash,
        "dashDotDotHeavy" => UnderlineStyle::DashDotDotHeavy,
        "wave" => UnderlineStyle::Wave,
        "wavyHeavy" => UnderlineStyle::WavyHeavy,
        "wavyDouble" => UnderlineStyle::WavyDouble,
        "none" => UnderlineStyle::None,
        other => {
            warn!("unknown underline style: {other}");
            UnderlineStyle::None
        }
    }
}

fn parse_vertical_align(val: &str) -> VerticalAlign {
    match val {
        "superscript" => VerticalAlign::Superscript,
        "subscript" => VerticalAlign::Subscript,
        "baseline" => VerticalAlign::Baseline,
        other => {
            warn!("unknown vertical align: {other}");
            VerticalAlign::Baseline
        }
    }
}

fn parse_highlight_color(val: &str) -> Option<HighlightColor> {
    Some(match val {
        "black" => HighlightColor::Black,
        "blue" => HighlightColor::Blue,
        "cyan" => HighlightColor::Cyan,
        "darkBlue" => HighlightColor::DarkBlue,
        "darkCyan" => HighlightColor::DarkCyan,
        "darkGray" => HighlightColor::DarkGray,
        "darkGreen" => HighlightColor::DarkGreen,
        "darkMagenta" => HighlightColor::DarkMagenta,
        "darkRed" => HighlightColor::DarkRed,
        "darkYellow" => HighlightColor::DarkYellow,
        "green" => HighlightColor::Green,
        "lightGray" => HighlightColor::LightGray,
        "magenta" => HighlightColor::Magenta,
        "red" => HighlightColor::Red,
        "white" => HighlightColor::White,
        "yellow" => HighlightColor::Yellow,
        "none" => return None,
        other => {
            warn!("unknown highlight color: {other}");
            return None;
        }
    })
}

fn parse_table_measure(e: &BytesStart<'_>) -> Result<TableMeasure> {
    let w = xml::optional_attr_i64(e, b"w")?.unwrap_or(0);
    let measure_type = xml::optional_attr(e, b"type")?;

    Ok(match measure_type.as_deref() {
        Some("dxa") => TableMeasure::Twips(Dimension::new(w)),
        Some("pct") => TableMeasure::Pct(Dimension::new(w)),
        Some("nil") => TableMeasure::Nil,
        Some("auto") | None => TableMeasure::Auto,
        Some(other) => {
            warn!("unknown table measure type: {other}");
            TableMeasure::Auto
        }
    })
}

fn parse_table_look(e: &BytesStart<'_>) -> Result<TableLook> {
    Ok(TableLook {
        first_row: xml::optional_attr_bool(e, b"firstRow")?,
        last_row: xml::optional_attr_bool(e, b"lastRow")?,
        first_column: xml::optional_attr_bool(e, b"firstColumn")?,
        last_column: xml::optional_attr_bool(e, b"lastColumn")?,
        no_h_band: xml::optional_attr_bool(e, b"noHBand")?,
        no_v_band: xml::optional_attr_bool(e, b"noVBand")?,
    })
}

fn parse_table_borders(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<TableBorders> {
    let mut borders = TableBorders {
        top: None,
        bottom: None,
        left: None,
        right: None,
        inside_h: None,
        inside_v: None,
    };

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                let border = parse_border(e)?;
                match local.as_slice() {
                    b"top" => borders.top = Some(border),
                    b"bottom" => borders.bottom = Some(border),
                    b"left" | b"start" => borders.left = Some(border),
                    b"right" | b"end" => borders.right = Some(border),
                    b"insideH" => borders.inside_h = Some(border),
                    b"insideV" => borders.inside_v = Some(border),
                    _ => xml::warn_unsupported_element("tblBorders", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tblBorders" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(borders)
}

fn parse_table_cell_borders(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<TableCellBorders> {
    let mut borders = TableCellBorders {
        top: None,
        bottom: None,
        left: None,
        right: None,
        inside_h: None,
        inside_v: None,
        tl2br: None,
        tr2bl: None,
    };

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                let border = parse_border(e)?;
                match local.as_slice() {
                    b"top" => borders.top = Some(border),
                    b"bottom" => borders.bottom = Some(border),
                    b"left" | b"start" => borders.left = Some(border),
                    b"right" | b"end" => borders.right = Some(border),
                    b"insideH" => borders.inside_h = Some(border),
                    b"insideV" => borders.inside_v = Some(border),
                    b"tl2br" => borders.tl2br = Some(border),
                    b"tr2bl" => borders.tr2bl = Some(border),
                    _ => xml::warn_unsupported_element("tcBorders", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tcBorders" => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(borders)
}

fn parse_edge_insets_twips(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<EdgeInsets<crate::dimension::Twips>> {
    let mut insets = EdgeInsets::ZERO;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                let w = xml::optional_attr_i64(e, b"w")?.unwrap_or(0);
                match local.as_slice() {
                    b"top" => insets.top = Dimension::new(w),
                    b"bottom" => insets.bottom = Dimension::new(w),
                    b"left" | b"start" => insets.left = Dimension::new(w),
                    b"right" | b"end" => insets.right = Dimension::new(w),
                    _ => xml::warn_unsupported_element("margins", &local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(insets)
}

fn parse_cell_vertical_align(val: &str) -> CellVerticalAlign {
    match val {
        "top" => CellVerticalAlign::Top,
        "center" => CellVerticalAlign::Center,
        "bottom" => CellVerticalAlign::Bottom,
        "both" => CellVerticalAlign::Both,
        other => {
            warn!("unknown cell vertical align: {other}");
            CellVerticalAlign::Top
        }
    }
}

fn parse_text_direction(val: &str) -> TextDirection {
    match val {
        "lrTb" => TextDirection::LeftToRightTopToBottom,
        "tbRl" => TextDirection::TopToBottomRightToLeft,
        "btLr" => TextDirection::BottomToTopLeftToRight,
        "lrTbV" => TextDirection::LeftToRightTopToBottomRotated,
        "tbRlV" => TextDirection::TopToBottomRightToLeftRotated,
        "tbLrV" => TextDirection::TopToBottomLeftToRightRotated,
        other => {
            warn!("unknown text direction: {other}");
            TextDirection::LeftToRightTopToBottom
        }
    }
}

fn parse_section_type(val: &str) -> SectionType {
    match val {
        "nextPage" => SectionType::NextPage,
        "continuous" => SectionType::Continuous,
        "evenPage" => SectionType::EvenPage,
        "oddPage" => SectionType::OddPage,
        "nextColumn" => SectionType::NextColumn,
        other => {
            warn!("unknown section type: {other}");
            SectionType::NextPage
        }
    }
}
