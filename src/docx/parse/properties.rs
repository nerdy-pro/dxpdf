//! Parsers for OOXML property elements: pPr, rPr, tblPr, trPr, tcPr, sectPr.
//!
//! Each parser consumes events from the reader until the corresponding End event,
//! returning a fully-populated properties struct.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::dimension::Dimension;
use crate::docx::error::{ParseError, Result};
use crate::docx::geometry::EdgeInsets;
use crate::docx::model::*;

fn invalid_value(attr: &str, value: &str) -> ParseError {
    ParseError::InvalidAttributeValue {
        attr: attr.to_string(),
        value: value.to_string(),
        reason: "unsupported value per OOXML spec".to_string(),
    }
}
use crate::docx::xml;

/// Read a `w:val` boolean attribute, defaulting to `true` when the attribute is
/// absent (OOXML §17.17.4 "toggle property" semantics: presence alone means on).
///
/// Returns `Ok(Some(bool))` always; `Ok(None)` is never produced because toggle
/// properties with no `val` mean `true`, and absence of the element itself is
/// handled at the call site (the arm simply won't match).
#[inline]
fn toggle_attr(e: &BytesStart<'_>) -> Result<Option<bool>> {
    Ok(Some(xml::optional_attr_bool(e, b"val")?.unwrap_or(true)))
}

/// Read a `w:val` string attribute and map it through `f`, returning
/// `Ok(None)` when the attribute is absent.
///
/// Equivalent to `xml::optional_attr(e, b"val")?.map(|v| f(&v)).transpose()`.
#[inline]
fn opt_val<T, F>(e: &BytesStart<'_>, f: F) -> Result<Option<T>>
where
    F: FnOnce(&str) -> Result<T>,
{
    xml::optional_attr(e, b"val")?.map(|v| f(&v)).transpose()
}

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
                    b"jc" => { props.alignment = opt_val(e, parse_alignment)?; }
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
                    b"keepNext"           => { props.keep_next           = toggle_attr(e)?; }
                    b"keepLines"          => { props.keep_lines          = toggle_attr(e)?; }
                    b"widowControl"       => { props.widow_control       = toggle_attr(e)?; }
                    b"pageBreakBefore"    => { props.page_break_before   = toggle_attr(e)?; }
                    b"suppressAutoHyphens"=> { props.suppress_auto_hyphens = toggle_attr(e)?; }
                    b"contextualSpacing"  => { props.contextual_spacing  = toggle_attr(e)?; }
                    b"bidi"               => { props.bidi                = toggle_attr(e)?; }
                    b"wordWrap"           => { props.word_wrap           = toggle_attr(e)?; }
                    b"textAlignment"      => { props.text_alignment      = opt_val(e, parse_text_alignment)?; }
                    b"cnfStyle"           => { props.cnf_style           = Some(parse_cnf_style(e)?); }
                    b"framePr"            => { props.frame_properties    = Some(parse_frame_kind(e)?); }
                    b"autoSpaceDE"        => { props.auto_space_de       = toggle_attr(e)?; }
                    b"autoSpaceDN"        => { props.auto_space_dn       = toggle_attr(e)?; }
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
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"rStyle" => {
                        style_id = xml::optional_attr(e, b"val")?.map(StyleId::new);
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
                    b"b"   => { props.bold   = toggle_attr(e)?; }
                    b"bCs" => {}
                    b"i"   => { props.italic = toggle_attr(e)?; }
                    b"iCs" => {}
                    b"u" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            props.underline = Some(parse_underline_style(&val)?);
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
                            props.highlight = parse_highlight_color(&val)?;
                        }
                    }
                    b"shd" => {
                        props.shading = Some(parse_shading(e)?);
                    }
                    b"vertAlign" => { props.vertical_align = opt_val(e, parse_vertical_align)?; }
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
                    b"caps"      => { props.all_caps   = toggle_attr(e)?; }
                    b"smallCaps" => { props.small_caps = toggle_attr(e)?; }
                    b"vanish"    => { props.vanish     = toggle_attr(e)?; }
                    b"noProof"   => { props.no_proof   = toggle_attr(e)?; }
                    b"webHidden" => { props.web_hidden = toggle_attr(e)?; }
                    b"position" => {
                        if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                            props.position = Some(Dimension::new(val));
                        }
                    }
                    b"lang" => {
                        props.lang = Some(Lang {
                            val: xml::optional_attr(e, b"val")?,
                            east_asia: xml::optional_attr(e, b"eastAsia")?,
                            bidi: xml::optional_attr(e, b"bidi")?,
                        });
                    }
                    b"rtl"     => { props.rtl     = toggle_attr(e)?; }
                    b"emboss"  => { props.emboss  = toggle_attr(e)?; }
                    b"imprint" => { props.imprint = toggle_attr(e)?; }
                    b"outline" => { props.outline = toggle_attr(e)?; }
                    b"shadow"  => { props.shadow  = toggle_attr(e)?; }
                    b"bdr" => {
                        props.border = Some(parse_border(e)?);
                    }
                    _ => xml::warn_unsupported_element("rPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"rPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"rPr")),
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
        let event = xml::next_event(reader, buf)?;
        let is_start = matches!(event, Event::Start(_));
        match event {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"tblStyle" => {
                        props.style_id = xml::optional_attr(e, b"val")?.map(StyleId::new);
                        style_id.clone_from(&props.style_id);
                    }
                    // Start-only: have child elements
                    b"tblBorders" if is_start => {
                        props.borders = Some(parse_table_borders(reader, buf)?);
                    }
                    b"tblCellMar" if is_start => {
                        props.cell_margins =
                            Some(parse_edge_insets_twips(reader, buf, b"tblCellMar")?);
                    }
                    // attrs-only — valid as Start or Empty
                    b"jc"               => { props.alignment          = opt_val(e, parse_alignment)?; }
                    b"tblW"             => { props.width               = Some(parse_table_measure(e)?); }
                    b"tblLayout"        => { props.layout              = opt_val(e, |v| Ok(match v {
                        "fixed" => TableLayout::Fixed,
                        "autofit" | "auto" => TableLayout::Auto,
                        other => return Err(invalid_value("tblLayout/val", other)),
                    }))?; }
                    b"tblInd"           => { props.indent              = Some(parse_table_measure(e)?); }
                    b"tblCellSpacing"   => { props.cell_spacing        = Some(parse_table_measure(e)?); }
                    b"tblLook"          => { props.look                = Some(parse_table_look(e)?); }
                    b"tblStyleRowBandSize" => { props.style_row_band_size = xml::optional_attr_u32(e, b"val")?; }
                    b"tblStyleColBandSize" => { props.style_col_band_size = xml::optional_attr_u32(e, b"val")?; }
                    b"tblpPr"           => { props.positioning         = Some(parse_table_positioning(e)?); }
                    b"tblOverlap"       => { props.overlap             = opt_val(e, parse_table_overlap)?; }
                    _ => xml::warn_unsupported_element("tblPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tblPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tblPr")),
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
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
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
                    b"tblHeader" => { props.is_header   = toggle_attr(e)?; }
                    b"cantSplit"  => { props.cant_split  = toggle_attr(e)?; }
                    b"jc"         => { props.justification = opt_val(e, parse_alignment)?; }
                    b"cnfStyle" => {
                        props.cnf_style = Some(parse_cnf_style(e)?);
                    }
                    b"gridAfter" => {
                        props.grid_after = xml::optional_attr_u32(e, b"val")?;
                    }
                    b"wAfter" => {
                        props.w_after = Some(parse_table_measure(e)?);
                    }
                    _ => xml::warn_unsupported_element("trPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"trPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"trPr")),
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
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"tcBorders" => {
                        props.borders = Some(parse_table_cell_borders(reader, buf)?);
                    }
                    b"tcMar" => {
                        props.margins = Some(parse_edge_insets_twips(reader, buf, b"tcMar")?);
                    }
                    _ => xml::warn_unsupported_element("tcPr", local),
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"tcW" => {
                        props.width = Some(parse_table_measure(e)?);
                    }
                    b"shd" => {
                        props.shading = Some(parse_shading(e)?);
                    }
                    b"vAlign" => { props.vertical_align = opt_val(e, parse_cell_vertical_align)?; }
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
                    b"textDirection" => { props.text_direction = opt_val(e, parse_text_direction)?; }
                    b"noWrap"        => { props.no_wrap        = toggle_attr(e)?; }
                    b"cnfStyle" => {
                        props.cnf_style = Some(parse_cnf_style(e)?);
                    }
                    _ => xml::warn_unsupported_element("tcPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tcPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tcPr")),
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
        let event = xml::next_event(reader, buf)?;
        let is_start = matches!(event, Event::Start(_));
        match event {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"pgSz" => {
                        let orientation = match xml::optional_attr(e, b"orient")?.as_deref() {
                            Some("landscape") => Some(PageOrientation::Landscape),
                            Some("portrait") => Some(PageOrientation::Portrait),
                            Some(other) => return Err(invalid_value("orient", other)),
                            None => None,
                        };
                        props.page_size = Some(PageSize {
                            width: xml::optional_attr_i64(e, b"w")?.map(Dimension::new),
                            height: xml::optional_attr_i64(e, b"h")?.map(Dimension::new),
                            orientation,
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
                        let columns = if is_start {
                            parse_column_definitions(reader, buf)?
                        } else {
                            Vec::new()
                        };
                        props.columns = Some(Columns {
                            count: xml::optional_attr_u32(e, b"num")?,
                            space: xml::optional_attr_i64(e, b"space")?.map(Dimension::new),
                            equal_width: xml::optional_attr_bool(e, b"equalWidth")?,
                            columns,
                        });
                    }
                    b"headerReference" => {
                        if let Some(r_id) = xml::optional_attr(e, b"id")? {
                            let hf_type = xml::optional_attr(e, b"type")?;
                            let rel = RelId::new(r_id);
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
                            let rel = RelId::new(r_id);
                            match hf_type.as_deref() {
                                Some("first") => props.footer_refs.first = Some(rel),
                                Some("even") => props.footer_refs.even = Some(rel),
                                _ => props.footer_refs.default = Some(rel),
                            }
                        }
                    }
                    b"docGrid" => {
                        let grid_type = match xml::optional_attr(e, b"type")?.as_deref() {
                            Some("default") => Some(DocGridType::Default),
                            Some("lines") => Some(DocGridType::Lines),
                            Some("linesAndChars") => Some(DocGridType::LinesAndChars),
                            Some("snapToChars") => Some(DocGridType::SnapToChars),
                            Some(other) => return Err(invalid_value("docGrid/type", other)),
                            None => None,
                        };
                        props.doc_grid = Some(DocGrid {
                            grid_type,
                            line_pitch: xml::optional_attr_i64(e, b"linePitch")?
                                .map(Dimension::new),
                            char_space: xml::optional_attr_i64(e, b"charSpace")?
                                .map(Dimension::new),
                        });
                    }
                    b"titlePg"   => { props.title_page      = toggle_attr(e)?; }
                    b"pgNumType" => { props.page_number_type = Some(parse_page_number_type(e)?); }
                    b"type"      => { props.section_type     = opt_val(e, parse_section_type)?; }
                    _ => xml::warn_unsupported_element("sectPr", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"sectPr" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"sectPr")),
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

/// §17.18.44 ST_Jc
pub fn parse_alignment(val: &str) -> Result<Alignment> {
    match val {
        "start" | "left" => Ok(Alignment::Start),
        "center" => Ok(Alignment::Center),
        "end" | "right" => Ok(Alignment::End),
        "both" | "justify" => Ok(Alignment::Both),
        "distribute" => Ok(Alignment::Distribute),
        "thaiDistribute" => Ok(Alignment::Thai),
        other => Err(invalid_value("jc", other)),
    }
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
        Some(other) => return Err(invalid_value("border/val", other)),
    };

    let sz = xml::optional_attr_i64(e, b"sz")?.unwrap_or(0);
    let space = xml::optional_attr_i64(e, b"space")?.unwrap_or(0);
    let color = parse_color_from_attr(e)?;

    Ok(Border {
        style,
        width: Dimension::new(sz),
        // §17.3.4: w:space is ST_PointMeasure (§17.18.68).
        space: Dimension::new(space),
        color,
    })
}

pub fn parse_shading(e: &BytesStart<'_>) -> Result<Shading> {
    let fill = match xml::optional_attr(e, b"fill")? {
        Some(ref s) if s.eq_ignore_ascii_case("auto") => Color::Auto,
        Some(ref s) => xml::parse_hex_color(s)
            .map(Color::Rgb)
            .ok_or_else(|| invalid_value("color", s))?,
        None => Color::Auto,
    };

    let color = match xml::optional_attr(e, b"color")? {
        Some(ref s) if s.eq_ignore_ascii_case("auto") => Color::Auto,
        Some(ref s) => xml::parse_hex_color(s)
            .map(Color::Rgb)
            .ok_or_else(|| invalid_value("color", s))?,
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
        Some(other) => return Err(invalid_value("shd/val", other)),
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
            .ok_or_else(|| invalid_value("color", s))?),
        None => Ok(Color::Auto),
    }
}

fn parse_color_from_attr(e: &BytesStart<'_>) -> Result<Color> {
    match xml::optional_attr(e, b"color")? {
        Some(ref s) if s.eq_ignore_ascii_case("auto") => Ok(Color::Auto),
        Some(ref s) => Ok(xml::parse_hex_color(s)
            .map(Color::Rgb)
            .ok_or_else(|| invalid_value("color", s))?),
        None => Ok(Color::Auto),
    }
}

fn parse_font_set(e: &BytesStart<'_>) -> Result<FontSet> {
    Ok(FontSet {
        ascii: FontSlot {
            explicit: xml::optional_attr(e, b"ascii")?,
            theme: xml::optional_attr(e, b"asciiTheme")?.and_then(|v| parse_theme_font_ref(&v)),
        },
        high_ansi: FontSlot {
            explicit: xml::optional_attr(e, b"hAnsi")?,
            theme: xml::optional_attr(e, b"hAnsiTheme")?.and_then(|v| parse_theme_font_ref(&v)),
        },
        east_asian: FontSlot {
            explicit: xml::optional_attr(e, b"eastAsia")?,
            theme: xml::optional_attr(e, b"eastAsiaTheme")?.and_then(|v| parse_theme_font_ref(&v)),
        },
        complex_script: FontSlot {
            explicit: xml::optional_attr(e, b"cs")?,
            theme: xml::optional_attr(e, b"cstheme")?.and_then(|v| parse_theme_font_ref(&v)),
        },
    })
}

/// §17.18.84 ST_Theme
fn parse_theme_font_ref(val: &str) -> Option<ThemeFontRef> {
    match val {
        "majorHAnsi" => Some(ThemeFontRef::MajorHAnsi),
        "majorEastAsia" => Some(ThemeFontRef::MajorEastAsia),
        "majorBidi" => Some(ThemeFontRef::MajorBidi),
        "minorHAnsi" => Some(ThemeFontRef::MinorHAnsi),
        "minorEastAsia" => Some(ThemeFontRef::MinorEastAsia),
        "minorBidi" => Some(ThemeFontRef::MinorBidi),
        _ => None,
    }
}

/// §17.18.99 ST_Underline
fn parse_underline_style(val: &str) -> Result<UnderlineStyle> {
    match val {
        "single" => Ok(UnderlineStyle::Single),
        "words" => Ok(UnderlineStyle::Words),
        "double" => Ok(UnderlineStyle::Double),
        "thick" => Ok(UnderlineStyle::Thick),
        "dotted" => Ok(UnderlineStyle::Dotted),
        "dottedHeavy" => Ok(UnderlineStyle::DottedHeavy),
        "dash" => Ok(UnderlineStyle::Dash),
        "dashedHeavy" => Ok(UnderlineStyle::DashedHeavy),
        "dashLong" => Ok(UnderlineStyle::DashLong),
        "dashLongHeavy" => Ok(UnderlineStyle::DashLongHeavy),
        "dotDash" => Ok(UnderlineStyle::DotDash),
        "dashDotHeavy" => Ok(UnderlineStyle::DashDotHeavy),
        "dotDotDash" => Ok(UnderlineStyle::DotDotDash),
        "dashDotDotHeavy" => Ok(UnderlineStyle::DashDotDotHeavy),
        "wave" => Ok(UnderlineStyle::Wave),
        "wavyHeavy" => Ok(UnderlineStyle::WavyHeavy),
        "wavyDouble" => Ok(UnderlineStyle::WavyDouble),
        "none" => Ok(UnderlineStyle::None),
        other => Err(invalid_value("u/val", other)),
    }
}

/// §17.18.100 ST_VerticalAlignRun
fn parse_vertical_align(val: &str) -> Result<VerticalAlign> {
    match val {
        "baseline" => Ok(VerticalAlign::Baseline),
        "superscript" => Ok(VerticalAlign::Superscript),
        "subscript" => Ok(VerticalAlign::Subscript),
        other => Err(invalid_value("vertAlign/val", other)),
    }
}

/// §17.18.40 ST_HighlightColor
fn parse_highlight_color(val: &str) -> Result<Option<HighlightColor>> {
    match val {
        "black" => Ok(Some(HighlightColor::Black)),
        "blue" => Ok(Some(HighlightColor::Blue)),
        "cyan" => Ok(Some(HighlightColor::Cyan)),
        "darkBlue" => Ok(Some(HighlightColor::DarkBlue)),
        "darkCyan" => Ok(Some(HighlightColor::DarkCyan)),
        "darkGray" => Ok(Some(HighlightColor::DarkGray)),
        "darkGreen" => Ok(Some(HighlightColor::DarkGreen)),
        "darkMagenta" => Ok(Some(HighlightColor::DarkMagenta)),
        "darkRed" => Ok(Some(HighlightColor::DarkRed)),
        "darkYellow" => Ok(Some(HighlightColor::DarkYellow)),
        "green" => Ok(Some(HighlightColor::Green)),
        "lightGray" => Ok(Some(HighlightColor::LightGray)),
        "magenta" => Ok(Some(HighlightColor::Magenta)),
        "red" => Ok(Some(HighlightColor::Red)),
        "white" => Ok(Some(HighlightColor::White)),
        "yellow" => Ok(Some(HighlightColor::Yellow)),
        "none" => Ok(None),
        other => Err(invalid_value("highlight/val", other)),
    }
}

/// §17.18.96 ST_TblWidth
fn parse_table_measure(e: &BytesStart<'_>) -> Result<TableMeasure> {
    let w = xml::optional_attr_i64(e, b"w")?.unwrap_or(0);
    let measure_type = xml::optional_attr(e, b"type")?;

    match measure_type.as_deref() {
        Some("dxa") => Ok(TableMeasure::Twips(Dimension::new(w))),
        Some("pct") => Ok(TableMeasure::Pct(Dimension::new(w))),
        Some("nil") => Ok(TableMeasure::Nil),
        Some("auto") | None => Ok(TableMeasure::Auto),
        Some(other) => Err(invalid_value("type", other)),
    }
}

/// §17.6.3: parse `w:col` children inside `w:cols`.
fn parse_column_definitions(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Vec<ColumnDefinition>> {
    let mut cols = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"col" => {
                        cols.push(ColumnDefinition {
                            width: xml::optional_attr_i64(e, b"w")?.map(Dimension::new),
                            space: xml::optional_attr_i64(e, b"space")?.map(Dimension::new),
                        });
                    }
                    _ => xml::warn_unsupported_element("cols", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"cols" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"cols")),
            _ => {}
        }
    }

    Ok(cols)
}

/// §17.4.56: parse tblLook. Supports both individual attributes (OOXML 2010+)
/// and the legacy `val` bit field format.
fn parse_table_look(e: &BytesStart<'_>) -> Result<TableLook> {
    let first_row = xml::optional_attr_bool(e, b"firstRow")?;
    let last_row = xml::optional_attr_bool(e, b"lastRow")?;
    let first_column = xml::optional_attr_bool(e, b"firstColumn")?;
    let last_column = xml::optional_attr_bool(e, b"lastColumn")?;
    let no_h_band = xml::optional_attr_bool(e, b"noHBand")?;
    let no_v_band = xml::optional_attr_bool(e, b"noVBand")?;

    // If individual attributes are present, use them directly.
    if first_row.is_some()
        || last_row.is_some()
        || first_column.is_some()
        || last_column.is_some()
        || no_h_band.is_some()
        || no_v_band.is_some()
    {
        return Ok(TableLook {
            first_row,
            last_row,
            first_column,
            last_column,
            no_h_band,
            no_v_band,
        });
    }

    // §17.4.56: legacy `val` attribute is a hex bit field.
    // Bit 0x0020: firstRow
    // Bit 0x0040: lastRow
    // Bit 0x0080: firstColumn
    // Bit 0x0100: lastColumn
    // Bit 0x0200: noHBand (horizontal banding disabled)
    // Bit 0x0400: noVBand (vertical banding disabled)
    if let Some(val_str) = xml::optional_attr(e, b"val")? {
        let val = u32::from_str_radix(&val_str, 16).unwrap_or(0);
        return Ok(TableLook {
            first_row: Some(val & 0x0020 != 0),
            last_row: Some(val & 0x0040 != 0),
            first_column: Some(val & 0x0080 != 0),
            last_column: Some(val & 0x0100 != 0),
            no_h_band: Some(val & 0x0200 != 0),
            no_v_band: Some(val & 0x0400 != 0),
        });
    }

    Ok(TableLook::default())
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
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                let border = parse_border(e)?;
                match local {
                    b"top" => borders.top = Some(border),
                    b"bottom" => borders.bottom = Some(border),
                    b"left" | b"start" => borders.left = Some(border),
                    b"right" | b"end" => borders.right = Some(border),
                    b"insideH" => borders.inside_h = Some(border),
                    b"insideV" => borders.inside_v = Some(border),
                    _ => xml::warn_unsupported_element("tblBorders", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tblBorders" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tblBorders")),
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
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                let border = parse_border(e)?;
                match local {
                    b"top" => borders.top = Some(border),
                    b"bottom" => borders.bottom = Some(border),
                    b"left" | b"start" => borders.left = Some(border),
                    b"right" | b"end" => borders.right = Some(border),
                    b"insideH" => borders.inside_h = Some(border),
                    b"insideV" => borders.inside_v = Some(border),
                    b"tl2br" => borders.tl2br = Some(border),
                    b"tr2bl" => borders.tr2bl = Some(border),
                    _ => xml::warn_unsupported_element("tcBorders", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"tcBorders" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"tcBorders")),
            _ => {}
        }
    }

    Ok(borders)
}

fn parse_edge_insets_twips(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<EdgeInsets<crate::docx::dimension::Twips>> {
    let mut insets = EdgeInsets::ZERO;

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                let w = xml::optional_attr_i64(e, b"w")?.unwrap_or(0);
                match local {
                    b"top" => insets.top = Dimension::new(w),
                    b"bottom" => insets.bottom = Dimension::new(w),
                    b"left" | b"start" => insets.left = Dimension::new(w),
                    b"right" | b"end" => insets.right = Dimension::new(w),
                    _ => xml::warn_unsupported_element("margins", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => return Err(xml::unexpected_eof(b"container")),
            _ => {}
        }
    }

    Ok(insets)
}

/// §17.18.101 ST_VerticalJc
fn parse_cell_vertical_align(val: &str) -> Result<CellVerticalAlign> {
    match val {
        "top" => Ok(CellVerticalAlign::Top),
        "center" => Ok(CellVerticalAlign::Center),
        "bottom" => Ok(CellVerticalAlign::Bottom),
        "both" => Ok(CellVerticalAlign::Both),
        other => Err(invalid_value("vAlign/val", other)),
    }
}

/// §17.18.93 ST_TextDirection
fn parse_text_direction(val: &str) -> Result<TextDirection> {
    match val {
        "lrTb" => Ok(TextDirection::LeftToRightTopToBottom),
        "tbRl" => Ok(TextDirection::TopToBottomRightToLeft),
        "btLr" => Ok(TextDirection::BottomToTopLeftToRight),
        "lrTbV" => Ok(TextDirection::LeftToRightTopToBottomRotated),
        "tbRlV" => Ok(TextDirection::TopToBottomRightToLeftRotated),
        "tbLrV" => Ok(TextDirection::TopToBottomLeftToRightRotated),
        other => Err(invalid_value("textDirection/val", other)),
    }
}

/// §17.18.77 ST_SectionMark
fn parse_section_type(val: &str) -> Result<SectionType> {
    match val {
        "nextPage" => Ok(SectionType::NextPage),
        "continuous" => Ok(SectionType::Continuous),
        "evenPage" => Ok(SectionType::EvenPage),
        "oddPage" => Ok(SectionType::OddPage),
        "nextColumn" => Ok(SectionType::NextColumn),
        other => Err(invalid_value("type/val", other)),
    }
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

/// §17.3.1.8: parse `w:cnfStyle` element attributes into a [`CnfStyle`] flag set.
///
/// The `val` binary string is parsed first (positions 0–11 → flags), then any
/// explicit individual attributes override the corresponding bits, allowing
/// producers that omit `val` to be handled correctly.
fn parse_cnf_style(e: &BytesStart<'_>) -> Result<CnfStyle> {
    // Seed from the legacy 12-char binary string if present.
    let mut flags = match xml::optional_attr(e, b"val")? {
        Some(s) => CnfStyle::from_val_str(&s),
        None => CnfStyle::empty(),
    };

    // Individual attributes take precedence over the `val` string.
    let pairs: &[(&[u8], CnfStyle)] = &[
        (b"firstRow",            CnfStyle::FIRST_ROW),
        (b"lastRow",             CnfStyle::LAST_ROW),
        (b"firstColumn",         CnfStyle::FIRST_COLUMN),
        (b"lastColumn",          CnfStyle::LAST_COLUMN),
        (b"oddVBand",            CnfStyle::ODD_V_BAND),
        (b"evenVBand",           CnfStyle::EVEN_V_BAND),
        (b"oddHBand",            CnfStyle::ODD_H_BAND),
        (b"evenHBand",           CnfStyle::EVEN_H_BAND),
        (b"firstRowFirstColumn", CnfStyle::FIRST_ROW_FIRST_COLUMN),
        (b"firstRowLastColumn",  CnfStyle::FIRST_ROW_LAST_COLUMN),
        (b"lastRowFirstColumn",  CnfStyle::LAST_ROW_FIRST_COLUMN),
        (b"lastRowLastColumn",   CnfStyle::LAST_ROW_LAST_COLUMN),
    ];
    for &(attr, flag) in pairs {
        match xml::optional_attr_bool(e, attr)? {
            Some(true)  => flags |= flag,
            Some(false) => flags &= !flag,
            None        => {}   // absent — leave the val-seeded bit unchanged
        }
    }

    Ok(flags)
}

/// §17.4.58: parse `w:tblpPr` attributes.
fn parse_table_positioning(e: &BytesStart<'_>) -> Result<TablePositioning> {
    Ok(TablePositioning {
        left_from_text: xml::optional_attr_i64(e, b"leftFromText")?.map(Dimension::new),
        right_from_text: xml::optional_attr_i64(e, b"rightFromText")?.map(Dimension::new),
        top_from_text: xml::optional_attr_i64(e, b"topFromText")?.map(Dimension::new),
        bottom_from_text: xml::optional_attr_i64(e, b"bottomFromText")?.map(Dimension::new),
        vert_anchor: match xml::optional_attr(e, b"vertAnchor")?.as_deref() {
            Some("text") => Some(TableAnchor::Text),
            Some("margin") => Some(TableAnchor::Margin),
            Some("page") => Some(TableAnchor::Page),
            Some(other) => return Err(invalid_value("tblpPr/vertAnchor", other)),
            None => None,
        },
        horz_anchor: match xml::optional_attr(e, b"horzAnchor")?.as_deref() {
            Some("text") => Some(TableAnchor::Text),
            Some("margin") => Some(TableAnchor::Margin),
            Some("page") => Some(TableAnchor::Page),
            Some(other) => return Err(invalid_value("tblpPr/horzAnchor", other)),
            None => None,
        },
        x_align: match xml::optional_attr(e, b"tblpXSpec")?.as_deref() {
            Some("left") => Some(TableXAlign::Left),
            Some("center") => Some(TableXAlign::Center),
            Some("right") => Some(TableXAlign::Right),
            Some("inside") => Some(TableXAlign::Inside),
            Some("outside") => Some(TableXAlign::Outside),
            Some(other) => return Err(invalid_value("tblpPr/tblpXSpec", other)),
            None => None,
        },
        y_align: match xml::optional_attr(e, b"tblpYSpec")?.as_deref() {
            Some("top") => Some(TableYAlign::Top),
            Some("center") => Some(TableYAlign::Center),
            Some("bottom") => Some(TableYAlign::Bottom),
            Some("inside") => Some(TableYAlign::Inside),
            Some("outside") => Some(TableYAlign::Outside),
            Some("inline") => Some(TableYAlign::Inline),
            Some(other) => return Err(invalid_value("tblpPr/tblpYSpec", other)),
            None => None,
        },
        x: xml::optional_attr_i64(e, b"tblpX")?.map(Dimension::new),
        y: xml::optional_attr_i64(e, b"tblpY")?.map(Dimension::new),
    })
}

/// §17.4.56 ST_TblOverlap
fn parse_table_overlap(val: &str) -> Result<TableOverlap> {
    match val {
        "overlap" => Ok(TableOverlap::Overlap),
        "never" => Ok(TableOverlap::Never),
        other => Err(invalid_value("tblOverlap/val", other)),
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

/// §17.6.12: parse `w:pgNumType` attributes.
fn parse_page_number_type(e: &BytesStart<'_>) -> Result<PageNumberType> {
    Ok(PageNumberType {
        format: match xml::optional_attr(e, b"fmt")?.as_deref() {
            Some("decimal") => Some(NumberFormat::Decimal),
            Some("upperRoman") => Some(NumberFormat::UpperRoman),
            Some("lowerRoman") => Some(NumberFormat::LowerRoman),
            Some("upperLetter") => Some(NumberFormat::UpperLetter),
            Some("lowerLetter") => Some(NumberFormat::LowerLetter),
            Some("ordinal") => Some(NumberFormat::Ordinal),
            Some("cardinalText") => Some(NumberFormat::CardinalText),
            Some("ordinalText") => Some(NumberFormat::OrdinalText),
            Some("none") => Some(NumberFormat::None),
            Some(other) => return Err(invalid_value("pgNumType/fmt", other)),
            None => None,
        },
        start: xml::optional_attr_u32(e, b"start")?,
        chap_style: xml::optional_attr_u32(e, b"chapStyle")?,
        chap_sep: match xml::optional_attr(e, b"chapSep")?.as_deref() {
            Some("hyphen") => Some(ChapterSeparator::Hyphen),
            Some("period") => Some(ChapterSeparator::Period),
            Some("colon") => Some(ChapterSeparator::Colon),
            Some("emDash") => Some(ChapterSeparator::EmDash),
            Some("enDash") => Some(ChapterSeparator::EnDash),
            Some(other) => return Err(invalid_value("pgNumType/chapSep", other)),
            None => None,
        },
    })
}
