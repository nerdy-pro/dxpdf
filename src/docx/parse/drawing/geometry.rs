//! Parser for `<a:custGeom>` (§20.1.9.8 CT_CustomGeometry2D).
//!
//! Reads `avLst` / `gdLst` / `ahLst` / `cxnLst` / `rect` / `pathLst` children
//! into a `CustomGeometry`. Path verbs (`moveTo`, `lnTo`, `cubicBezTo`,
//! `quadBezTo`, `arcTo`, `close`) are preserved in document order.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::docx::dimension::Dimension;
use crate::docx::error::{ParseError, Result};
use crate::docx::model::{
    AdjCoord, AdjPoint, AdjustHandle, ConnectionSite, CustomGeometry, GeomGuide, PathCommand,
    PathDef, PathFillMode, TextRect,
};
use crate::docx::xml;

/// Parse `<a:custGeom>…</a:custGeom>`. Reader is positioned after the Start
/// event.
pub fn parse_custom_geometry(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<CustomGeometry> {
    let mut geom = CustomGeometry::default();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local_owned = xml::local_name_owned(e.name().as_ref());
                match &*local_owned {
                    b"avLst" => geom.av_list = parse_guide_list(reader, buf, b"avLst")?,
                    b"gdLst" => geom.gd_list = parse_guide_list(reader, buf, b"gdLst")?,
                    b"ahLst" => geom.ah_list = parse_adjust_handle_list(reader, buf)?,
                    b"cxnLst" => geom.cxn_list = parse_connection_site_list(reader, buf)?,
                    b"rect" => {
                        geom.rect = Some(parse_text_rect(e)?);
                        xml::skip_to_end(reader, buf, b"rect")?;
                    }
                    b"pathLst" => geom.paths = parse_path_list(reader, buf)?,
                    _ => {
                        xml::warn_unsupported_element("custGeom", &local_owned);
                        xml::skip_to_end(reader, buf, &local_owned)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"avLst" | b"gdLst" | b"ahLst" | b"cxnLst" | b"pathLst" => {
                        // Empty container — nothing to collect.
                    }
                    b"rect" => geom.rect = Some(parse_text_rect(e)?),
                    _ => xml::warn_unsupported_element("custGeom", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"custGeom" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"custGeom")),
            _ => {}
        }
    }

    Ok(geom)
}

// ── Guide list (avLst / gdLst share shape) ──────────────────────────────────

fn parse_guide_list(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<Vec<GeomGuide>> {
    let mut guides = Vec::new();
    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e)
                if xml::local_name(e.name().as_ref()) == b"gd" =>
            {
                guides.push(parse_geom_guide(e)?);
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => return Err(xml::unexpected_eof(end_tag)),
            _ => {}
        }
    }
    Ok(guides)
}

fn parse_geom_guide(e: &BytesStart<'_>) -> Result<GeomGuide> {
    Ok(GeomGuide {
        name: xml::required_attr(e, b"name")?,
        formula: xml::required_attr(e, b"fmla")?,
    })
}

// ── Adjust handles (§20.1.9.1) ──────────────────────────────────────────────

fn parse_adjust_handle_list(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Vec<AdjustHandle>> {
    let mut handles = Vec::new();
    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local_owned = xml::local_name_owned(e.name().as_ref());
                match &*local_owned {
                    b"ahXY" => handles.push(parse_ah_xy(reader, buf, e)?),
                    b"ahPolar" => handles.push(parse_ah_polar(reader, buf, e)?),
                    _ => {
                        xml::warn_unsupported_element("ahLst", &local_owned);
                        xml::skip_to_end(reader, buf, &local_owned)?;
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"ahLst" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"ahLst")),
            _ => {}
        }
    }
    Ok(handles)
}

fn parse_ah_xy(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    start: &BytesStart<'_>,
) -> Result<AdjustHandle> {
    let guide_ref_x = xml::optional_attr(start, b"gdRefX")?;
    let guide_ref_y = xml::optional_attr(start, b"gdRefY")?;
    let min_x = xml::optional_attr(start, b"minX")?
        .map(|s| parse_adj_coord(&s))
        .transpose()?;
    let max_x = xml::optional_attr(start, b"maxX")?
        .map(|s| parse_adj_coord(&s))
        .transpose()?;
    let min_y = xml::optional_attr(start, b"minY")?
        .map(|s| parse_adj_coord(&s))
        .transpose()?;
    let max_y = xml::optional_attr(start, b"maxY")?
        .map(|s| parse_adj_coord(&s))
        .transpose()?;
    let position = parse_position_child(reader, buf, b"ahXY")?;
    Ok(AdjustHandle::XY {
        guide_ref_x,
        guide_ref_y,
        min_x,
        max_x,
        min_y,
        max_y,
        position,
    })
}

fn parse_ah_polar(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    start: &BytesStart<'_>,
) -> Result<AdjustHandle> {
    let guide_ref_r = xml::optional_attr(start, b"gdRefR")?;
    let guide_ref_ang = xml::optional_attr(start, b"gdRefAng")?;
    let min_r = xml::optional_attr(start, b"minR")?
        .map(|s| parse_adj_coord(&s))
        .transpose()?;
    let max_r = xml::optional_attr(start, b"maxR")?
        .map(|s| parse_adj_coord(&s))
        .transpose()?;
    let min_ang = xml::optional_attr(start, b"minAng")?
        .map(|s| parse_adj_coord(&s))
        .transpose()?;
    let max_ang = xml::optional_attr(start, b"maxAng")?
        .map(|s| parse_adj_coord(&s))
        .transpose()?;
    let position = parse_position_child(reader, buf, b"ahPolar")?;
    Ok(AdjustHandle::Polar {
        guide_ref_r,
        guide_ref_ang,
        min_r,
        max_r,
        min_ang,
        max_ang,
        position,
    })
}

fn parse_position_child(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<AdjPoint> {
    let mut position: Option<AdjPoint> = None;
    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) if xml::local_name(e.name().as_ref()) == b"pos" => {
                position = Some(parse_adj_point(e)?);
            }
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"pos" => {
                position = Some(parse_adj_point(e)?);
                xml::skip_to_end(reader, buf, b"pos")?;
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => return Err(xml::unexpected_eof(end_tag)),
            _ => {}
        }
    }
    position.ok_or_else(|| ParseError::MissingElement {
        parent: String::from_utf8_lossy(end_tag).into_owned(),
        child: "pos".into(),
    })
}

// ── Connection sites (§20.1.9.7) ────────────────────────────────────────────

fn parse_connection_site_list(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Vec<ConnectionSite>> {
    let mut sites = Vec::new();
    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"cxn" => {
                sites.push(parse_connection_site(reader, buf, e)?);
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"cxnLst" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"cxnLst")),
            _ => {}
        }
    }
    Ok(sites)
}

fn parse_connection_site(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    start: &BytesStart<'_>,
) -> Result<ConnectionSite> {
    let angle = parse_adj_coord(&xml::required_attr(start, b"ang")?)?;
    let position = parse_position_child(reader, buf, b"cxn")?;
    Ok(ConnectionSite { angle, position })
}

// ── TextRect (§20.1.9.22) ───────────────────────────────────────────────────

fn parse_text_rect(e: &BytesStart<'_>) -> Result<TextRect> {
    Ok(TextRect {
        left: parse_adj_coord(&xml::required_attr(e, b"l")?)?,
        top: parse_adj_coord(&xml::required_attr(e, b"t")?)?,
        right: parse_adj_coord(&xml::required_attr(e, b"r")?)?,
        bottom: parse_adj_coord(&xml::required_attr(e, b"b")?)?,
    })
}

// ── Path list (§20.1.9.15) ──────────────────────────────────────────────────

fn parse_path_list(reader: &mut Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Vec<PathDef>> {
    let mut paths = Vec::new();
    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"path" => {
                paths.push(parse_path(reader, buf, e)?);
            }
            Event::Empty(ref e) if xml::local_name(e.name().as_ref()) == b"path" => {
                paths.push(parse_empty_path(e)?);
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"pathLst" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"pathLst")),
            _ => {}
        }
    }
    Ok(paths)
}

fn parse_empty_path(e: &BytesStart<'_>) -> Result<PathDef> {
    let a = parse_path_attrs(e)?;
    Ok(PathDef {
        w: a.w,
        h: a.h,
        fill: a.fill,
        stroke: a.stroke,
        extrusion_ok: a.extrusion_ok,
        commands: Vec::new(),
    })
}

fn parse_path(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    start: &BytesStart<'_>,
) -> Result<PathDef> {
    let a = parse_path_attrs(start)?;
    let mut commands = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Start(ref e) => {
                let local_owned = xml::local_name_owned(e.name().as_ref());
                match &*local_owned {
                    b"moveTo" => commands.push(PathCommand::MoveTo(parse_single_point_child(
                        reader, buf, b"moveTo",
                    )?)),
                    b"lnTo" => commands.push(PathCommand::LineTo(parse_single_point_child(
                        reader, buf, b"lnTo",
                    )?)),
                    b"cubicBezTo" => {
                        let points = parse_point_children(reader, buf, b"cubicBezTo", 3)?;
                        commands.push(PathCommand::CubicBezTo(
                            points[0].clone(),
                            points[1].clone(),
                            points[2].clone(),
                        ));
                    }
                    b"quadBezTo" => {
                        let points = parse_point_children(reader, buf, b"quadBezTo", 2)?;
                        commands.push(PathCommand::QuadBezTo(points[0].clone(), points[1].clone()));
                    }
                    b"arcTo" => {
                        commands.push(parse_arc_to(e)?);
                        xml::skip_to_end(reader, buf, b"arcTo")?;
                    }
                    b"close" => {
                        commands.push(PathCommand::Close);
                        xml::skip_to_end(reader, buf, b"close")?;
                    }
                    _ => {
                        xml::warn_unsupported_element("path", &local_owned);
                        xml::skip_to_end(reader, buf, &local_owned)?;
                    }
                }
            }
            Event::Empty(ref e) => {
                let qn = e.name();
                let local = xml::local_name(qn.as_ref());
                match local {
                    b"arcTo" => commands.push(parse_arc_to(e)?),
                    b"close" => commands.push(PathCommand::Close),
                    _ => xml::warn_unsupported_element("path", local),
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"path" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"path")),
            _ => {}
        }
    }

    Ok(PathDef {
        w: a.w,
        h: a.h,
        fill: a.fill,
        stroke: a.stroke,
        extrusion_ok: a.extrusion_ok,
        commands,
    })
}

struct PathAttrs {
    w: Dimension<crate::docx::dimension::Emu>,
    h: Dimension<crate::docx::dimension::Emu>,
    fill: PathFillMode,
    stroke: bool,
    extrusion_ok: bool,
}

fn parse_path_attrs(e: &BytesStart<'_>) -> Result<PathAttrs> {
    Ok(PathAttrs {
        w: Dimension::new(xml::optional_attr_i64(e, b"w")?.unwrap_or(0)),
        h: Dimension::new(xml::optional_attr_i64(e, b"h")?.unwrap_or(0)),
        fill: xml::optional_attr(e, b"fill")?
            .map(|s| parse_path_fill_mode(&s))
            .transpose()?
            .unwrap_or(PathFillMode::Norm),
        // §20.1.9.15: defaults per spec — stroke=true, extrusionOk=true.
        stroke: xml::optional_attr_bool(e, b"stroke")?.unwrap_or(true),
        extrusion_ok: xml::optional_attr_bool(e, b"extrusionOk")?.unwrap_or(true),
    })
}

fn parse_path_fill_mode(val: &str) -> Result<PathFillMode> {
    match val {
        "none" => Ok(PathFillMode::None),
        "norm" => Ok(PathFillMode::Norm),
        "lighten" => Ok(PathFillMode::Lighten),
        "lightenLess" => Ok(PathFillMode::LightenLess),
        "darken" => Ok(PathFillMode::Darken),
        "darkenLess" => Ok(PathFillMode::DarkenLess),
        other => Err(ParseError::InvalidAttributeValue {
            attr: "fill".into(),
            value: other.into(),
            reason: "expected value per §20.1.10.45 ST_PathFillMode".into(),
        }),
    }
}

// ── Point children (pt) ─────────────────────────────────────────────────────

fn parse_single_point_child(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
) -> Result<AdjPoint> {
    let points = parse_point_children(reader, buf, end_tag, 1)?;
    Ok(points.into_iter().next().unwrap())
}

fn parse_point_children(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    end_tag: &[u8],
    expected: usize,
) -> Result<Vec<AdjPoint>> {
    let mut points = Vec::with_capacity(expected);
    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e)
                if xml::local_name(e.name().as_ref()) == b"pt" =>
            {
                points.push(parse_adj_point(e)?);
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => return Err(xml::unexpected_eof(end_tag)),
            _ => {}
        }
    }
    if points.len() != expected {
        return Err(ParseError::InvalidAttributeValue {
            attr: "pt-count".into(),
            value: points.len().to_string(),
            reason: format!(
                "expected {expected} pt children under {}",
                String::from_utf8_lossy(end_tag)
            ),
        });
    }
    Ok(points)
}

fn parse_adj_point(e: &BytesStart<'_>) -> Result<AdjPoint> {
    Ok(AdjPoint {
        x: parse_adj_coord(&xml::required_attr(e, b"x")?)?,
        y: parse_adj_coord(&xml::required_attr(e, b"y")?)?,
    })
}

fn parse_arc_to(e: &BytesStart<'_>) -> Result<PathCommand> {
    Ok(PathCommand::ArcTo {
        wr: parse_adj_coord(&xml::required_attr(e, b"wR")?)?,
        hr: parse_adj_coord(&xml::required_attr(e, b"hR")?)?,
        start_angle: parse_adj_coord(&xml::required_attr(e, b"stAng")?)?,
        swing_angle: parse_adj_coord(&xml::required_attr(e, b"swAng")?)?,
    })
}

// ── AdjCoord lexing ────────────────────────────────────────────────────────

/// Parse an `AdjCoord` attribute value: a decimal integer literal or a
/// guide-name reference.
fn parse_adj_coord(val: &str) -> Result<AdjCoord> {
    if let Ok(n) = val.parse::<i64>() {
        return Ok(AdjCoord::Lit(n));
    }
    // Accept any non-empty non-numeric value as a guide name.
    if val.is_empty() {
        return Err(ParseError::InvalidAttributeValue {
            attr: "AdjCoord".into(),
            value: val.into(),
            reason: "expected decimal integer or guide name".into(),
        });
    }
    Ok(AdjCoord::Guide(val.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml_src: &str) -> CustomGeometry {
        let mut reader = Reader::from_reader(xml_src.as_bytes());
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match xml::next_event(&mut reader, &mut buf).unwrap() {
                Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"custGeom" => {
                    return parse_custom_geometry(&mut reader, &mut buf).unwrap();
                }
                Event::Eof => panic!("no custGeom"),
                _ => {}
            }
        }
    }

    #[test]
    fn empty_geom() {
        let g = parse(r#"<a:custGeom xmlns:a="urn:a"></a:custGeom>"#);
        assert!(g.av_list.is_empty());
        assert!(g.gd_list.is_empty());
        assert!(g.paths.is_empty());
        assert!(g.rect.is_none());
    }

    #[test]
    fn guide_lists() {
        let g = parse(
            r#"<a:custGeom xmlns:a="urn:a">
                <a:avLst>
                    <a:gd name="adj1" fmla="val 50000"/>
                </a:avLst>
                <a:gdLst>
                    <a:gd name="half_w" fmla="*/ w 1 2"/>
                    <a:gd name="half_h" fmla="*/ h 1 2"/>
                </a:gdLst>
            </a:custGeom>"#,
        );
        assert_eq!(g.av_list.len(), 1);
        assert_eq!(g.av_list[0].name, "adj1");
        assert_eq!(g.av_list[0].formula, "val 50000");
        assert_eq!(g.gd_list.len(), 2);
    }

    #[test]
    fn rect_with_literal_coords() {
        let g = parse(
            r#"<a:custGeom xmlns:a="urn:a">
                <a:rect l="0" t="0" r="100" b="50"/>
            </a:custGeom>"#,
        );
        let r = g.rect.unwrap();
        assert_eq!(r.left, AdjCoord::Lit(0));
        assert_eq!(r.right, AdjCoord::Lit(100));
        assert_eq!(r.bottom, AdjCoord::Lit(50));
    }

    #[test]
    fn rect_with_guide_refs() {
        let g = parse(
            r#"<a:custGeom xmlns:a="urn:a">
                <a:rect l="l" t="t" r="r" b="b"/>
            </a:custGeom>"#,
        );
        let r = g.rect.unwrap();
        assert_eq!(r.left, AdjCoord::Guide("l".into()));
        assert_eq!(r.bottom, AdjCoord::Guide("b".into()));
    }

    #[test]
    fn path_with_rect_shape() {
        let g = parse(
            r#"<a:custGeom xmlns:a="urn:a">
                <a:pathLst>
                    <a:path w="100" h="50">
                        <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
                        <a:lnTo><a:pt x="100" y="0"/></a:lnTo>
                        <a:lnTo><a:pt x="100" y="50"/></a:lnTo>
                        <a:lnTo><a:pt x="0" y="50"/></a:lnTo>
                        <a:close/>
                    </a:path>
                </a:pathLst>
            </a:custGeom>"#,
        );
        assert_eq!(g.paths.len(), 1);
        let p = &g.paths[0];
        assert_eq!(p.w.raw(), 100);
        assert_eq!(p.h.raw(), 50);
        assert_eq!(p.commands.len(), 5);
        assert!(matches!(p.commands[0], PathCommand::MoveTo(_)));
        assert!(matches!(p.commands[1], PathCommand::LineTo(_)));
        assert!(matches!(p.commands[4], PathCommand::Close));
    }

    #[test]
    fn path_with_cubic_and_quad_bez() {
        let g = parse(
            r#"<a:custGeom xmlns:a="urn:a">
                <a:pathLst>
                    <a:path w="100" h="100">
                        <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
                        <a:cubicBezTo>
                            <a:pt x="10" y="10"/>
                            <a:pt x="20" y="20"/>
                            <a:pt x="30" y="30"/>
                        </a:cubicBezTo>
                        <a:quadBezTo>
                            <a:pt x="40" y="40"/>
                            <a:pt x="50" y="50"/>
                        </a:quadBezTo>
                    </a:path>
                </a:pathLst>
            </a:custGeom>"#,
        );
        let p = &g.paths[0];
        assert_eq!(p.commands.len(), 3);
        assert!(matches!(p.commands[1], PathCommand::CubicBezTo(_, _, _)));
        assert!(matches!(p.commands[2], PathCommand::QuadBezTo(_, _)));
    }

    #[test]
    fn path_arc_to() {
        let g = parse(
            r#"<a:custGeom xmlns:a="urn:a">
                <a:pathLst>
                    <a:path w="100" h="100">
                        <a:arcTo wR="25" hR="25" stAng="0" swAng="5400000"/>
                    </a:path>
                </a:pathLst>
            </a:custGeom>"#,
        );
        let p = &g.paths[0];
        let PathCommand::ArcTo {
            wr,
            hr,
            start_angle,
            swing_angle,
        } = &p.commands[0]
        else {
            panic!()
        };
        assert_eq!(*wr, AdjCoord::Lit(25));
        assert_eq!(*hr, AdjCoord::Lit(25));
        assert_eq!(*start_angle, AdjCoord::Lit(0));
        assert_eq!(*swing_angle, AdjCoord::Lit(5_400_000));
    }

    #[test]
    fn path_attrs_defaults() {
        let g = parse(
            r#"<a:custGeom xmlns:a="urn:a">
                <a:pathLst>
                    <a:path w="10" h="20"/>
                </a:pathLst>
            </a:custGeom>"#,
        );
        let p = &g.paths[0];
        assert_eq!(p.fill, PathFillMode::Norm);
        assert!(p.stroke);
        assert!(p.extrusion_ok);
    }

    #[test]
    fn path_fill_mode_lighten() {
        let g = parse(
            r#"<a:custGeom xmlns:a="urn:a">
                <a:pathLst>
                    <a:path w="10" h="20" fill="lighten" stroke="0" extrusionOk="0"/>
                </a:pathLst>
            </a:custGeom>"#,
        );
        let p = &g.paths[0];
        assert_eq!(p.fill, PathFillMode::Lighten);
        assert!(!p.stroke);
        assert!(!p.extrusion_ok);
    }

    #[test]
    fn invalid_path_fill_mode_errors() {
        let mut reader = Reader::from_reader(
            r#"<a:custGeom xmlns:a="urn:a">
                <a:pathLst><a:path fill="bogus"/></a:pathLst>
            </a:custGeom>"#
                .as_bytes(),
        );
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match xml::next_event(&mut reader, &mut buf).unwrap() {
                Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"custGeom" => {
                    assert!(parse_custom_geometry(&mut reader, &mut buf).is_err());
                    return;
                }
                Event::Eof => panic!(),
                _ => {}
            }
        }
    }

    #[test]
    fn cubic_bez_with_wrong_point_count_errors() {
        let mut reader = Reader::from_reader(
            r#"<a:custGeom xmlns:a="urn:a">
                <a:pathLst>
                    <a:path w="10" h="10">
                        <a:cubicBezTo>
                            <a:pt x="0" y="0"/>
                            <a:pt x="1" y="1"/>
                        </a:cubicBezTo>
                    </a:path>
                </a:pathLst>
            </a:custGeom>"#
                .as_bytes(),
        );
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        loop {
            match xml::next_event(&mut reader, &mut buf).unwrap() {
                Event::Start(ref e) if xml::local_name(e.name().as_ref()) == b"custGeom" => {
                    assert!(parse_custom_geometry(&mut reader, &mut buf).is_err());
                    return;
                }
                Event::Eof => panic!(),
                _ => {}
            }
        }
    }

    #[test]
    fn adj_point_with_guide_reference() {
        let g = parse(
            r#"<a:custGeom xmlns:a="urn:a">
                <a:pathLst>
                    <a:path w="10" h="10">
                        <a:moveTo><a:pt x="hc" y="vc"/></a:moveTo>
                    </a:path>
                </a:pathLst>
            </a:custGeom>"#,
        );
        let PathCommand::MoveTo(pt) = &g.paths[0].commands[0] else {
            panic!()
        };
        assert_eq!(pt.x, AdjCoord::Guide("hc".into()));
        assert_eq!(pt.y, AdjCoord::Guide("vc".into()));
    }

    #[test]
    fn connection_sites_parsed() {
        let g = parse(
            r#"<a:custGeom xmlns:a="urn:a">
                <a:cxnLst>
                    <a:cxn ang="0">
                        <a:pos x="0" y="0"/>
                    </a:cxn>
                </a:cxnLst>
            </a:custGeom>"#,
        );
        assert_eq!(g.cxn_list.len(), 1);
        assert_eq!(g.cxn_list[0].angle, AdjCoord::Lit(0));
    }
}
