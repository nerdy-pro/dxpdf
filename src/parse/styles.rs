use quick_xml::events::Event;
use quick_xml::Reader;

use super::archive::{attr_val, local_name};

pub(super) fn parse_styles(xml: &str) -> crate::model::StyleMap {
    use std::rc::Rc;

    let mut styles = crate::model::StyleMap::new();

    // Built-in character styles that Word defines implicitly
    styles.insert(
        "Hyperlink".to_string(),
        crate::model::ResolvedParagraphStyle {
            run_props: crate::model::ResolvedRunStyle {
                color: crate::model::Color::from_hex("0563C1"),
                underline: Some(true),
                ..Default::default()
            },
            ..Default::default()
        },
    );
    let mut reader = Reader::from_str(xml);
    let mut in_style = false;
    let mut style_id = String::new();
    let mut style_type = String::new();
    let mut in_ppr = false;
    let mut in_rpr = false;
    let mut based_on: Option<String> = None;

    // Current style properties being collected
    let mut alignment = None;
    let mut spacing = None;
    let mut indentation = None;
    let mut bold = None;
    let mut italic = None;
    let mut underline = None;
    let mut font_size: Option<crate::dimension::HalfPoints> = None;
    let mut font_family: Option<Rc<str>> = None;
    let mut color = None;

    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"style" => {
                        in_style = true;
                        style_id.clear();
                        style_type.clear();
                        based_on = None;
                        alignment = None;
                        spacing = None;
                        indentation = None;
                        bold = None;
                        italic = None;
                        underline = None;
                        font_size = None;
                        font_family = None;
                        color = None;
                        in_ppr = false;
                        in_rpr = false;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key {
                                b"styleId" => style_id = val.into_owned(),
                                b"type" => style_type = val.into_owned(),
                                _ => {}
                            }
                        }
                    }
                    b"pPr" if in_style => in_ppr = true,
                    b"rPr" if in_style => in_rpr = true,
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"style" => {
                        if in_style && (style_type == "paragraph" || style_type == "character") {
                            let mut resolved = crate::model::ResolvedParagraphStyle {
                                alignment,
                                spacing,
                                indentation,
                                run_props: crate::model::ResolvedRunStyle {
                                    bold,
                                    italic,
                                    underline,
                                    font_size,
                                    font_family: font_family.clone(),
                                    color,
                                },
                            };
                            // Inherit from basedOn style
                            if let Some(ref base_id) = based_on {
                                if let Some(base) = styles.get(base_id) {
                                    if resolved.alignment.is_none() {
                                        resolved.alignment = base.alignment;
                                    }
                                    if resolved.spacing.is_none() {
                                        resolved.spacing = base.spacing;
                                    }
                                    if resolved.indentation.is_none() {
                                        resolved.indentation = base.indentation;
                                    }
                                    if resolved.run_props.bold.is_none() {
                                        resolved.run_props.bold = base.run_props.bold;
                                    }
                                    if resolved.run_props.italic.is_none() {
                                        resolved.run_props.italic = base.run_props.italic;
                                    }
                                    if resolved.run_props.font_size.is_none() {
                                        resolved.run_props.font_size = base.run_props.font_size;
                                    }
                                    if resolved.run_props.font_family.is_none() {
                                        resolved.run_props.font_family =
                                            base.run_props.font_family.clone();
                                    }
                                    if resolved.run_props.color.is_none() {
                                        resolved.run_props.color = base.run_props.color;
                                    }
                                }
                            }
                            styles.insert(style_id.clone(), resolved);
                        }
                        in_style = false;
                    }
                    b"pPr" => in_ppr = false,
                    b"rPr" => in_rpr = false,
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) if in_style => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if in_ppr {
                    match local {
                        b"jc" => {
                            if let Some(val) = attr_val(e, b"val") {
                                alignment = match val.as_str() {
                                    "left" | "start" => Some(crate::model::Alignment::Left),
                                    "center" => Some(crate::model::Alignment::Center),
                                    "right" | "end" => Some(crate::model::Alignment::Right),
                                    "both" | "justify" => Some(crate::model::Alignment::Justify),
                                    _ => None,
                                };
                            }
                        }
                        b"spacing" => {
                            let mut sp = crate::model::Spacing::default();
                            if let Some(v) = attr_val(e, b"before") {
                                sp.before = v.parse::<i64>().ok().map(crate::dimension::Twips::new);
                            }
                            if let Some(v) = attr_val(e, b"after") {
                                sp.after = v.parse::<i64>().ok().map(crate::dimension::Twips::new);
                            }
                            if let Some(v) = attr_val(e, b"line") {
                                sp.line = v.parse::<i64>().ok().map(crate::dimension::Twips::new);
                            }
                            if let Some(v) = attr_val(e, b"lineRule") {
                                sp.line_rule = match v.as_str() {
                                    "auto" => crate::model::LineRule::Auto,
                                    "exact" => crate::model::LineRule::Exact,
                                    "atLeast" => crate::model::LineRule::AtLeast,
                                    _ => crate::model::LineRule::Auto,
                                };
                            }
                            spacing = Some(sp);
                        }
                        b"ind" => {
                            let mut ind = crate::model::Indentation::default();
                            if let Some(v) = attr_val(e, b"left") {
                                ind.left = v.parse::<i64>().ok().map(crate::dimension::Twips::new);
                            }
                            if let Some(v) = attr_val(e, b"right") {
                                ind.right = v.parse::<i64>().ok().map(crate::dimension::Twips::new);
                            }
                            if let Some(v) = attr_val(e, b"firstLine") {
                                ind.first_line =
                                    v.parse::<i64>().ok().map(crate::dimension::Twips::new);
                            }
                            if let Some(v) = attr_val(e, b"hanging") {
                                if let Ok(h) = v.parse::<i64>() {
                                    ind.first_line = Some(crate::dimension::Twips::new(-h));
                                }
                            }
                            indentation = Some(ind);
                        }
                        _ => {}
                    }
                }
                if in_rpr || (!in_ppr && in_style) {
                    match local {
                        b"b" => bold = Some(true),
                        b"i" => italic = Some(true),
                        b"u" => underline = Some(true),
                        b"sz" => {
                            if let Some(v) = attr_val(e, b"val") {
                                font_size =
                                    v.parse::<i64>().ok().map(crate::dimension::HalfPoints::new);
                            }
                        }
                        b"rFonts" => {
                            if let Some(v) = attr_val(e, b"ascii") {
                                font_family = Some(Rc::from(v.as_str()));
                            } else if let Some(v) = attr_val(e, b"hAnsi") {
                                font_family = Some(Rc::from(v.as_str()));
                            }
                        }
                        b"color" => {
                            if let Some(v) = attr_val(e, b"val") {
                                color = crate::model::Color::from_hex(&v);
                            }
                        }
                        b"basedOn" => {
                            based_on = attr_val(e, b"val");
                        }
                        _ => {}
                    }
                }
                // basedOn can be at style level (not inside pPr/rPr)
                if !in_ppr && !in_rpr && local == b"basedOn" {
                    based_on = attr_val(e, b"val");
                }
            }
            Err(_) => break,
            _ => {}
        }
    }

    styles
}
