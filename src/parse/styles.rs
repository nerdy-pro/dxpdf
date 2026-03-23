use std::rc::Rc;

use quick_xml::events::Event;
use quick_xml::Reader;

use super::xml::helpers::{get_attr_lossy as attr_val, local_name};

/// Per-style accumulator for properties parsed from `w:style` elements.
struct StyleBuilder {
    style_id: String,
    style_type: String,
    based_on: Option<String>,
    in_ppr: bool,
    in_rpr: bool,
    alignment: Option<crate::model::Alignment>,
    spacing: Option<crate::model::Spacing>,
    indentation: Option<crate::model::Indentation>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    font_size: Option<crate::dimension::HalfPoints>,
    font_family: Option<Rc<str>>,
    color: Option<crate::model::Color>,
}

impl StyleBuilder {
    fn new(e: &quick_xml::events::BytesStart<'_>) -> Self {
        let mut style_id = String::new();
        let mut style_type = String::new();
        for attr in e.attributes().flatten() {
            let key = local_name(attr.key.as_ref());
            let val = String::from_utf8_lossy(&attr.value);
            match key {
                b"styleId" => style_id = val.into_owned(),
                b"type" => style_type = val.into_owned(),
                _ => {}
            }
        }
        Self {
            style_id,
            style_type,
            based_on: None,
            in_ppr: false,
            in_rpr: false,
            alignment: None,
            spacing: None,
            indentation: None,
            bold: None,
            italic: None,
            underline: None,
            font_size: None,
            font_family: None,
            color: None,
        }
    }

    fn finish(self, styles: &mut crate::model::StyleMap) {
        if self.style_type != "paragraph" && self.style_type != "character" {
            return;
        }
        let mut resolved = crate::model::ResolvedParagraphStyle {
            alignment: self.alignment,
            spacing: self.spacing,
            indentation: self.indentation,
            run_props: crate::model::ResolvedRunStyle {
                bold: self.bold,
                italic: self.italic,
                underline: self.underline,
                font_size: self.font_size,
                font_family: self.font_family,
                color: self.color,
            },
        };
        if let Some(ref base_id) = self.based_on {
            if let Some(base) = styles.get(base_id) {
                resolved.merge_from(base);
            }
        }
        styles.insert(self.style_id, resolved);
    }
}

pub(super) fn parse_styles(xml: &str) -> crate::model::StyleMap {
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
    let mut builder: Option<StyleBuilder> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"style" => {
                        builder = Some(StyleBuilder::new(e));
                    }
                    b"pPr" if builder.is_some() => {
                        if let Some(ref mut b) = builder {
                            b.in_ppr = true;
                        }
                    }
                    b"rPr" if builder.is_some() => {
                        if let Some(ref mut b) = builder {
                            b.in_rpr = true;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                match local {
                    b"style" => {
                        if let Some(b) = builder.take() {
                            b.finish(&mut styles);
                        }
                    }
                    b"pPr" => {
                        if let Some(ref mut b) = builder {
                            b.in_ppr = false;
                        }
                    }
                    b"rPr" => {
                        if let Some(ref mut b) = builder {
                            b.in_rpr = false;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) if builder.is_some() => {
                let b = builder.as_mut().unwrap();
                let name = e.name();
                let local = local_name(name.as_ref());
                if b.in_ppr {
                    match local {
                        b"jc" => {
                            if let Some(val) = attr_val(e, b"val") {
                                b.alignment = match val.as_str() {
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
                            b.spacing = Some(sp);
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
                            b.indentation = Some(ind);
                        }
                        _ => {}
                    }
                }
                if b.in_rpr || !b.in_ppr {
                    match local {
                        b"b" => b.bold = Some(true),
                        b"i" => b.italic = Some(true),
                        b"u" => b.underline = Some(true),
                        b"sz" => {
                            if let Some(v) = attr_val(e, b"val") {
                                b.font_size =
                                    v.parse::<i64>().ok().map(crate::dimension::HalfPoints::new);
                            }
                        }
                        b"rFonts" => {
                            if let Some(v) = attr_val(e, b"ascii") {
                                b.font_family = Some(Rc::from(v.as_str()));
                            } else if let Some(v) = attr_val(e, b"hAnsi") {
                                b.font_family = Some(Rc::from(v.as_str()));
                            }
                        }
                        b"color" => {
                            if let Some(v) = attr_val(e, b"val") {
                                b.color = crate::model::Color::from_hex(&v);
                            }
                        }
                        b"basedOn" => {
                            b.based_on = attr_val(e, b"val");
                        }
                        _ => {}
                    }
                }
                // basedOn can be at style level (not inside pPr/rPr)
                if !b.in_ppr && !b.in_rpr && local == b"basedOn" {
                    b.based_on = attr_val(e, b"val");
                }
            }
            Err(_) => break,
            _ => {}
        }
    }

    styles
}
