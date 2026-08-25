//! §21.1.2.1 CT_TextBody — DrawingML text (`a:txBody` and the chart part's
//! `c:rich`, which is the same type).
//!
//! Distinct from the WordprocessingML text a `wps:txbx` carries: paragraphs
//! are `a:p`, runs `a:r` with `a:rPr` attributes in DrawingML units, and the
//! whole body is literal — SmartArt bakes its text-fit result into each
//! run's `@sz`, so what parses here is what renders.

use serde::Deserialize;

use crate::docx::model::dimension::{CentiPoints, Dimension};
use crate::docx::model::{
    Alignment, DrawingRunProps, DrawingTextBody, DrawingTextParagraph, DrawingTextRun,
};
use crate::model::Dup;

use super::shape::BodyPrXml;

/// `a:txBody` / `c:rich`.
#[derive(Debug, Default, Deserialize)]
pub(crate) struct TextBodyXml {
    #[serde(rename = "bodyPr", default)]
    pub(crate) body_pr: Vec<BodyPrXml>,
    #[serde(rename = "p", default)]
    pub(crate) paragraphs: Vec<TextParagraphXml>,
}

/// §21.1.2.2.6 `a:p`. Children are collected in order so interleaved
/// `a:r`/`a:br`/`a:fld` keep their sequence.
#[derive(Debug, Default, Deserialize)]
pub(crate) struct TextParagraphXml {
    #[serde(rename = "pPr", default)]
    p_pr: Vec<TextParagraphPrXml>,
    #[serde(rename = "$value", default)]
    children: Vec<TextParagraphChildXml>,
}

#[derive(Debug, Deserialize)]
pub(crate) enum TextParagraphChildXml {
    #[serde(rename = "r")]
    Run(TextRunXml),
    #[serde(rename = "br")]
    Break(super::fill::Empty),
    /// §21.1.2.2.4 `a:fld` — generated text (chart axis units, slide
    /// numbers). Its cached `a:t` is literal, so it reads as a run.
    #[serde(rename = "fld")]
    Field(TextRunXml),
    #[serde(other)]
    Other,
}

/// §21.1.2.2.7 `a:pPr`.
#[derive(Debug, Default, Deserialize)]
pub(crate) struct TextParagraphPrXml {
    #[serde(rename = "@algn", default)]
    algn: Option<StTextAlignType>,
    #[serde(rename = "defRPr", default)]
    def_r_pr: Vec<TextRunPrXml>,
}

/// §20.1.10.59 ST_TextAlignType.
#[derive(Clone, Copy, Debug, Deserialize)]
enum StTextAlignType {
    #[serde(rename = "l")]
    Left,
    #[serde(rename = "ctr")]
    Center,
    #[serde(rename = "r")]
    Right,
    #[serde(rename = "just")]
    Justify,
    /// The DrawingML-only distributed variants; folded into justify.
    #[serde(rename = "justLow")]
    JustLow,
    #[serde(rename = "dist")]
    Dist,
    #[serde(rename = "thaiDist")]
    ThaiDist,
}

/// `a:r` / `a:fld`.
#[derive(Debug, Default, Deserialize)]
pub(crate) struct TextRunXml {
    #[serde(rename = "rPr", default)]
    r_pr: Vec<TextRunPrXml>,
    #[serde(rename = "t", default)]
    text: Vec<String>,
}

/// §21.1.2.3.9 `a:rPr` (and `a:defRPr`, the same CT_TextCharacterProperties).
#[derive(Debug, Default, Deserialize)]
pub(crate) struct TextRunPrXml {
    #[serde(rename = "@sz", default)]
    sz: Option<Dimension<CentiPoints>>,
    #[serde(rename = "@b", default)]
    b: Option<super::super::super::primitives::AttrBool>,
    #[serde(rename = "@i", default)]
    i: Option<super::super::super::primitives::AttrBool>,
    #[serde(rename = "solidFill", default)]
    solid_fill: Vec<super::fill::SolidFillXml>,
    #[serde(rename = "latin", default)]
    latin: Vec<LatinFontXml>,
}

#[derive(Debug, Deserialize)]
pub(crate) struct LatinFontXml {
    #[serde(rename = "@typeface", default)]
    typeface: Option<String>,
}

impl From<StTextAlignType> for Alignment {
    fn from(a: StTextAlignType) -> Self {
        // The shared enum is logical (start/end); DrawingML's `l`/`r` are
        // physical, but a diagram label or chart title has no w:bidi context
        // to flip under, so start/end is the faithful mapping.
        match a {
            StTextAlignType::Left => Self::Start,
            StTextAlignType::Center => Self::Center,
            StTextAlignType::Right => Self::End,
            StTextAlignType::Justify | StTextAlignType::JustLow => Self::Both,
            StTextAlignType::Dist | StTextAlignType::ThaiDist => Self::Distribute,
        }
    }
}

impl From<TextRunPrXml> for DrawingRunProps {
    fn from(x: TextRunPrXml) -> Self {
        Self {
            size: x.sz,
            bold: x.b.map(|v| v.0),
            italic: x.i.map(|v| v.0),
            color: Dup::from(x.solid_fill)
                .into_value()
                .and_then(|f| f.color)
                .map(Into::into),
            family: Dup::from(x.latin).into_value().and_then(|l| l.typeface),
        }
    }
}

impl From<TextBodyXml> for DrawingTextBody {
    fn from(x: TextBodyXml) -> Self {
        Self {
            body_pr: Dup::from(x.body_pr).into_value().map(Into::into),
            paragraphs: x.paragraphs.into_iter().map(Into::into).collect(),
        }
    }
}

impl From<TextParagraphXml> for DrawingTextParagraph {
    fn from(x: TextParagraphXml) -> Self {
        let p_pr = Dup::from(x.p_pr).into_value();
        let (alignment, default_run) = match p_pr {
            Some(p) => (
                p.algn.map(Into::into),
                Dup::from(p.def_r_pr).into_value().map(Into::into),
            ),
            None => (None, None),
        };
        let runs = x
            .children
            .into_iter()
            .filter_map(|c| match c {
                TextParagraphChildXml::Run(r) | TextParagraphChildXml::Field(r) => {
                    Some(DrawingTextRun::Text {
                        text: r.text.concat(),
                        props: Dup::from(r.r_pr)
                            .into_value()
                            .map(Into::into)
                            .unwrap_or_default(),
                    })
                }
                TextParagraphChildXml::Break(_) => Some(DrawingTextRun::Break),
                TextParagraphChildXml::Other => None,
            })
            .collect();
        Self {
            alignment,
            default_run,
            runs,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> DrawingTextBody {
        let x: TextBodyXml = quick_xml::de::from_str(xml).unwrap();
        x.into()
    }

    #[test]
    fn runs_keep_their_baked_size_and_style() {
        let b = parse(
            r#"<txBody><bodyPr/><p><pPr algn="ctr"/>
                 <r><rPr sz="2300" b="1"><solidFill><srgbClr val="FF0000"/></solidFill>
                    <latin typeface="Calibri"/></rPr><t>Step one</t></r>
               </p></txBody>"#,
        );
        assert_eq!(b.paragraphs.len(), 1);
        let p = &b.paragraphs[0];
        assert_eq!(p.alignment, Some(Alignment::Center));
        let DrawingTextRun::Text { text, props } = &p.runs[0] else {
            panic!()
        };
        assert_eq!(text, "Step one");
        assert_eq!(props.size.unwrap().raw(), 2300, "hundredths of a point");
        assert_eq!(props.bold, Some(true));
        assert_eq!(props.family.as_deref(), Some("Calibri"));
        assert!(props.color.is_some());
    }

    /// Chart text often puts everything on `a:defRPr` and nothing on runs.
    #[test]
    fn paragraph_default_run_properties_parse() {
        let b = parse(
            r#"<txBody><bodyPr/><p><pPr><defRPr sz="900"/></pPr>
                 <r><t>42</t></r></p></txBody>"#,
        );
        let p = &b.paragraphs[0];
        assert_eq!(p.default_run.as_ref().unwrap().size.unwrap().raw(), 900);
        let DrawingTextRun::Text { props, .. } = &p.runs[0] else {
            panic!()
        };
        assert_eq!(props.size, None, "the run itself stays bare");
    }

    /// `a:br` splits lines; `a:fld` reads as its cached text.
    #[test]
    fn breaks_and_fields_keep_their_order() {
        let b = parse(
            r#"<txBody><bodyPr/><p>
                 <r><t>one</t></r><br/><fld id="{X}" type="slidenum"><t>2</t></fld>
               </p></txBody>"#,
        );
        let kinds: Vec<&'static str> = b.paragraphs[0]
            .runs
            .iter()
            .map(|r| match r {
                DrawingTextRun::Text { .. } => "t",
                DrawingTextRun::Break => "br",
            })
            .collect();
        assert_eq!(kinds, vec!["t", "br", "t"]);
    }
}
