//! Serde schema for OMML (Office Math, ECMA-376 part 1 §22.1) — the minimal
//! subset this engine renders: `m:r` runs, `m:sSup` superscripts, `m:f`
//! fractions. quick-xml matches on local names with the prefix stripped, so
//! `m:oMath` is `"oMath"` here, exactly as `mc:AlternateContent` is
//! `"AlternateContent"` in `body_schema`.

use serde::Deserialize;

use crate::docx::parse::body_schema::TextXml;
use crate::docx::whitespace_workaround::restore_whitespace_sentinels;
use crate::model::{MathBlock, MathElement, MathRun};

/// `<m:oMath>` — inline math content inside a paragraph.
#[derive(Deserialize, Default)]
pub(crate) struct OMathXml {
    #[serde(rename = "$value", default)]
    pub content: Vec<MathChildXml>,
}

/// `<m:oMathPara>` — the display-equation wrapper: one or more `m:oMath`
/// children plus an `m:oMathParaPr` this engine drops (a plain struct with no
/// `$value` field discards unlisted children).
#[derive(Deserialize, Default)]
pub(crate) struct OMathParaXml {
    #[serde(rename = "oMath", default)]
    pub math: Vec<OMathXml>,
}

#[derive(Deserialize)]
pub(crate) enum MathChildXml {
    #[serde(rename = "r")]
    Run(MathRunXml),
    #[serde(rename = "sSup")]
    SSup(SSupXml),
    #[serde(rename = "f")]
    Frac(FracXml),
    /// `sSub`, `nary`, `rad`, `d`, … — dropped with a warning at conversion.
    #[serde(other)]
    Other,
}

/// `<m:r>` — a math run. Both `m:rPr` and `w:rPr` strip to `"rPr"` and are
/// deliberately unmodelled: a plain struct discards them, and the math face
/// supplies the formatting.
#[derive(Deserialize, Default)]
pub(crate) struct MathRunXml {
    #[serde(rename = "t", default)]
    pub texts: Vec<TextXml>,
}

/// `<m:sSup>` — `m:sSupPr` is discarded.
#[derive(Deserialize, Default)]
pub(crate) struct SSupXml {
    #[serde(rename = "e", default)]
    pub base: Vec<MathArgXml>,
    #[serde(rename = "sup", default)]
    pub sup: Vec<MathArgXml>,
}

/// `<m:f>` — `m:fPr` is discarded.
#[derive(Deserialize, Default)]
pub(crate) struct FracXml {
    #[serde(rename = "num", default)]
    pub num: Vec<MathArgXml>,
    #[serde(rename = "den", default)]
    pub den: Vec<MathArgXml>,
}

/// CT_OMathArg — the same content model as `m:oMath` itself.
#[derive(Deserialize, Default)]
pub(crate) struct MathArgXml {
    #[serde(rename = "$value", default)]
    pub content: Vec<MathChildXml>,
}

impl From<OMathXml> for MathBlock {
    fn from(x: OMathXml) -> Self {
        MathBlock {
            content: convert_children(x.content),
        }
    }
}

fn convert_children(children: Vec<MathChildXml>) -> Vec<MathElement> {
    let mut out = Vec::new();
    for child in children {
        match child {
            MathChildXml::Run(r) => {
                let text: String = r
                    .texts
                    .iter()
                    .map(|t| restore_whitespace_sentinels(&t.content))
                    .collect();
                if !text.is_empty() {
                    out.push(MathElement::Run(MathRun { text }));
                }
            }
            MathChildXml::SSup(s) => out.push(MathElement::Superscript {
                base: convert_args(s.base),
                sup: convert_args(s.sup),
            }),
            MathChildXml::Frac(f) => out.push(MathElement::Fraction {
                num: convert_args(f.num),
                den: convert_args(f.den),
            }),
            MathChildXml::Other => {
                log::warn!("unsupported OMML construct skipped (only m:r, m:sSup, m:f render)");
            }
        }
    }
    out
}

fn convert_args(args: Vec<MathArgXml>) -> Vec<MathElement> {
    args.into_iter()
        .flat_map(|a| convert_children(a.content))
        .collect()
}
