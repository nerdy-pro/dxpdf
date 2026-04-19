//! `<w:rPr>` schema (§17.3.2 run properties).
//!
//! Carries every direct run-formatting element plus the shared sub-schemas
//! from sibling modules. Deserializes to `(RunProperties, Option<StyleId>)`
//! via the `split` method — the style id is routed separately because the
//! property cascade applies it before direct formatting.

use serde::Deserialize;

use crate::docx::model::dimension::{Dimension, HalfPoints, Twips};
use crate::docx::model::{RunProperties, StrikeStyle, StyleId, UnderlineStyle};
use crate::docx::parse::primitives::st_enums::{StHighlightColor, StUnderline, StVerticalAlignRun};
use crate::docx::parse::primitives::{HexColor, OnOff};

use super::border::BorderXml;
use super::fonts::RFontsXml;
use super::lang::LangXml;
use super::shading::ShdXml;

/// Schema for the `<w:rPr>` element. All fields optional.
#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct RPrXml {
    #[serde(rename = "rStyle", default)]
    r_style: Option<ValString>,
    #[serde(rename = "rFonts", default)]
    r_fonts: Option<RFontsXml>,

    #[serde(rename = "sz", default)]
    sz: Option<ValAttr<Dimension<HalfPoints>>>,
    // Complex-script counterparts are intentionally ignored — renderer uses a single size.
    #[serde(rename = "b", default)]
    b: Option<OnOff>,
    #[serde(rename = "i", default)]
    i: Option<OnOff>,
    #[serde(rename = "u", default)]
    u: Option<UnderlineXml>,
    #[serde(rename = "strike", default)]
    strike: Option<OnOff>,
    #[serde(rename = "dstrike", default)]
    dstrike: Option<OnOff>,

    #[serde(rename = "color", default)]
    color: Option<ColorXml>,
    #[serde(rename = "highlight", default)]
    highlight: Option<ValAttr<StHighlightColor>>,
    #[serde(default)]
    shd: Option<ShdXml>,

    #[serde(rename = "vertAlign", default)]
    vert_align: Option<ValAttr<StVerticalAlignRun>>,

    #[serde(rename = "spacing", default)]
    spacing: Option<ValAttr<Dimension<Twips>>>,
    #[serde(rename = "kern", default)]
    kern: Option<ValAttr<Dimension<HalfPoints>>>,

    #[serde(rename = "caps", default)]
    caps: Option<OnOff>,
    #[serde(rename = "smallCaps", default)]
    small_caps: Option<OnOff>,
    #[serde(rename = "vanish", default)]
    vanish: Option<OnOff>,
    #[serde(rename = "noProof", default)]
    no_proof: Option<OnOff>,
    #[serde(rename = "webHidden", default)]
    web_hidden: Option<OnOff>,
    #[serde(rename = "rtl", default)]
    rtl: Option<OnOff>,
    #[serde(rename = "emboss", default)]
    emboss: Option<OnOff>,
    #[serde(rename = "imprint", default)]
    imprint: Option<OnOff>,
    #[serde(rename = "outline", default)]
    outline: Option<OnOff>,
    #[serde(rename = "shadow", default)]
    shadow: Option<OnOff>,

    #[serde(rename = "position", default)]
    position: Option<ValAttr<Dimension<HalfPoints>>>,

    #[serde(rename = "lang", default)]
    lang: Option<LangXml>,
    #[serde(rename = "bdr", default)]
    bdr: Option<BorderXml>,
}

/// `<w:u w:val="..."/>` — underline. Unlike other ST-enum wrappers we can't
/// use a bare `ValAttr<StUnderline>` because the attribute is optional; an
/// underline element with no `@val` means "Single" per §17.3.2.40.
#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct UnderlineXml {
    #[serde(rename = "@val", default)]
    val: Option<StUnderline>,
}

/// `<w:color w:val="RRGGBB" ... />` — run color. The spec also allows
/// theme-color fields (`@themeColor`, `@themeTint`, `@themeShade`) which we
/// don't yet resolve; recorded here as raw strings in case a future pass
/// wants them.
#[derive(Clone, Debug, Deserialize)]
pub(crate) struct ColorXml {
    #[serde(rename = "@val")]
    val: HexColor,
}

#[derive(Clone, Debug, Deserialize)]
pub(crate) struct ValString {
    #[serde(rename = "@val")]
    val: String,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(bound(deserialize = "T: serde::Deserialize<'de>"))]
pub(crate) struct ValAttr<T> {
    #[serde(rename = "@val")]
    val: T,
}

impl RPrXml {
    /// Split into `(properties, style_id)`. The style id applies first in
    /// the cascade (§17.7.2), so it stays separate from the direct-formatting
    /// `RunProperties`.
    pub(crate) fn split(self) -> (RunProperties, Option<StyleId>) {
        let style_id = self.r_style.map(|v| StyleId::new(v.val));
        let props = RunProperties {
            fonts: self.r_fonts.map(Into::into).unwrap_or_default(),
            font_size: self.sz.map(|s| s.val),
            bold: self.b.map(|OnOff(b)| b),
            italic: self.i.map(|OnOff(b)| b),
            underline: self.u.map(resolve_underline),
            strike: resolve_strike(self.strike, self.dstrike),
            color: self.color.map(|c| c.val.into()),
            highlight: self.highlight.map(|h| h.val.into()),
            shading: self.shd.map(Into::into),
            vertical_align: self.vert_align.map(|v| v.val.into()),
            spacing: self.spacing.map(|s| s.val),
            kerning: self.kern.map(|k| k.val),
            all_caps: self.caps.map(|OnOff(b)| b),
            small_caps: self.small_caps.map(|OnOff(b)| b),
            vanish: self.vanish.map(|OnOff(b)| b),
            no_proof: self.no_proof.map(|OnOff(b)| b),
            web_hidden: self.web_hidden.map(|OnOff(b)| b),
            rtl: self.rtl.map(|OnOff(b)| b),
            emboss: self.emboss.map(|OnOff(b)| b),
            imprint: self.imprint.map(|OnOff(b)| b),
            outline: self.outline.map(|OnOff(b)| b),
            shadow: self.shadow.map(|OnOff(b)| b),
            position: self.position.map(|p| p.val),
            lang: self.lang.map(Into::into),
            border: self.bdr.map(Into::into),
        };
        (props, style_id)
    }
}

/// `<w:u/>` with no `@val` means Single per §17.3.2.40; otherwise use the
/// named style.
fn resolve_underline(u: UnderlineXml) -> UnderlineStyle {
    match u.val {
        Some(v) => v.into(),
        None => UnderlineStyle::Single,
    }
}

/// `<w:strike/>` and `<w:dstrike/>` are separate OnOff toggles; dstrike
/// takes precedence when both are on.
fn resolve_strike(strike: Option<OnOff>, dstrike: Option<OnOff>) -> Option<StrikeStyle> {
    let d = dstrike.map(|OnOff(b)| b).unwrap_or(false);
    let s = strike.map(|OnOff(b)| b).unwrap_or(false);
    match (d, s) {
        (true, _) => Some(StrikeStyle::Double),
        (false, true) => Some(StrikeStyle::Single),
        (false, false) => {
            // explicit off → Some(None), absent → None
            if strike.is_some() || dstrike.is_some() {
                Some(StrikeStyle::None)
            } else {
                None
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::{BorderStyle, Color, HighlightColor, UnderlineStyle, VerticalAlign};

    fn parse(xml: &str) -> (RunProperties, Option<StyleId>) {
        let r: RPrXml = quick_xml::de::from_str(xml).expect("deserialize rPr");
        r.split()
    }

    #[test]
    fn empty_rpr_default_run_properties() {
        let (rp, sid) = parse(r#"<rPr/>"#);
        assert!(sid.is_none());
        assert!(rp.bold.is_none());
        assert!(rp.italic.is_none());
    }

    #[test]
    fn style_ref_extracted() {
        let (rp, sid) = parse(r#"<rPr><rStyle val="Emphasis"/></rPr>"#);
        assert_eq!(sid.map(|s| s.as_str().to_string()), Some("Emphasis".into()));
        assert!(rp.bold.is_none());
    }

    #[test]
    fn basic_toggles() {
        let (rp, _) = parse(r#"<rPr><b/><i/><caps/></rPr>"#);
        assert_eq!(rp.bold, Some(true));
        assert_eq!(rp.italic, Some(true));
        assert_eq!(rp.all_caps, Some(true));
    }

    #[test]
    fn toggle_off_is_false() {
        let (rp, _) = parse(r#"<rPr><b val="false"/></rPr>"#);
        assert_eq!(rp.bold, Some(false));
    }

    #[test]
    fn font_size_is_half_points() {
        let (rp, _) = parse(r#"<rPr><sz val="22"/></rPr>"#);
        assert_eq!(rp.font_size.map(|d| d.raw()), Some(22));
    }

    #[test]
    fn underline_with_val() {
        let (rp, _) = parse(r#"<rPr><u val="double"/></rPr>"#);
        assert_eq!(rp.underline, Some(UnderlineStyle::Double));
    }

    #[test]
    fn underline_without_val_defaults_single() {
        let (rp, _) = parse(r#"<rPr><u/></rPr>"#);
        assert_eq!(rp.underline, Some(UnderlineStyle::Single));
    }

    #[test]
    fn strike_single() {
        let (rp, _) = parse(r#"<rPr><strike/></rPr>"#);
        assert_eq!(rp.strike, Some(StrikeStyle::Single));
    }

    #[test]
    fn dstrike_wins_over_strike() {
        let (rp, _) = parse(r#"<rPr><strike/><dstrike/></rPr>"#);
        assert_eq!(rp.strike, Some(StrikeStyle::Double));
    }

    #[test]
    fn strike_explicit_off() {
        let (rp, _) = parse(r#"<rPr><strike val="0"/></rPr>"#);
        assert_eq!(rp.strike, Some(StrikeStyle::None));
    }

    #[test]
    fn color_rgb_and_auto() {
        let (rp, _) = parse(r#"<rPr><color val="FF0000"/></rPr>"#);
        assert_eq!(rp.color, Some(Color::Rgb(0xFF0000)));

        let (rp, _) = parse(r#"<rPr><color val="auto"/></rPr>"#);
        assert_eq!(rp.color, Some(Color::Auto));
    }

    #[test]
    fn highlight_via_st_enum() {
        let (rp, _) = parse(r#"<rPr><highlight val="yellow"/></rPr>"#);
        assert_eq!(rp.highlight, Some(HighlightColor::Yellow));
    }

    #[test]
    fn vertical_align_superscript() {
        let (rp, _) = parse(r#"<rPr><vertAlign val="superscript"/></rPr>"#);
        assert_eq!(rp.vertical_align, Some(VerticalAlign::Superscript));
    }

    #[test]
    fn spacing_and_kern_and_position() {
        let (rp, _) = parse(
            r#"<rPr>
                <spacing val="40"/>
                <kern val="20"/>
                <position val="-4"/>
            </rPr>"#,
        );
        assert_eq!(rp.spacing.map(|d| d.raw()), Some(40));
        assert_eq!(rp.kerning.map(|d| d.raw()), Some(20));
        assert_eq!(rp.position.map(|d| d.raw()), Some(-4));
    }

    #[test]
    fn lang_tri_mode() {
        let (rp, _) = parse(r#"<rPr><lang val="en-US" eastAsia="ja-JP"/></rPr>"#);
        let l = rp.lang.unwrap();
        assert_eq!(l.val.as_deref(), Some("en-US"));
        assert_eq!(l.east_asia.as_deref(), Some("ja-JP"));
    }

    #[test]
    fn border_via_bdr() {
        let (rp, _) = parse(r#"<rPr><bdr val="single" sz="4" color="000000"/></rPr>"#);
        let b = rp.border.unwrap();
        assert_eq!(b.style, BorderStyle::Single);
        assert_eq!(b.width.raw(), 4);
    }

    #[test]
    fn fonts_explicit_and_theme_mix() {
        let (rp, _) = parse(r#"<rPr><rFonts ascii="Calibri" hAnsiTheme="minorHAnsi"/></rPr>"#);
        assert_eq!(rp.fonts.ascii.explicit.as_deref(), Some("Calibri"));
        assert!(rp.fonts.high_ansi.theme.is_some());
    }

    #[test]
    fn full_rpr_end_to_end() {
        let xml = r#"<rPr>
            <rStyle val="Heading1Char"/>
            <rFonts ascii="Arial" hAnsi="Arial"/>
            <b/>
            <i/>
            <sz val="28"/>
            <color val="2E74B5"/>
            <u val="single"/>
            <lang val="en-US"/>
        </rPr>"#;
        let (rp, sid) = parse(xml);
        assert_eq!(
            sid.map(|s| s.as_str().to_string()),
            Some("Heading1Char".into())
        );
        assert_eq!(rp.fonts.ascii.explicit.as_deref(), Some("Arial"));
        assert_eq!(rp.bold, Some(true));
        assert_eq!(rp.italic, Some(true));
        assert_eq!(rp.font_size.map(|d| d.raw()), Some(28));
        assert_eq!(rp.color, Some(Color::Rgb(0x2E74B5)));
        assert_eq!(rp.underline, Some(UnderlineStyle::Single));
    }
}
