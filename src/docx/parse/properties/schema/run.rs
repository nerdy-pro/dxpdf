//! `<w:rPr>` schema (§17.3.2 run properties).
//!
//! Carries every direct run-formatting element plus the shared sub-schemas
//! from sibling modules. Deserializes to `(RunProperties, Option<StyleId>)`
//! via the `split` method — the style id is routed separately because the
//! property cascade applies it before direct formatting.

use crate::model::Dup;
use serde::Deserialize;

use crate::docx::model::dimension::{Dimension, HalfPoints, Twips, Unit};
use crate::docx::model::{RunProperties, StrikeStyle, StyleId, TextScale, UnderlineStyle};
use crate::docx::parse::primitives::st_enums::{StHighlightColor, StUnderline, StVerticalAlignRun};
use crate::docx::parse::primitives::units::deserialize_nonnegative_dimension;
use crate::docx::parse::primitives::{last_toggle, HexColor, OnOff};

use super::border::BorderXml;
use super::fonts::RFontsXml;
use super::lang::LangXml;
use super::shading::ShdXml;

/// Schema for the `<w:rPr>` element. All fields optional.
///
/// Every child is typed `Vec<T>`, not `Option<T>`, so a producer that repeats
/// one cannot fail the parse; `split` collapses each with `last`/`last_toggle`.
/// The policy and the reasoning behind "last wins" are in
/// `crate::docx::parse::primitives::duplicates`.
#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct RPrXml {
    /// §17.13.5.15 `<w:del>` inside a paragraph mark's `rPr` — the mark
    /// itself is deleted, which is how Word spells a whole-paragraph tracked
    /// delete. Presence-only; the attributes are the wrapper's usual id/
    /// author/date and nothing here reads them.
    #[serde(rename = "del", default)]
    del: Vec<MarkDelXml>,
    #[serde(rename = "rStyle", default)]
    r_style: Vec<ValString>,
    #[serde(rename = "rFonts", default)]
    r_fonts: Vec<RFontsXml>,

    #[serde(rename = "sz", default)]
    sz: Vec<NonNegativeDimensionVal<HalfPoints>>,
    // Complex-script counterparts are intentionally ignored — renderer uses a single size.
    #[serde(rename = "b", default)]
    b: Vec<OnOff>,
    #[serde(rename = "i", default)]
    i: Vec<OnOff>,
    #[serde(rename = "u", default)]
    u: Vec<UnderlineXml>,
    #[serde(rename = "strike", default)]
    strike: Vec<OnOff>,
    #[serde(rename = "dstrike", default)]
    dstrike: Vec<OnOff>,

    #[serde(rename = "color", default)]
    color: Vec<ColorXml>,
    #[serde(rename = "highlight", default)]
    highlight: Vec<ValAttr<StHighlightColor>>,
    #[serde(default)]
    shd: Vec<ShdXml>,

    #[serde(rename = "vertAlign", default)]
    vert_align: Vec<ValAttr<StVerticalAlignRun>>,

    #[serde(rename = "spacing", default)]
    spacing: Vec<ValAttr<Dimension<Twips>>>,
    #[serde(rename = "kern", default)]
    kern: Vec<NonNegativeDimensionVal<HalfPoints>>,
    /// §17.3.2.45 — `<w:w w:val="80"/>`: horizontal character scale in percent.
    #[serde(rename = "w", default)]
    char_scale: Vec<ValAttr<u16>>,

    #[serde(rename = "caps", default)]
    caps: Vec<OnOff>,
    #[serde(rename = "smallCaps", default)]
    small_caps: Vec<OnOff>,
    #[serde(rename = "vanish", default)]
    vanish: Vec<OnOff>,
    #[serde(rename = "noProof", default)]
    no_proof: Vec<OnOff>,
    #[serde(rename = "webHidden", default)]
    web_hidden: Vec<OnOff>,
    #[serde(rename = "rtl", default)]
    rtl: Vec<OnOff>,
    #[serde(rename = "emboss", default)]
    emboss: Vec<OnOff>,
    #[serde(rename = "imprint", default)]
    imprint: Vec<OnOff>,
    #[serde(rename = "outline", default)]
    outline: Vec<OnOff>,
    #[serde(rename = "shadow", default)]
    shadow: Vec<OnOff>,

    #[serde(rename = "position", default)]
    position: Vec<ValAttr<Dimension<HalfPoints>>>,

    #[serde(rename = "lang", default)]
    lang: Vec<LangXml>,
    #[serde(rename = "bdr", default)]
    bdr: Vec<BorderXml>,
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
/// don't yet resolve — they are currently ignored (only `@val` is read).
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

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(bound(deserialize = "U: Unit"))]
struct NonNegativeDimensionVal<U: Unit> {
    #[serde(
        rename = "@val",
        deserialize_with = "deserialize_nonnegative_dimension"
    )]
    val: Dimension<U>,
}

/// Payload of `rPr/del` — attributes (id/author/date) accepted and ignored.
#[derive(Clone, Debug, Deserialize, Default)]
pub(crate) struct MarkDelXml {}

impl RPrXml {
    /// §17.13.5.15: whether this property bag (as a paragraph mark's `rPr`)
    /// carries a `<w:del>` — the mark is deleted.
    pub(crate) fn mark_deleted(&self) -> bool {
        !self.del.is_empty()
    }

    /// Split into `(properties, style_id)`. The style id applies first in
    /// the cascade (§17.7.2), so it stays separate from the direct-formatting
    /// `RunProperties`.
    pub(crate) fn split(self) -> (RunProperties, Option<StyleId>) {
        let style_id = Dup::from(self.r_style)
            .into_value()
            .map(|v| StyleId::new(v.val));
        let props = RunProperties {
            fonts: Dup::from(self.r_fonts)
                .into_value()
                .map(Into::into)
                .unwrap_or_default(),
            font_size: Dup::from(self.sz).map(|s| s.val),
            bold: last_toggle(self.b),
            italic: last_toggle(self.i),
            underline: Dup::from(self.u).filter_map(resolve_underline),
            strike: resolve_strike(self.strike, self.dstrike),
            color: Dup::from(self.color).map(|c| c.val.into()),
            highlight: Dup::from(self.highlight).map(|h| h.val.into()),
            shading: Dup::from(self.shd).map(Into::into),
            vertical_align: Dup::from(self.vert_align).map(|v| v.val.into()),
            spacing: Dup::from(self.spacing).map(|s| s.val),
            kerning: Dup::from(self.kern).map(|k| k.val),
            all_caps: last_toggle(self.caps),
            small_caps: last_toggle(self.small_caps),
            vanish: last_toggle(self.vanish),
            no_proof: last_toggle(self.no_proof),
            web_hidden: last_toggle(self.web_hidden),
            rtl: last_toggle(self.rtl),
            emboss: last_toggle(self.emboss),
            imprint: last_toggle(self.imprint),
            outline: last_toggle(self.outline),
            shadow: last_toggle(self.shadow),
            position: Dup::from(self.position).map(|p| p.val),
            lang: Dup::from(self.lang).map(Into::into),
            border: Dup::from(self.bdr).map(Into::into),
            text_scale: Dup::from(self.char_scale).map(|v| TextScale::new(v.val)),
        };
        (props, style_id)
    }
}

/// Resolve `<w:u .../>` to an `UnderlineStyle` if — and only if — `@val` is
/// present. A `<w:u>` element without `@val` is silent in the cascade
/// (returns `None`) so it doesn't override an inherited style and doesn't
/// force an underline of its own.
///
/// §17.3.2.40 documents `@val` defaulting to `single` when omitted, but real
/// Word output emits `<w:u w:color="…"/>` (no `@val`) merely to remember a
/// chosen underline color even when the user has *not* turned underline on.
/// Treating that as "single" makes every such run render underlined — which
/// neither Word nor LibreOffice does. Matching Word's observable behaviour
/// is the right call here; the literal spec interpretation is wrong about
/// real-world documents.
fn resolve_underline(u: UnderlineXml) -> Option<UnderlineStyle> {
    u.val.map(Into::into)
}

/// `<w:strike/>` and `<w:dstrike/>` are separate OnOff toggles; dstrike
/// takes precedence when both are on. Each input is the full list of repeated
/// occurrences inside the parent `<w:rPr>` — by §17.7.2 last-wins cascade,
/// only the final element of each list is observable, so we collapse before
/// resolving precedence.
fn resolve_strike(strike: Vec<OnOff>, dstrike: Vec<OnOff>) -> Option<StrikeStyle> {
    let strike_present = !strike.is_empty();
    let dstrike_present = !dstrike.is_empty();
    let s = last_toggle(strike).unwrap_or(false);
    let d = last_toggle(dstrike).unwrap_or(false);
    match (d, s) {
        (true, _) => Some(StrikeStyle::Double),
        (false, true) => Some(StrikeStyle::Single),
        (false, false) => {
            // explicit off → Some(None), absent → None
            if strike_present || dstrike_present {
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
    use crate::docx::model::{
        BorderStyle, Color, HighlightColor, TextScale, UnderlineStyle, VerticalAlign,
    };

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
        assert_eq!(rp.font_size.map(|d| d.raw()), Dup::from(Some(22)));
    }

    #[test]
    fn underline_with_val() {
        let (rp, _) = parse(r#"<rPr><u val="double"/></rPr>"#);
        assert_eq!(rp.underline, Dup::from(Some(UnderlineStyle::Double)));
    }

    #[test]
    fn underline_without_val_is_silent_in_cascade() {
        // Real Word emits `<w:u w:color="…"/>` — no `@val` — to remember a
        // chosen underline color even when underline is *not* on. Treating
        // that as "single" caused every such run to render underlined, which
        // doesn't match Word's actual rendering. So `<w:u>` without `@val`
        // contributes nothing to the cascade (parser returns None), letting
        // any inherited underline win.
        let (rp, _) = parse(r#"<rPr><u/></rPr>"#);
        assert_eq!(rp.underline, Dup::from(None));
    }

    #[test]
    fn underline_with_color_but_no_val_is_silent() {
        // Same shape Word actually emits — color attribute alone, no `@val`.
        let (rp, _) = parse(r#"<rPr><u color="000000"/></rPr>"#);
        assert_eq!(rp.underline, Dup::from(None));
    }

    #[test]
    fn underline_val_none_is_explicit_override() {
        // §17.3.2.40: w:val="none" is the explicit "no underline" override —
        // it must round-trip as `Some(UnderlineStyle::None)`, distinct from
        // both an absent <w:u/> element (None) and an inherited underline.
        let (rp, _) = parse(r#"<rPr><u val="none"/></rPr>"#);
        assert_eq!(rp.underline, Dup::from(Some(UnderlineStyle::None)));
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
        assert_eq!(rp.color, Dup::from(Some(Color::Rgb(0xFF0000))));

        let (rp, _) = parse(r#"<rPr><color val="auto"/></rPr>"#);
        assert_eq!(rp.color, Dup::from(Some(Color::Auto)));
    }

    #[test]
    fn highlight_via_st_enum() {
        let (rp, _) = parse(r#"<rPr><highlight val="yellow"/></rPr>"#);
        assert_eq!(rp.highlight, Dup::from(Some(HighlightColor::Yellow)));
    }

    #[test]
    fn highlight_val_none_is_explicit_override() {
        // §17.3.2.15 / §17.18.40: <w:highlight w:val="none"/> is the spec's
        // explicit "no highlight" override — must round-trip to
        // `Some(HighlightColor::None)`, not a parse error.
        let (rp, _) = parse(r#"<rPr><highlight val="none"/></rPr>"#);
        assert_eq!(rp.highlight, Dup::from(Some(HighlightColor::None)));
    }

    #[test]
    fn vertical_align_superscript() {
        let (rp, _) = parse(r#"<rPr><vertAlign val="superscript"/></rPr>"#);
        assert_eq!(
            rp.vertical_align,
            Dup::from(Some(VerticalAlign::Superscript))
        );
    }

    #[test]
    fn text_scale_parsed() {
        // §17.3.2.45: <w:w w:val="80"/> compresses character width to 80%.
        let (rp, _) = parse(r#"<rPr><w val="80"/></rPr>"#);
        assert_eq!(rp.text_scale, Dup::from(Some(TextScale::new(80))));
        assert_eq!(rp.text_scale.cloned().unwrap().percent(), 80);
    }

    #[test]
    fn text_scale_absent_is_none() {
        // No <w:w> element → inherit from style cascade.
        let (rp, _) = parse(r#"<rPr><b/></rPr>"#);
        assert_eq!(rp.text_scale, Dup::from(None));
    }

    #[test]
    fn text_scale_clamps_above_600() {
        // §17.18.81: ST_TextScale max is 600.
        let (rp, _) = parse(r#"<rPr><w val="999"/></rPr>"#);
        assert_eq!(rp.text_scale, Dup::from(Some(TextScale::new(600))));
    }

    #[test]
    fn text_scale_zero_normalizes_to_100() {
        // Word treats <w:w w:val="0"/> as the default 100%.
        let (rp, _) = parse(r#"<rPr><w val="0"/></rPr>"#);
        assert_eq!(rp.text_scale, Dup::from(Some(TextScale::NORMAL)));
    }

    #[test]
    fn negative_decimal_font_size_is_rejected() {
        let parsed: Result<RPrXml, _> = quick_xml::de::from_str(r#"<rPr><sz val="-1.5"/></rPr>"#);
        assert!(parsed.is_err(), "negative font sizes must be rejected");
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
        assert_eq!(rp.spacing.map(|d| d.raw()), Dup::from(Some(40)));
        assert_eq!(rp.kerning.map(|d| d.raw()), Dup::from(Some(20)));
        assert_eq!(rp.position.map(|d| d.raw()), Dup::from(Some(-4)));
    }

    #[test]
    fn lang_tri_mode() {
        let (rp, _) = parse(r#"<rPr><lang val="en-US" eastAsia="ja-JP"/></rPr>"#);
        let l = rp.lang.cloned().unwrap();
        assert_eq!(l.val.as_deref(), Some("en-US"));
        assert_eq!(l.east_asia.as_deref(), Some("ja-JP"));
    }

    #[test]
    fn border_via_bdr() {
        let (rp, _) = parse(r#"<rPr><bdr val="single" sz="4" color="000000"/></rPr>"#);
        let b = rp.border.cloned().unwrap();
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
    fn duplicate_toggle_is_tolerated_last_wins() {
        // Real-world LibreOffice DOCX writers occasionally emit duplicate
        // self-closing toggles like `<w:b/><w:b/>`. Word renders these without
        // complaint — last-wins semantics means the second copy is a no-op.
        // The derived serde impl would error with `duplicate field`; the
        // manual Deserialize impl on RPrXml must accept it.
        let (rp, _) = parse(r#"<rPr><b/><b/></rPr>"#);
        assert_eq!(rp.bold, Some(true));
    }

    #[test]
    fn duplicate_toggle_last_wins_when_values_differ() {
        // If two duplicate toggles disagree, last wins.
        let (rp, _) = parse(r#"<rPr><b val="0"/><b/></rPr>"#);
        assert_eq!(rp.bold, Some(true));
        let (rp, _) = parse(r#"<rPr><b/><b val="0"/></rPr>"#);
        assert_eq!(rp.bold, Some(false));
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
        assert_eq!(rp.font_size.map(|d| d.raw()), Dup::from(Some(28)));
        assert_eq!(rp.color, Dup::from(Some(Color::Rgb(0x2E74B5))));
        assert_eq!(rp.underline, Dup::from(Some(UnderlineStyle::Single)));
    }

    /// A duplicated **non-toggle** child is schema-invalid and Word opens it
    /// anyway; see `primitives::duplicates` for why the last one wins.
    #[test]
    fn duplicate_non_toggle_children_are_tolerated_last_wins() {
        let (rp, _) = parse(
            r#"<rPr>
                 <sz val="20"/><sz val="28"/>
                 <color val="FF0000"/><color val="2E74B5"/>
               </rPr>"#,
        );
        assert_eq!(rp.font_size.get().map(|d| d.raw()), Some(28), "§17.3.2.38");
        assert_eq!(rp.color.get(), Some(&Color::Rgb(0x2E74B5)), "§17.3.2.6");
        // Both occurrences reach the model; only the read resolves.
        assert_eq!(rp.font_size.all().len(), 2);
        assert_eq!(rp.color.all().len(), 2);
    }
}
