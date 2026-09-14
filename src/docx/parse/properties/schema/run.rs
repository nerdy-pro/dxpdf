//! `<w:rPr>` schema (§17.3.2 run properties).
//!
//! Carries every direct run-formatting element plus the shared sub-schemas
//! from sibling modules. Deserializes to `(RunProperties, Option<StyleId>)`
//! via the `split` method — the style id is routed separately because the
//! property cascade applies it before direct formatting.
//!
//! `w14_shadow` is the one field here that isn't `w:`-namespaced: a
//! Word-2010 `<w14:shadow>` text-effect extension (blur/offset/color, not
//! modeled at the run level) shares the local name `shadow` with §17.3.2.37's
//! ordinary boolean toggle below. `quick_xml::de` ordinarily matches struct
//! fields by local name only, discarding the namespace, which would collide
//! the two into one `Vec` and fail the parse the moment they're non-adjacent
//! — confirmed against a real document (issue reproduction
//! `joern.hendrich@vdwbayern.de.docx`). `vendor/quick-xml-0.41.0` is a
//! dxpdf-patched quick-xml (see that directory's `PATCH.md`) that recognizes
//! a `#[serde(rename)]` written in Clark notation (`"{uri}local"`) and
//! matches it against the element's *resolved* namespace instead — so this
//! field claims only the Word-2010 extension, and the plain `shadow` field
//! below claims only the real toggle, however either producer bound its own
//! prefix. Every other field here is unaffected: none of their renames use
//! that notation, so they still match by local name exactly as before.

use crate::model::Dup;
use serde::de::IgnoredAny;
use serde::{Deserialize, Deserializer};

/// §17.18.81 ST_TextScale: a bare percent number (`80`, the Transitional
/// spelling) or the same number with a `%` sign (`80%`, what a Strict export
/// writes). Both land on the one 0..=600 scale `TextScale::new` clamps.
#[derive(Clone, Copy, Debug)]
struct TextScaleXml(u16);

impl<'de> Deserialize<'de> for TextScaleXml {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        let raw = String::deserialize(deserializer)?;
        let head = raw.strip_suffix('%').unwrap_or(&raw);
        // §22.9.2.9/§17.18.81 admit no `+` sign — matches the rejection
        // every other percent/universal-measure spelling in this codebase
        // applies (`parse_decimal_scaled`'s `allow_plus: false`).
        if head.starts_with('+') {
            return Err(serde::de::Error::custom(
                "expected a §17.18.81 text-scale percentage",
            ));
        }
        head.parse::<u16>()
            .map(TextScaleXml)
            .map_err(|_| serde::de::Error::custom("expected a §17.18.81 text-scale percentage"))
    }
}

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
    char_scale: Vec<ValAttr<TextScaleXml>>,

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
    /// Word-2010 `<w14:shadow>` — see the module doc. Recognized only so it
    /// stops colliding with `shadow` above; its content (blur/offset/color)
    /// is not modeled and is discarded.
    #[serde(
        rename = "{http://schemas.microsoft.com/office/word/2010/wordml}shadow",
        default
    )]
    w14_shadow: Vec<IgnoredAny>,

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
            text_scale: Dup::from(self.char_scale).map(|v| TextScale::new(v.val.0)),
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
        // §17.18.81 also admits the percent spelling a Strict export writes.
        let (rp, _) = parse(r#"<rPr><w val="80%"/></rPr>"#);
        assert_eq!(rp.text_scale, Dup::from(Some(TextScale::new(80))));

        // §17.3.2.45: <w:w w:val="80"/> compresses character width to 80%.
        let (rp, _) = parse(r#"<rPr><w val="80"/></rPr>"#);
        assert_eq!(rp.text_scale, Dup::from(Some(TextScale::new(80))));
        assert_eq!(rp.text_scale.cloned().unwrap().percent(), 80);
    }

    #[test]
    fn text_scale_rejects_plus_sign() {
        // §17.18.81/§22.9.2.9 admit no `+` sign — matches the rejection
        // every other percent/universal-measure spelling in this codebase
        // applies (see units.rs's `percent_sign_travels_and_plus_is_rejected`).
        let parsed: Result<RPrXml, _> = quick_xml::de::from_str(r#"<rPr><w val="+80%"/></rPr>"#);
        assert!(parsed.is_err(), "a leading + must be rejected");
        let parsed: Result<RPrXml, _> = quick_xml::de::from_str(r#"<rPr><w val="+80"/></rPr>"#);
        assert!(parsed.is_err(), "a leading + must be rejected");
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

    /// The collision the module doc describes: a real `<w:shadow>` toggle
    /// and a Word-2010 `<w14:shadow>` extension, non-adjacent, in one
    /// `<rPr>`. Without the dxpdf-patched quick-xml's namespace-qualified
    /// matching this is a hard parse error (`duplicate field "shadow"`);
    /// with it, each element reaches its own field regardless of which
    /// prefix the producer chose for either namespace.
    #[test]
    fn w14_shadow_extension_does_not_collide_with_the_shadow_toggle() {
        let xml = r#"<rPr xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
            <shadow val="0"/>
            <outline val="0"/>
            <w14:shadow w14:blurRad="0"><w14:srgbClr w14:val="000000"/></w14:shadow>
        </rPr>"#;
        let (rp, _) = parse(xml);
        assert_eq!(rp.shadow, Some(false));
    }

    /// The same collision, but with the Word-2010 namespace bound to a
    /// producer-chosen prefix other than the conventional `w14` — matching
    /// happens on the *resolved namespace URI*, not the literal prefix text.
    #[test]
    fn w14_shadow_extension_matches_by_namespace_not_by_prefix_spelling() {
        let xml = r#"<rPr xmlns:ext="http://schemas.microsoft.com/office/word/2010/wordml">
            <shadow val="0"/>
            <ext:shadow ext:blurRad="0"/>
        </rPr>"#;
        let (rp, _) = parse(xml);
        assert_eq!(rp.shadow, Some(false));
    }
}
