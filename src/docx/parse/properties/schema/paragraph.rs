//! `<w:pPr>` schema (§17.3.1 paragraph properties).
//!
//! Entry point: `PPrXml::split()` returns `ParsedParagraphProperties` —
//! direct formatting, style id, mark-run properties, and an optional
//! nested `<w:sectPr>` (§17.6.18 "last-paragraph-of-section" marker).

use crate::model::Dup;
use serde::Deserialize;

use crate::docx::model::dimension::{Dimension, Twips};
use crate::docx::model::{
    Alignment, CnfStyle, DropCap, FirstLineIndent, FrameKind, FrameWrap, HeightRule, Indentation,
    LineSpacing, NumberingReference, OutlineLevel, ParagraphBorders, ParagraphProperties,
    ParagraphSpacing, RunProperties, Shading, StyleId, TabStop, TextAlignment, TextBoxPositioning,
};
use crate::docx::parse::primitives::st_enums::{
    StAnchor, StFrameWrap, StHeightRule, StJc, StLineSpacingRule, StTextAlignment, StXAlign,
    StYAlign,
};
use crate::docx::parse::primitives::units::deserialize_optional_nonnegative_dimension;
use crate::docx::parse::primitives::{last_toggle, OnOff};

use super::border::ParagraphBordersXml;
use super::cnf_style::CnfStyleXml;
use super::run::RPrXml;
use super::section::SectPrXml;
use super::shading::ShdXml;
use super::tabs::TabsXml;

/// All the artifacts produced by deserializing a `<w:pPr>`. The split
/// mirrors the legacy `ParsedParagraphProperties` so it plugs into the
/// existing resolve pipeline unchanged.
pub(crate) struct ParsedPPr {
    pub properties: ParagraphProperties,
    pub style_id: Option<StyleId>,
    pub run_properties: Option<RunProperties>,
    pub section_properties: Option<crate::docx::model::SectionProperties>,
    /// §17.13.5.15: the paragraph mark itself is tracked-deleted — how Word
    /// spells a whole-paragraph delete (issue #154).
    pub mark_deleted: bool,
}

/// Schema for the `<w:pPr>` element (§17.3.1).
///
/// Every child is typed `Vec<T>`, not `Option<T>`, so a producer that repeats
/// one cannot fail the parse. `split` carries the non-toggle children into the
/// model as `Dup<T>` — every occurrence survives, and the reader picks — while
/// the toggles collapse here via `last_toggle`, which is §17.7.2's own rule
/// rather than this parser's. The policy and the reasoning behind "last wins"
/// for everything else are in `crate::docx::parse::primitives::duplicates`.
#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct PPrXml {
    #[serde(rename = "pStyle", default)]
    p_style: Vec<ValString>,
    #[serde(default)]
    ind: Vec<IndXml>,
    #[serde(default)]
    spacing: Vec<SpacingXml>,
    #[serde(default)]
    jc: Vec<ValAttr<StJc>>,
    #[serde(default)]
    shd: Vec<ShdXml>,
    #[serde(rename = "outlineLvl", default)]
    outline_lvl: Vec<ValAttr<u8>>,
    #[serde(rename = "numPr", default)]
    num_pr: Vec<NumPrXml>,
    #[serde(default)]
    tabs: Vec<TabsXml>,
    #[serde(rename = "pBdr", default)]
    p_bdr: Vec<ParagraphBordersXml>,
    #[serde(rename = "rPr", default)]
    r_pr: Vec<RPrXml>,
    #[serde(rename = "sectPr", default)]
    sect_pr: Vec<SectPrXml>,
    #[serde(rename = "textAlignment", default)]
    text_alignment: Vec<ValAttr<StTextAlignment>>,
    #[serde(rename = "cnfStyle", default)]
    cnf_style: Vec<CnfStyleXml>,
    #[serde(rename = "framePr", default)]
    frame_pr: Vec<FramePrXml>,

    // OnOff toggles, collapsed by `last_toggle` rather than `last` so the
    // `OnOff` wrapper comes off at the same time. §17.7.2 is the citation for
    // these and only these; every other child above is `Vec` under this
    // parser's own policy, in `primitives::duplicates`.
    #[serde(rename = "keepNext", default)]
    keep_next: Vec<OnOff>,
    #[serde(rename = "keepLines", default)]
    keep_lines: Vec<OnOff>,
    #[serde(rename = "widowControl", default)]
    widow_control: Vec<OnOff>,
    #[serde(rename = "pageBreakBefore", default)]
    page_break_before: Vec<OnOff>,
    #[serde(rename = "suppressAutoHyphens", default)]
    suppress_auto_hyphens: Vec<OnOff>,
    #[serde(rename = "contextualSpacing", default)]
    contextual_spacing: Vec<OnOff>,
    #[serde(default)]
    bidi: Vec<OnOff>,
    #[serde(rename = "wordWrap", default)]
    word_wrap: Vec<OnOff>,
    #[serde(rename = "autoSpaceDE", default)]
    auto_space_de: Vec<OnOff>,
    #[serde(rename = "autoSpaceDN", default)]
    auto_space_dn: Vec<OnOff>,
}

#[derive(Clone, Debug, Deserialize)]
struct ValString {
    #[serde(rename = "@val")]
    val: String,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(bound(deserialize = "T: serde::Deserialize<'de>"))]
struct ValAttr<T> {
    #[serde(rename = "@val")]
    val: T,
}

/// `<w:ind>` — indentation. Legacy `@left`/`@right` alias `@start`/`@end`.
/// `@firstLine` and `@hanging` are mutually exclusive; when both present,
/// hanging wins per renderer convention (legacy parser matched this).
#[derive(Clone, Copy, Debug, Deserialize)]
struct IndXml {
    #[serde(rename = "@start", alias = "@left", default)]
    start: Option<Dimension<Twips>>,
    #[serde(rename = "@end", alias = "@right", default)]
    end: Option<Dimension<Twips>>,
    #[serde(rename = "@firstLine", default)]
    first_line: Option<Dimension<Twips>>,
    #[serde(rename = "@hanging", default)]
    hanging: Option<Dimension<Twips>>,
    #[serde(rename = "@mirrorIndents", default)]
    mirror: Option<AttrBool>,
}

impl From<IndXml> for Indentation {
    fn from(x: IndXml) -> Self {
        let first_line = match (x.first_line, x.hanging) {
            (_, Some(h)) => Some(FirstLineIndent::Hanging(h)),
            (Some(f), None) => Some(FirstLineIndent::FirstLine(f)),
            (None, None) => None,
        };
        Self {
            start: x.start,
            end: x.end,
            first_line,
            mirror: x.mirror.map(|b| b.0),
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
struct SpacingXml {
    #[serde(
        rename = "@before",
        default,
        deserialize_with = "deserialize_optional_nonnegative_dimension"
    )]
    before: Option<Dimension<Twips>>,
    #[serde(
        rename = "@after",
        default,
        deserialize_with = "deserialize_optional_nonnegative_dimension"
    )]
    after: Option<Dimension<Twips>>,
    #[serde(rename = "@line", default)]
    line: Option<Dimension<Twips>>,
    #[serde(rename = "@lineRule", default)]
    line_rule: Option<StLineSpacingRule>,
    #[serde(rename = "@beforeAutospacing", default)]
    before_auto: Option<AttrBool>,
    #[serde(rename = "@afterAutospacing", default)]
    after_auto: Option<AttrBool>,
}

impl From<SpacingXml> for ParagraphSpacing {
    fn from(x: SpacingXml) -> Self {
        let line = x
            .line
            .map(|v| match x.line_rule.unwrap_or(StLineSpacingRule::Auto) {
                StLineSpacingRule::Auto => LineSpacing::Auto(v),
                StLineSpacingRule::Exact => LineSpacing::Exact(v),
                StLineSpacingRule::AtLeast => LineSpacing::AtLeast(v),
            });
        Self {
            before: x.before,
            after: x.after,
            line,
            before_auto_spacing: x.before_auto.map(|b| b.0),
            after_auto_spacing: x.after_auto.map(|b| b.0),
        }
    }
}

#[derive(Clone, Debug, Deserialize)]
struct NumPrXml {
    #[serde(default)]
    ilvl: Vec<ValAttr<u8>>,
    #[serde(rename = "numId", default)]
    num_id: Vec<ValAttr<i64>>,
}

/// `<w:framePr>` — legacy frame positioning. Splits by `@dropCap`:
/// `drop`/`margin` → `FrameKind::DropCap`; absent or `none` → `TextBox`.
#[derive(Clone, Copy, Debug, Deserialize)]
struct FramePrXml {
    #[serde(rename = "@dropCap", default)]
    drop_cap: Option<StDropCap>,
    #[serde(rename = "@lines", default)]
    lines: Option<u32>,
    #[serde(
        rename = "@hSpace",
        default,
        deserialize_with = "deserialize_optional_nonnegative_dimension"
    )]
    h_space: Option<Dimension<Twips>>,
    #[serde(
        rename = "@vSpace",
        default,
        deserialize_with = "deserialize_optional_nonnegative_dimension"
    )]
    v_space: Option<Dimension<Twips>>,
    #[serde(
        rename = "@w",
        default,
        deserialize_with = "deserialize_optional_nonnegative_dimension"
    )]
    w: Option<Dimension<Twips>>,
    #[serde(
        rename = "@h",
        default,
        deserialize_with = "deserialize_optional_nonnegative_dimension"
    )]
    h: Option<Dimension<Twips>>,
    #[serde(rename = "@hRule", default)]
    h_rule: Option<StHeightRule>,
    #[serde(rename = "@wrap", default)]
    wrap: Option<StFrameWrap>,
    #[serde(rename = "@hAnchor", default)]
    h_anchor: Option<StAnchor>,
    #[serde(rename = "@vAnchor", default)]
    v_anchor: Option<StAnchor>,
    #[serde(rename = "@x", default)]
    x: Option<Dimension<Twips>>,
    #[serde(rename = "@y", default)]
    y: Option<Dimension<Twips>>,
    #[serde(rename = "@xAlign", default)]
    x_align: Option<StXAlign>,
    #[serde(rename = "@yAlign", default)]
    y_align: Option<StYAlign>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
enum StDropCap {
    None,
    Drop,
    Margin,
}

impl From<FramePrXml> for FrameKind {
    fn from(x: FramePrXml) -> Self {
        match x.drop_cap {
            Some(StDropCap::Drop) => Self::DropCap {
                style: DropCap::Drop,
                lines: x.lines.unwrap_or(3),
                h_space: x.h_space,
            },
            Some(StDropCap::Margin) => Self::DropCap {
                style: DropCap::Margin,
                lines: x.lines.unwrap_or(3),
                h_space: x.h_space,
            },
            Some(StDropCap::None) | None => Self::TextBox(TextBoxPositioning {
                width: x.w,
                height: x.h,
                height_rule: x.h_rule.map(HeightRule::from),
                h_space: x.h_space,
                v_space: x.v_space,
                wrap: x.wrap.map(FrameWrap::from),
                h_anchor: x.h_anchor.map(Into::into),
                v_anchor: x.v_anchor.map(Into::into),
                x: x.x,
                y: x.y,
                x_align: x.x_align.map(Into::into),
                y_align: x.y_align.map(Into::into),
            }),
        }
    }
}

use crate::docx::parse::primitives::AttrBool;

impl PPrXml {
    pub(crate) fn split(self) -> ParsedPPr {
        let style_id = Dup::from(self.p_style)
            .into_value()
            .map(|v| StyleId::new(v.val));

        let (run_properties, mark_deleted) = match Dup::from(self.r_pr).into_value() {
            Some(r) => {
                // §17.13.5.15: read the mark deletion before `split` consumes
                // the bag.
                let mark_deleted = r.mark_deleted();
                let (rp, _sid) = r.split();
                (Some(rp), mark_deleted)
            }
            None => (None, false),
        };
        // rStyle inside pPr/rPr applies to the paragraph mark only; the
        // legacy parser discards this style id too.

        let section_properties = Dup::from(self.sect_pr).into_value().map(Into::into);

        let properties = ParagraphProperties {
            alignment: Dup::from(self.jc).map(|j| Alignment::from(j.val)),
            indentation: Dup::from(self.ind).map(Into::into),
            spacing: Dup::from(self.spacing).map(Into::into),
            numbering: Dup::from(self.num_pr).filter_map(numbering_ref),
            // The one child that resolves here rather than at the read; the
            // field's doc comment on `ParagraphProperties` says why.
            tabs: Dup::from(self.tabs)
                .into_value()
                .map(<Vec<TabStop>>::from)
                .unwrap_or_default(),
            borders: Dup::from(self.p_bdr).map(ParagraphBorders::from),
            shading: Dup::from(self.shd).map(Shading::from),
            keep_next: last_toggle(self.keep_next),
            keep_lines: last_toggle(self.keep_lines),
            widow_control: last_toggle(self.widow_control),
            page_break_before: last_toggle(self.page_break_before),
            suppress_auto_hyphens: last_toggle(self.suppress_auto_hyphens),
            contextual_spacing: last_toggle(self.contextual_spacing),
            bidi: last_toggle(self.bidi),
            word_wrap: last_toggle(self.word_wrap),
            outline_level: Dup::from(self.outline_lvl)
                .filter_map(|v| OutlineLevel::from_ooxml(v.val)),
            text_alignment: Dup::from(self.text_alignment).map(|v| TextAlignment::from(v.val)),
            cnf_style: Dup::from(self.cnf_style).map(CnfStyle::from),
            frame_properties: Dup::from(self.frame_pr).map(FrameKind::from),
            auto_space_de: last_toggle(self.auto_space_de),
            auto_space_dn: last_toggle(self.auto_space_dn),
        };

        ParsedPPr {
            properties,
            style_id,
            run_properties,
            section_properties,
            mark_deleted,
        }
    }
}

fn numbering_ref(x: NumPrXml) -> Option<NumberingReference> {
    let num_id = Dup::from(x.num_id).into_value()?;
    Some(NumberingReference {
        num_id: num_id.val,
        level: Dup::from(x.ilvl).into_value().map(|v| v.val).unwrap_or(0),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::{Alignment, BorderStyle, DropCap, ShadingPattern, TextAlignment};

    fn parse(xml: &str) -> ParsedPPr {
        let x: PPrXml = quick_xml::de::from_str(xml).unwrap();
        x.split()
    }

    #[test]
    fn empty_pprx_produces_defaults() {
        let r = parse(r#"<pPr/>"#);
        assert_eq!(r.properties.alignment, Dup::from(None));
        assert!(r.style_id.is_none());
        assert!(r.run_properties.is_none());
        assert!(r.section_properties.is_none());
    }

    #[test]
    fn p_style_routed_separately() {
        let r = parse(r#"<pPr><pStyle val="Heading1"/></pPr>"#);
        assert_eq!(
            r.style_id.map(|s| s.as_str().to_string()),
            Some("Heading1".into())
        );
        assert_eq!(r.properties.alignment, Dup::from(None));
    }

    #[test]
    fn direct_formatting_batch() {
        let r = parse(
            r#"<pPr>
                <jc val="both"/>
                <ind start="720" firstLine="360"/>
                <spacing before="120" after="240" line="360" lineRule="auto"/>
                <keepNext/>
                <keepLines val="false"/>
                <outlineLvl val="0"/>
                <textAlignment val="center"/>
            </pPr>"#,
        );
        let p = r.properties;
        assert_eq!(p.alignment, Dup::from(Some(Alignment::Both)));
        assert_eq!(p.indentation.get().unwrap().start.unwrap().raw(), 720);
        match p.indentation.get().unwrap().first_line {
            Some(FirstLineIndent::FirstLine(d)) => assert_eq!(d.raw(), 360),
            other => panic!("expected FirstLine, got {other:?}"),
        }
        match p.spacing.get().unwrap().line {
            Some(LineSpacing::Auto(d)) => assert_eq!(d.raw(), 360),
            other => panic!("expected Auto, got {other:?}"),
        }
        assert_eq!(p.keep_next, Some(true));
        assert_eq!(p.keep_lines, Some(false));
        assert_eq!(p.outline_level.map(|o| o.value()), Dup::from(Some(1)));
        assert_eq!(p.text_alignment, Dup::from(Some(TextAlignment::Center)));
    }

    #[test]
    fn indentation_legacy_left_right_aliases() {
        let r = parse(r#"<pPr><ind left="720" right="360"/></pPr>"#);
        let ind = r.properties.indentation.get().unwrap();
        assert_eq!(ind.start.unwrap().raw(), 720);
        assert_eq!(ind.end.unwrap().raw(), 360);
    }

    #[test]
    fn negative_decimal_indentation_remains_valid() {
        let r = parse(r#"<pPr><ind start="-1.5"/></pPr>"#);
        assert_eq!(
            r.properties.indentation.get().unwrap().start.unwrap().raw(),
            -2
        );
    }

    #[test]
    fn num_pr_both_ilvl_and_num_id() {
        let r = parse(r#"<pPr><numPr><ilvl val="2"/><numId val="5"/></numPr></pPr>"#);
        let n = r.properties.numbering.get().unwrap();
        assert_eq!(n.level, 2);
        assert_eq!(n.num_id, 5);
    }

    #[test]
    fn num_pr_without_num_id_is_none() {
        let r = parse(r#"<pPr><numPr><ilvl val="1"/></numPr></pPr>"#);
        assert!(r.properties.numbering.get().is_none());
    }

    #[test]
    fn borders_shading_and_tabs() {
        let r = parse(
            r#"<pPr>
                <pBdr><top val="single"/></pBdr>
                <shd val="solid" fill="FFFF00"/>
                <tabs><tab pos="1440" val="center"/></tabs>
            </pPr>"#,
        );
        let p = r.properties;
        assert_eq!(
            p.borders.get().unwrap().top.unwrap().style,
            BorderStyle::Single
        );
        assert_eq!(p.shading.get().unwrap().pattern, ShadingPattern::Solid);
        assert_eq!(p.tabs.len(), 1);
        assert_eq!(p.tabs[0].position.raw(), 1440);
    }

    #[test]
    fn mark_run_properties_split_out() {
        let r = parse(r#"<pPr><rPr><b/><color val="FF0000"/></rPr></pPr>"#);
        let rp = r.run_properties.unwrap();
        assert_eq!(rp.bold, Some(true));
    }

    #[test]
    fn nested_sect_pr_routed_separately() {
        let r = parse(r#"<pPr><sectPr><pgSz w="12240" h="15840"/></sectPr></pPr>"#);
        let sp = r.section_properties.unwrap();
        assert_eq!(sp.page_size.get().unwrap().width.unwrap().raw(), 12240);
    }

    #[test]
    fn frame_pr_drop_cap() {
        let r = parse(r#"<pPr><framePr dropCap="drop" lines="2"/></pPr>"#);
        match r.properties.frame_properties.get() {
            Some(FrameKind::DropCap { style, lines, .. }) => {
                assert_eq!(*style, DropCap::Drop);
                assert_eq!(*lines, 2);
            }
            other => panic!("expected DropCap, got {other:?}"),
        }
    }

    #[test]
    fn frame_pr_text_box_default() {
        let r = parse(r#"<pPr><framePr w="5000" h="3000" hAnchor="margin"/></pPr>"#);
        match r.properties.frame_properties.get() {
            Some(FrameKind::TextBox(tb)) => {
                assert_eq!(tb.width.unwrap().raw(), 5000);
                assert_eq!(tb.height.unwrap().raw(), 3000);
            }
            other => panic!("expected TextBox, got {other:?}"),
        }
    }

    #[test]
    fn cnf_style_binary_val() {
        let r = parse(r#"<pPr><cnfStyle val="100000000000"/></pPr>"#);
        assert_eq!(r.properties.cnf_style, Dup::from(Some(CnfStyle::FIRST_ROW)));
    }

    #[test]
    fn all_ten_toggles() {
        let r = parse(
            r#"<pPr>
                <keepNext/><keepLines/><widowControl/><pageBreakBefore/>
                <suppressAutoHyphens/><contextualSpacing/><bidi/><wordWrap/>
                <autoSpaceDE/><autoSpaceDN/>
            </pPr>"#,
        );
        let p = r.properties;
        assert_eq!(p.keep_next, Some(true));
        assert_eq!(p.keep_lines, Some(true));
        assert_eq!(p.widow_control, Some(true));
        assert_eq!(p.page_break_before, Some(true));
        assert_eq!(p.suppress_auto_hyphens, Some(true));
        assert_eq!(p.contextual_spacing, Some(true));
        assert_eq!(p.bidi, Some(true));
        assert_eq!(p.word_wrap, Some(true));
        assert_eq!(p.auto_space_de, Some(true));
        assert_eq!(p.auto_space_dn, Some(true));
    }

    #[test]
    fn unknown_jc_is_strict() {
        let r: Result<PPrXml, _> = quick_xml::de::from_str(r#"<pPr><jc val="bogus"/></pPr>"#);
        assert!(r.is_err());
    }

    #[test]
    fn duplicate_toggles_are_tolerated_last_wins() {
        // LibreOffice/AOO emit redundant duplicate toggles. With `Option<OnOff>`
        // serde would fail with "duplicate field" and take down the whole parse;
        // `Vec<OnOff>` + last_toggle accepts them (§17.7.2 last wins).
        let r = parse(r#"<pPr><keepNext/><keepNext/></pPr>"#);
        assert_eq!(r.properties.keep_next, Some(true));

        // When duplicates disagree, the last one wins.
        let r = parse(r#"<pPr><widowControl val="1"/><widowControl val="0"/></pPr>"#);
        assert_eq!(r.properties.widow_control, Some(false));
    }

    /// A duplicated **non-toggle** child is schema-invalid and Word opens it
    /// anyway; see `primitives::duplicates` for why the last one wins.
    #[test]
    fn duplicate_non_toggle_children_are_tolerated_last_wins() {
        let r = parse(
            r#"<pPr>
                 <jc val="left"/><jc val="center"/>
                 <ind left="100"/><ind left="720"/>
               </pPr>"#,
        );
        assert_eq!(
            r.properties.alignment.get(),
            Some(&Alignment::Center),
            "§17.3.1.13"
        );
        assert_eq!(
            r.properties
                .indentation
                .get()
                .and_then(|i| i.start)
                .map(|d| d.raw()),
            Some(720),
            "§17.3.1.12"
        );
        // Resolving at the read is the point: both occurrences reached the
        // model, so a consumer that wants the first one still can.
        assert_eq!(
            r.properties.alignment.all(),
            &[Alignment::Start, Alignment::Center]
        );
        assert_eq!(r.properties.indentation.all().len(), 2);
    }
}
