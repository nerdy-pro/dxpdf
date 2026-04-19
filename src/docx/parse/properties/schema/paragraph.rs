//! `<w:pPr>` schema (§17.3.1 paragraph properties).
//!
//! Entry point: `PPrXml::split()` returns `ParsedParagraphProperties` —
//! direct formatting, style id, mark-run properties, and an optional
//! nested `<w:sectPr>` (§17.6.18 "last-paragraph-of-section" marker).

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
use crate::docx::parse::primitives::OnOff;

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
}

#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct PPrXml {
    #[serde(rename = "pStyle", default)]
    p_style: Option<ValString>,
    #[serde(default)]
    ind: Option<IndXml>,
    #[serde(default)]
    spacing: Option<SpacingXml>,
    #[serde(default)]
    jc: Option<ValAttr<StJc>>,
    #[serde(default)]
    shd: Option<ShdXml>,
    #[serde(rename = "outlineLvl", default)]
    outline_lvl: Option<ValAttr<u8>>,
    #[serde(rename = "numPr", default)]
    num_pr: Option<NumPrXml>,
    #[serde(default)]
    tabs: Option<TabsXml>,
    #[serde(rename = "pBdr", default)]
    p_bdr: Option<ParagraphBordersXml>,
    #[serde(rename = "rPr", default)]
    r_pr: Option<RPrXml>,
    #[serde(rename = "sectPr", default)]
    sect_pr: Option<SectPrXml>,
    #[serde(rename = "textAlignment", default)]
    text_alignment: Option<ValAttr<StTextAlignment>>,
    #[serde(rename = "cnfStyle", default)]
    cnf_style: Option<CnfStyleXml>,
    #[serde(rename = "framePr", default)]
    frame_pr: Option<FramePrXml>,

    // OnOff toggles
    #[serde(rename = "keepNext", default)]
    keep_next: Option<OnOff>,
    #[serde(rename = "keepLines", default)]
    keep_lines: Option<OnOff>,
    #[serde(rename = "widowControl", default)]
    widow_control: Option<OnOff>,
    #[serde(rename = "pageBreakBefore", default)]
    page_break_before: Option<OnOff>,
    #[serde(rename = "suppressAutoHyphens", default)]
    suppress_auto_hyphens: Option<OnOff>,
    #[serde(rename = "contextualSpacing", default)]
    contextual_spacing: Option<OnOff>,
    #[serde(default)]
    bidi: Option<OnOff>,
    #[serde(rename = "wordWrap", default)]
    word_wrap: Option<OnOff>,
    #[serde(rename = "autoSpaceDE", default)]
    auto_space_de: Option<OnOff>,
    #[serde(rename = "autoSpaceDN", default)]
    auto_space_dn: Option<OnOff>,
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
    #[serde(rename = "@before", default)]
    before: Option<Dimension<Twips>>,
    #[serde(rename = "@after", default)]
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

#[derive(Clone, Copy, Debug, Deserialize)]
struct NumPrXml {
    #[serde(default)]
    ilvl: Option<ValAttr<u8>>,
    #[serde(rename = "numId", default)]
    num_id: Option<ValAttr<i64>>,
}

/// `<w:framePr>` — legacy frame positioning. Splits by `@dropCap`:
/// `drop`/`margin` → `FrameKind::DropCap`; absent or `none` → `TextBox`.
#[derive(Clone, Copy, Debug, Deserialize)]
struct FramePrXml {
    #[serde(rename = "@dropCap", default)]
    drop_cap: Option<StDropCap>,
    #[serde(rename = "@lines", default)]
    lines: Option<u32>,
    #[serde(rename = "@hSpace", default)]
    h_space: Option<Dimension<Twips>>,
    #[serde(rename = "@vSpace", default)]
    v_space: Option<Dimension<Twips>>,
    #[serde(rename = "@w", default)]
    w: Option<Dimension<Twips>>,
    #[serde(rename = "@h", default)]
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
        let style_id = self.p_style.map(|v| StyleId::new(v.val));

        let (run_properties, _run_style_id) = match self.r_pr {
            Some(r) => {
                let (rp, sid) = r.split();
                (Some(rp), sid)
            }
            None => (None, None),
        };
        // rStyle inside pPr/rPr applies to the paragraph mark only; the
        // legacy parser discards this style id too.

        let section_properties = self.sect_pr.map(Into::into);

        let properties = ParagraphProperties {
            alignment: self.jc.map(|j| Alignment::from(j.val)),
            indentation: self.ind.map(Into::into),
            spacing: self.spacing.map(Into::into),
            numbering: self.num_pr.and_then(numbering_ref),
            tabs: self.tabs.map(<Vec<TabStop>>::from).unwrap_or_default(),
            borders: self.p_bdr.map(ParagraphBorders::from),
            shading: self.shd.map(Shading::from),
            keep_next: self.keep_next.map(|OnOff(b)| b),
            keep_lines: self.keep_lines.map(|OnOff(b)| b),
            widow_control: self.widow_control.map(|OnOff(b)| b),
            page_break_before: self.page_break_before.map(|OnOff(b)| b),
            suppress_auto_hyphens: self.suppress_auto_hyphens.map(|OnOff(b)| b),
            contextual_spacing: self.contextual_spacing.map(|OnOff(b)| b),
            bidi: self.bidi.map(|OnOff(b)| b),
            word_wrap: self.word_wrap.map(|OnOff(b)| b),
            outline_level: self
                .outline_lvl
                .and_then(|v| OutlineLevel::from_ooxml(v.val)),
            text_alignment: self.text_alignment.map(|v| TextAlignment::from(v.val)),
            cnf_style: self.cnf_style.map(CnfStyle::from),
            frame_properties: self.frame_pr.map(FrameKind::from),
            auto_space_de: self.auto_space_de.map(|OnOff(b)| b),
            auto_space_dn: self.auto_space_dn.map(|OnOff(b)| b),
        };

        ParsedPPr {
            properties,
            style_id,
            run_properties,
            section_properties,
        }
    }
}

fn numbering_ref(x: NumPrXml) -> Option<NumberingReference> {
    let num_id = x.num_id?;
    Some(NumberingReference {
        num_id: num_id.val,
        level: x.ilvl.map(|v| v.val).unwrap_or(0),
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
        assert_eq!(r.properties.alignment, None);
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
        assert_eq!(r.properties.alignment, None);
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
        assert_eq!(p.alignment, Some(Alignment::Both));
        assert_eq!(p.indentation.unwrap().start.unwrap().raw(), 720);
        match p.indentation.unwrap().first_line {
            Some(FirstLineIndent::FirstLine(d)) => assert_eq!(d.raw(), 360),
            other => panic!("expected FirstLine, got {other:?}"),
        }
        match p.spacing.unwrap().line {
            Some(LineSpacing::Auto(d)) => assert_eq!(d.raw(), 360),
            other => panic!("expected Auto, got {other:?}"),
        }
        assert_eq!(p.keep_next, Some(true));
        assert_eq!(p.keep_lines, Some(false));
        assert_eq!(p.outline_level.map(|o| o.value()), Some(1));
        assert_eq!(p.text_alignment, Some(TextAlignment::Center));
    }

    #[test]
    fn indentation_legacy_left_right_aliases() {
        let r = parse(r#"<pPr><ind left="720" right="360"/></pPr>"#);
        let ind = r.properties.indentation.unwrap();
        assert_eq!(ind.start.unwrap().raw(), 720);
        assert_eq!(ind.end.unwrap().raw(), 360);
    }

    #[test]
    fn num_pr_both_ilvl_and_num_id() {
        let r = parse(r#"<pPr><numPr><ilvl val="2"/><numId val="5"/></numPr></pPr>"#);
        let n = r.properties.numbering.unwrap();
        assert_eq!(n.level, 2);
        assert_eq!(n.num_id, 5);
    }

    #[test]
    fn num_pr_without_num_id_is_none() {
        let r = parse(r#"<pPr><numPr><ilvl val="1"/></numPr></pPr>"#);
        assert!(r.properties.numbering.is_none());
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
        assert_eq!(p.borders.unwrap().top.unwrap().style, BorderStyle::Single);
        assert_eq!(p.shading.unwrap().pattern, ShadingPattern::Solid);
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
        assert_eq!(sp.page_size.unwrap().width.unwrap().raw(), 12240);
    }

    #[test]
    fn frame_pr_drop_cap() {
        let r = parse(r#"<pPr><framePr dropCap="drop" lines="2"/></pPr>"#);
        match r.properties.frame_properties {
            Some(FrameKind::DropCap { style, lines, .. }) => {
                assert_eq!(style, DropCap::Drop);
                assert_eq!(lines, 2);
            }
            other => panic!("expected DropCap, got {other:?}"),
        }
    }

    #[test]
    fn frame_pr_text_box_default() {
        let r = parse(r#"<pPr><framePr w="5000" h="3000" hAnchor="margin"/></pPr>"#);
        match r.properties.frame_properties {
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
        assert_eq!(r.properties.cnf_style, Some(CnfStyle::FIRST_ROW));
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
}
