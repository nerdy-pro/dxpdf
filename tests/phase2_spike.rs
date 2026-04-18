//! Phase 2 spike per `docs/serde-migration-plan.md`.
//!
//! A narrow `PPrXml` with five representative children that exercise every
//! mechanical pattern the full property port will rely on:
//!
//! 1. **ST_* enum** — `<w:jc w:val="center"/>` via `StJc` → `Alignment`.
//! 2. **`Dimension<U>` attribute** — `<w:ind w:start="720" .../>`.
//! 3. **Toggle with OnOff semantics** — `<w:keepNext/>`.
//! 4. **Constructor-validated primitive** — `<w:outlineLvl w:val="0"/>`
//!    maps via `OutlineLevel::from_ooxml`.
//! 5. **Composite discriminated union** — `<w:spacing w:line="360"
//!    w:lineRule="auto"/>` collapses rule + value into `LineSpacing::Auto`.
//!
//! If any of these fail, the full Phase 2 port is blocked.

use dxpdf::docx::model::{
    Alignment, FirstLineIndent, Indentation, LineSpacing, OutlineLevel, ParagraphProperties,
    ParagraphSpacing,
};
use dxpdf::model::dimension::{Dimension, Twips};
use dxpdf::docx::parse::primitives::st_enums::{StJc, StLineSpacingRule};
use dxpdf::docx::parse::primitives::OnOff;
use serde::Deserialize;

#[derive(Debug, Deserialize, Default)]
struct PPrXml {
    #[serde(default)]
    jc: Option<ValAttr<StJc>>,
    #[serde(default)]
    ind: Option<IndXml>,
    #[serde(default)]
    spacing: Option<SpacingXml>,
    #[serde(rename = "keepNext", default)]
    keep_next: Option<OnOff>,
    #[serde(rename = "outlineLvl", default)]
    outline_lvl: Option<ValAttr<u8>>,
}

#[derive(Debug, Deserialize)]
struct ValAttr<T> {
    #[serde(rename = "@val")]
    val: T,
}

#[derive(Debug, Deserialize, Default)]
struct IndXml {
    #[serde(rename = "@start", alias = "@left", default)]
    start: Option<Dimension<Twips>>,
    #[serde(rename = "@end", alias = "@right", default)]
    end: Option<Dimension<Twips>>,
    #[serde(rename = "@firstLine", default)]
    first_line: Option<Dimension<Twips>>,
    #[serde(rename = "@hanging", default)]
    hanging: Option<Dimension<Twips>>,
}

#[derive(Debug, Deserialize, Default)]
struct SpacingXml {
    #[serde(rename = "@before", default)]
    before: Option<Dimension<Twips>>,
    #[serde(rename = "@after", default)]
    after: Option<Dimension<Twips>>,
    #[serde(rename = "@line", default)]
    line: Option<Dimension<Twips>>,
    #[serde(rename = "@lineRule", default)]
    line_rule: Option<StLineSpacingRule>,
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
            mirror: None,
        }
    }
}

impl From<SpacingXml> for ParagraphSpacing {
    fn from(x: SpacingXml) -> Self {
        let line = x.line.map(|v| match x.line_rule.unwrap_or(StLineSpacingRule::Auto) {
            StLineSpacingRule::Auto => LineSpacing::Auto(v),
            StLineSpacingRule::Exact => LineSpacing::Exact(v),
            StLineSpacingRule::AtLeast => LineSpacing::AtLeast(v),
        });
        Self {
            before: x.before,
            after: x.after,
            line,
            before_auto_spacing: None,
            after_auto_spacing: None,
        }
    }
}

impl From<PPrXml> for ParagraphProperties {
    fn from(x: PPrXml) -> Self {
        let mut p = ParagraphProperties::default();
        p.alignment = x.jc.map(|j| j.val.into());
        p.indentation = x.ind.map(Into::into);
        p.spacing = x.spacing.map(Into::into);
        p.keep_next = x.keep_next.map(|OnOff(b)| b);
        p.outline_level = x
            .outline_lvl
            .and_then(|v| OutlineLevel::from_ooxml(v.val));
        p
    }
}

// ── Tests ────────────────────────────────────────────────────────────────

fn parse(xml: &str) -> ParagraphProperties {
    quick_xml::de::from_str::<PPrXml>(xml)
        .expect("deserialize pPr")
        .into()
}

#[test]
fn empty_ppr_yields_default() {
    let p = parse(r#"<w:pPr xmlns:w="urn:w"/>"#);
    assert_eq!(p.alignment, None);
    assert_eq!(p.indentation, None);
    assert_eq!(p.spacing, None);
    assert_eq!(p.keep_next, None);
    assert_eq!(p.outline_level, None);
}

#[test]
fn jc_round_trips_through_st_enum() {
    let p = parse(r#"<w:pPr xmlns:w="urn:w"><w:jc w:val="center"/></w:pPr>"#);
    assert_eq!(p.alignment, Some(Alignment::Center));

    // spec alias: "justify" → Alignment::Both
    let p = parse(r#"<w:pPr xmlns:w="urn:w"><w:jc w:val="justify"/></w:pPr>"#);
    assert_eq!(p.alignment, Some(Alignment::Both));

    // Start/End: model uses directional names, OOXML uses presentation names.
    let p = parse(r#"<w:pPr xmlns:w="urn:w"><w:jc w:val="left"/></w:pPr>"#);
    assert_eq!(p.alignment, Some(Alignment::Start));
}

#[test]
fn indentation_captures_twips() {
    let p = parse(
        r#"<w:pPr xmlns:w="urn:w">
            <w:ind w:start="720" w:end="360" w:firstLine="180"/>
        </w:pPr>"#,
    );
    let ind = p.indentation.expect("indentation present");
    assert_eq!(ind.start.map(|d| d.raw()), Some(720));
    assert_eq!(ind.end.map(|d| d.raw()), Some(360));
    match ind.first_line {
        Some(FirstLineIndent::FirstLine(d)) => assert_eq!(d.raw(), 180),
        other => panic!("expected FirstLine, got {other:?}"),
    }
}

#[test]
fn indentation_hanging_takes_precedence() {
    let p = parse(r#"<w:pPr xmlns:w="urn:w"><w:ind w:hanging="360"/></w:pPr>"#);
    let ind = p.indentation.unwrap();
    match ind.first_line {
        Some(FirstLineIndent::Hanging(d)) => assert_eq!(d.raw(), 360),
        other => panic!("expected Hanging, got {other:?}"),
    }
}

#[test]
fn spacing_auto_composite() {
    let p = parse(
        r#"<w:pPr xmlns:w="urn:w">
            <w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>
        </w:pPr>"#,
    );
    let s = p.spacing.expect("spacing present");
    assert_eq!(s.before.map(|d| d.raw()), Some(120));
    assert_eq!(s.after.map(|d| d.raw()), Some(240));
    match s.line {
        Some(LineSpacing::Auto(d)) => assert_eq!(d.raw(), 360),
        other => panic!("expected LineSpacing::Auto, got {other:?}"),
    }
}

#[test]
fn spacing_exact_composite() {
    let p = parse(
        r#"<w:pPr xmlns:w="urn:w">
            <w:spacing w:line="480" w:lineRule="exact"/>
        </w:pPr>"#,
    );
    match p.spacing.unwrap().line {
        Some(LineSpacing::Exact(d)) => assert_eq!(d.raw(), 480),
        other => panic!("expected LineSpacing::Exact, got {other:?}"),
    }
}

#[test]
fn spacing_at_least_composite() {
    let p = parse(
        r#"<w:pPr xmlns:w="urn:w">
            <w:spacing w:line="400" w:lineRule="atLeast"/>
        </w:pPr>"#,
    );
    match p.spacing.unwrap().line {
        Some(LineSpacing::AtLeast(d)) => assert_eq!(d.raw(), 400),
        other => panic!("expected LineSpacing::AtLeast, got {other:?}"),
    }
}

#[test]
fn keep_next_toggle_present_is_true() {
    let p = parse(r#"<w:pPr xmlns:w="urn:w"><w:keepNext/></w:pPr>"#);
    assert_eq!(p.keep_next, Some(true));
}

#[test]
fn keep_next_toggle_off_is_false() {
    let p = parse(r#"<w:pPr xmlns:w="urn:w"><w:keepNext w:val="false"/></w:pPr>"#);
    assert_eq!(p.keep_next, Some(false));
}

#[test]
fn outline_lvl_constructor_conversion() {
    // OOXML value 0 → Heading 1 → OutlineLevel(1)
    let p = parse(r#"<w:pPr xmlns:w="urn:w"><w:outlineLvl w:val="0"/></w:pPr>"#);
    assert_eq!(p.outline_level.map(|o| o.value()), Some(1));

    // OOXML value 8 → Heading 9 → OutlineLevel(9)
    let p = parse(r#"<w:pPr xmlns:w="urn:w"><w:outlineLvl w:val="8"/></w:pPr>"#);
    assert_eq!(p.outline_level.map(|o| o.value()), Some(9));

    // Out-of-range 9 → None (OutlineLevel::from_ooxml refuses)
    let p = parse(r#"<w:pPr xmlns:w="urn:w"><w:outlineLvl w:val="9"/></w:pPr>"#);
    assert_eq!(p.outline_level, None);
}

#[test]
fn end_to_end_all_five_children() {
    let xml = r#"<w:pPr xmlns:w="urn:w">
        <w:jc w:val="both"/>
        <w:ind w:start="720" w:firstLine="360"/>
        <w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>
        <w:keepNext/>
        <w:outlineLvl w:val="1"/>
    </w:pPr>"#;
    let p = parse(xml);

    assert_eq!(p.alignment, Some(Alignment::Both));
    assert_eq!(
        p.indentation.map(|i| (i.start.map(|d| d.raw()), i.first_line)),
        Some((
            Some(720),
            Some(FirstLineIndent::FirstLine(Dimension::<Twips>::new(360)))
        ))
    );
    match p.spacing.and_then(|s| s.line) {
        Some(LineSpacing::Auto(d)) => assert_eq!(d.raw(), 240),
        other => panic!("expected Auto(240), got {other:?}"),
    }
    assert_eq!(p.keep_next, Some(true));
    assert_eq!(p.outline_level.map(|o| o.value()), Some(2));
}

#[test]
fn unknown_jc_value_is_strict() {
    let r: Result<PPrXml, _> =
        quick_xml::de::from_str(r#"<w:pPr xmlns:w="urn:w"><w:jc w:val="bogus"/></w:pPr>"#);
    assert!(r.is_err(), "expected strict rejection, got {r:?}");
}
