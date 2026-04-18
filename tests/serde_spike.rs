//! Mixed-content spike per `docs/serde-migration-plan.md` pre-work.
//!
//! Validates that quick-xml 0.37's serde layer handles three patterns that
//! the plan's Phase 4 depends on:
//!
//! 1. `#[serde(rename = "$value")]` + untagged enum of child element types
//!    preserves **document order** of heterogeneous children.
//! 2. A struct containing both `$value` children and `$text` content works
//!    alongside sibling child elements (for `<w:t>`-style text runs).
//! 3. Attribute-bearing leaf elements deserialize correctly as enum variants.
//!
//! If any assertion here fails, Phase 2 of the plan is blocked.

use serde::Deserialize;

// ── Pattern 1 ──────────────────────────────────────────────────────────────
// A paragraph containing runs and hyperlinks in arbitrary order. Mimics the
// <w:p> shape without any properties.

#[derive(Debug, Deserialize)]
struct ParagraphXml {
    #[serde(rename = "$value", default)]
    children: Vec<ParagraphChild>,
}

#[derive(Debug, Deserialize)]
enum ParagraphChild {
    #[serde(rename = "r")]
    Run(RunXml),
    #[serde(rename = "hyperlink")]
    Hyperlink(HyperlinkXml),
}

#[derive(Debug, Deserialize)]
struct RunXml {
    #[serde(rename = "@id", default)]
    id: String,
}

#[derive(Debug, Deserialize)]
struct HyperlinkXml {
    #[serde(rename = "@target", default)]
    target: String,
}

#[test]
fn mixed_children_preserve_order() {
    let xml = r#"<p>
        <r id="a"/>
        <hyperlink target="first"/>
        <r id="b"/>
        <r id="c"/>
        <hyperlink target="second"/>
        <r id="d"/>
    </p>"#;

    let p: ParagraphXml = quick_xml::de::from_str(xml).unwrap();

    assert_eq!(p.children.len(), 6);
    match &p.children[0] {
        ParagraphChild::Run(r) => assert_eq!(r.id, "a"),
        other => panic!("child 0: expected Run, got {other:?}"),
    }
    match &p.children[1] {
        ParagraphChild::Hyperlink(h) => assert_eq!(h.target, "first"),
        other => panic!("child 1: expected Hyperlink, got {other:?}"),
    }
    match &p.children[2] {
        ParagraphChild::Run(r) => assert_eq!(r.id, "b"),
        other => panic!("child 2: expected Run, got {other:?}"),
    }
    match &p.children[3] {
        ParagraphChild::Run(r) => assert_eq!(r.id, "c"),
        other => panic!("child 3: expected Run, got {other:?}"),
    }
    match &p.children[4] {
        ParagraphChild::Hyperlink(h) => assert_eq!(h.target, "second"),
        other => panic!("child 4: expected Hyperlink, got {other:?}"),
    }
    match &p.children[5] {
        ParagraphChild::Run(r) => assert_eq!(r.id, "d"),
        other => panic!("child 5: expected Run, got {other:?}"),
    }
}

// ── Pattern 2 ──────────────────────────────────────────────────────────────
// A run containing text nodes, tabs, and breaks interleaved. Mimics <w:r>
// shape. Text nodes are wrapped in `<t>`.

#[derive(Debug, Deserialize)]
struct Run2Xml {
    #[serde(rename = "$value", default)]
    children: Vec<RunChild>,
}

#[derive(Debug, Deserialize)]
enum RunChild {
    #[serde(rename = "t")]
    Text(TextXml),
    #[serde(rename = "tab")]
    Tab,
    #[serde(rename = "br")]
    Br,
}

#[derive(Debug, Deserialize)]
struct TextXml {
    #[serde(rename = "$text", default)]
    content: String,
    #[serde(rename = "@xml:space", default)]
    space: Option<String>,
}

#[test]
fn run_text_tabs_breaks_preserve_order() {
    let xml = r#"<r>
        <t>Hello</t>
        <tab/>
        <t xml:space="preserve"> world</t>
        <br/>
        <t>next line</t>
    </r>"#;

    let r: Run2Xml = quick_xml::de::from_str(xml).unwrap();
    assert_eq!(r.children.len(), 5);

    match &r.children[0] {
        RunChild::Text(t) => {
            assert_eq!(t.content, "Hello");
            assert!(t.space.is_none());
        }
        other => panic!("child 0: expected Text, got {other:?}"),
    }
    assert!(matches!(&r.children[1], RunChild::Tab));
    match &r.children[2] {
        RunChild::Text(t) => {
            assert_eq!(t.content, " world");
            assert_eq!(t.space.as_deref(), Some("preserve"));
        }
        other => panic!("child 2: expected Text, got {other:?}"),
    }
    assert!(matches!(&r.children[3], RunChild::Br));
    match &r.children[4] {
        RunChild::Text(t) => assert_eq!(t.content, "next line"),
        other => panic!("child 4: expected Text, got {other:?}"),
    }
}

// ── Pattern 3 ──────────────────────────────────────────────────────────────
// Paragraph with a leading property element (pPr analogue) plus the mixed
// children after it. Validates that `$value` coexists with named siblings.

#[derive(Debug, Deserialize)]
struct Paragraph3Xml {
    #[serde(rename = "pPr", default)]
    p_pr: Option<PPrXml>,
    #[serde(rename = "$value", default)]
    children: Vec<ParagraphChild>,
}

#[derive(Debug, Deserialize)]
struct PPrXml {
    #[serde(rename = "@style", default)]
    style: String,
}

#[test]
fn named_sibling_plus_value_children() {
    let xml = r#"<p>
        <pPr style="Heading1"/>
        <r id="x"/>
        <hyperlink target="home"/>
        <r id="y"/>
    </p>"#;

    let p: Paragraph3Xml = quick_xml::de::from_str(xml).unwrap();
    assert_eq!(p.p_pr.as_ref().map(|x| x.style.as_str()), Some("Heading1"));
    assert_eq!(p.children.len(), 3);
    assert!(matches!(&p.children[0], ParagraphChild::Run(r) if r.id == "x"));
    assert!(matches!(&p.children[1], ParagraphChild::Hyperlink(h) if h.target == "home"));
    assert!(matches!(&p.children[2], ParagraphChild::Run(r) if r.id == "y"));
}
