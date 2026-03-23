//! API compatibility tests.
//!
//! These tests verify that the public API surface of dxpdf remains stable.
//! They don't test behavior — they test that types, methods, functions, and
//! constants exist with the expected signatures. If any of these tests fail
//! after a change, it means the public API has changed in a breaking way.

use std::collections::HashMap;
use std::rc::Rc;

// ---------------------------------------------------------------------------
// Top-level re-exports & entry-point functions
// ---------------------------------------------------------------------------

#[test]
fn top_level_convert_signature() {
    let _: fn(&[u8]) -> Result<Vec<u8>, dxpdf::Error> = dxpdf::convert;
}

#[test]
fn top_level_convert_document_signature() {
    let _: fn(&dxpdf::model::Document) -> Result<Vec<u8>, dxpdf::Error> = dxpdf::convert_document;
}

// ---------------------------------------------------------------------------
// error module
// ---------------------------------------------------------------------------

#[test]
fn error_variants_exist() {
    // Verify all Error variants are constructible.
    let _ = dxpdf::Error::MissingEntry(String::new());
    let _ = dxpdf::Error::Render(String::new());
    // Io, Zip, Xml, XmlAttr are From-converted; just verify the enum is Debug + Display.
    let err = dxpdf::Error::Render("test".into());
    let _ = format!("{err}");
    let _ = format!("{err:?}");
}

#[test]
fn error_is_std_error() {
    fn assert_std_error<T: std::error::Error>() {}
    assert_std_error::<dxpdf::Error>();
}

// ---------------------------------------------------------------------------
// model module — DEFAULT_FONT_FAMILY constant
// ---------------------------------------------------------------------------

#[test]
fn model_default_font_family_constant() {
    let _: &str = dxpdf::model::DEFAULT_FONT_FAMILY;
}

// ---------------------------------------------------------------------------
// model module — type aliases
// ---------------------------------------------------------------------------

#[test]
fn model_type_aliases() {
    let _: dxpdf::model::ImageStore = HashMap::new();
    let _: dxpdf::model::StyleMap = HashMap::new();
    let _: dxpdf::model::NumberingMap = HashMap::new();
}

// ---------------------------------------------------------------------------
// model module — Document
// ---------------------------------------------------------------------------

#[test]
fn document_default_and_fields() {
    let doc = dxpdf::model::Document::default();

    let _: &Vec<dxpdf::model::Block> = &doc.blocks;
    let _: &Option<dxpdf::model::SectionProperties> = &doc.final_section;
    let _: dxpdf::dimension::Twips = doc.default_tab_stop;
    let _: dxpdf::dimension::HalfPoints = doc.default_font_size;
    let _: &Rc<str> = &doc.default_font_family;
    let _: &dxpdf::model::Spacing = &doc.default_spacing;
    let _: &dxpdf::model::CellMargins = &doc.default_cell_margins;
    let _: &dxpdf::model::Spacing = &doc.table_cell_spacing;
    let _: &dxpdf::model::TableBorders = &doc.default_table_borders;
    let _: &dxpdf::model::StyleMap = &doc.styles;
    let _: &dxpdf::model::NumberingMap = &doc.numbering;
    let _: &Option<dxpdf::model::HeaderFooter> = &doc.default_header;
    let _: &Option<dxpdf::model::HeaderFooter> = &doc.default_footer;
    let _: &dxpdf::model::ImageStore = &doc.images;
}

#[test]
fn document_font_families_method() {
    let doc = dxpdf::model::Document::default();
    let _: Vec<Rc<str>> = doc.font_families();
}

// ---------------------------------------------------------------------------
// model module — Block & Paragraph
// ---------------------------------------------------------------------------

#[test]
fn block_variants() {
    use dxpdf::model::*;

    let p = Paragraph {
        properties: ParagraphProperties::default(),
        runs: vec![],
        floats: vec![],
        section_properties: None,
    };
    let _ = Block::Paragraph(Box::new(p));

    let t = Table {
        rows: vec![],
        grid_cols: vec![],
        default_cell_margins: None,
        cell_spacing: None,
        borders: None,
    };
    let _ = Block::Table(Box::new(t));
}

#[test]
fn paragraph_fields() {
    use dxpdf::model::*;
    let p = Paragraph {
        properties: ParagraphProperties::default(),
        runs: vec![],
        floats: vec![],
        section_properties: None,
    };
    let _: &ParagraphProperties = &p.properties;
    let _: &Vec<Inline> = &p.runs;
    let _: &Vec<FloatingImage> = &p.floats;
    let _: &Option<SectionProperties> = &p.section_properties;
}

#[test]
fn paragraph_properties_fields() {
    use dxpdf::model::*;
    let pp = ParagraphProperties::default();
    let _: &Option<Alignment> = &pp.alignment;
    let _: &Option<Spacing> = &pp.spacing;
    let _: &Option<Indentation> = &pp.indentation;
    let _: &Vec<TabStop> = &pp.tab_stops;
    let _: &Option<Color> = &pp.shading;
    let _: &Option<String> = &pp.style_id;
    let _: &Option<ListRef> = &pp.list_ref;
    let _: &Option<ParagraphBorders> = &pp.paragraph_borders;
}

// ---------------------------------------------------------------------------
// model module — Inline & TextRun
// ---------------------------------------------------------------------------

#[test]
fn inline_variants() {
    use dxpdf::model::*;
    let _ = Inline::LineBreak;
    let _ = Inline::Tab;
    let _ = Inline::TextRun(TextRun {
        text: String::new(),
        properties: RunProperties::default(),
        hyperlink_url: None,
    });
    let _ = Inline::Image(InlineImage {
        rel_id: RelId::from("rId1"),
        size: dxpdf::geometry::PtSize::new(
            dxpdf::dimension::Pt::new(0.0),
            dxpdf::dimension::Pt::new(0.0),
        ),
    });
    let _ = Inline::Field(FieldCode {
        field_type: FieldType::Page,
        properties: RunProperties::default(),
    });
}

#[test]
fn text_run_fields() {
    use dxpdf::model::*;
    let tr = TextRun {
        text: "hello".into(),
        properties: RunProperties::default(),
        hyperlink_url: Some("https://example.com".into()),
    };
    let _: &str = &tr.text;
    let _: &RunProperties = &tr.properties;
    let _: &Option<String> = &tr.hyperlink_url;
}

#[test]
fn run_properties_fields_and_methods() {
    use dxpdf::model::*;
    let rp = RunProperties::default();
    let _: bool = rp.bold;
    let _: bool = rp.italic;
    let _: bool = rp.underline;
    let _: &Option<dxpdf::dimension::HalfPoints> = &rp.font_size;
    let _: &Option<Rc<str>> = &rp.font_family;
    let _: &Option<Color> = &rp.color;
    let _: &Option<dxpdf::dimension::Twips> = &rp.char_spacing;
    let _: &Option<Color> = &rp.shading;
    let _: &Option<VertAlign> = &rp.vert_align;
    let _: &Option<String> = &rp.style_id;
}

// ---------------------------------------------------------------------------
// model module — Resolved styles
// ---------------------------------------------------------------------------

#[test]
fn resolved_paragraph_style_fields() {
    use dxpdf::model::*;
    let s = ResolvedParagraphStyle::default();
    let _: &Option<Alignment> = &s.alignment;
    let _: &Option<Spacing> = &s.spacing;
    let _: &Option<Indentation> = &s.indentation;
    let _: &ResolvedRunStyle = &s.run_props;
}

#[test]
fn resolved_run_style_fields() {
    use dxpdf::model::*;
    let s = ResolvedRunStyle::default();
    let _: &Option<bool> = &s.bold;
    let _: &Option<bool> = &s.italic;
    let _: &Option<bool> = &s.underline;
    let _: &Option<dxpdf::dimension::HalfPoints> = &s.font_size;
    let _: &Option<Rc<str>> = &s.font_family;
    let _: &Option<Color> = &s.color;
}

// ---------------------------------------------------------------------------
// model module — Spacing, Indentation, Alignment, LineRule, LineSpacing
// ---------------------------------------------------------------------------

#[test]
fn spacing_fields_and_methods() {
    use dxpdf::model::*;
    let s = Spacing {
        before: Some(dxpdf::dimension::Twips::new(0)),
        after: Some(dxpdf::dimension::Twips::new(0)),
        line: Some(dxpdf::dimension::Twips::new(0)),
        line_rule: LineRule::Auto,
    };
    let _: Option<LineSpacing> = s.line_spacing();
    let _: dxpdf::dimension::Pt = s.line_pt();
}

#[test]
fn alignment_variants() {
    use dxpdf::model::Alignment;
    let _ = Alignment::Left;
    let _ = Alignment::Center;
    let _ = Alignment::Right;
    let _ = Alignment::Justify;
}

#[test]
fn line_rule_variants() {
    use dxpdf::model::LineRule;
    let _ = LineRule::Auto;
    let _ = LineRule::Exact;
    let _ = LineRule::AtLeast;
}

#[test]
fn line_spacing_variants() {
    use dxpdf::model::LineSpacing;
    let _ = LineSpacing::Multiplier(1.0);
    let _ = LineSpacing::Fixed(dxpdf::dimension::Pt::new(12.0));
    let _ = LineSpacing::AtLeast(dxpdf::dimension::Pt::new(12.0));
}

// ---------------------------------------------------------------------------
// model module — TabStop, TabStopType
// ---------------------------------------------------------------------------

#[test]
fn tab_stop_fields_and_methods() {
    use dxpdf::model::*;
    let ts = TabStop {
        position: dxpdf::dimension::Twips::new(720),
        stop_type: TabStopType::Left,
    };
    let _: dxpdf::dimension::Twips = ts.position;
    let _: f32 = f32::from(ts.position);
}

#[test]
fn tab_stop_type_variants() {
    use dxpdf::model::TabStopType;
    let _ = TabStopType::Left;
    let _ = TabStopType::Center;
    let _ = TabStopType::Right;
    let _ = TabStopType::Decimal;
}

// ---------------------------------------------------------------------------
// model module — Page & Section
// ---------------------------------------------------------------------------

#[test]
fn page_size_fields_and_methods() {
    use dxpdf::model::*;
    let ps = PageSize::new(
        dxpdf::dimension::Twips::new(12240),
        dxpdf::dimension::Twips::new(15840),
    );
    let _: f32 = f32::from(ps.width);
    let _: f32 = f32::from(ps.height);
}

#[test]
fn page_margins_fields_and_methods() {
    use dxpdf::model::*;
    let pm = PageMargins {
        top: dxpdf::dimension::Twips::new(1440),
        right: dxpdf::dimension::Twips::new(1440),
        bottom: dxpdf::dimension::Twips::new(1440),
        left: dxpdf::dimension::Twips::new(1440),
        header: dxpdf::dimension::Twips::new(720),
        footer: dxpdf::dimension::Twips::new(720),
    };
    let _: f32 = f32::from(pm.top);
    let _: f32 = f32::from(pm.right);
    let _: f32 = f32::from(pm.bottom);
    let _: f32 = f32::from(pm.left);
    let _: f32 = f32::from(pm.header);
    let _: f32 = f32::from(pm.footer);
}

#[test]
fn section_properties_fields() {
    use dxpdf::model::*;
    let _doc = Document::default();
    // SectionProperties is Option on Document; test fields via construction.
    let sp = SectionProperties {
        page_size: None,
        page_margins: None,
        header: None,
        footer: None,
        header_rel_id: None,
        footer_rel_id: None,
    };
    let _: &Option<PageSize> = &sp.page_size;
    let _: &Option<PageMargins> = &sp.page_margins;
    let _: &Option<HeaderFooter> = &sp.header;
    let _: &Option<HeaderFooter> = &sp.footer;
    let _: &Option<String> = &sp.header_rel_id;
    let _: &Option<String> = &sp.footer_rel_id;
}

#[test]
fn header_footer_fields() {
    use dxpdf::model::*;
    let hf = HeaderFooter { blocks: vec![] };
    let _: &Vec<Block> = &hf.blocks;
}

// ---------------------------------------------------------------------------
// model module — Color
// ---------------------------------------------------------------------------

#[test]
fn color_fields_and_methods() {
    use dxpdf::model::Color;
    let c = Color {
        r: 255,
        g: 0,
        b: 128,
    };
    let _: u8 = c.r;
    let _: u8 = c.g;
    let _: u8 = c.b;

    let parsed = Color::from_hex("FF0080");
    assert!(parsed.is_some());
}

// ---------------------------------------------------------------------------
// model module — Borders
// ---------------------------------------------------------------------------

#[test]
fn border_style_variants() {
    use dxpdf::model::BorderStyle;
    let _ = BorderStyle::None;
    let _ = BorderStyle::Single;
    let _ = BorderStyle::Double;
    let _ = BorderStyle::Dashed;
    let _ = BorderStyle::Dotted;
}

#[test]
fn border_def_fields_and_methods() {
    use dxpdf::model::*;
    let bd = BorderDef::single(8, Color::BLACK);
    let _: &BorderStyle = &bd.style;
    let _: dxpdf::dimension::EighthPoints = bd.size;
    let _: &Color = &bd.color;
    let _: bool = bd.is_visible();
}

#[test]
fn table_borders_fields() {
    use dxpdf::model::*;
    let bd = BorderDef::single(4, Color::BLACK);
    let tb = TableBorders {
        top: bd,
        bottom: bd,
        left: bd,
        right: bd,
        inside_h: bd,
        inside_v: bd,
    };
    let _ = &tb.top;
    let _ = &tb.bottom;
    let _ = &tb.left;
    let _ = &tb.right;
    let _ = &tb.inside_h;
    let _ = &tb.inside_v;
}

#[test]
fn cell_borders_fields() {
    use dxpdf::model::*;
    let cb = CellBorders {
        top: None,
        bottom: None,
        left: None,
        right: None,
    };
    let _: &Option<BorderDef> = &cb.top;
    let _: &Option<BorderDef> = &cb.bottom;
    let _: &Option<BorderDef> = &cb.left;
    let _: &Option<BorderDef> = &cb.right;
}

#[test]
fn paragraph_borders_fields() {
    use dxpdf::model::*;
    let pb = ParagraphBorders {
        top: None,
        bottom: None,
        left: None,
        right: None,
    };
    let _: &Option<BorderDef> = &pb.top;
    let _: &Option<BorderDef> = &pb.bottom;
    let _: &Option<BorderDef> = &pb.left;
    let _: &Option<BorderDef> = &pb.right;
}

// ---------------------------------------------------------------------------
// model module — Tables
// ---------------------------------------------------------------------------

#[test]
fn table_fields() {
    use dxpdf::model::*;
    let t = Table {
        rows: vec![],
        grid_cols: vec![],
        default_cell_margins: None,
        cell_spacing: None,
        borders: None,
    };
    let _: &Vec<TableRow> = &t.rows;
    let _: &Vec<dxpdf::dimension::Twips> = &t.grid_cols;
    let _: &Option<CellMargins> = &t.default_cell_margins;
    let _: &Option<dxpdf::model::Spacing> = &t.cell_spacing;
    let _: &Option<TableBorders> = &t.borders;
}

#[test]
fn table_row_fields() {
    use dxpdf::model::*;
    let row = TableRow {
        cells: vec![],
        height: Some(dxpdf::dimension::Twips::new(720)),
    };
    let _: &Vec<TableCell> = &row.cells;
    let _: &Option<dxpdf::dimension::Twips> = &row.height;
}

#[test]
fn table_cell_fields_and_methods() {
    use dxpdf::model::*;
    let cell = TableCell {
        blocks: vec![],
        width: Some(dxpdf::dimension::Twips::new(3000)),
        grid_span: 1,
        vertical_merge: None,
        cell_margins: None,
        cell_borders: None,
        shading: None,
    };
    let _: &Vec<Block> = &cell.blocks;
    let _: &Option<dxpdf::dimension::Twips> = &cell.width;
    let _: u32 = cell.grid_span;
    let _: &Option<VerticalMerge> = &cell.vertical_merge;
    let _: &Option<CellMargins> = &cell.cell_margins;
    let _: &Option<CellBorders> = &cell.cell_borders;
    let _: &Option<Color> = &cell.shading;
    let _: Option<f32> = cell.width.map(f32::from);
    let _: bool = cell.is_vmerge_continue();
}

#[test]
fn cell_margins_fields_and_methods() {
    use dxpdf::model::*;
    let cm = CellMargins::new(
        dxpdf::dimension::Twips::new(0),
        dxpdf::dimension::Twips::new(108),
        dxpdf::dimension::Twips::new(0),
        dxpdf::dimension::Twips::new(108),
    );
    let _: f32 = f32::from(cm.top);
    let _: f32 = f32::from(cm.right);
    let _: f32 = f32::from(cm.bottom);
    let _: f32 = f32::from(cm.left);
}

#[test]
fn vertical_merge_variants() {
    use dxpdf::model::VerticalMerge;
    let _ = VerticalMerge::Restart;
    let _ = VerticalMerge::Continue;
}

// ---------------------------------------------------------------------------
// model module — Images, Fields, Lists
// ---------------------------------------------------------------------------

#[test]
fn inline_image_fields() {
    use dxpdf::model::*;
    let img = InlineImage {
        rel_id: RelId::from("rId1"),
        size: dxpdf::geometry::PtSize::new(
            dxpdf::dimension::Pt::new(100.0),
            dxpdf::dimension::Pt::new(100.0),
        ),
    };
    let _: &RelId = &img.rel_id;
    let _: dxpdf::dimension::Pt = img.size.width;
    let _: dxpdf::dimension::Pt = img.size.height;
}

#[test]
fn floating_image_fields() {
    use dxpdf::model::*;
    let img = FloatingImage {
        rel_id: RelId::from("rId2"),
        size: dxpdf::geometry::PtSize::new(
            dxpdf::dimension::Pt::new(200.0),
            dxpdf::dimension::Pt::new(150.0),
        ),
        offset: dxpdf::geometry::PtOffset::new(
            dxpdf::dimension::Pt::new(10.0),
            dxpdf::dimension::Pt::new(20.0),
        ),
        align_h: None,
        align_v: None,
        wrap_side: WrapSide::BothSides,
        pct_pos_h: None,
        pct_pos_v: None,
    };
    let _: &RelId = &img.rel_id;
    let _: dxpdf::dimension::Pt = img.size.width;
    let _: dxpdf::dimension::Pt = img.size.height;
    let _: dxpdf::dimension::Pt = img.offset.x;
    let _: dxpdf::dimension::Pt = img.offset.y;
    let _: &Option<String> = &img.align_h;
    let _: &Option<String> = &img.align_v;
    let _: &WrapSide = &img.wrap_side;
    let _: &Option<i32> = &img.pct_pos_h;
    let _: &Option<i32> = &img.pct_pos_v;
}

#[test]
fn wrap_side_variants() {
    use dxpdf::model::WrapSide;
    let _ = WrapSide::Left;
    let _ = WrapSide::Right;
    let _ = WrapSide::BothSides;
}

#[test]
fn field_code_fields() {
    use dxpdf::model::*;
    let fc = FieldCode {
        field_type: FieldType::Page,
        properties: RunProperties::default(),
    };
    let _: &FieldType = &fc.field_type;
    let _: &RunProperties = &fc.properties;
}

#[test]
fn field_type_variants() {
    use dxpdf::model::FieldType;
    let _ = FieldType::Page;
    let _ = FieldType::NumPages;
}

#[test]
fn list_ref_fields() {
    use dxpdf::model::*;
    let lr = ListRef {
        num_id: 1,
        level: 0,
    };
    let _: u32 = lr.num_id;
    let _: u32 = lr.level;
}

#[test]
fn numbering_def_fields() {
    use dxpdf::model::*;
    let nd = NumberingDef { levels: vec![] };
    let _: &Vec<NumberingLevel> = &nd.levels;
}

#[test]
fn numbering_level_fields() {
    use dxpdf::model::*;
    let nl = NumberingLevel {
        format: NumberFormat::Decimal,
        level_text: "%1.".into(),
        start: 1,
        indent_left: dxpdf::dimension::Twips::new(720),
        indent_hanging: dxpdf::dimension::Twips::new(360),
    };
    let _: &NumberFormat = &nl.format;
    let _: &str = &nl.level_text;
    let _: u32 = nl.start;
    let _: dxpdf::dimension::Twips = nl.indent_left;
    let _: dxpdf::dimension::Twips = nl.indent_hanging;
}

#[test]
fn number_format_variants() {
    use dxpdf::model::NumberFormat;
    let _ = NumberFormat::Bullet("•".into());
    let _ = NumberFormat::Decimal;
    let _ = NumberFormat::LowerLetter;
    let _ = NumberFormat::UpperLetter;
    let _ = NumberFormat::LowerRoman;
    let _ = NumberFormat::UpperRoman;
}

// ---------------------------------------------------------------------------
// model module — VertAlign, RelId
// ---------------------------------------------------------------------------

#[test]
fn vert_align_variants() {
    use dxpdf::model::VertAlign;
    let _ = VertAlign::Superscript;
    let _ = VertAlign::Subscript;
}

#[test]
fn rel_id_api() {
    use dxpdf::model::RelId;
    let rid = RelId::from("rId5");
    let _: &str = rid.as_str();
    // Deref to str
    let _: &str = &rid;
}

// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// parse module
// ---------------------------------------------------------------------------

#[test]
fn parse_function_signature() {
    let _: fn(&[u8]) -> Result<dxpdf::model::Document, dxpdf::Error> = dxpdf::parse::parse;
}

#[test]
fn parse_xml_functions_exist() {
    // Verify the xml submodule functions are accessible.
    let _: fn(&str) -> Result<dxpdf::model::Document, dxpdf::Error> =
        dxpdf::parse::xml::parse_document_xml;

    let _: fn(&str, &HashMap<String, String>) -> Result<dxpdf::model::Document, dxpdf::Error> =
        dxpdf::parse::xml::parse_document_xml_with_rels;

    let _: fn(&str) -> Result<dxpdf::model::HeaderFooter, dxpdf::Error> =
        dxpdf::parse::xml::parse_header_footer_xml;

    let _: fn(&str, &HashMap<String, String>) -> Result<dxpdf::model::HeaderFooter, dxpdf::Error> =
        dxpdf::parse::xml::parse_header_footer_xml_with_rels;
}

// ---------------------------------------------------------------------------
// render module — fonts
// ---------------------------------------------------------------------------

#[test]
fn render_fonts_signatures() {
    use skia_safe::{Font, FontMgr, FontStyle, Typeface};

    let _: fn(&FontMgr, &str, FontStyle) -> Typeface = dxpdf::render::fonts::resolve_typeface;
    let _: fn(&FontMgr, &[Rc<str>]) = dxpdf::render::fonts::preload_fonts;
    let _: fn(&FontMgr, &str, dxpdf::dimension::Pt, bool, bool) -> Font =
        dxpdf::render::fonts::make_font;
}

// ---------------------------------------------------------------------------
// render module — layout
// ---------------------------------------------------------------------------

#[test]
fn layouted_page_fields() {
    use dxpdf::geometry::PtSize;
    use dxpdf::render::layout::{DrawCommand, LayoutedPage};
    let page = LayoutedPage {
        commands: vec![],
        page_size: PtSize::new(
            dxpdf::dimension::Pt::new(612.0),
            dxpdf::dimension::Pt::new(792.0),
        ),
    };
    let _: &Vec<DrawCommand> = &page.commands;
    let _: PtSize = page.page_size;
}

#[test]
fn draw_command_variants() {
    use dxpdf::dimension::Pt;
    use dxpdf::geometry::{PtLineSegment, PtOffset, PtRect};
    use dxpdf::model::Color;
    use dxpdf::render::layout::DrawCommand;
    let _ = DrawCommand::Text {
        position: PtOffset::new(Pt::new(0.0), Pt::new(0.0)),
        text: String::new(),
        font_family: Rc::from("Helvetica"),
        char_spacing_pt: Pt::new(0.0),
        font_size: Pt::new(12.0),
        bold: false,
        italic: false,
        color: Color::BLACK,
    };
    let _ = DrawCommand::Underline {
        line: PtLineSegment::new(
            PtOffset::new(Pt::new(0.0), Pt::new(0.0)),
            PtOffset::new(Pt::new(100.0), Pt::new(0.0)),
        ),
        color: Color::BLACK,
        width: Pt::new(1.0),
    };
    let _ = DrawCommand::Line {
        line: PtLineSegment::new(
            PtOffset::new(Pt::new(0.0), Pt::new(0.0)),
            PtOffset::new(Pt::new(100.0), Pt::new(100.0)),
        ),
        color: Color::BLACK,
        width: Pt::new(1.0),
    };
    let _ = DrawCommand::Rect {
        rect: PtRect::from_xywh(Pt::new(0.0), Pt::new(0.0), Pt::new(100.0), Pt::new(50.0)),
        color: Color::new(200, 200, 200),
    };
    let _ = DrawCommand::LinkAnnotation {
        rect: PtRect::from_xywh(Pt::new(0.0), Pt::new(0.0), Pt::new(100.0), Pt::new(12.0)),
        url: "https://example.com".into(),
    };
    // DrawCommand::Image requires an Rc<skia_safe::Image> — skip construction
    // but verify the variant exists via pattern matching.
}

#[test]
fn layout_function_exists() {
    // Verify layout + measure functions are accessible.
    use dxpdf::render::layout::layout;
    use dxpdf::render::layout::measure::{measure, MeasuredDocument};
    let _ = measure as fn(&dxpdf::model::Document, &skia_safe::FontMgr) -> MeasuredDocument;
    let _ = layout
        as fn(&MeasuredDocument, &skia_safe::FontMgr) -> Vec<dxpdf::render::layout::LayoutedPage>;
}

// ---------------------------------------------------------------------------
// render module — layout::measurer
// ---------------------------------------------------------------------------

#[test]
fn text_measurer_api() {
    use dxpdf::render::layout::TextMeasurer;
    let fm = skia_safe::FontMgr::new();
    let m = TextMeasurer::new(fm);
    let f = m.font("Helvetica", dxpdf::dimension::Pt::new(12.0), false, false);
    let _: dxpdf::dimension::Pt = f.measure_width("hello");
    let fm = f.metrics();
    let _: dxpdf::dimension::Pt = fm.line_height;
    let _: dxpdf::dimension::Pt = fm.ascent;
}

// ---------------------------------------------------------------------------
// render module — painter
// ---------------------------------------------------------------------------

#[test]
fn painter_function_signatures() {
    use dxpdf::render::layout::LayoutedPage;
    use dxpdf::render::painter::{render_to_pdf, render_to_pdf_with_font_mgr};

    let _: fn(&[LayoutedPage]) -> Result<Vec<u8>, dxpdf::Error> = render_to_pdf;
    let _: fn(&[LayoutedPage], &skia_safe::FontMgr) -> Result<Vec<u8>, dxpdf::Error> =
        render_to_pdf_with_font_mgr;
}
