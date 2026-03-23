use super::*;
use crate::dimension::Twips;

fn wrap_body(content: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:body>{content}</w:body>
            </w:document>"#
    )
}

#[test]
fn parse_empty_document() {
    let xml = wrap_body("");
    let doc = parse_document_xml(&xml).unwrap();
    assert!(doc.sections.iter().all(|s| s.blocks.is_empty()));
}

#[test]
fn parse_single_paragraph() {
    let xml = wrap_body(r#"<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>"#);
    let doc = parse_document_xml(&xml).unwrap();
    assert_eq!(doc.sections[0].blocks.len(), 1);
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    assert_eq!(p.runs.len(), 1);
    let Inline::TextRun(tr) = &p.runs[0] else {
        panic!()
    };
    assert_eq!(tr.text, "Hello World");
}

#[test]
fn parse_bold_italic() {
    let xml =
        wrap_body(r#"<w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Bold Italic</w:t></w:r></w:p>"#);
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    let Inline::TextRun(tr) = &p.runs[0] else {
        panic!()
    };
    assert!(tr.properties.bold);
    assert!(tr.properties.italic);
    assert!(!tr.properties.underline);
}

#[test]
fn parse_font_size_and_color() {
    let xml = wrap_body(
        r#"<w:p><w:r><w:rPr><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr><w:t>Red 14pt</w:t></w:r></w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    let Inline::TextRun(tr) = &p.runs[0] else {
        panic!()
    };
    assert_eq!(
        tr.properties.font_size,
        Some(crate::dimension::HalfPoints::new(28))
    );
    assert_eq!(f32::from(tr.properties.font_size.unwrap()), 14.0);
    assert_eq!(tr.properties.color, Some(Color { r: 255, g: 0, b: 0 }));
}

#[test]
fn parse_alignment() {
    let xml = wrap_body(
        r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Centered</w:t></w:r></w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    assert_eq!(p.properties.alignment, Some(Alignment::Center));
}

#[test]
fn parse_spacing_and_indentation() {
    let xml = wrap_body(
        r#"<w:p><w:pPr>
                <w:spacing w:before="240" w:after="120" w:line="360"/>
                <w:ind w:left="720" w:firstLine="360"/>
            </w:pPr><w:r><w:t>Indented</w:t></w:r></w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    let spacing = p.properties.spacing.unwrap();
    assert_eq!(spacing.before, Some(Twips::new(240)));
    assert_eq!(spacing.after, Some(Twips::new(120)));
    assert_eq!(spacing.line, Some(Twips::new(360)));
    let indent = p.properties.indentation.unwrap();
    assert_eq!(indent.left, Some(Twips::new(720)));
    assert_eq!(indent.first_line, Some(Twips::new(360)));
}

#[test]
fn parse_table() {
    let xml = wrap_body(
        r#"<w:tbl>
                <w:tr>
                    <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
                </w:tr>
                <w:tr>
                    <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
                </w:tr>
            </w:tbl>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    assert_eq!(doc.sections[0].blocks.len(), 1);
    let Block::Table(table) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    assert_eq!(table.rows.len(), 2);
    assert_eq!(table.rows[0].cells.len(), 2);
    let Block::Paragraph(p) = &table.rows[0].cells[0].blocks[0] else {
        panic!()
    };
    let Inline::TextRun(tr) = &p.runs[0] else {
        panic!()
    };
    assert_eq!(tr.text, "A1");
}

#[test]
fn parse_multiple_runs() {
    let xml = wrap_body(
        r#"<w:p>
                <w:r><w:t>Hello </w:t></w:r>
                <w:r><w:rPr><w:b/></w:rPr><w:t>World</w:t></w:r>
            </w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    assert_eq!(p.runs.len(), 2);
    let Inline::TextRun(tr1) = &p.runs[0] else {
        panic!()
    };
    let Inline::TextRun(tr2) = &p.runs[1] else {
        panic!()
    };
    assert_eq!(tr1.text, "Hello ");
    assert!(!tr1.properties.bold);
    assert_eq!(tr2.text, "World");
    assert!(tr2.properties.bold);
}

#[test]
fn parse_font_family() {
    let xml = wrap_body(
        r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>Arial text</w:t></w:r></w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    let Inline::TextRun(tr) = &p.runs[0] else {
        panic!()
    };
    assert_eq!(tr.properties.font_family.as_deref(), Some("Arial"));
}

#[test]
fn parse_inline_image() {
    let xml = wrap_body(
        r#"<w:p><w:r><w:drawing>
                <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                    <wp:extent cx="914400" cy="457200"/>
                    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                        <a:graphicData>
                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:blipFill>
                                    <a:blip r:embed="rId5" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                </pic:blipFill>
                            </pic:pic>
                        </a:graphicData>
                    </a:graphic>
                </wp:inline>
            </w:drawing></w:r></w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    let Inline::Image(img) = &p.runs[0] else {
        panic!()
    };
    assert_eq!(img.rel_id, RelId::from("rId5"));
    assert!((f32::from(img.size.width) - 72.0).abs() < 0.1);
    assert!((f32::from(img.size.height) - 36.0).abs() < 0.1);
}

#[test]
fn parse_image_with_text() {
    let xml = wrap_body(
        r#"<w:p>
                <w:r><w:t>Before </w:t></w:r>
                <w:r><w:drawing>
                    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                        <wp:extent cx="914400" cy="914400"/>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                            <a:graphicData>
                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:blipFill>
                                        <a:blip r:embed="rId3" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                    </pic:blipFill>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing></w:r>
                <w:r><w:t> After</w:t></w:r>
            </w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    assert_eq!(p.runs.len(), 3);
    assert!(matches!(p.runs[0], Inline::TextRun(_)));
    assert!(matches!(p.runs[1], Inline::Image(_)));
    assert!(matches!(p.runs[2], Inline::TextRun(_)));
}

#[test]
fn parse_section_properties() {
    let xml = wrap_body(
        r#"<w:p><w:pPr><w:sectPr>
                    <w:pgSz w:w="11906" w:h="16838"/>
                    <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>
                </w:sectPr></w:pPr>
                <w:r><w:t>Section 1</w:t></w:r>
            </w:p>
            <w:p><w:r><w:t>Section 2</w:t></w:r></w:p>
            <w:sectPr>
                <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
                <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
            </w:sectPr>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    assert_eq!(doc.sections.len(), 2);

    // Section 1: contains "Section 1" paragraph, with the inline sectPr properties
    assert_eq!(doc.sections[0].blocks.len(), 1);
    let Block::Paragraph(p1) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    let Inline::TextRun(tr1) = &p1.runs[0] else {
        panic!()
    };
    assert_eq!(tr1.text, "Section 1");
    let sect1 = &doc.sections[0].properties;
    assert_eq!(sect1.page_size.unwrap().width, Twips::new(11906));
    assert_eq!(sect1.page_size.unwrap().height, Twips::new(16838));
    assert_eq!(sect1.page_margins.unwrap().top, Twips::new(720));

    // Section 2: contains "Section 2" paragraph, with the final body sectPr properties
    assert_eq!(doc.sections[1].blocks.len(), 1);
    let Block::Paragraph(p2) = &doc.sections[1].blocks[0] else {
        panic!()
    };
    let Inline::TextRun(tr2) = &p2.runs[0] else {
        panic!()
    };
    assert_eq!(tr2.text, "Section 2");
    let sect2 = &doc.sections[1].properties;
    assert_eq!(sect2.page_size.unwrap().width, Twips::new(16838));
    assert_eq!(sect2.page_margins.unwrap().top, Twips::new(1440));
}

#[test]
fn parse_tab_stops() {
    let xml = wrap_body(
        r#"<w:p><w:pPr><w:tabs>
                    <w:tab w:val="left" w:pos="2880"/>
                    <w:tab w:val="center" w:pos="4320"/>
                    <w:tab w:val="right" w:pos="9360"/>
                </w:tabs></w:pPr>
                <w:r><w:t>Col1</w:t></w:r>
            </w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    assert_eq!(p.properties.tab_stops.len(), 3);
    assert_eq!(p.properties.tab_stops[0].position, Twips::new(2880));
    assert_eq!(p.properties.tab_stops[0].stop_type, TabStopType::Left);
    assert_eq!(p.properties.tab_stops[1].position, Twips::new(4320));
    assert_eq!(p.properties.tab_stops[1].stop_type, TabStopType::Center);
    assert_eq!(p.properties.tab_stops[2].position, Twips::new(9360));
    assert_eq!(p.properties.tab_stops[2].stop_type, TabStopType::Right);
}

#[test]
fn parse_floating_image_with_pct_pos_offset() {
    let xml = wrap_body(
        r#"<w:p><w:r><w:drawing>
                <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                           xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing">
                    <wp:extent cx="914400" cy="457200"/>
                    <wp:positionH relativeFrom="page">
                        <wp14:pctPosHOffset>5000</wp14:pctPosHOffset>
                    </wp:positionH>
                    <wp:positionV relativeFrom="page">
                        <wp14:pctPosVOffset>3000</wp14:pctPosVOffset>
                    </wp:positionV>
                    <wp:wrapNone/>
                    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                        <a:graphicData>
                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:blipFill>
                                    <a:blip r:embed="rId7" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                </pic:blipFill>
                            </pic:pic>
                        </a:graphicData>
                    </a:graphic>
                </wp:anchor>
            </w:drawing></w:r></w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    assert_eq!(p.floats.len(), 1);
    let float = &p.floats[0];
    assert_eq!(float.rel_id, RelId::from("rId7"));
    // pctPosHOffset = 5000 → 5% of page width
    assert_eq!(float.pct_pos_h, Some(5000));
    // pctPosVOffset = 3000 → 3% of page height
    assert_eq!(float.pct_pos_v, Some(3000));
}

#[test]
fn parse_floating_image_without_pct_pos_has_none() {
    let xml = wrap_body(
        r#"<w:p><w:r><w:drawing>
                <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                    <wp:extent cx="914400" cy="457200"/>
                    <wp:positionH relativeFrom="margin">
                        <wp:posOffset>320675</wp:posOffset>
                    </wp:positionH>
                    <wp:positionV relativeFrom="margin">
                        <wp:posOffset>114300</wp:posOffset>
                    </wp:positionV>
                    <wp:wrapNone/>
                    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                        <a:graphicData>
                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:blipFill>
                                    <a:blip r:embed="rId8" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                                </pic:blipFill>
                            </pic:pic>
                        </a:graphicData>
                    </a:graphic>
                </wp:anchor>
            </w:drawing></w:r></w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    let float = &p.floats[0];
    assert_eq!(float.pct_pos_h, None);
    assert_eq!(float.pct_pos_v, None);
    // Should still have absolute offset
    assert!(
        (f32::from(float.offset.x) - 25.25).abs() < 0.5,
        "offset_x={}",
        float.offset.x
    );
}

#[test]
fn parse_hyperlink_resolves_url() {
    let xml = wrap_body(
        r#"<w:p>
                <w:hyperlink r:id="rId6" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                    <w:r><w:rPr><w:color w:val="0563C1"/><w:u/></w:rPr>
                        <w:t>Click here</w:t>
                    </w:r>
                </w:hyperlink>
            </w:p>"#,
    );
    let mut rels = std::collections::HashMap::new();
    rels.insert("rId6".to_string(), "https://example.com".to_string());
    let doc = parse_document_xml_with_rels(&xml, &rels).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    assert_eq!(p.runs.len(), 1);
    let Inline::TextRun(tr) = &p.runs[0] else {
        panic!()
    };
    assert_eq!(tr.text, "Click here");
    assert_eq!(tr.hyperlink_url.as_deref(), Some("https://example.com"));
    assert!(tr.properties.underline);
}

#[test]
fn parse_vert_align_superscript() {
    let xml = wrap_body(
        r#"<w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r></w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    let Inline::TextRun(tr) = &p.runs[0] else {
        panic!()
    };
    assert_eq!(tr.properties.vert_align, Some(VertAlign::Superscript));
}

#[test]
fn parse_paragraph_bottom_border() {
    let xml = wrap_body(
        r#"<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr></w:pPr>
                <w:r><w:t>Bordered</w:t></w:r></w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    let borders = p
        .properties
        .paragraph_borders
        .as_ref()
        .expect("should have paragraph borders");
    assert!(borders.bottom.is_some(), "should have bottom border");
    assert!(borders.top.is_none(), "should not have top border");
}

#[test]
fn parse_vert_align_subscript() {
    let xml = wrap_body(
        r#"<w:p><w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>2</w:t></w:r></w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    let Inline::TextRun(tr) = &p.runs[0] else {
        panic!()
    };
    assert_eq!(tr.properties.vert_align, Some(VertAlign::Subscript));
}

#[test]
fn parse_page_field_code() {
    let xml = wrap_body(
        r#"<w:p>
                <w:r><w:t>Page </w:t></w:r>
                <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
                <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                <w:r><w:t>3</w:t></w:r>
                <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    // Should have "Page " text and a PAGE field
    assert!(
        p.runs
            .iter()
            .any(|r| matches!(r, Inline::Field(fc) if fc.field_type == FieldType::Page)),
        "Should contain a PAGE field, got: {:?}",
        p.runs
    );
}

#[test]
fn parse_numpages_field_code() {
    let xml = wrap_body(
        r#"<w:p>
                <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                <w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
                <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                <w:r><w:t>5</w:t></w:r>
                <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    assert!(
        p.runs
            .iter()
            .any(|r| matches!(r, Inline::Field(fc) if fc.field_type == FieldType::NumPages)),
        "Should contain a NUMPAGES field, got: {:?}",
        p.runs
    );
}

#[test]
fn parse_unknown_field_uses_cached_value() {
    let xml = wrap_body(
        r#"<w:p>
                <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                <w:r><w:instrText> MERGEFIELD Name </w:instrText></w:r>
                <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                <w:r><w:t>John</w:t></w:r>
                <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    // Unknown fields should render the cached value as plain text
    assert!(
        p.runs
            .iter()
            .any(|r| matches!(r, Inline::TextRun(tr) if tr.text == "John")),
        "Unknown field should use cached value 'John', got: {:?}",
        p.runs
    );
}

#[test]
fn parse_hyperlink_without_rel_id_keeps_text() {
    let xml = wrap_body(
        r#"<w:p>
                <w:hyperlink>
                    <w:r><w:t>No link</w:t></w:r>
                </w:hyperlink>
            </w:p>"#,
    );
    let doc = parse_document_xml(&xml).unwrap();
    let Block::Paragraph(p) = &doc.sections[0].blocks[0] else {
        panic!()
    };
    assert_eq!(p.runs.len(), 1);
    let Inline::TextRun(tr) = &p.runs[0] else {
        panic!()
    };
    assert_eq!(tr.text, "No link");
    assert_eq!(tr.hyperlink_url, None);
}
