fn main() {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"), "/../dxpdf-docx/test-files/sample-docx-files-sample1.docx"
    )).unwrap();
    let doc = dxpdf_docx::parse(&bytes).unwrap();

    // Find the paragraph with colored text ("red, green")
    for block in &doc.body {
        if let dxpdf_docx_model::model::Block::Paragraph(p) = block {
            let full_text: String = p.content.iter().filter_map(|i| {
                if let dxpdf_docx_model::model::Inline::TextRun(tr) = i {
                    Some(tr.text.as_str())
                } else { None }
            }).collect();
            if full_text.contains("red") && full_text.contains("green") && full_text.contains("blue") {
                println!("Found paragraph: {:?}", &full_text[..full_text.len().min(80)]);
                for inline in &p.content {
                    if let dxpdf_docx_model::model::Inline::TextRun(tr) = inline {
                        let color = tr.properties.color;
                        println!("  Run: {:?} color={:?}", tr.text, color);
                    }
                }
                break;
            }
        }
    }
}
