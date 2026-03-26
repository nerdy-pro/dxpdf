fn main() {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"), "/../dxpdf-docx/test-files/sample-docx-files-sample1.docx"
    )).unwrap();
    let doc = dxpdf_docx::parse(&bytes).unwrap();

    for (i, block) in doc.body.iter().enumerate().take(10) {
        match block {
            dxpdf_docx_model::model::Block::Paragraph(p) => {
                let text: String = p.content.iter().filter_map(|i| {
                    if let dxpdf_docx_model::model::Inline::TextRun(tr) = i {
                        Some(tr.text.as_str())
                    } else { None }
                }).collect();
                let preview: String = text.chars().take(60).collect();
                println!("[{i}] Paragraph: \"{preview}\"");
            }
            dxpdf_docx_model::model::Block::Table(t) => {
                let rows = t.rows.len();
                let cols = t.rows.first().map(|r| r.cells.len()).unwrap_or(0);
                let grid: Vec<_> = t.grid.iter().map(|g| g.width.raw()).collect();
                // Check if any cell has content
                let has_content = t.rows.iter().any(|r| r.cells.iter().any(|c| {
                    c.content.iter().any(|b| matches!(b, dxpdf_docx_model::model::Block::Paragraph(p) if !p.content.is_empty()))
                }));
                println!("[{i}] Table: {rows}x{cols} grid={grid:?} has_content={has_content}");
            }
            dxpdf_docx_model::model::Block::SectionBreak(_) => {
                println!("[{i}] SectionBreak");
            }
        }
    }
}
