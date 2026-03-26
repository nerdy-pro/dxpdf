fn main() {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"), "/../dxpdf-docx/test-files/sample-docx-files-sample1.docx"
    )).unwrap();
    let doc = dxpdf_docx::parse(&bytes).unwrap();

    let mut table_idx = 0;
    for (i, block) in doc.body.iter().enumerate() {
        if let dxpdf_docx_model::model::Block::Table(t) = block {
            let rows = t.rows.len();
            let cols = t.rows.first().map(|r| r.cells.len()).unwrap_or(0);
            let grid: Vec<_> = t.grid.iter().map(|g| g.width.raw()).collect();
            println!("Table #{table_idx} at block [{i}]: {rows}x{cols} grid={grid:?}");
            for (ri, row) in t.rows.iter().enumerate() {
                for (ci, cell) in row.cells.iter().enumerate() {
                    let text: String = cell.content.iter().filter_map(|b| {
                        if let dxpdf_docx_model::model::Block::Paragraph(p) = b {
                            let t: String = p.content.iter().filter_map(|i| {
                                if let dxpdf_docx_model::model::Inline::TextRun(tr) = i {
                                    Some(tr.text.as_str())
                                } else { None }
                            }).collect();
                            Some(t)
                        } else { None }
                    }).collect::<Vec<_>>().join("|");
                    if !text.is_empty() {
                        let preview: String = text.chars().take(30).collect();
                        println!("  [{ri},{ci}] \"{preview}\"");
                    }
                }
            }
            table_idx += 1;
        }
    }
    println!("Total tables: {table_idx}");
}
