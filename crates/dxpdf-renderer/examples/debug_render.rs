fn main() {
    let docx_bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../dxpdf-docx/test-files/sample-docx-files-sample1.docx"
    )).unwrap();
    let doc = dxpdf_docx::parse(&docx_bytes).unwrap();
    let (resolved, pages) = dxpdf_renderer::resolve_and_layout(&doc);

    println!("Sections: {}", resolved.sections.len());
    println!("Pages: {}", pages.len());
    println!("Page size: {:?}", pages[0].page_size);

    for (pi, page) in pages.iter().enumerate() {
        let texts: Vec<_> = page.commands.iter().filter_map(|c| {
            if let dxpdf_renderer::layout::draw_command::DrawCommand::Text { position, text, .. } = c {
                Some((position, text))
            } else { None }
        }).collect();
        println!("\nPage {pi}: {} text commands", texts.len());
        for (pos, text) in texts.iter().take(10) {
            let preview: String = text.chars().take(50).collect();
            println!("  x={:.1} y={:.1} \"{}\"", pos.x.raw(), pos.y.raw(), preview);
        }
    }
}
