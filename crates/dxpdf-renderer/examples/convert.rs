/// Convert all sample DOCX files to PDF using the new renderer.
fn main() {
    env_logger::init();
    let test_dir = concat!(env!("CARGO_MANIFEST_DIR"), "/../dxpdf-docx/test-files");
    let out_dir = concat!(env!("CARGO_MANIFEST_DIR"), "/../../target/rendered");
    std::fs::create_dir_all(out_dir).unwrap();

    let font_mgr = skia_safe::FontMgr::new();

    for entry in std::fs::read_dir(test_dir).unwrap() {
        let entry = entry.unwrap();
        let path = entry.path();
        if path.extension().is_none_or(|e| e != "docx") {
            continue;
        }

        let filename = path.file_stem().unwrap().to_str().unwrap();
        let docx_bytes = std::fs::read(&path).unwrap();

        let t0 = std::time::Instant::now();
        let doc = match dxpdf_docx::parse(&docx_bytes) {
            Ok(d) => d,
            Err(e) => {
                eprintln!("{filename}: parse error: {e}");
                continue;
            }
        };
        let parse_ms = t0.elapsed().as_millis();

        let t1 = std::time::Instant::now();
        let pdf_bytes = match dxpdf_renderer::render_with_font_mgr(&doc, &font_mgr) {
            Ok(b) => b,
            Err(e) => {
                eprintln!("{filename}: render error: {e}");
                continue;
            }
        };
        let render_ms = t1.elapsed().as_millis();

        let out_path = format!("{out_dir}/{filename}.pdf");
        std::fs::write(&out_path, &pdf_bytes).unwrap();

        println!(
            "{filename}.docx -> {:.1}KB PDF ({parse_ms}ms parse, {render_ms}ms render)",
            pdf_bytes.len() as f64 / 1024.0,
        );
    }
}
