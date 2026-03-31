/// Convert DOCX files to PDF using the renderer.
///
/// With no arguments: converts all files in test-files/ and test-cases/.
/// With arguments: converts the specified files.
fn main() {
    env_logger::init();
    let out_dir = concat!(env!("CARGO_MANIFEST_DIR"), "/target/rendered");
    std::fs::create_dir_all(out_dir).unwrap();

    let font_mgr = skia_safe::FontMgr::new();

    let args: Vec<String> = std::env::args().skip(1).collect();
    let paths: Vec<std::path::PathBuf> = if args.is_empty() {
        // Default: scan both directories.
        let dirs = [
            concat!(env!("CARGO_MANIFEST_DIR"), "/test-files"),
            concat!(env!("CARGO_MANIFEST_DIR"), "/test-cases"),
        ];
        dirs.iter()
            .filter_map(|d| std::fs::read_dir(d).ok())
            .flat_map(|rd| rd.filter_map(|e| e.ok()).map(|e| e.path()))
            .filter(|p| p.extension().is_some_and(|e| e == "docx"))
            .collect()
    } else {
        args.iter().map(std::path::PathBuf::from).collect()
    };

    for path in &paths {
        let filename = path.file_stem().unwrap().to_str().unwrap();
        let docx_bytes = std::fs::read(path).unwrap();

        let t0 = std::time::Instant::now();
        let doc = match dxpdf::docx::parse(&docx_bytes) {
            Ok(d) => d,
            Err(e) => {
                eprintln!("{filename}: parse error: {e}");
                continue;
            }
        };
        let parse_ms = t0.elapsed().as_millis();

        let t1 = std::time::Instant::now();
        let pdf_bytes = match dxpdf::render::render_with_font_mgr(&doc, &font_mgr) {
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
