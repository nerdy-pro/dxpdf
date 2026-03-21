use std::path::PathBuf;

use clap::Parser;

#[derive(Parser)]
#[command(name = "dxpdf", about = "Convert DOCX files to PDF")]
struct Cli {
    /// Input .docx file
    input: PathBuf,

    /// Output .pdf file (defaults to input path with .pdf extension)
    #[arg(short, long)]
    output: Option<PathBuf>,
}

fn main() -> Result<(), dxpdf::Error> {
    env_logger::Builder::from_env(env_logger::Env::default().default_filter_or("warn")).init();

    let cli = Cli::parse();

    let output = cli
        .output
        .unwrap_or_else(|| cli.input.with_extension("pdf"));

    let docx_bytes = std::fs::read(&cli.input)?;

    if std::env::var("DXPDF_BENCH").is_ok() {
        use std::time::Instant;
        let t0 = Instant::now();
        let document = dxpdf::parse::parse(&docx_bytes)?;
        let t1 = Instant::now();
        let font_mgr = skia_safe::FontMgr::new();
        dxpdf::render::fonts::preload_fonts(&font_mgr, &document.font_families());
        let t_fonts = Instant::now();
        let config = dxpdf::render::layout::LayoutConfig::default();
        let pages = dxpdf::render::layout::layout(&document, &config, &font_mgr);
        let t2 = Instant::now();
        let pdf_bytes = dxpdf::render::painter::render_to_pdf_with_font_mgr(&pages, &font_mgr)?;
        let t3 = Instant::now();
        std::fs::write(&output, &pdf_bytes)?;
        let mut texts = 0u32;
        let mut lines = 0u32;
        let mut rects = 0u32;
        let mut imgs = 0u32;
        let mut links = 0u32;
        for p in &pages {
            for cmd in &p.commands {
                match cmd {
                    dxpdf::render::layout::DrawCommand::Text { .. } => texts += 1,
                    dxpdf::render::layout::DrawCommand::Line { .. }
                    | dxpdf::render::layout::DrawCommand::Underline { .. } => lines += 1,
                    dxpdf::render::layout::DrawCommand::Rect { .. } => rects += 1,
                    dxpdf::render::layout::DrawCommand::Image { .. } => imgs += 1,
                    dxpdf::render::layout::DrawCommand::LinkAnnotation { .. } => links += 1,
                }
            }
        }
        eprintln!("Parse:  {:?}", t1 - t0);
        eprintln!("Fonts:  {:?}", t_fonts - t1);
        eprintln!("Layout: {:?}", t2 - t_fonts);
        eprintln!("Paint:  {:?}", t3 - t2);
        eprintln!("Total:  {:?}", t3 - t0);
        eprintln!(
            "Pages: {} | Text: {} | Lines: {} | Rects: {} | Images: {} | Links: {}",
            pages.len(),
            texts,
            lines,
            rects,
            imgs,
            links
        );
    } else {
        let pdf_bytes = dxpdf::convert(&docx_bytes)?;
        std::fs::write(&output, &pdf_bytes)?;
        eprintln!("Converted {} -> {}", cli.input.display(), output.display());
    }

    Ok(())
}
