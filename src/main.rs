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
    let pdf_bytes = dxpdf::convert(&docx_bytes)?;
    std::fs::write(&output, &pdf_bytes)?;
    eprintln!("Converted {} -> {}", cli.input.display(), output.display());

    Ok(())
}
