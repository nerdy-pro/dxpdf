use std::path::PathBuf;
use std::process::ExitCode;

use clap::Parser;
use dxpdf::{RenderOptions, DEFAULT_IMAGE_DPI};

#[derive(Parser)]
#[command(
    name = "dxpdf",
    about = "DOCX files to PDF converter",
    // Without this, clap defines no `--version` at all. A packaged CLI is
    // expected to answer it — it is the first thing a user runs after `apt
    // install`, the first thing a bug report quotes, and what the Debian
    // install test in CI uses as its smoke check.
    version,
    allow_negative_numbers = true
)]
struct Cli {
    /// Input .docx file
    input: PathBuf,

    /// Output .pdf file (defaults to input path with .pdf extension)
    #[arg(short, long)]
    output: Option<PathBuf>,

    /// Target resolution (pixels per inch) for embedded raster images.
    /// Higher values yield crisper images and larger PDFs.
    #[arg(long, default_value_t = DEFAULT_IMAGE_DPI, value_parser = parse_image_dpi)]
    image_dpi: f32,
}

/// Lowest / highest `--image-dpi` the CLI accepts. Bounds only catch typos; the
/// library floor lives in `RenderOptions`.
const MIN_CLI_IMAGE_DPI: f32 = 1.0;
const MAX_CLI_IMAGE_DPI: f32 = 2400.0;

/// Parse and range-check `--image-dpi`, rejecting non-numeric, non-finite, and
/// out-of-range values instead of silently clamping them — so a typo like
/// `--image-dpi -300` surfaces as a CLI error rather than unreadable output.
///
/// This deliberately differs from the library and Python APIs, which *clamp*
/// the same input to `MIN_IMAGE_DPI` (see `RenderOptions::with_image_dpi`). A
/// programmatic caller passing an out-of-range value has usually computed it and
/// wants a usable render; a human typing one has usually made a typo and wants
/// to be told. The divergence is stated in both places so neither reads as an
/// oversight.
fn parse_image_dpi(s: &str) -> Result<f32, String> {
    let dpi: f32 = s
        .parse()
        .map_err(|_| format!("`{s}` is not a valid number"))?;
    if !dpi.is_finite() || !(MIN_CLI_IMAGE_DPI..=MAX_CLI_IMAGE_DPI).contains(&dpi) {
        return Err(format!(
            "must be between {MIN_CLI_IMAGE_DPI} and {MAX_CLI_IMAGE_DPI} DPI (got `{s}`)"
        ));
    }
    Ok(dpi)
}

/// Returning `Result` from `main` would report failures through `Termination`,
/// which formats with **`Debug`** — so a missing input printed
/// `Error: Io(Os { code: 2, kind: NotFound, .. })` and every `Display` message
/// the crate defines went unseen. Printing it here is the only way the CLI gets
/// the same message the library and the Python bindings already produce.
fn main() -> ExitCode {
    env_logger::Builder::from_env(env_logger::Env::default().default_filter_or("warn")).init();

    match run() {
        Ok(()) => ExitCode::SUCCESS,
        Err(err) => {
            eprintln!("error: {err}");
            ExitCode::FAILURE
        }
    }
}

fn run() -> Result<(), dxpdf::Error> {
    let cli = Cli::parse();

    let output = cli
        .output
        .unwrap_or_else(|| cli.input.with_extension("pdf"));

    let options = RenderOptions::default().with_image_dpi(cli.image_dpi);

    let docx_bytes = options
        .resource_limits()
        .read_input(std::fs::File::open(&cli.input)?)?;
    let pdf_bytes = dxpdf::convert_with_options(&docx_bytes, &options)?;
    std::fs::write(&output, &pdf_bytes)?;
    eprintln!("Converted {} -> {}", cli.input.display(), output.display());

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn image_dpi_accepts_in_range_and_the_default() {
        assert_eq!(parse_image_dpi("220").unwrap(), 220.0);
        assert_eq!(parse_image_dpi("1").unwrap(), MIN_CLI_IMAGE_DPI);
        assert_eq!(parse_image_dpi("2400").unwrap(), MAX_CLI_IMAGE_DPI);
        // `default_value_t` renders the default and re-parses it — must round-trip.
        assert_eq!(
            parse_image_dpi(&DEFAULT_IMAGE_DPI.to_string()).unwrap(),
            DEFAULT_IMAGE_DPI
        );
    }

    #[test]
    fn image_dpi_rejects_out_of_range_and_non_numeric() {
        for bad in ["0", "-5", "-300", "2401", "nan", "inf", "-inf", "abc", ""] {
            assert!(
                parse_image_dpi(bad).is_err(),
                "`{bad}` should be rejected, not silently clamped"
            );
        }
    }
}
