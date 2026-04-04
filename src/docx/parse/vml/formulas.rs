//! VML formula parsing.

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::docx::error::Result;
use crate::docx::model::*;
use crate::docx::xml;

/// VML §14.1.2.6: parse `v:formulas` — list of `v:f` formula equations.
pub(super) fn parse_formulas(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<Vec<VmlFormula>> {
    let mut formulas = Vec::new();

    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) if xml::local_name(e.name().as_ref()) == b"f" => {
                if let Some(eqn) = xml::optional_attr(e, b"eqn")? {
                    if let Some(f) = parse_formula(&eqn) {
                        formulas.push(f);
                    } else {
                        log::warn!("vml-formula: failed to parse {:?}", eqn);
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"formulas" => break,
            Event::Eof => return Err(xml::unexpected_eof(b"formulas")),
            _ => {}
        }
    }

    Ok(formulas)
}

/// Parse a single VML formula equation string (e.g., "sum #0 0 10800").
fn parse_formula(eqn: &str) -> Option<VmlFormula> {
    let parts: Vec<&str> = eqn.split_whitespace().collect();
    if parts.is_empty() {
        return None;
    }

    let operation = match parts[0] {
        "val" => VmlFormulaOp::Val,
        "sum" => VmlFormulaOp::Sum,
        "prod" => VmlFormulaOp::Product,
        "mid" => VmlFormulaOp::Mid,
        "abs" => VmlFormulaOp::Abs,
        "min" => VmlFormulaOp::Min,
        "max" => VmlFormulaOp::Max,
        "if" => VmlFormulaOp::If,
        "sqrt" => VmlFormulaOp::Sqrt,
        "mod" => VmlFormulaOp::Mod,
        "sin" => VmlFormulaOp::Sin,
        "cos" => VmlFormulaOp::Cos,
        "tan" => VmlFormulaOp::Tan,
        "atan2" => VmlFormulaOp::Atan2,
        "sinatan2" => VmlFormulaOp::SinAtan2,
        "cosatan2" => VmlFormulaOp::CosAtan2,
        "sumangle" => VmlFormulaOp::SumAngle,
        "ellipse" => VmlFormulaOp::Ellipse,
        other => {
            log::warn!("vml-formula: unsupported operation {:?}", other);
            return None;
        }
    };

    let arg = |i: usize| -> VmlFormulaArg {
        parts
            .get(i)
            .and_then(|s| parse_formula_arg(s))
            .unwrap_or(VmlFormulaArg::Literal(0))
    };

    Some(VmlFormula {
        operation,
        args: [arg(1), arg(2), arg(3)],
    })
}

/// Parse a single VML formula argument.
fn parse_formula_arg(s: &str) -> Option<VmlFormulaArg> {
    if let Some(rest) = s.strip_prefix('#') {
        return rest.parse::<u32>().ok().map(VmlFormulaArg::AdjRef);
    }
    if let Some(rest) = s.strip_prefix('@') {
        return rest.parse::<u32>().ok().map(VmlFormulaArg::FormulaRef);
    }
    let guide = match s {
        "width" => Some(VmlGuide::Width),
        "height" => Some(VmlGuide::Height),
        "xcenter" => Some(VmlGuide::XCenter),
        "ycenter" => Some(VmlGuide::YCenter),
        "xrange" => Some(VmlGuide::XRange),
        "yrange" => Some(VmlGuide::YRange),
        "pixelWidth" => Some(VmlGuide::PixelWidth),
        "pixelHeight" => Some(VmlGuide::PixelHeight),
        "pixelLineWidth" => Some(VmlGuide::PixelLineWidth),
        "emuWidth" => Some(VmlGuide::EmuWidth),
        "emuHeight" => Some(VmlGuide::EmuHeight),
        "emuWidth2" => Some(VmlGuide::EmuWidth2),
        "emuHeight2" => Some(VmlGuide::EmuHeight2),
        _ => None,
    };
    if let Some(g) = guide {
        return Some(VmlFormulaArg::Guide(g));
    }
    s.parse::<i64>().ok().map(VmlFormulaArg::Literal)
}
