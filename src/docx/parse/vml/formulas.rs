//! VML formula parsing.

use crate::docx::model::*;

/// Parse a single VML formula equation string (e.g., "sum #0 0 10800").
pub(super) fn parse_formula(eqn: &str) -> Option<VmlFormula> {
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
