//! VML path command parsing.

use crate::docx::model::*;

/// Parse VML path commands from the `path` attribute string (§14.2.1.6).
pub(super) fn parse_path_commands(s: Option<String>) -> Vec<VmlPathCommand> {
    let s = match s {
        Some(s) => s,
        None => return Vec::new(),
    };

    let mut cmds = Vec::new();
    // Tokenize: split on commas and whitespace, but keep command letters as separate tokens.
    let mut tokens: Vec<&str> = Vec::new();
    let mut rest = s.as_str();
    while !rest.is_empty() {
        // Skip whitespace and commas.
        rest = rest.trim_start_matches(|c: char| c == ',' || c.is_ascii_whitespace());
        if rest.is_empty() {
            break;
        }
        // Check for multi-char commands first.
        let cmd_len = if rest.starts_with("wa")
            || rest.starts_with("wr")
            || rest.starts_with("at")
            || rest.starts_with("ar")
            || rest.starts_with("qx")
            || rest.starts_with("qy")
            || rest.starts_with("nf")
            || rest.starts_with("ns")
            || rest.starts_with("hа")
        // ha..hh variants
        {
            2
        } else if rest.starts_with(|c: char| c.is_ascii_alphabetic()) {
            1
        } else {
            0
        };

        if cmd_len > 0 {
            tokens.push(&rest[..cmd_len]);
            rest = &rest[cmd_len..];
        } else if rest.starts_with('@') {
            // §14.2.1.6: @n formula reference — consume @digits.
            let end = rest[1..]
                .find(|c: char| !c.is_ascii_digit())
                .map(|p| p + 1)
                .unwrap_or(rest.len());
            tokens.push(&rest[..end]);
            rest = &rest[end..];
        } else {
            // Numeric token: consume until delimiter.
            let end = rest
                .find(|c: char| {
                    c == ',' || c == '@' || c.is_ascii_whitespace() || c.is_ascii_alphabetic()
                })
                .unwrap_or(rest.len());
            if end > 0 {
                tokens.push(&rest[..end]);
                rest = &rest[end..];
            } else {
                rest = &rest[1..]; // skip unrecognized char
            }
        }
    }

    let mut i = 0;
    while i < tokens.len() {
        let tok = tokens[i];
        i += 1;
        match tok {
            "m" => {
                if let Some((x, y)) = take_2_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::MoveTo { x, y });
                }
            }
            "l" => {
                if let Some((x, y)) = take_2_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::LineTo { x, y });
                }
            }
            "c" => {
                if let Some((x1, y1, x2, y2, x, y)) = take_6_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::CurveTo {
                        x1,
                        y1,
                        x2,
                        y2,
                        x,
                        y,
                    });
                }
            }
            "r" => {
                if let Some((dx, dy)) = take_2_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::RLineTo { dx, dy });
                }
            }
            "v" => {
                if let Some((dx1, dy1, dx2, dy2, dx, dy)) = take_6_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::RCurveTo {
                        dx1,
                        dy1,
                        dx2,
                        dy2,
                        dx,
                        dy,
                    });
                }
            }
            "t" => {
                if let Some((dx, dy)) = take_2_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::RMoveTo { dx, dy });
                }
            }
            "x" => cmds.push(VmlPathCommand::Close),
            "e" => cmds.push(VmlPathCommand::End),
            "qx" => {
                if let Some((x, y)) = take_2_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::QuadrantX { x, y });
                }
            }
            "qy" => {
                if let Some((x, y)) = take_2_coord(&tokens, &mut i) {
                    cmds.push(VmlPathCommand::QuadrantY { x, y });
                }
            }
            "nf" => cmds.push(VmlPathCommand::NoFill),
            "ns" => cmds.push(VmlPathCommand::NoStroke),
            "wa" | "wr" | "at" | "ar" => {
                let kind = match tok {
                    "wa" => VmlArcKind::WA,
                    "wr" => VmlArcKind::WR,
                    "at" => VmlArcKind::AT,
                    _ => VmlArcKind::AR,
                };
                let args = (|| {
                    Some(VmlPathCommand::Arc {
                        kind,
                        bounding_x1: take_coord(&tokens, &mut i)?,
                        bounding_y1: take_coord(&tokens, &mut i)?,
                        bounding_x2: take_coord(&tokens, &mut i)?,
                        bounding_y2: take_coord(&tokens, &mut i)?,
                        start_x: take_coord(&tokens, &mut i)?,
                        start_y: take_coord(&tokens, &mut i)?,
                        end_x: take_coord(&tokens, &mut i)?,
                        end_y: take_coord(&tokens, &mut i)?,
                    })
                })();
                if let Some(cmd) = args {
                    cmds.push(cmd);
                }
            }
            _ => {
                // §14.2.1.6: bare coordinate in command position — implicit lineto.
                let x = if let Some(rest) = tok.strip_prefix('@') {
                    rest.parse::<u32>().ok().map(VmlPathCoord::FormulaRef)
                } else {
                    tok.parse::<i64>().ok().map(VmlPathCoord::Literal)
                };
                if let Some(x) = x {
                    if let Some(y) = take_coord(&tokens, &mut i) {
                        cmds.push(VmlPathCommand::LineTo { x, y });
                    }
                } else {
                    log::warn!("vml-path: unsupported command {:?}", tok);
                }
            }
        }
    }

    cmds
}

fn take_coord(tokens: &[&str], i: &mut usize) -> Option<VmlPathCoord> {
    if *i >= tokens.len() {
        return None;
    }
    let tok = tokens[*i];
    let coord = if let Some(rest) = tok.strip_prefix('@') {
        VmlPathCoord::FormulaRef(rest.parse::<u32>().ok()?)
    } else {
        VmlPathCoord::Literal(tok.parse::<i64>().ok()?)
    };
    *i += 1;
    Some(coord)
}

fn take_2_coord(tokens: &[&str], i: &mut usize) -> Option<(VmlPathCoord, VmlPathCoord)> {
    let a = take_coord(tokens, i)?;
    let b = take_coord(tokens, i)?;
    Some((a, b))
}

fn take_6_coord(
    tokens: &[&str],
    i: &mut usize,
) -> Option<(
    VmlPathCoord,
    VmlPathCoord,
    VmlPathCoord,
    VmlPathCoord,
    VmlPathCoord,
    VmlPathCoord,
)> {
    let a = take_coord(tokens, i)?;
    let b = take_coord(tokens, i)?;
    let c = take_coord(tokens, i)?;
    let d = take_coord(tokens, i)?;
    let e = take_coord(tokens, i)?;
    let f = take_coord(tokens, i)?;
    Some((a, b, c, d, e, f))
}

/// Parse a VML `adj` attribute — comma-separated integer adjustment values.
pub(super) fn parse_adj(s: Option<String>) -> Vec<i64> {
    match s {
        Some(s) => s
            .split(',')
            .filter_map(|v| v.trim().parse::<i64>().ok())
            .collect(),
        None => Vec::new(),
    }
}

/// Parse a VML Vector2D string ("x,y") into `VmlVector2D`.
pub(super) fn parse_vector2d(s: Option<String>) -> Option<VmlVector2D> {
    let s = s?;
    let (x_str, y_str) = s.split_once(',')?;
    let x = x_str.trim().parse::<i64>().ok()?;
    let y = y_str.trim().parse::<i64>().ok()?;
    Some(VmlVector2D { x, y })
}
