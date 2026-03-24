use crate::ast::{ComparisonOp, FieldInstruction};
use crate::context::FieldContext;
use crate::format;

/// The result of evaluating a field instruction.
#[derive(Clone, Debug, PartialEq)]
pub enum FieldValue {
    /// A text string result.
    Text(String),
    /// A numeric result.
    Number(f64),
    /// Field cannot be evaluated with the current context.
    Unevaluable,
}

/// Evaluate a parsed field instruction against a context.
pub fn evaluate(instr: &FieldInstruction, ctx: &FieldContext) -> FieldValue {
    match instr {
        FieldInstruction::Page { switches } => {
            match ctx.page_number {
                Some(n) => format_result(n.to_string(), switches),
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::NumPages { switches } => {
            match ctx.total_pages {
                Some(n) => format_result(n.to_string(), switches),
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::Section { switches } => {
            match ctx.section_number {
                Some(n) => format_result(n.to_string(), switches),
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::SectionPages { switches } => {
            match ctx.section_pages {
                Some(n) => format_result(n.to_string(), switches),
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::Date { format: _, switches } => {
            match &ctx.date {
                Some(date) => {
                    let formatted = if let Some(pattern) = &switches.date_format {
                        format::format_date(date, pattern)
                    } else {
                        // Default: M/d/yyyy
                        format::format_date(date, "M/d/yyyy")
                    };
                    format_result(formatted, switches)
                }
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::Time { format: _, switches } => {
            match &ctx.time {
                Some(time) => {
                    let formatted = if let Some(pattern) = &switches.date_format {
                        format::format_time(time, pattern)
                    } else {
                        format::format_time(time, "h:mm AM/PM")
                    };
                    format_result(formatted, switches)
                }
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::Author { format: _, switches } => {
            match ctx.document_properties.get("author") {
                Some(author) => format_result(author.clone(), switches),
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::Title { switches } => {
            match ctx.document_properties.get("title") {
                Some(title) => format_result(title.clone(), switches),
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::Subject { switches } => {
            match ctx.document_properties.get("subject") {
                Some(subject) => format_result(subject.clone(), switches),
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::FileName { path, switches } => {
            let value = if *path {
                ctx.file_path.as_ref()
            } else {
                ctx.file_name.as_ref()
            };
            match value {
                Some(name) => format_result(name.clone(), switches),
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::Ref {
            bookmark,
            footnote_ref,
            paragraph_number,
            hyperlink: _,
            switches,
        } => {
            match ctx.bookmarks.get(bookmark) {
                Some(loc) => {
                    let text = if *footnote_ref {
                        loc.note_ref.as_deref().unwrap_or("").to_string()
                    } else if *paragraph_number {
                        loc.paragraph_number.as_deref().unwrap_or("").to_string()
                    } else {
                        loc.text.as_deref().unwrap_or("").to_string()
                    };
                    format_result(text, switches)
                }
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::PageRef {
            bookmark,
            hyperlink: _,
            switches,
        } => {
            match ctx.bookmarks.get(bookmark) {
                Some(loc) => match loc.page {
                    Some(page) => format_result(page.to_string(), switches),
                    None => FieldValue::Unevaluable,
                },
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::NoteRef {
            bookmark,
            superscript: _,
            hyperlink: _,
            switches,
        } => {
            match ctx.bookmarks.get(bookmark) {
                Some(loc) => match &loc.note_ref {
                    Some(nr) => format_result(nr.clone(), switches),
                    None => FieldValue::Unevaluable,
                },
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::Hyperlink { target, .. } => {
            // Hyperlinks produce their target as text — the display text
            // is typically in the field result, not the instruction.
            FieldValue::Text(target.clone())
        }

        FieldInstruction::Toc { .. } => {
            // TOC requires full document analysis — unevaluable at field level.
            FieldValue::Unevaluable
        }

        FieldInstruction::Seq {
            identifier,
            reset_to,
            repeat,
            next: _,
            switches,
        } => {
            // SEQ evaluation is stateful. We use the snapshot from context.
            if let Some(reset) = reset_to {
                if let Ok(n) = reset.parse::<u32>() {
                    return format_result(n.to_string(), switches);
                }
            }
            if *repeat {
                let n = ctx.sequences.get(identifier.as_str()).copied().unwrap_or(0);
                return format_result(n.to_string(), switches);
            }
            // Default: next value (current + 1)
            let n = ctx.sequences.get(identifier.as_str()).copied().unwrap_or(0) + 1;
            format_result(n.to_string(), switches)
        }

        FieldInstruction::If {
            left,
            operator,
            right,
            then_text,
            else_text,
            switches,
        } => {
            let result = if evaluate_comparison(left, *operator, right) {
                then_text.clone()
            } else {
                else_text.clone()
            };
            format_result(result, switches)
        }

        FieldInstruction::MergeField {
            name,
            text_before,
            text_after,
            switches,
        } => {
            match ctx.merge_data.as_ref().and_then(|d| d.get(name.as_str())) {
                Some(value) if !value.is_empty() => {
                    let mut result = String::new();
                    if let Some(before) = text_before {
                        result.push_str(before);
                    }
                    result.push_str(value);
                    if let Some(after) = text_after {
                        result.push_str(after);
                    }
                    format_result(result, switches)
                }
                _ => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::DocProperty { name, switches } => {
            // Try case-insensitive lookup
            let key = name.to_ascii_lowercase();
            let value = ctx
                .document_properties
                .iter()
                .find(|(k, _)| k.to_ascii_lowercase() == key)
                .map(|(_, v)| v.clone());
            match value {
                Some(v) => format_result(v, switches),
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::Symbol {
            char_code, font: _, switches,
        } => {
            match char::from_u32(*char_code) {
                Some(ch) => format_result(ch.to_string(), switches),
                None => FieldValue::Unevaluable,
            }
        }

        FieldInstruction::IncludePicture { .. } => {
            // Picture inclusion requires image loading — unevaluable at field level.
            FieldValue::Unevaluable
        }

        FieldInstruction::Eq { equation, switches } => {
            format_result(equation.clone(), switches)
        }

        FieldInstruction::Unknown { .. } => FieldValue::Unevaluable,
    }
}

/// Apply formatting switches to a text result.
fn format_result(value: String, switches: &crate::ast::CommonSwitches) -> FieldValue {
    let mut text = value;

    // Apply numeric format
    if let Some(ref num_fmt) = switches.numeric_format {
        if let Ok(n) = text.parse::<f64>() {
            text = format::format_number(n, num_fmt);
        }
    }

    // Apply general format
    if let Some(ref fmt) = switches.format {
        text = format::apply_general_format(&text, fmt);
    }

    FieldValue::Text(text)
}

/// Evaluate a comparison for IF fields.
fn evaluate_comparison(left: &str, op: ComparisonOp, right: &str) -> bool {
    // Try numeric comparison first
    if let (Ok(l), Ok(r)) = (left.parse::<f64>(), right.parse::<f64>()) {
        return match op {
            ComparisonOp::Equal => (l - r).abs() < f64::EPSILON,
            ComparisonOp::NotEqual => (l - r).abs() >= f64::EPSILON,
            ComparisonOp::LessThan => l < r,
            ComparisonOp::LessEqual => l <= r,
            ComparisonOp::GreaterThan => l > r,
            ComparisonOp::GreaterEqual => l >= r,
        };
    }

    // Fall back to string comparison (case-insensitive per OOXML spec)
    let l = left.to_ascii_lowercase();
    let r = right.to_ascii_lowercase();
    match op {
        ComparisonOp::Equal => l == r,
        ComparisonOp::NotEqual => l != r,
        ComparisonOp::LessThan => l < r,
        ComparisonOp::LessEqual => l <= r,
        ComparisonOp::GreaterThan => l > r,
        ComparisonOp::GreaterEqual => l >= r,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ast::CommonSwitches;
    use crate::context::{BookmarkLocation, Date, Time};
    use std::collections::HashMap;

    fn sw() -> CommonSwitches {
        CommonSwitches::default()
    }

    fn ctx_with<F: FnOnce(&mut FieldContext)>(f: F) -> FieldContext {
        let mut ctx = FieldContext::default();
        f(&mut ctx);
        ctx
    }

    #[test]
    fn eval_page() {
        let ctx = ctx_with(|c| c.page_number = Some(7));
        let instr = FieldInstruction::Page { switches: sw() };
        assert_eq!(evaluate(&instr, &ctx), FieldValue::Text("7".into()));
    }

    #[test]
    fn eval_page_missing() {
        let ctx = FieldContext::default();
        let instr = FieldInstruction::Page { switches: sw() };
        assert_eq!(evaluate(&instr, &ctx), FieldValue::Unevaluable);
    }

    #[test]
    fn eval_numpages() {
        let ctx = ctx_with(|c| c.total_pages = Some(42));
        let instr = FieldInstruction::NumPages { switches: sw() };
        assert_eq!(evaluate(&instr, &ctx), FieldValue::Text("42".into()));
    }

    #[test]
    fn eval_date_default_format() {
        let ctx = ctx_with(|c| {
            c.date = Some(Date {
                year: 2026,
                month: 3,
                day: 24,
            })
        });
        let instr = FieldInstruction::Date {
            format: None,
            switches: sw(),
        };
        assert_eq!(
            evaluate(&instr, &ctx),
            FieldValue::Text("3/24/2026".into())
        );
    }

    #[test]
    fn eval_date_custom_format() {
        let ctx = ctx_with(|c| {
            c.date = Some(Date {
                year: 2026,
                month: 12,
                day: 5,
            })
        });
        let instr = FieldInstruction::Date {
            format: None,
            switches: CommonSwitches {
                date_format: Some("dd/MM/yyyy".into()),
                ..Default::default()
            },
        };
        assert_eq!(
            evaluate(&instr, &ctx),
            FieldValue::Text("05/12/2026".into())
        );
    }

    #[test]
    fn eval_time() {
        let ctx = ctx_with(|c| {
            c.time = Some(Time {
                hour: 14,
                minute: 30,
                second: 0,
            })
        });
        let instr = FieldInstruction::Time {
            format: None,
            switches: sw(),
        };
        assert_eq!(
            evaluate(&instr, &ctx),
            FieldValue::Text("2:30 PM".into())
        );
    }

    #[test]
    fn eval_author() {
        let ctx = ctx_with(|c| {
            c.document_properties
                .insert("author".into(), "Jane Doe".into());
        });
        let instr = FieldInstruction::Author {
            format: None,
            switches: sw(),
        };
        assert_eq!(
            evaluate(&instr, &ctx),
            FieldValue::Text("Jane Doe".into())
        );
    }

    #[test]
    fn eval_mergefield() {
        let ctx = ctx_with(|c| {
            c.merge_data = Some(HashMap::from([("FirstName".into(), "Alice".into())]));
        });
        let instr = FieldInstruction::MergeField {
            name: "FirstName".into(),
            text_before: Some("Dear ".into()),
            text_after: Some(",".into()),
            switches: sw(),
        };
        assert_eq!(
            evaluate(&instr, &ctx),
            FieldValue::Text("Dear Alice,".into())
        );
    }

    #[test]
    fn eval_mergefield_missing() {
        let ctx = FieldContext::default();
        let instr = FieldInstruction::MergeField {
            name: "FirstName".into(),
            text_before: None,
            text_after: None,
            switches: sw(),
        };
        assert_eq!(evaluate(&instr, &ctx), FieldValue::Unevaluable);
    }

    #[test]
    fn eval_if_equal() {
        let ctx = FieldContext::default();
        let instr = FieldInstruction::If {
            left: "hello".into(),
            operator: ComparisonOp::Equal,
            right: "hello".into(),
            then_text: "yes".into(),
            else_text: "no".into(),
            switches: sw(),
        };
        assert_eq!(evaluate(&instr, &ctx), FieldValue::Text("yes".into()));
    }

    #[test]
    fn eval_if_numeric() {
        let ctx = FieldContext::default();
        let instr = FieldInstruction::If {
            left: "10".into(),
            operator: ComparisonOp::GreaterThan,
            right: "5".into(),
            then_text: "big".into(),
            else_text: "small".into(),
            switches: sw(),
        };
        assert_eq!(evaluate(&instr, &ctx), FieldValue::Text("big".into()));
    }

    #[test]
    fn eval_pageref() {
        let ctx = ctx_with(|c| {
            c.bookmarks.insert(
                "_Ref123".into(),
                BookmarkLocation {
                    page: Some(5),
                    text: None,
                    paragraph_number: None,
                    note_ref: None,
                },
            );
        });
        let instr = FieldInstruction::PageRef {
            bookmark: "_Ref123".into(),
            hyperlink: false,
            switches: sw(),
        };
        assert_eq!(evaluate(&instr, &ctx), FieldValue::Text("5".into()));
    }

    #[test]
    fn eval_symbol() {
        let ctx = FieldContext::default();
        let instr = FieldInstruction::Symbol {
            char_code: 169,
            font: None,
            switches: sw(),
        };
        assert_eq!(evaluate(&instr, &ctx), FieldValue::Text("\u{a9}".into()));
    }

    #[test]
    fn eval_with_numeric_format() {
        let ctx = ctx_with(|c| c.page_number = Some(7));
        let instr = FieldInstruction::Page {
            switches: CommonSwitches {
                numeric_format: Some("0.00".into()),
                ..Default::default()
            },
        };
        assert_eq!(evaluate(&instr, &ctx), FieldValue::Text("7.00".into()));
    }

    #[test]
    fn eval_with_upper_format() {
        let ctx = ctx_with(|c| {
            c.document_properties
                .insert("author".into(), "jane doe".into());
        });
        let instr = FieldInstruction::Author {
            format: None,
            switches: CommonSwitches {
                format: Some("Upper".into()),
                ..Default::default()
            },
        };
        assert_eq!(
            evaluate(&instr, &ctx),
            FieldValue::Text("JANE DOE".into())
        );
    }

    #[test]
    fn eval_seq_next() {
        let ctx = ctx_with(|c| {
            c.sequences.insert("Figure".into(), 3);
        });
        let instr = FieldInstruction::Seq {
            identifier: "Figure".into(),
            reset_to: None,
            repeat: false,
            next: true,
            switches: sw(),
        };
        assert_eq!(evaluate(&instr, &ctx), FieldValue::Text("4".into()));
    }

    #[test]
    fn eval_seq_reset() {
        let ctx = FieldContext::default();
        let instr = FieldInstruction::Seq {
            identifier: "Table".into(),
            reset_to: Some("1".into()),
            repeat: false,
            next: false,
            switches: sw(),
        };
        assert_eq!(evaluate(&instr, &ctx), FieldValue::Text("1".into()));
    }

    #[test]
    fn eval_filename() {
        let ctx = ctx_with(|c| {
            c.file_name = Some("document.docx".into());
            c.file_path = Some("/home/user/document.docx".into());
        });

        let instr_name = FieldInstruction::FileName {
            path: false,
            switches: sw(),
        };
        assert_eq!(
            evaluate(&instr_name, &ctx),
            FieldValue::Text("document.docx".into())
        );

        let instr_path = FieldInstruction::FileName {
            path: true,
            switches: sw(),
        };
        assert_eq!(
            evaluate(&instr_path, &ctx),
            FieldValue::Text("/home/user/document.docx".into())
        );
    }

    #[test]
    fn eval_unknown() {
        let ctx = FieldContext::default();
        let instr = FieldInstruction::Unknown {
            field_type: "FOOBAR".into(),
            raw: " FOOBAR ".into(),
        };
        assert_eq!(evaluate(&instr, &ctx), FieldValue::Unevaluable);
    }
}
