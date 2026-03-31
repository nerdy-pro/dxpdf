use crate::field::ast::{CommonSwitches, ComparisonOp, FieldInstruction};
use crate::field::error::FieldParseError;
use crate::field::switches;

// ── Tokenizer ────────────────────────────────────────────────────────

/// A token produced by the field instruction tokenizer.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Token {
    /// An unquoted word.
    Word(String, usize),
    /// A double-quoted string (contents only, quotes stripped).
    Quoted(String, usize),
    /// A switch: backslash followed by a character, e.g. `\o`, `\*`.
    Switch(char, usize),
}

/// Tokenize a field instruction string into a sequence of tokens.
pub(crate) fn tokenize(input: &str) -> Result<Vec<Token>, FieldParseError> {
    let mut tokens = Vec::new();
    let bytes = input.as_bytes();
    let len = bytes.len();
    let mut i = 0;

    while i < len {
        // Skip whitespace
        if bytes[i].is_ascii_whitespace() {
            i += 1;
            continue;
        }

        // Quoted string
        if bytes[i] == b'"' {
            let start = i;
            i += 1; // skip opening quote
            let mut s = String::new();
            while i < len && bytes[i] != b'"' {
                s.push(bytes[i] as char);
                i += 1;
            }
            if i >= len {
                return Err(FieldParseError::UnterminatedString { position: start });
            }
            i += 1; // skip closing quote
                    // Handle Unicode: re-decode from the original slice for non-ASCII
            let content = &input[start + 1..i - 1];
            tokens.push(Token::Quoted(content.to_string(), start));
            continue;
        }

        // Switch: backslash followed by a non-whitespace character
        if bytes[i] == b'\\' {
            let start = i;
            i += 1;
            if i < len && !bytes[i].is_ascii_whitespace() {
                let ch = bytes[i] as char;
                i += 1;
                tokens.push(Token::Switch(ch, start));
            } else {
                return Err(FieldParseError::InvalidSwitch {
                    position: start,
                    found: "\\".to_string(),
                });
            }
            continue;
        }

        // Word: sequence of non-whitespace, non-quote, non-backslash characters
        let start = i;
        while i < len && !bytes[i].is_ascii_whitespace() && bytes[i] != b'"' && bytes[i] != b'\\' {
            i += 1;
        }
        let word = &input[start..i];
        tokens.push(Token::Word(word.to_string(), start));
    }

    Ok(tokens)
}

// ── Parser ───────────────────────────────────────────────────────────

/// Parse a field instruction string into a typed `FieldInstruction`.
pub fn parse(input: &str) -> Result<FieldInstruction, FieldParseError> {
    let tokens = tokenize(input)?;
    if tokens.is_empty() {
        return Err(FieldParseError::Empty);
    }

    // First token must be a word (the field type keyword)
    let (field_type_str, rest) = match &tokens[0] {
        Token::Word(w, _) => (w.as_str(), &tokens[1..]),
        Token::Quoted(_, pos) | Token::Switch(_, pos) => {
            return Err(FieldParseError::MissingArgument {
                field_type: String::new(),
                argument: format!("field type keyword at position {pos}"),
            });
        }
    };

    // Extract common switches first
    let (common, mut remaining) = switches::extract_common_switches(rest)?;

    match field_type_str.to_ascii_uppercase().as_str() {
        "PAGE" => Ok(FieldInstruction::Page { switches: common }),
        "NUMPAGES" => Ok(FieldInstruction::NumPages { switches: common }),
        "SECTION" => Ok(FieldInstruction::Section { switches: common }),
        "SECTIONPAGES" => Ok(FieldInstruction::SectionPages { switches: common }),

        "DATE" => {
            let format = take_first_arg(&mut remaining);
            Ok(FieldInstruction::Date {
                format,
                switches: common,
            })
        }
        "TIME" => {
            let format = take_first_arg(&mut remaining);
            Ok(FieldInstruction::Time {
                format,
                switches: common,
            })
        }
        "AUTHOR" => {
            let format = take_first_arg(&mut remaining);
            Ok(FieldInstruction::Author {
                format,
                switches: common,
            })
        }
        "TITLE" => Ok(FieldInstruction::Title { switches: common }),
        "SUBJECT" => Ok(FieldInstruction::Subject { switches: common }),

        "FILENAME" => {
            let path = switches::has_flag(&remaining, 'p');
            Ok(FieldInstruction::FileName {
                path,
                switches: common,
            })
        }

        "REF" => {
            let bookmark = require_first_arg(&mut remaining, "REF", "bookmark")?;
            let footnote_ref = switches::has_flag(&remaining, 'f');
            let paragraph_number = switches::has_flag(&remaining, 'n');
            let hyperlink = switches::has_flag(&remaining, 'h');
            Ok(FieldInstruction::Ref {
                bookmark,
                footnote_ref,
                paragraph_number,
                hyperlink,
                switches: common,
            })
        }

        "PAGEREF" => {
            let bookmark = require_first_arg(&mut remaining, "PAGEREF", "bookmark")?;
            let hyperlink = switches::has_flag(&remaining, 'h');
            Ok(FieldInstruction::PageRef {
                bookmark,
                hyperlink,
                switches: common,
            })
        }

        "NOTEREF" => {
            let bookmark = require_first_arg(&mut remaining, "NOTEREF", "bookmark")?;
            let superscript = switches::has_flag(&remaining, 'f');
            let hyperlink = switches::has_flag(&remaining, 'h');
            Ok(FieldInstruction::NoteRef {
                bookmark,
                superscript,
                hyperlink,
                switches: common,
            })
        }

        "HYPERLINK" => {
            let target = require_first_arg(&mut remaining, "HYPERLINK", "target URL")?;
            let anchor = switches::take_switch_with_value(&mut remaining, 'l')?;
            let target_frame = switches::take_switch_with_value(&mut remaining, 't')?;
            let image_map = switches::has_flag(&remaining, 'm');
            Ok(FieldInstruction::Hyperlink {
                target,
                anchor,
                target_frame,
                image_map,
                switches: common,
            })
        }

        "TOC" => {
            let outline_levels = switches::take_switch_with_value(&mut remaining, 'o')?;
            let hyperlinks = switches::has_flag(&remaining, 'h');
            let no_page_numbers = switches::has_flag(&remaining, 'n');
            let custom_styles = switches::take_switch_with_value(&mut remaining, 't')?;
            let tc_identifier = switches::take_switch_with_value(&mut remaining, 'f')?;
            Ok(FieldInstruction::Toc {
                outline_levels,
                hyperlinks,
                no_page_numbers,
                custom_styles,
                tc_identifier,
                switches: common,
            })
        }

        "SEQ" => {
            let identifier = require_first_arg(&mut remaining, "SEQ", "identifier")?;
            let reset_to = switches::take_switch_with_value(&mut remaining, 'r')?;
            let repeat = switches::has_flag(&remaining, 'c');
            let next = switches::has_flag(&remaining, 'n');
            Ok(FieldInstruction::Seq {
                identifier,
                reset_to,
                repeat,
                next,
                switches: common,
            })
        }

        "IF" => parse_if_field(&mut remaining, common),

        "MERGEFIELD" => {
            let name = require_first_arg(&mut remaining, "MERGEFIELD", "field name")?;
            let text_before = switches::take_switch_with_value(&mut remaining, 'b')?;
            let text_after = switches::take_switch_with_value(&mut remaining, 'f')?;
            Ok(FieldInstruction::MergeField {
                name,
                text_before,
                text_after,
                switches: common,
            })
        }

        "DOCPROPERTY" => {
            let name = require_first_arg(&mut remaining, "DOCPROPERTY", "property name")?;
            Ok(FieldInstruction::DocProperty {
                name,
                switches: common,
            })
        }

        "SYMBOL" => {
            let code_str = require_first_arg(&mut remaining, "SYMBOL", "character code")?;
            let char_code =
                code_str
                    .parse::<u32>()
                    .map_err(|e| FieldParseError::InvalidNumber {
                        value: code_str,
                        reason: e.to_string(),
                    })?;
            let font = switches::take_switch_with_value(&mut remaining, 'f')?;
            Ok(FieldInstruction::Symbol {
                char_code,
                font,
                switches: common,
            })
        }

        "INCLUDEPICTURE" => {
            let path = require_first_arg(&mut remaining, "INCLUDEPICTURE", "file path")?;
            let no_store = switches::has_flag(&remaining, 'd');
            Ok(FieldInstruction::IncludePicture {
                path,
                no_store,
                switches: common,
            })
        }

        "EQ" => {
            // Collect all remaining tokens as the equation text
            let equation = collect_remaining_as_text(&remaining);
            Ok(FieldInstruction::Eq {
                equation,
                switches: common,
            })
        }

        _ => Ok(FieldInstruction::Unknown {
            field_type: field_type_str.to_string(),
            raw: input.to_string(),
        }),
    }
}

// ── Helpers ──────────────────────────────────────────────────────────

/// Take the first word/quoted argument from the remaining tokens, if present.
fn take_first_arg(tokens: &mut Vec<Token>) -> Option<String> {
    if tokens.is_empty() {
        return None;
    }
    match &tokens[0] {
        Token::Word(s, _) | Token::Quoted(s, _) => {
            let val = s.clone();
            tokens.remove(0);
            Some(val)
        }
        _ => None,
    }
}

/// Require the first word/quoted argument, returning an error if missing.
fn require_first_arg(
    tokens: &mut Vec<Token>,
    field_type: &str,
    argument: &str,
) -> Result<String, FieldParseError> {
    take_first_arg(tokens).ok_or_else(|| FieldParseError::MissingArgument {
        field_type: field_type.to_string(),
        argument: argument.to_string(),
    })
}

/// Collect all remaining tokens back into a single string.
fn collect_remaining_as_text(tokens: &[Token]) -> String {
    let mut parts = Vec::new();
    for token in tokens {
        match token {
            Token::Word(s, _) => parts.push(s.clone()),
            Token::Quoted(s, _) => {
                parts.push(format!("\"{s}\""));
            }
            Token::Switch(ch, _) => parts.push(format!("\\{ch}")),
        }
    }
    parts.join(" ")
}

/// Parse an IF field: `IF expr op expr "then" "else"`
fn parse_if_field(
    tokens: &mut Vec<Token>,
    switches: CommonSwitches,
) -> Result<FieldInstruction, FieldParseError> {
    let left = require_first_arg(tokens, "IF", "left operand")?;

    let op_str = require_first_arg(tokens, "IF", "comparison operator")?;
    let operator = parse_comparison_op(&op_str)?;

    let right = require_first_arg(tokens, "IF", "right operand")?;

    let then_text = take_first_arg(tokens).unwrap_or_default();
    let else_text = take_first_arg(tokens).unwrap_or_default();

    Ok(FieldInstruction::If {
        left,
        operator,
        right,
        then_text,
        else_text,
        switches,
    })
}

fn parse_comparison_op(s: &str) -> Result<ComparisonOp, FieldParseError> {
    match s {
        "=" => Ok(ComparisonOp::Equal),
        "<>" => Ok(ComparisonOp::NotEqual),
        "<" => Ok(ComparisonOp::LessThan),
        "<=" => Ok(ComparisonOp::LessEqual),
        ">" => Ok(ComparisonOp::GreaterThan),
        ">=" => Ok(ComparisonOp::GreaterEqual),
        _ => Err(FieldParseError::InvalidOperator {
            found: s.to_string(),
        }),
    }
}

// ── Tests ────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use crate::field::ast::*;

    #[test]
    fn tokenize_simple() {
        let tokens = tokenize(" PAGE ").unwrap();
        assert_eq!(tokens, vec![Token::Word("PAGE".into(), 1)]);
    }

    #[test]
    fn tokenize_switches_and_quoted() {
        let tokens = tokenize(r#" TOC \o "1-3" \h "#).unwrap();
        assert_eq!(
            tokens,
            vec![
                Token::Word("TOC".into(), 1),
                Token::Switch('o', 5),
                Token::Quoted("1-3".into(), 8),
                Token::Switch('h', 14),
            ]
        );
    }

    #[test]
    fn tokenize_unterminated_string() {
        let err = tokenize(r#" REF "unterminated "#).unwrap_err();
        assert!(matches!(err, FieldParseError::UnterminatedString { .. }));
    }

    #[test]
    fn parse_page() {
        let instr = parse(" PAGE ").unwrap();
        assert_eq!(
            instr,
            FieldInstruction::Page {
                switches: CommonSwitches::default()
            }
        );
    }

    #[test]
    fn parse_page_case_insensitive() {
        let instr = parse(" page ").unwrap();
        assert_eq!(
            instr,
            FieldInstruction::Page {
                switches: CommonSwitches::default()
            }
        );
    }

    #[test]
    fn parse_page_with_format_switch() {
        let instr = parse(r#" PAGE \* MERGEFORMAT "#).unwrap();
        assert_eq!(
            instr,
            FieldInstruction::Page {
                switches: CommonSwitches {
                    format: Some("MERGEFORMAT".into()),
                    numeric_format: None,
                    date_format: None,
                }
            }
        );
    }

    #[test]
    fn parse_numpages() {
        let instr = parse(" NUMPAGES ").unwrap();
        assert_eq!(
            instr,
            FieldInstruction::NumPages {
                switches: CommonSwitches::default()
            }
        );
    }

    #[test]
    fn parse_date_with_format() {
        let instr = parse(r#" DATE \@ "dd/MM/yyyy" "#).unwrap();
        assert_eq!(
            instr,
            FieldInstruction::Date {
                format: None,
                switches: CommonSwitches {
                    format: None,
                    numeric_format: None,
                    date_format: Some("dd/MM/yyyy".into()),
                },
            }
        );
    }

    #[test]
    fn parse_ref_with_switches() {
        let instr = parse(r" REF _Toc123 \h \n ").unwrap();
        assert_eq!(
            instr,
            FieldInstruction::Ref {
                bookmark: "_Toc123".into(),
                footnote_ref: false,
                paragraph_number: true,
                hyperlink: true,
                switches: CommonSwitches::default(),
            }
        );
    }

    #[test]
    fn parse_ref_missing_bookmark() {
        let err = parse(" REF ").unwrap_err();
        assert!(matches!(err, FieldParseError::MissingArgument { .. }));
    }

    #[test]
    fn parse_hyperlink() {
        let instr = parse(r#" HYPERLINK "https://example.com" \l "section1" "#).unwrap();
        assert_eq!(
            instr,
            FieldInstruction::Hyperlink {
                target: "https://example.com".into(),
                anchor: Some("section1".into()),
                target_frame: None,
                image_map: false,
                switches: CommonSwitches::default(),
            }
        );
    }

    #[test]
    fn parse_toc() {
        let instr = parse(r#" TOC \o "1-3" \h "#).unwrap();
        assert_eq!(
            instr,
            FieldInstruction::Toc {
                outline_levels: Some("1-3".into()),
                hyperlinks: true,
                no_page_numbers: false,
                custom_styles: None,
                tc_identifier: None,
                switches: CommonSwitches::default(),
            }
        );
    }

    #[test]
    fn parse_if_field() {
        let instr = parse(r#" IF x = y "yes" "no" "#).unwrap();
        assert_eq!(
            instr,
            FieldInstruction::If {
                left: "x".into(),
                operator: ComparisonOp::Equal,
                right: "y".into(),
                then_text: "yes".into(),
                else_text: "no".into(),
                switches: CommonSwitches::default(),
            }
        );
    }

    #[test]
    fn parse_mergefield() {
        let instr = parse(r#" MERGEFIELD FirstName \* Upper "#).unwrap();
        assert_eq!(
            instr,
            FieldInstruction::MergeField {
                name: "FirstName".into(),
                text_before: None,
                text_after: None,
                switches: CommonSwitches {
                    format: Some("Upper".into()),
                    numeric_format: None,
                    date_format: None,
                },
            }
        );
    }

    #[test]
    fn parse_symbol() {
        let instr = parse(r#" SYMBOL 169 \f "Symbol" "#).unwrap();
        assert_eq!(
            instr,
            FieldInstruction::Symbol {
                char_code: 169,
                font: Some("Symbol".into()),
                switches: CommonSwitches::default(),
            }
        );
    }

    #[test]
    fn parse_symbol_invalid_code() {
        let err = parse(" SYMBOL abc ").unwrap_err();
        assert!(matches!(err, FieldParseError::InvalidNumber { .. }));
    }

    #[test]
    fn parse_unknown_field() {
        let input = " FOOBAR something ";
        let instr = parse(input).unwrap();
        assert_eq!(
            instr,
            FieldInstruction::Unknown {
                field_type: "FOOBAR".into(),
                raw: input.to_string(),
            }
        );
    }

    #[test]
    fn parse_empty() {
        let err = parse("   ").unwrap_err();
        assert_eq!(err, FieldParseError::Empty);
    }

    #[test]
    fn parse_filename_with_path() {
        let instr = parse(r" FILENAME \p ").unwrap();
        assert_eq!(
            instr,
            FieldInstruction::FileName {
                path: true,
                switches: CommonSwitches::default(),
            }
        );
    }

    #[test]
    fn parse_seq() {
        let instr = parse(r" SEQ Figure \r 1 ").unwrap();
        assert_eq!(
            instr,
            FieldInstruction::Seq {
                identifier: "Figure".into(),
                reset_to: Some("1".into()),
                repeat: false,
                next: false,
                switches: CommonSwitches::default(),
            }
        );
    }

    #[test]
    fn parse_pageref() {
        let instr = parse(r" PAGEREF _Ref123 \h ").unwrap();
        assert_eq!(
            instr,
            FieldInstruction::PageRef {
                bookmark: "_Ref123".into(),
                hyperlink: true,
                switches: CommonSwitches::default(),
            }
        );
    }

    #[test]
    fn parse_docproperty() {
        let instr = parse(" DOCPROPERTY Author ").unwrap();
        assert_eq!(
            instr,
            FieldInstruction::DocProperty {
                name: "Author".into(),
                switches: CommonSwitches::default(),
            }
        );
    }

    #[test]
    fn parse_includepicture() {
        let instr = parse(r#" INCLUDEPICTURE "image.png" \d "#).unwrap();
        assert_eq!(
            instr,
            FieldInstruction::IncludePicture {
                path: "image.png".into(),
                no_store: true,
                switches: CommonSwitches::default(),
            }
        );
    }

    #[test]
    fn parse_noteref() {
        let instr = parse(r" NOTEREF _Ref456 \f \h ").unwrap();
        assert_eq!(
            instr,
            FieldInstruction::NoteRef {
                bookmark: "_Ref456".into(),
                superscript: true,
                hyperlink: true,
                switches: CommonSwitches::default(),
            }
        );
    }

    #[test]
    fn parse_numeric_format() {
        let instr = parse(r#" PAGE \# "0.00" "#).unwrap();
        assert_eq!(
            instr,
            FieldInstruction::Page {
                switches: CommonSwitches {
                    format: None,
                    numeric_format: Some("0.00".into()),
                    date_format: None,
                }
            }
        );
    }

    #[test]
    fn parse_if_not_equal() {
        let instr = parse(r#" IF a <> b "different" "same" "#).unwrap();
        assert_eq!(
            instr,
            FieldInstruction::If {
                left: "a".into(),
                operator: ComparisonOp::NotEqual,
                right: "b".into(),
                then_text: "different".into(),
                else_text: "same".into(),
                switches: CommonSwitches::default(),
            }
        );
    }
}
