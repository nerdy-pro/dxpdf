use crate::ast::CommonSwitches;
use crate::error::FieldParseError;
use crate::parse::Token;

/// Consume common formatting switches (`\*`, `\#`, `\@`) from a token stream.
///
/// Returns the consumed switches and any tokens that were not common switches
/// (i.e., field-specific switches or arguments).
pub(crate) fn extract_common_switches(
    tokens: &[Token],
) -> Result<(CommonSwitches, Vec<Token>), FieldParseError> {
    let mut switches = CommonSwitches::default();
    let mut remaining = Vec::new();
    let mut i = 0;

    while i < tokens.len() {
        match &tokens[i] {
            Token::Switch(ch, pos) if *ch == '*' => {
                i += 1;
                let value = take_switch_value(tokens, &mut i, "\\*", *pos)?;
                switches.format = Some(value);
            }
            Token::Switch(ch, pos) if *ch == '#' => {
                i += 1;
                let value = take_switch_value(tokens, &mut i, "\\#", *pos)?;
                switches.numeric_format = Some(value);
            }
            Token::Switch(ch, pos) if *ch == '@' => {
                i += 1;
                let value = take_switch_value(tokens, &mut i, "\\@", *pos)?;
                switches.date_format = Some(value);
            }
            _ => {
                remaining.push(tokens[i].clone());
                i += 1;
            }
        }
    }

    Ok((switches, remaining))
}

/// Take a switch value — either the next quoted string or word token.
fn take_switch_value(
    tokens: &[Token],
    i: &mut usize,
    switch_name: &str,
    _pos: usize,
) -> Result<String, FieldParseError> {
    if *i < tokens.len() {
        match &tokens[*i] {
            Token::Quoted(s, _) | Token::Word(s, _) => {
                let val = s.clone();
                *i += 1;
                Ok(val)
            }
            _ => Err(FieldParseError::MissingSwitchValue {
                switch: switch_name.to_string(),
            }),
        }
    } else {
        Err(FieldParseError::MissingSwitchValue {
            switch: switch_name.to_string(),
        })
    }
}

/// Check if the next token at index `i` is a switch with the given character.
/// If so, consume it and return true, advancing `i`.
pub(crate) fn has_flag(tokens: &[Token], ch: char) -> bool {
    tokens
        .iter()
        .any(|t| matches!(t, Token::Switch(c, _) if *c == ch))
}

/// Find a switch with the given character and consume its value argument.
/// Returns `Some(value)` if found, `None` if not present.
pub(crate) fn take_switch_with_value(
    tokens: &mut Vec<Token>,
    ch: char,
) -> Result<Option<String>, FieldParseError> {
    let pos = tokens
        .iter()
        .position(|t| matches!(t, Token::Switch(c, _) if *c == ch));

    let Some(idx) = pos else {
        return Ok(None);
    };

    let switch_pos = match &tokens[idx] {
        Token::Switch(_, p) => *p,
        _ => unreachable!(),
    };

    // Remove the switch token
    tokens.remove(idx);

    // The value should now be at `idx` (shifted down)
    if idx < tokens.len() {
        match &tokens[idx] {
            Token::Quoted(s, _) | Token::Word(s, _) => {
                let val = s.clone();
                tokens.remove(idx);
                Ok(Some(val))
            }
            _ => Err(FieldParseError::MissingSwitchValue {
                switch: format!("\\{ch}"),
            }),
        }
    } else {
        Err(FieldParseError::MissingSwitchValue {
            switch: format!("\\{ch} at position {switch_pos}"),
        })
    }
}

/// Remove all flag switches (switches without values) with the given character.
#[allow(dead_code)]
pub(crate) fn remove_flags(tokens: &mut Vec<Token>, ch: char) {
    tokens.retain(|t| !matches!(t, Token::Switch(c, _) if *c == ch));
}
