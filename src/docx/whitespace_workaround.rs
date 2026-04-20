//! Workaround for quick-xml's silent whitespace stripping.
//!
//! # The problem
//!
//! quick-xml's serde [`Deserializer`](quick_xml::de::Deserializer) wraps every
//! input reader in a private trimmer that drops `Text` events containing only
//! whitespace. Even when the source XML carries `xml:space="preserve"`, the
//! trimmer eats the content before any visitor sees it. Custom `Deserialize`
//! impls cannot intercept — the relevant constructors and trim flag are private,
//! and `XmlRead` cannot be plugged in via any public API in 0.39.
//!
//! In Word documents this typically appears as `<w:r>` runs that contain a
//! single whitespace-only `<w:t xml:space="preserve"> </w:t>` between two
//! other runs (e.g. a label followed by a value with different formatting).
//! Without this workaround, the space disappears and rendered text reads
//! `Label:Value` instead of `Label: Value`.
//!
//! # The hack
//!
//! Before handing each XML part to quick-xml we scan for whitespace-only
//! `<w:t...>...</w:t>` content and substitute each whitespace byte with a
//! Private Use Area codepoint. quick-xml then sees a non-whitespace text node
//! and preserves it. The body parser ([`crate::docx::parse::body`]) reverses
//! the substitution when emitting [`RunElement::Text`].
//!
//! Sentinel mapping (BYTE → CHAR):
//!
//! | source byte | sentinel    |
//! |-------------|-------------|
//! | `0x20` SP   | `U+E020`    |
//! | `0x09` HT   | `U+E009`    |
//! | `0x0A` LF   | `U+E00A`    |
//! | `0x0D` CR   | `U+E00D`    |
//!
//! The pattern `0xE000 + <ASCII byte>` keeps the reversal trivial and ensures
//! sentinels never collide with anything OOXML legitimately emits — Word never
//! writes Private Use Area codepoints.
//!
//! TODO(quick-xml): drop this module once one of the following lands upstream
//! and we adopt it:
//!     * a public `Deserializer` constructor that accepts a custom `XmlRead`,
//!     * a public knob to disable `StartTrimmer`, or
//!     * `xml:space="preserve"` honored for whitespace-only `Text` events.
//!
//! See <https://github.com/tafia/quick-xml/issues/> — file/link before merge.

/// Sentinel for ASCII space (`0x20`).
pub(crate) const WS_SENTINEL_SPACE: char = '\u{E020}';
/// Sentinel for ASCII tab (`0x09`).
pub(crate) const WS_SENTINEL_TAB: char = '\u{E009}';
/// Sentinel for ASCII line feed (`0x0A`).
pub(crate) const WS_SENTINEL_LF: char = '\u{E00A}';
/// Sentinel for ASCII carriage return (`0x0D`).
pub(crate) const WS_SENTINEL_CR: char = '\u{E00D}';

/// Pre-process an XML byte buffer so whitespace-only `<w:t>` content survives
/// quick-xml's trimmer. Returns the original buffer unchanged when no
/// substitution is needed (common case for binary parts).
///
/// See module-level docs for why this is necessary.
pub(crate) fn substitute_whitespace_only_runs(xml: &[u8]) -> Vec<u8> {
    // Fast path: if no `<w:t` substring appears in the buffer there is nothing
    // to scan. This avoids paying the per-byte loop cost on parts (images,
    // theme XML, etc.) that have no run text.
    if !contains_subslice(xml, b"<w:t") {
        return xml.to_vec();
    }

    let mut out: Vec<u8> = Vec::with_capacity(xml.len());
    let mut i = 0;
    while i < xml.len() {
        if let Some((content_start, content_end, close_tag_end)) = match_w_t_with_text(xml, i) {
            let content = &xml[content_start..content_end];
            if !content.is_empty() && content.iter().all(is_xml_whitespace_byte) {
                // Copy bytes up to the start of the content (i.e. through `>`),
                // emit sentinels for each whitespace byte, then resume after
                // the original content (the closing `</w:t>` will be copied
                // on the next iteration).
                out.extend_from_slice(&xml[i..content_start]);
                for &b in content {
                    let sentinel = match b {
                        b' ' => WS_SENTINEL_SPACE,
                        b'\t' => WS_SENTINEL_TAB,
                        b'\n' => WS_SENTINEL_LF,
                        b'\r' => WS_SENTINEL_CR,
                        _ => unreachable!("guarded by is_xml_whitespace_byte"),
                    };
                    let mut buf = [0u8; 4];
                    out.extend_from_slice(sentinel.encode_utf8(&mut buf).as_bytes());
                }
                i = content_end;
                continue;
            }
            // Not whitespace-only — leave the element alone but skip past the
            // closing tag in one step so we don't rescan attribute bytes.
            out.extend_from_slice(&xml[i..close_tag_end]);
            i = close_tag_end;
            continue;
        }
        out.push(xml[i]);
        i += 1;
    }
    out
}

/// Reverse the sentinel substitution. Returns the original string when no
/// sentinels are present (zero-allocation common path).
pub(crate) fn restore_whitespace_sentinels(s: &str) -> String {
    if !s.chars().any(is_ws_sentinel) {
        return s.to_string();
    }
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        out.push(match ch {
            WS_SENTINEL_SPACE => ' ',
            WS_SENTINEL_TAB => '\t',
            WS_SENTINEL_LF => '\n',
            WS_SENTINEL_CR => '\r',
            other => other,
        });
    }
    out
}

#[inline]
fn is_xml_whitespace_byte(b: &u8) -> bool {
    matches!(*b, b' ' | b'\t' | b'\n' | b'\r')
}

#[inline]
fn is_ws_sentinel(c: char) -> bool {
    matches!(
        c,
        WS_SENTINEL_SPACE | WS_SENTINEL_TAB | WS_SENTINEL_LF | WS_SENTINEL_CR
    )
}

/// True if `needle` appears anywhere in `haystack`.
fn contains_subslice(haystack: &[u8], needle: &[u8]) -> bool {
    if needle.is_empty() {
        return true;
    }
    haystack.windows(needle.len()).any(|w| w == needle)
}

/// Try to match `<w:t...>CONTENT</w:t>` starting at `pos`.
/// On success, returns `(content_start, content_end, close_tag_end)` where
/// `close_tag_end` is the byte index just past `</w:t>`.
///
/// Self-closing `<w:t/>` is intentionally ignored — there is no content to
/// preserve. Other `<w:t...` prefixes (`<w:tbl`, `<w:tab`) are rejected by the
/// element-name boundary check.
fn match_w_t_with_text(xml: &[u8], pos: usize) -> Option<(usize, usize, usize)> {
    let prefix = b"<w:t";
    if !xml[pos..].starts_with(prefix) {
        return None;
    }
    let after_name = pos + prefix.len();
    let next = *xml.get(after_name)?;
    // `<w:t` must be followed by `>`, `/`, or whitespace. Anything else means
    // we matched the prefix of a longer element name like `<w:tbl`.
    if !matches!(next, b'>' | b'/' | b' ' | b'\t' | b'\n' | b'\r') {
        return None;
    }
    // Find the byte index just past the start tag's closing `>`. We walk
    // forward respecting attribute quoting so a `>` inside an attribute value
    // doesn't fool us. (XML allows unescaped `>` in attribute values; OOXML
    // never emits one but cheap to handle.)
    let mut j = after_name;
    let mut quote: Option<u8> = None;
    let start_tag_end = loop {
        let b = *xml.get(j)?;
        match (quote, b) {
            (None, b'>') => break j + 1,
            (None, b'/') if xml.get(j + 1) == Some(&b'>') => return None, // self-closing
            (None, b'"') | (None, b'\'') => quote = Some(b),
            (Some(q), b) if b == q => quote = None,
            _ => {}
        }
        j += 1;
    };

    // Find the next `<` — that's the start of the closing tag (or a child
    // element, but `<w:t>` per OOXML never has element children).
    let lt = xml[start_tag_end..].iter().position(|&b| b == b'<')?;
    let content_end = start_tag_end + lt;

    let close = b"</w:t>";
    if !xml[content_end..].starts_with(close) {
        return None;
    }
    let close_tag_end = content_end + close.len();
    Some((start_tag_end, content_end, close_tag_end))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn s(b: Vec<u8>) -> String {
        String::from_utf8(b).unwrap()
    }

    // ── substitute_whitespace_only_runs ──────────────────────────────────────

    #[test]
    fn single_space_preserve_is_substituted() {
        let xml = br#"<w:r><w:t xml:space="preserve"> </w:t></w:r>"#;
        let out = s(substitute_whitespace_only_runs(xml));
        assert_eq!(
            out,
            format!(
                r#"<w:r><w:t xml:space="preserve">{}</w:t></w:r>"#,
                WS_SENTINEL_SPACE
            )
        );
    }

    #[test]
    fn three_spaces_preserve_each_substituted() {
        let xml = br#"<w:t xml:space="preserve">   </w:t>"#;
        let out = s(substitute_whitespace_only_runs(xml));
        let expected = format!(
            r#"<w:t xml:space="preserve">{0}{0}{0}</w:t>"#,
            WS_SENTINEL_SPACE
        );
        assert_eq!(out, expected);
    }

    #[test]
    fn whitespace_only_without_preserve_is_also_substituted() {
        // Many third-party DOCX writers emit whitespace-only <w:t> without the
        // xml:space attribute. quick-xml strips them all the same.
        let xml = br#"<w:t> </w:t>"#;
        let out = s(substitute_whitespace_only_runs(xml));
        assert_eq!(out, format!("<w:t>{}</w:t>", WS_SENTINEL_SPACE));
    }

    #[test]
    fn tab_only_is_substituted() {
        let xml = b"<w:t xml:space=\"preserve\">\t</w:t>";
        let out = s(substitute_whitespace_only_runs(xml));
        assert_eq!(
            out,
            format!("<w:t xml:space=\"preserve\">{}</w:t>", WS_SENTINEL_TAB)
        );
    }

    #[test]
    fn lf_and_cr_substituted() {
        let xml = b"<w:t xml:space=\"preserve\">\r\n</w:t>";
        let out = s(substitute_whitespace_only_runs(xml));
        assert_eq!(
            out,
            format!(
                "<w:t xml:space=\"preserve\">{}{}</w:t>",
                WS_SENTINEL_CR, WS_SENTINEL_LF
            )
        );
    }

    #[test]
    fn mixed_content_is_left_untouched() {
        // Content has a non-whitespace char → quick-xml preserves it as-is,
        // we must not modify.
        let xml = br#"<w:t xml:space="preserve">hello world </w:t>"#;
        let out = s(substitute_whitespace_only_runs(xml));
        assert_eq!(out, std::str::from_utf8(xml).unwrap());
    }

    #[test]
    fn empty_w_t_is_left_untouched() {
        let xml = br#"<w:t></w:t>"#;
        let out = s(substitute_whitespace_only_runs(xml));
        assert_eq!(out, std::str::from_utf8(xml).unwrap());
    }

    #[test]
    fn self_closing_w_t_is_left_untouched() {
        let xml = br#"<w:t xml:space="preserve"/>"#;
        let out = s(substitute_whitespace_only_runs(xml));
        assert_eq!(out, std::str::from_utf8(xml).unwrap());
    }

    #[test]
    fn longer_element_names_are_not_matched() {
        // `<w:tbl>`, `<w:tab/>`, `<w:tc>` all start with `<w:t` but should not
        // trigger substitution.
        let xml = br#"<w:tbl><w:tr><w:tc><w:tab/></w:tc></w:tr></w:tbl>"#;
        let out = s(substitute_whitespace_only_runs(xml));
        assert_eq!(out, std::str::from_utf8(xml).unwrap());
    }

    #[test]
    fn multiple_runs_in_one_buffer() {
        let xml = br#"<w:p><w:r><w:t>A</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>"#;
        let out = s(substitute_whitespace_only_runs(xml));
        let expected = format!(
            r#"<w:p><w:r><w:t>A</w:t></w:r><w:r><w:t xml:space="preserve">{}</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>"#,
            WS_SENTINEL_SPACE
        );
        assert_eq!(out, expected);
    }

    #[test]
    fn buffer_without_w_t_is_unchanged_via_fast_path() {
        let xml = br#"<a:theme><a:clrScheme/></a:theme>"#;
        let out = substitute_whitespace_only_runs(xml);
        assert_eq!(&out[..], &xml[..]);
    }

    #[test]
    fn empty_buffer_is_handled() {
        let xml: &[u8] = b"";
        let out = substitute_whitespace_only_runs(xml);
        assert!(out.is_empty());
    }

    #[test]
    fn attribute_with_quoted_gt_is_handled() {
        // XML technically allows `>` in attribute values without escaping.
        // The scanner must respect quoting so the inner `>` doesn't end the
        // start tag prematurely.
        let xml = br#"<w:t weird="a>b" xml:space="preserve"> </w:t>"#;
        let out = s(substitute_whitespace_only_runs(xml));
        assert_eq!(
            out,
            format!(
                r#"<w:t weird="a>b" xml:space="preserve">{}</w:t>"#,
                WS_SENTINEL_SPACE
            )
        );
    }

    // ── restore_whitespace_sentinels ─────────────────────────────────────────

    #[test]
    fn restore_replaces_each_sentinel_with_original_byte() {
        let input = format!(
            "{}{}{}{}",
            WS_SENTINEL_SPACE, WS_SENTINEL_TAB, WS_SENTINEL_LF, WS_SENTINEL_CR
        );
        assert_eq!(restore_whitespace_sentinels(&input), " \t\n\r");
    }

    #[test]
    fn restore_passes_normal_text_through() {
        let input = "hello world";
        assert_eq!(restore_whitespace_sentinels(input), "hello world");
    }

    #[test]
    fn restore_handles_mixed_text_and_sentinels() {
        let input = format!("a{}b", WS_SENTINEL_SPACE);
        assert_eq!(restore_whitespace_sentinels(&input), "a b");
    }

    // ── round-trip: substitute then parse via quick-xml then restore ────────

    #[test]
    fn round_trip_through_quick_xml_recovers_original_whitespace() {
        use serde::Deserialize;

        #[derive(Deserialize)]
        struct TextXml {
            #[serde(rename = "$text", default)]
            content: String,
        }
        #[derive(Deserialize)]
        struct R {
            #[serde(rename = "t")]
            t: TextXml,
        }

        // The exact problem case from the Protokoll DOCX. We strip the `w:`
        // namespace prefix here only because this test deserializes outside
        // the real document context where namespaces are bound — the
        // substitution function still operates on the canonical `<w:t>` form.
        let original = br#"<r><w:t xml:space="preserve"> </w:t></r>"#;
        let preprocessed = substitute_whitespace_only_runs(original);
        let parsed: R = quick_xml::de::from_str(std::str::from_utf8(&preprocessed).unwrap())
            .expect("quick-xml parse");
        let restored = restore_whitespace_sentinels(&parsed.t.content);
        assert_eq!(restored, " ", "expected single literal space to survive");
    }
}
