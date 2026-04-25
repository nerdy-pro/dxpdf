//! Grapheme cluster classification for inline text.
//!
//! Splits a run's text into a typed sequence of [`InlineCluster`]s. Cluster
//! boundaries follow UAX #29 (Extended Grapheme Cluster). Emoji classification
//! follows UTS #51 (Unicode Emoji): the decision is made from codepoint
//! properties, never from the run's font name.
//!
//! Spec references for individual rules are inline in [`classify_cluster`].

use unicode_properties::{EmojiStatus, UnicodeEmoji};
use unicode_segmentation::UnicodeSegmentation;

// ─── Public ADTs ─────────────────────────────────────────────────────────────

/// One indivisible rendering unit produced from a run's text.
///
/// Adjacent non-emoji clusters are merged into a single [`InlineCluster::Text`]
/// span so the existing text path can measure them in one call. Emoji clusters
/// are surfaced individually because each is rasterized independently.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum InlineCluster<'a> {
    /// Plain text — render via the existing text path. The slice is one or
    /// more contiguous grapheme clusters from the source string.
    Text(&'a str),
    /// Emoji — render via the raster-and-embed path.
    Emoji(EmojiCluster<'a>),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmojiCluster<'a> {
    /// The cluster text exactly as it appeared in the source (no normalization
    /// at this stage — normalization happens at the rasterizer cache key).
    pub text: &'a str,
    pub presentation: EmojiPresentation,
    pub structure: EmojiStructure,
}

/// UTS #51 §2 presentation style. Determined from the default-presentation
/// property of the base scalar, modulated by VS-15 (U+FE0E) / VS-16 (U+FE0F).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EmojiPresentation {
    /// Default-text or VS-15 forced. The codepoint should be rendered as a
    /// monochrome glyph using the run's text font.
    Text,
    /// Default-emoji or VS-16 forced. The codepoint should be rendered via
    /// the color emoji typeface.
    Emoji,
}

/// UTS #51 §2 emoji sequence taxonomy.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EmojiStructure {
    /// Single emoji codepoint, optionally followed by a presentation selector.
    Single,
    /// Keycap: `[0-9#*]` + VS-16 + U+20E3 (COMBINING ENCLOSING KEYCAP).
    KeycapSequence { base: char },
    /// Emoji_Modifier_Base + Emoji_Modifier (skin tone).
    ModifierSequence { base: char, tone: SkinTone },
    /// Flag sequence — either RIS pair or black-flag tag sequence.
    FlagSequence(FlagKind),
    /// Two or more emoji elements joined by U+200D (ZWJ).
    ZwjSequence,
}

/// UTS #51 §1.4.5 emoji modifiers (Fitzpatrick skin tone).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SkinTone {
    Light,       // U+1F3FB
    MediumLight, // U+1F3FC
    Medium,      // U+1F3FD
    MediumDark,  // U+1F3FE
    Dark,        // U+1F3FF
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FlagKind {
    /// Pair of Regional Indicator Symbols (UTS #51 §1.4.7).
    Regional,
    /// Black flag (U+1F3F4) + tag sequence + END_TAG (U+E007F).
    Subdivision,
}

// ─── Public API ──────────────────────────────────────────────────────────────

/// Classify `text` into a sequence of clusters.
///
/// Adjacent text clusters are merged into one [`InlineCluster::Text`] span.
/// Empty input yields an empty vector.
pub fn classify(text: &str) -> Vec<InlineCluster<'_>> {
    let mut out = Vec::new();
    let mut text_start: Option<usize> = None;
    let mut text_end = 0usize;
    let mut byte_offset = 0usize;

    for cluster in text.graphemes(true) {
        let cluster_len = cluster.len();
        match classify_cluster(cluster) {
            ClusterClass::Text => {
                if text_start.is_none() {
                    text_start = Some(byte_offset);
                }
                text_end = byte_offset + cluster_len;
            }
            ClusterClass::Emoji {
                presentation,
                structure,
            } => {
                if let Some(s) = text_start.take() {
                    out.push(InlineCluster::Text(&text[s..text_end]));
                }
                out.push(InlineCluster::Emoji(EmojiCluster {
                    text: cluster,
                    presentation,
                    structure,
                }));
            }
        }
        byte_offset += cluster_len;
    }
    if let Some(s) = text_start {
        out.push(InlineCluster::Text(&text[s..text_end]));
    }
    out
}

// ─── Internal classification ────────────────────────────────────────────────

/// Outcome of classifying one grapheme cluster. Internal — the public API
/// produces [`InlineCluster`]s with text slices.
enum ClusterClass {
    Text,
    Emoji {
        presentation: EmojiPresentation,
        structure: EmojiStructure,
    },
}

fn classify_cluster(cluster: &str) -> ClusterClass {
    // UAX #29 grapheme iteration never yields an empty cluster, but be
    // defensive — an empty input string at the API boundary skips the loop.
    let mut chars = cluster.chars();
    let Some(first) = chars.next() else {
        return ClusterClass::Text;
    };
    let chars: Vec<char> = std::iter::once(first).chain(chars).collect();

    // 1. Subdivision flag — UTS #51 §1.4.7. Black flag base + tag sequence.
    //    Tag chars are U+E0020..U+E007E with U+E007F as terminator.
    if first == '\u{1F3F4}' && chars.iter().any(|&c| is_tag(c)) {
        return ClusterClass::Emoji {
            presentation: EmojiPresentation::Emoji,
            structure: EmojiStructure::FlagSequence(FlagKind::Subdivision),
        };
    }

    // 2. Regional indicator pair — UTS #51 §1.4.7. Exactly two RIS chars.
    let ris_count = chars.iter().filter(|&&c| is_regional_indicator(c)).count();
    if ris_count == 2 && ris_count == chars.len() {
        return ClusterClass::Emoji {
            presentation: EmojiPresentation::Emoji,
            structure: EmojiStructure::FlagSequence(FlagKind::Regional),
        };
    }

    // 3. Keycap — UTS #51 §1.4.4. Base char + VS-16 + U+20E3.
    if chars.contains(&'\u{20E3}') && is_keycap_base(first) {
        return ClusterClass::Emoji {
            presentation: EmojiPresentation::Emoji,
            structure: EmojiStructure::KeycapSequence { base: first },
        };
    }

    // After special sequences, an Emoji=YES char must appear in the cluster
    // before we treat it as emoji at all. Without this guard, "a\u{200D}b"
    // (a non-emoji ZWJ pairing — illegal but representable) would be
    // misclassified as a ZwjSequence.
    let Some(emoji_base) = chars.iter().copied().find(|&c| char_is_emoji(c)) else {
        return ClusterClass::Text;
    };

    let has_zwj = chars.contains(&'\u{200D}');
    let has_vs15 = chars.contains(&'\u{FE0E}');
    let has_vs16 = chars.contains(&'\u{FE0F}');

    // 4. ZWJ sequence — UTS #51 §1.4.6.
    if has_zwj {
        return ClusterClass::Emoji {
            presentation: EmojiPresentation::Emoji,
            structure: EmojiStructure::ZwjSequence,
        };
    }

    // 5. Modifier sequence — UTS #51 §1.4.5.
    if let Some(tone) = chars.iter().find_map(|&c| skin_tone_for(c)) {
        return ClusterClass::Emoji {
            presentation: EmojiPresentation::Emoji,
            structure: EmojiStructure::ModifierSequence {
                base: emoji_base,
                tone,
            },
        };
    }

    // 6. Single emoji codepoint, with optional presentation selector.
    let presentation = if has_vs15 {
        EmojiPresentation::Text
    } else if has_vs16 || has_default_emoji_presentation(emoji_base) {
        EmojiPresentation::Emoji
    } else {
        EmojiPresentation::Text
    };
    ClusterClass::Emoji {
        presentation,
        structure: EmojiStructure::Single,
    }
}

// ─── Codepoint property helpers ──────────────────────────────────────────────

/// `Emoji=YES` per UTS #51. Pure emoji components (RIS, ZWJ, VS-16, digits as
/// keycap bases) are *not* covered here — those are handled by the sequence
/// detectors above.
fn char_is_emoji(c: char) -> bool {
    use EmojiStatus::*;
    matches!(
        c.emoji_status(),
        EmojiPresentation
            | EmojiModifierBase
            | EmojiPresentationAndModifierBase
            | EmojiOther
            | EmojiPresentationAndEmojiComponent
            | EmojiPresentationAndModifierAndEmojiComponent
            | EmojiOtherAndEmojiComponent
    )
}

/// `Emoji_Presentation=YES` per UTS #51 — the codepoint defaults to emoji
/// (color) presentation in the absence of a VS selector.
fn has_default_emoji_presentation(c: char) -> bool {
    use EmojiStatus::*;
    matches!(
        c.emoji_status(),
        EmojiPresentation
            | EmojiPresentationAndModifierBase
            | EmojiPresentationAndEmojiComponent
            | EmojiPresentationAndModifierAndEmojiComponent
    )
}

fn is_regional_indicator(c: char) -> bool {
    matches!(c, '\u{1F1E6}'..='\u{1F1FF}')
}

fn is_tag(c: char) -> bool {
    // UTS #51: tag base U+E0020..U+E007E and END_TAG U+E007F.
    matches!(c, '\u{E0020}'..='\u{E007F}')
}

fn is_keycap_base(c: char) -> bool {
    matches!(c, '0'..='9' | '#' | '*')
}

/// Return the [`SkinTone`] for emoji modifier codepoints U+1F3FB..U+1F3FF.
fn skin_tone_for(c: char) -> Option<SkinTone> {
    match c {
        '\u{1F3FB}' => Some(SkinTone::Light),
        '\u{1F3FC}' => Some(SkinTone::MediumLight),
        '\u{1F3FD}' => Some(SkinTone::Medium),
        '\u{1F3FE}' => Some(SkinTone::MediumDark),
        '\u{1F3FF}' => Some(SkinTone::Dark),
        _ => None,
    }
}

// ─── Tests (C1–C13 from docs/emoji-rendering.md) ─────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    fn emoji(
        text: &str,
        presentation: EmojiPresentation,
        structure: EmojiStructure,
    ) -> InlineCluster<'_> {
        InlineCluster::Emoji(EmojiCluster {
            text,
            presentation,
            structure,
        })
    }

    /// C1 — pure ASCII text, no emoji.
    #[test]
    fn c1_pure_text() {
        assert_eq!(classify("hello"), vec![InlineCluster::Text("hello")]);
    }

    /// C2 — emoji embedded in surrounding text. Adjacent text clusters merge
    /// into a single span; the emoji is surfaced on its own.
    #[test]
    fn c2_emoji_between_text() {
        assert_eq!(
            classify("hi 📞 there"),
            vec![
                InlineCluster::Text("hi "),
                emoji("📞", EmojiPresentation::Emoji, EmojiStructure::Single),
                InlineCluster::Text(" there"),
            ]
        );
    }

    /// C3 — U+260E BLACK TELEPHONE has Emoji=YES with default-text presentation.
    #[test]
    fn c3_default_text_presentation() {
        assert_eq!(
            classify("\u{260E}"),
            vec![emoji(
                "\u{260E}",
                EmojiPresentation::Text,
                EmojiStructure::Single
            )]
        );
    }

    /// C4 — VS-16 promotes a default-text emoji to emoji presentation.
    #[test]
    fn c4_vs16_promotes_to_emoji() {
        assert_eq!(
            classify("\u{260E}\u{FE0F}"),
            vec![emoji(
                "\u{260E}\u{FE0F}",
                EmojiPresentation::Emoji,
                EmojiStructure::Single
            )]
        );
    }

    /// C5 — VS-15 demotes any emoji to text presentation.
    #[test]
    fn c5_vs15_demotes_to_text() {
        assert_eq!(
            classify("\u{260E}\u{FE0E}"),
            vec![emoji(
                "\u{260E}\u{FE0E}",
                EmojiPresentation::Text,
                EmojiStructure::Single
            )]
        );
    }

    /// C6 — modifier sequence: 👍 (U+1F44D) + medium-dark skin tone (U+1F3FE).
    #[test]
    fn c6_modifier_sequence() {
        assert_eq!(
            classify("\u{1F44D}\u{1F3FE}"),
            vec![emoji(
                "\u{1F44D}\u{1F3FE}",
                EmojiPresentation::Emoji,
                EmojiStructure::ModifierSequence {
                    base: '\u{1F44D}',
                    tone: SkinTone::MediumDark,
                },
            )]
        );
    }

    /// C7 — ZWJ sequence: family of three (man + ZWJ + woman + ZWJ + girl).
    #[test]
    fn c7_zwj_sequence() {
        let text = "\u{1F468}\u{200D}\u{1F469}\u{200D}\u{1F467}";
        assert_eq!(
            classify(text),
            vec![emoji(
                text,
                EmojiPresentation::Emoji,
                EmojiStructure::ZwjSequence
            )]
        );
    }

    /// C8 — Regional indicator pair (German flag).
    #[test]
    fn c8_regional_flag() {
        let text = "\u{1F1E9}\u{1F1EA}";
        assert_eq!(
            classify(text),
            vec![emoji(
                text,
                EmojiPresentation::Emoji,
                EmojiStructure::FlagSequence(FlagKind::Regional),
            )]
        );
    }

    /// C9 — Subdivision flag (Scotland: black flag + tag sequence).
    #[test]
    fn c9_subdivision_flag() {
        let text = "\u{1F3F4}\u{E0067}\u{E0062}\u{E0073}\u{E0063}\u{E0074}\u{E007F}";
        assert_eq!(
            classify(text),
            vec![emoji(
                text,
                EmojiPresentation::Emoji,
                EmojiStructure::FlagSequence(FlagKind::Subdivision),
            )]
        );
    }

    /// C10 — keycap: '1' + VS-16 + U+20E3.
    #[test]
    fn c10_keycap() {
        let text = "1\u{FE0F}\u{20E3}";
        assert_eq!(
            classify(text),
            vec![emoji(
                text,
                EmojiPresentation::Emoji,
                EmojiStructure::KeycapSequence { base: '1' },
            )]
        );
    }

    /// C11 — two adjacent emoji split into two clusters per UAX #29.
    #[test]
    fn c11_adjacent_emojis() {
        assert_eq!(
            classify("\u{1F4DE}\u{1F4E7}"),
            vec![
                emoji(
                    "\u{1F4DE}",
                    EmojiPresentation::Emoji,
                    EmojiStructure::Single
                ),
                emoji(
                    "\u{1F4E7}",
                    EmojiPresentation::Emoji,
                    EmojiStructure::Single
                ),
            ]
        );
    }

    /// C12 — empty string yields no clusters.
    #[test]
    fn c12_empty_string() {
        assert_eq!(classify(""), Vec::<InlineCluster>::new());
    }

    /// C13 — combining mark on a base letter forms one text grapheme cluster
    /// preceding an emoji cluster. Tests that the merging logic flushes the
    /// pending text span before emitting the emoji.
    #[test]
    fn c13_combining_mark_then_emoji() {
        let text = "a\u{0301}\u{1F4DE}";
        assert_eq!(
            classify(text),
            vec![
                InlineCluster::Text("a\u{0301}"),
                emoji(
                    "\u{1F4DE}",
                    EmojiPresentation::Emoji,
                    EmojiStructure::Single
                ),
            ]
        );
    }

    // ─── Negative / robustness tests ─────────────────────────────────────────

    /// Non-emoji ZWJ pair must not be misclassified as a ZWJ emoji sequence.
    #[test]
    fn negative_zwj_between_letters_is_text() {
        let text = "a\u{200D}b";
        // UAX #29 may bind these as one cluster; either way, no emoji char
        // is present so the result must be Text only.
        let clusters = classify(text);
        for cluster in &clusters {
            assert!(
                matches!(cluster, InlineCluster::Text(_)),
                "non-emoji ZWJ pairing must classify as text, got {cluster:?}"
            );
        }
    }

    /// Single Regional Indicator (no pair) is a degenerate input. We accept
    /// either Text or Single-emoji classification but must never panic.
    #[test]
    fn negative_lone_regional_indicator_does_not_panic() {
        let _ = classify("\u{1F1E9}");
    }
}
