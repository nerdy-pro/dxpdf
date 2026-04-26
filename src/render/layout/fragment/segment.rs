//! Cross-run grapheme cluster reassembly.
//!
//! See `docs/cross-run-cluster-reassembly.md` for the full design.
//!
//! Word splits a UAX #29 grapheme cluster across `<w:r>` runs whenever
//! consecutive characters land in different `<w:rFonts>` slots
//! (§17.3.2.26). The classic example is the keycap `1️⃣` — digit `1`
//! lands in the `ascii` slot, `VS-16` + `U+20E3` land in the `hAnsi`
//! slot, so Word emits three separate runs even though the user authored
//! one emoji.
//!
//! This module pre-passes a paragraph's inline list, joining consecutive
//! text-only runs into a flat string with per-character provenance, then
//! classifies the joined string per UAX #29 + UTS #51 (via the existing
//! `cluster` module). Each resulting [`SegmentPiece`] is either a single-
//! run text span or a cross-run emoji cluster ready for the existing
//! emit-fragments path.

use unicode_segmentation::UnicodeSegmentation;

use crate::model::{Inline, RunElement, TextRun};
use crate::render::emoji::cluster::{self, EmojiPresentation, EmojiStructure, InlineCluster};

// ─── Public ADTs ─────────────────────────────────────────────────────────────

/// A maximal contiguous unit within a paragraph's inline list, after
/// stripping run boundaries that don't carry semantic meaning.
///
/// Joining stops at:
/// - Any inline that isn't a `Inline::TextRun` (Image, Hyperlink, Pict,
///   FieldChar, InstrText, BookmarkStart/End, FootnoteRef, EndnoteRef,
///   Symbol, Separator, AlternateContent, …).
/// - Any `Inline::TextRun` whose `content` contains a `RunElement` other
///   than `RunElement::Text` (Tab, LineBreak, PageBreak, ColumnBreak,
///   LastRenderedPageBreak).
///
/// Both classes surface as `Discrete` so the existing dispatch in
/// `collect_fragments` handles them unchanged.
#[derive(Debug)]
pub(super) enum InlineUnit<'a> {
    /// One or more consecutive text-only `Inline::TextRun`s, joined into
    /// a flat string with per-character provenance back to the originating
    /// run.
    TextSegment(JoinedTextSegment<'a>),
    /// Anything else — a borrow into the original `Inline` tree.
    Discrete(&'a Inline),
}

/// A flat text segment with per-character provenance.
///
/// Invariants:
/// - `text.chars().count() == char_runs.len()`
/// - `text` is C0-control-stripped (XML §2.1) except for TAB.
/// - Every `char_runs[i]` is a stable reference into the parsed model.
#[derive(Debug)]
pub(super) struct JoinedTextSegment<'a> {
    text: String,
    /// One entry per `char` in `text`. Tracks which originating
    /// [`TextRun`] produced that scalar.
    char_runs: Vec<&'a TextRun>,
}

impl<'a> JoinedTextSegment<'a> {
    /// Joined-and-stripped text. Used by tests to verify provenance and
    /// by future paragraph-level diagnostics — kept as part of the public
    /// API of this module.
    #[allow(dead_code)]
    pub(super) fn text(&self) -> &str {
        &self.text
    }

    pub(super) fn char_runs(&self) -> &[&'a TextRun] {
        &self.char_runs
    }

    /// Classify this segment per UAX #29 + UTS #51, splitting at run
    /// boundaries within text spans and packaging emoji clusters
    /// (possibly cross-run) as a single piece using the cluster's first
    /// run as the base.
    pub(super) fn classify(self) -> Vec<SegmentPiece<'a>> {
        let mut out = Vec::new();
        let mut buffer: Option<TextBuffer<'a>> = None;
        let mut char_idx = 0usize;

        for ic in cluster::classify(&self.text) {
            match ic {
                InlineCluster::Text(span) => {
                    // Walk graphemes inside this text span so we can
                    // split when the originating run changes.
                    for grapheme in UnicodeSegmentation::graphemes(span, true) {
                        let chars_in = grapheme.chars().count();
                        let run = self.char_runs[char_idx];
                        match buffer.as_mut() {
                            None => {
                                buffer = Some(TextBuffer {
                                    run,
                                    text: grapheme.to_string(),
                                });
                            }
                            Some(buf) if std::ptr::eq(buf.run, run) => {
                                buf.text.push_str(grapheme);
                            }
                            Some(_) => {
                                // Run boundary inside a text span — flush.
                                let prev = buffer.take().unwrap();
                                out.push(SegmentPiece::Text {
                                    run: prev.run,
                                    text: prev.text,
                                });
                                buffer = Some(TextBuffer {
                                    run,
                                    text: grapheme.to_string(),
                                });
                            }
                        }
                        char_idx += chars_in;
                    }
                }
                InlineCluster::Emoji(ec) => {
                    // Flush any pending text.
                    if let Some(prev) = buffer.take() {
                        out.push(SegmentPiece::Text {
                            run: prev.run,
                            text: prev.text,
                        });
                    }
                    let base_run = self.char_runs[char_idx];
                    let chars_in = ec.text.chars().count();
                    out.push(SegmentPiece::Emoji {
                        base_run,
                        text: ec.text.to_string(),
                        presentation: ec.presentation,
                        structure: ec.structure,
                    });
                    char_idx += chars_in;
                }
            }
        }

        if let Some(last) = buffer {
            out.push(SegmentPiece::Text {
                run: last.run,
                text: last.text,
            });
        }

        out
    }
}

struct TextBuffer<'a> {
    run: &'a TextRun,
    text: String,
}

/// One classified piece of a [`JoinedTextSegment`] after applying UAX #29
/// + UTS #51.
///
/// `Text` is split at run boundaries so each piece carries a single
/// originating run; that run's properties drive the whole piece. `Emoji`
/// is not split — a UAX #29 cluster is one visual unit at one size, so
/// when it spans runs we use the base run's formatting (font-name hint
/// per §17.3.2.26, color, baseline offset, font size).
#[derive(Debug)]
pub(super) enum SegmentPiece<'a> {
    Text {
        run: &'a TextRun,
        text: String,
    },
    Emoji {
        base_run: &'a TextRun,
        text: String,
        presentation: EmojiPresentation,
        structure: EmojiStructure,
    },
}

// ─── Builder ─────────────────────────────────────────────────────────────────

/// Walk a paragraph's inline list and produce [`InlineUnit`]s with
/// consecutive text-only runs joined into [`JoinedTextSegment`]s.
pub(super) fn build_inline_units(inlines: &[Inline]) -> Vec<InlineUnit<'_>> {
    let mut out = Vec::new();
    let mut buf = SegmentBuilder::default();

    for inline in inlines {
        match inline {
            Inline::TextRun(tr) if is_text_only_run(tr) => {
                for el in &tr.content {
                    if let RunElement::Text(s) = el {
                        buf.append(s, tr);
                    }
                }
            }
            other => {
                if let Some(seg) = buf.take() {
                    out.push(InlineUnit::TextSegment(seg));
                }
                out.push(InlineUnit::Discrete(other));
            }
        }
    }

    if let Some(seg) = buf.take() {
        out.push(InlineUnit::TextSegment(seg));
    }
    out
}

fn is_text_only_run(tr: &TextRun) -> bool {
    !tr.content.is_empty() && tr.content.iter().all(|e| matches!(e, RunElement::Text(_)))
}

#[derive(Default)]
struct SegmentBuilder<'a> {
    text: String,
    char_runs: Vec<&'a TextRun>,
}

impl<'a> SegmentBuilder<'a> {
    fn append(&mut self, raw: &str, run: &'a TextRun) {
        // XML §2.1: strip C0 control chars (other than TAB) at segment
        // build time. Matches the legacy `emit_text_fragments` contract.
        for c in raw.chars() {
            if c.is_control() && c != '\t' {
                continue;
            }
            self.text.push(c);
            self.char_runs.push(run);
        }
    }

    fn take(&mut self) -> Option<JoinedTextSegment<'a>> {
        if self.text.is_empty() {
            return None;
        }
        Some(JoinedTextSegment {
            text: std::mem::take(&mut self.text),
            char_runs: std::mem::take(&mut self.char_runs),
        })
    }
}

// ─── Tests (Phase α.0 + α.1) ─────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{RevisionIds, RunProperties, Symbol};

    // ── Helpers ────────────────────────────────────────────────────────

    fn text_run(text: &str) -> Inline {
        Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties::default(),
            content: vec![RunElement::Text(text.into())],
            rsids: RevisionIds::default(),
        }))
    }

    fn run_with_elements(elements: Vec<RunElement>) -> Inline {
        Inline::TextRun(Box::new(TextRun {
            style_id: None,
            properties: RunProperties::default(),
            content: elements,
            rsids: RevisionIds::default(),
        }))
    }

    /// Any non-TextRun inline that breaks segment joining. We use Symbol
    /// (small, simple) as the canonical "interrupter" in tests; the real
    /// pipeline treats every non-TextRun inline the same way.
    fn discrete_inline() -> Inline {
        Inline::Symbol(Symbol {
            font: "Wingdings".into(),
            char_code: 0xF0FE,
        })
    }

    fn extract_run(inline: &Inline) -> &TextRun {
        match inline {
            Inline::TextRun(tr) => tr,
            _ => panic!("expected TextRun"),
        }
    }

    fn segment_of<'a>(unit: &'a InlineUnit<'a>) -> &'a JoinedTextSegment<'a> {
        match unit {
            InlineUnit::TextSegment(s) => s,
            other => panic!("expected TextSegment, got {other:?}"),
        }
    }

    // ── Phase α.0: build_inline_units ──────────────────────────────────

    /// A1 — single TextRun produces one segment whose char_runs all
    /// point to that run.
    #[test]
    fn a1_single_text_run() {
        let inlines = vec![text_run("hello")];
        let r1 = extract_run(&inlines[0]);

        let units = build_inline_units(&inlines);
        assert_eq!(units.len(), 1);
        let seg = segment_of(&units[0]);
        assert_eq!(seg.text(), "hello");
        assert_eq!(seg.char_runs().len(), 5);
        for &cr in seg.char_runs() {
            assert!(std::ptr::eq(cr, r1));
        }
    }

    /// A2 — two consecutive TextRuns join into one segment with
    /// per-character provenance interleaved.
    #[test]
    fn a2_two_runs_join() {
        let inlines = vec![text_run("ab"), text_run("cd")];
        let r1 = extract_run(&inlines[0]);
        let r2 = extract_run(&inlines[1]);

        let units = build_inline_units(&inlines);
        assert_eq!(units.len(), 1);
        let seg = segment_of(&units[0]);
        assert_eq!(seg.text(), "abcd");
        assert_eq!(seg.char_runs().len(), 4);
        assert!(std::ptr::eq(seg.char_runs()[0], r1));
        assert!(std::ptr::eq(seg.char_runs()[1], r1));
        assert!(std::ptr::eq(seg.char_runs()[2], r2));
        assert!(std::ptr::eq(seg.char_runs()[3], r2));
    }

    /// A3 — a non-TextRun inline (Symbol stands in for any of Image,
    /// Hyperlink, Pict, FootnoteRef, …) breaks segment joining.
    #[test]
    fn a3_non_text_inline_breaks_join() {
        let inlines = vec![text_run("ab"), discrete_inline(), text_run("cd")];
        let units = build_inline_units(&inlines);
        assert_eq!(units.len(), 3);
        assert_eq!(segment_of(&units[0]).text(), "ab");
        assert!(matches!(units[1], InlineUnit::Discrete(Inline::Symbol(_))));
        assert_eq!(segment_of(&units[2]).text(), "cd");
    }

    /// A4 — a TextRun whose content contains a non-text element (Tab)
    /// is treated as Discrete; subsequent text-only runs start a fresh
    /// segment.
    #[test]
    fn a4_text_run_with_tab_is_discrete() {
        let r1 = run_with_elements(vec![
            RunElement::Text("a".into()),
            RunElement::Tab,
            RunElement::Text("b".into()),
        ]);
        let inlines = vec![text_run("hello"), r1, text_run("world")];
        let units = build_inline_units(&inlines);
        assert_eq!(units.len(), 3);
        assert_eq!(segment_of(&units[0]).text(), "hello");
        assert!(matches!(units[1], InlineUnit::Discrete(Inline::TextRun(_))));
        assert_eq!(segment_of(&units[2]).text(), "world");
    }

    /// A5 — empty TextRun produces no segment.
    #[test]
    fn a5_empty_text_run_is_skipped() {
        let inlines = vec![text_run("")];
        let units = build_inline_units(&inlines);
        assert!(units.is_empty(), "no segment from empty text");
    }

    /// A5b — purely-empty TextRun (no content elements) is treated as
    /// not-text-only and surfaces as Discrete.
    #[test]
    fn a5b_no_content_text_run_is_discrete() {
        let inlines = vec![run_with_elements(vec![])];
        let units = build_inline_units(&inlines);
        assert_eq!(units.len(), 1);
        assert!(matches!(units[0], InlineUnit::Discrete(_)));
    }

    /// A6 — XML §2.1: C0 control chars (other than TAB) are stripped at
    /// segment-build time. char_runs.len() reflects the cleaned text.
    #[test]
    fn a6_control_chars_stripped() {
        let inlines = vec![text_run("\u{0001}hi\u{0002}")];
        let units = build_inline_units(&inlines);
        assert_eq!(units.len(), 1);
        let seg = segment_of(&units[0]);
        assert_eq!(seg.text(), "hi");
        assert_eq!(seg.char_runs().len(), 2);
    }

    /// A6b — TAB chars in text are PRESERVED (legacy contract; tabs in
    /// `RunElement::Text` are different from `RunElement::Tab`).
    #[test]
    fn a6b_tab_preserved_in_text() {
        let inlines = vec![text_run("a\tb")];
        let units = build_inline_units(&inlines);
        let seg = segment_of(&units[0]);
        assert_eq!(seg.text(), "a\tb");
    }

    /// A7 — three runs all with text → one segment with full
    /// invariant `text.chars().count() == char_runs.len()`.
    #[test]
    fn a7_three_runs_invariant() {
        let inlines = vec![text_run("ab"), text_run("c"), text_run("def")];
        let units = build_inline_units(&inlines);
        assert_eq!(units.len(), 1);
        let seg = segment_of(&units[0]);
        assert_eq!(seg.text().chars().count(), seg.char_runs().len());
        assert_eq!(seg.text(), "abcdef");
    }

    /// A8 — a Hyperlink Inline breaks the join (hyperlink content is its
    /// own scope handled via recursion at the emit site).
    #[test]
    fn a8_hyperlink_breaks_join() {
        use crate::model::{Hyperlink, HyperlinkTarget, RelId};
        let inlines = vec![
            text_run("before "),
            Inline::Hyperlink(Hyperlink {
                target: HyperlinkTarget::External(RelId::new("rId1")),
                content: vec![text_run("link")],
            }),
            text_run(" after"),
        ];
        let units = build_inline_units(&inlines);
        assert_eq!(units.len(), 3);
        assert_eq!(segment_of(&units[0]).text(), "before ");
        assert!(matches!(
            units[1],
            InlineUnit::Discrete(Inline::Hyperlink(_))
        ));
        assert_eq!(segment_of(&units[2]).text(), " after");
    }

    /// A9 — Inline::FieldChar breaks the join (so the field state
    /// machine sees each text segment as a distinct field-state zone).
    #[test]
    fn a9_field_char_breaks_join() {
        use crate::model::{FieldChar, FieldCharType};
        let inlines = vec![
            text_run("page "),
            Inline::FieldChar(FieldChar {
                field_char_type: FieldCharType::Begin,
                dirty: None,
                fld_lock: None,
            }),
            text_run("instr"),
        ];
        let units = build_inline_units(&inlines);
        assert_eq!(units.len(), 3);
        assert_eq!(segment_of(&units[0]).text(), "page ");
        assert!(matches!(
            units[1],
            InlineUnit::Discrete(Inline::FieldChar(_))
        ));
        assert_eq!(segment_of(&units[2]).text(), "instr");
    }

    // ── Phase α.1: JoinedTextSegment::classify ─────────────────────────

    /// B1 — single-run "hello" → one Text piece.
    #[test]
    fn b1_single_run_text() {
        let inlines = vec![text_run("hello")];
        let r1 = extract_run(&inlines[0]);

        let units = build_inline_units(&inlines);
        let seg = match units.into_iter().next().unwrap() {
            InlineUnit::TextSegment(s) => s,
            _ => panic!(),
        };
        let pieces = seg.classify();
        assert_eq!(pieces.len(), 1);
        match &pieces[0] {
            SegmentPiece::Text { run, text } => {
                assert!(std::ptr::eq(*run, r1));
                assert_eq!(text, "hello");
            }
            other => panic!("expected Text, got {other:?}"),
        }
    }

    /// B2 — two-run "ab" + "cd" → two Text pieces, one per run.
    #[test]
    fn b2_two_runs_split_text() {
        let inlines = vec![text_run("ab"), text_run("cd")];
        let r1 = extract_run(&inlines[0]);
        let r2 = extract_run(&inlines[1]);

        let units = build_inline_units(&inlines);
        let seg = match units.into_iter().next().unwrap() {
            InlineUnit::TextSegment(s) => s,
            _ => panic!(),
        };
        let pieces = seg.classify();
        assert_eq!(pieces.len(), 2);
        match &pieces[0] {
            SegmentPiece::Text { run, text } => {
                assert!(std::ptr::eq(*run, r1));
                assert_eq!(text, "ab");
            }
            _ => panic!(),
        }
        match &pieces[1] {
            SegmentPiece::Text { run, text } => {
                assert!(std::ptr::eq(*run, r2));
                assert_eq!(text, "cd");
            }
            _ => panic!(),
        }
    }

    /// B3 — keycap split as ["1", VS-16, U+20E3] → one Emoji piece.
    /// **Direct fix for the user-visible keycap bug.**
    #[test]
    fn b3_cross_run_keycap_reassembles() {
        let inlines = vec![text_run("1"), text_run("\u{FE0F}"), text_run("\u{20E3}")];
        let r1 = extract_run(&inlines[0]);

        let units = build_inline_units(&inlines);
        let seg = match units.into_iter().next().unwrap() {
            InlineUnit::TextSegment(s) => s,
            _ => panic!(),
        };
        let pieces = seg.classify();
        assert_eq!(pieces.len(), 1);
        match &pieces[0] {
            SegmentPiece::Emoji {
                base_run,
                text,
                presentation,
                structure,
            } => {
                assert!(std::ptr::eq(*base_run, r1));
                assert_eq!(text, "1\u{FE0F}\u{20E3}");
                assert_eq!(*presentation, EmojiPresentation::Emoji);
                assert!(matches!(
                    structure,
                    EmojiStructure::KeycapSequence { base: '1' }
                ));
            }
            other => panic!("expected Emoji KeycapSequence, got {other:?}"),
        }
    }

    /// B4 — ZWJ family split per emoji → one Emoji piece.
    #[test]
    fn b4_cross_run_zwj_family() {
        let inlines = vec![
            text_run("\u{1F468}"),
            text_run("\u{200D}"),
            text_run("\u{1F469}"),
            text_run("\u{200D}"),
            text_run("\u{1F467}"),
        ];
        let units = build_inline_units(&inlines);
        let seg = match units.into_iter().next().unwrap() {
            InlineUnit::TextSegment(s) => s,
            _ => panic!(),
        };
        let pieces = seg.classify();
        assert_eq!(pieces.len(), 1);
        match &pieces[0] {
            SegmentPiece::Emoji {
                text, structure, ..
            } => {
                assert_eq!(text, "\u{1F468}\u{200D}\u{1F469}\u{200D}\u{1F467}");
                assert!(matches!(structure, EmojiStructure::ZwjSequence));
            }
            _ => panic!(),
        }
    }

    /// B5 — modifier sequence split as [👍, 🏿] → one Emoji piece.
    #[test]
    fn b5_cross_run_modifier_sequence() {
        let inlines = vec![text_run("\u{1F44D}"), text_run("\u{1F3FF}")];
        let units = build_inline_units(&inlines);
        let seg = match units.into_iter().next().unwrap() {
            InlineUnit::TextSegment(s) => s,
            _ => panic!(),
        };
        let pieces = seg.classify();
        assert_eq!(pieces.len(), 1);
        match &pieces[0] {
            SegmentPiece::Emoji {
                text, structure, ..
            } => {
                assert_eq!(text, "\u{1F44D}\u{1F3FF}");
                assert!(matches!(
                    structure,
                    EmojiStructure::ModifierSequence {
                        base: '\u{1F44D}',
                        ..
                    }
                ));
            }
            _ => panic!(),
        }
    }

    /// B6 — mixed: ["hi 1", VS-16+U+20E3, " there"] → three pieces:
    /// Text("hi "), Emoji("1️⃣"), Text(" there").
    #[test]
    fn b6_mixed_text_and_keycap() {
        let inlines = vec![
            text_run("hi 1"),
            text_run("\u{FE0F}\u{20E3}"),
            text_run(" there"),
        ];
        let r1 = extract_run(&inlines[0]);
        let r3 = extract_run(&inlines[2]);

        let units = build_inline_units(&inlines);
        let seg = match units.into_iter().next().unwrap() {
            InlineUnit::TextSegment(s) => s,
            _ => panic!(),
        };
        let pieces = seg.classify();
        assert_eq!(pieces.len(), 3);
        match &pieces[0] {
            SegmentPiece::Text { run, text } => {
                assert!(std::ptr::eq(*run, r1));
                assert_eq!(text, "hi ");
            }
            _ => panic!(),
        }
        match &pieces[1] {
            SegmentPiece::Emoji {
                base_run,
                text,
                structure,
                ..
            } => {
                assert!(std::ptr::eq(*base_run, r1));
                assert_eq!(text, "1\u{FE0F}\u{20E3}");
                assert!(matches!(
                    structure,
                    EmojiStructure::KeycapSequence { base: '1' }
                ));
            }
            _ => panic!(),
        }
        match &pieces[2] {
            SegmentPiece::Text { run, text } => {
                assert!(std::ptr::eq(*run, r3));
                assert_eq!(text, " there");
            }
            _ => panic!(),
        }
    }

    /// B7 — two distinct emoji "📞📧" in two runs → two Emoji pieces
    /// (UAX #29 clusters: each emoji is its own grapheme).
    #[test]
    fn b7_adjacent_distinct_emojis() {
        let inlines = vec![text_run("\u{1F4DE}"), text_run("\u{1F4E7}")];
        let units = build_inline_units(&inlines);
        let seg = match units.into_iter().next().unwrap() {
            InlineUnit::TextSegment(s) => s,
            _ => panic!(),
        };
        let pieces = seg.classify();
        assert_eq!(pieces.len(), 2);
        for p in &pieces {
            assert!(matches!(p, SegmentPiece::Emoji { .. }));
        }
    }

    /// B8 — combining mark cross-run: 'a' + U+0301 → one Text piece
    /// (UAX #29 binds the combining mark).
    #[test]
    fn b8_cross_run_combining_mark() {
        let inlines = vec![text_run("a"), text_run("\u{0301}")];
        let r1 = extract_run(&inlines[0]);

        let units = build_inline_units(&inlines);
        let seg = match units.into_iter().next().unwrap() {
            InlineUnit::TextSegment(s) => s,
            _ => panic!(),
        };
        let pieces = seg.classify();
        assert_eq!(pieces.len(), 1);
        match &pieces[0] {
            SegmentPiece::Text { run, text } => {
                assert!(std::ptr::eq(*run, r1));
                assert_eq!(text, "a\u{0301}");
            }
            _ => panic!(),
        }
    }

    /// B9 — RIS pair split: ["🇩", "🇪"] → one Emoji FlagSequence.
    #[test]
    fn b9_cross_run_ris_pair() {
        let inlines = vec![text_run("\u{1F1E9}"), text_run("\u{1F1EA}")];
        let units = build_inline_units(&inlines);
        let seg = match units.into_iter().next().unwrap() {
            InlineUnit::TextSegment(s) => s,
            _ => panic!(),
        };
        let pieces = seg.classify();
        assert_eq!(pieces.len(), 1);
        match &pieces[0] {
            SegmentPiece::Emoji {
                text, structure, ..
            } => {
                assert_eq!(text, "\u{1F1E9}\u{1F1EA}");
                assert!(matches!(
                    structure,
                    EmojiStructure::FlagSequence(crate::render::emoji::cluster::FlagKind::Regional)
                ));
            }
            _ => panic!(),
        }
    }

    /// B10 — default-text codepoints (digits) NOT promoted: regression
    /// for the digit fix. Single-run digits stay text.
    #[test]
    fn b10_digits_in_one_run_stay_text() {
        let inlines = vec![text_run("Numbers: 1, 2, 3")];
        let units = build_inline_units(&inlines);
        let seg = match units.into_iter().next().unwrap() {
            InlineUnit::TextSegment(s) => s,
            _ => panic!(),
        };
        let pieces = seg.classify();
        assert_eq!(pieces.len(), 1);
        assert!(matches!(pieces[0], SegmentPiece::Text { .. }));
    }

    /// B11 — field char between text-only runs has already produced a
    /// segment break at A9; classification of each segment is independent.
    #[test]
    fn b11_field_char_isolates_segments() {
        use crate::model::{FieldChar, FieldCharType};
        let inlines = vec![
            text_run("a"),
            text_run("b"),
            Inline::FieldChar(FieldChar {
                field_char_type: FieldCharType::Begin,
                dirty: None,
                fld_lock: None,
            }),
            text_run("c"),
            text_run("d"),
        ];
        let units = build_inline_units(&inlines);
        assert_eq!(units.len(), 3);
        // First segment "ab" classifies independently.
        let seg1 = match &units[0] {
            InlineUnit::TextSegment(s) => s,
            _ => panic!(),
        };
        assert_eq!(seg1.text(), "ab");
        // Second segment "cd" classifies independently.
        let seg2 = match &units[2] {
            InlineUnit::TextSegment(s) => s,
            _ => panic!(),
        };
        assert_eq!(seg2.text(), "cd");
    }
}
