//! §17.15.1.10 automatic hyphenation, on real laid-out pages (issue #158).
//!
//! The unit tests in `src/i18n/hyphenate.rs` prove the vendored patterns cut
//! a word where they should, and the ones in `src/render/layout/line.rs`
//! prove the fitter obeys the §17.15.1.53 zone, the §17.15.1.22 limit and
//! §17.15.1.27 against a table of cuts. Neither says the four `w:settings`
//! elements, the run's `w:lang`, the paragraph's §17.3.1.34 exemption and
//! the real measurer are wired to each other — which is what this file
//! asserts, over `test-files/auto-hyphenation-*.docx`
//! (`scripts/make_auto_hyphenation_fixtures.py`).
//!
//! Every assertion is a *property* of the rendered lines rather than a
//! transcript of them, so a change of host font moves no goalpost: an
//! inserted hyphen is a `-` the engine painted that the paragraph's own text
//! does not contain (the fixtures contain none), and it is read back from
//! the `DrawCommand::Text` stream grouped by baseline, the way
//! `tests/line_breaking.rs` reads breaks.

use std::collections::BTreeMap;

use dxpdf::i18n::hyphenate::{opportunities, HyphenLanguage};
use dxpdf::model::{Block, Inline, RunElement};
use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};

const BASE: &str = "test-files/auto-hyphenation-base.docx";
const ZONE_WIDE: &str = "test-files/auto-hyphenation-zone-wide.docx";
const LIMIT_1: &str = "test-files/auto-hyphenation-limit-1.docx";
const NOCAPS: &str = "test-files/auto-hyphenation-nocaps.docx";
const OFF: &str = "test-files/auto-hyphenation-off.docx";
const COLUMNS: &str = "test-files/auto-hyphenation-columns.docx";
/// Not one of the set: the reporter's document from issue #90, which also
/// declares `<w:autoHyphenation w:val="true"/>` — and, in `docDefaults`,
/// `<w:suppressAutoHyphens w:val="true"/>`, which §17.7.5.2 cascades onto
/// every paragraph. The §17.3.1.34 control at the style-hierarchy level.
const OUTLINE_CONTROL: &str = "test-files/document_outline.docx";

/// One body paragraph as the parse saw it: its text, and the text of the
/// runs tagged `en-US` inside it.
struct Paragraph {
    text: String,
    english: Vec<String>,
}

/// Parse a fixture and lay it out. The paragraph texts come from the parse,
/// not from this file, so the generator stays the single source of truth.
fn fixture(path: &str) -> (Vec<Paragraph>, Vec<LayoutedPage>) {
    let bytes = std::fs::read(path).unwrap_or_else(|e| panic!("{path}: {e}"));
    let doc = dxpdf::docx::parse(&bytes).unwrap_or_else(|e| panic!("{path} parses: {e}"));
    let paragraphs = doc
        .body
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(p) => {
                let mut text = String::new();
                let mut english = Vec::new();
                for inline in &p.content {
                    if let Inline::TextRun(run) = inline {
                        let run_text: String = run
                            .content
                            .iter()
                            .filter_map(|el| match el {
                                RunElement::Text(t) => Some(t.as_str()),
                                _ => None,
                            })
                            .collect();
                        let tag = run.properties.lang.get().and_then(|l| l.val.as_deref());
                        if tag == Some("en-US") {
                            english.push(run_text.clone());
                        }
                        text.push_str(&run_text);
                    }
                }
                Some(Paragraph { text, english })
            }
            _ => None,
        })
        .filter(|p| !p.text.is_empty())
        .collect();
    (paragraphs, dxpdf::render::resolve_and_layout(doc).1)
}

/// The paragraph text of every block, tables included, as one string.
fn all_text(blocks: &[Block], out: &mut String) {
    for block in blocks {
        match block {
            Block::Paragraph(p) => {
                for inline in &p.content {
                    if let Inline::TextRun(run) = inline {
                        for el in &run.content {
                            if let RunElement::Text(t) = el {
                                out.push_str(t);
                            }
                        }
                    }
                }
                out.push('\n');
            }
            Block::Table(t) => {
                for row in &t.rows {
                    for cell in &row.cells {
                        all_text(&cell.content, out);
                    }
                }
            }
            Block::SectionBreak(_) => {}
        }
    }
}

/// Every laid-out line as its text pieces in x order. A line is the set of
/// text commands on one page sharing a baseline — and, when `column_split`
/// is given, on one side of that x, because two columns share baselines.
fn lines(pages: &[LayoutedPage], column_split: Option<f32>) -> Vec<Vec<String>> {
    let mut out = Vec::new();
    for page in pages {
        let mut by_line: BTreeMap<(bool, i64), Vec<(i64, String)>> = BTreeMap::new();
        for command in &page.commands {
            if let DrawCommand::Text { position, text, .. } = command {
                let x = position.x.raw();
                let right_column = column_split.is_some_and(|split| x >= split);
                by_line
                    .entry((right_column, position.y.raw().round() as i64))
                    .or_default()
                    .push((x.round() as i64, text.to_string()));
            }
        }
        for (_, mut pieces) in by_line {
            pieces.sort_by_key(|(x, _)| *x);
            let line: Vec<String> = pieces.into_iter().map(|(_, t)| t).collect();
            if !line.concat().trim().is_empty() {
                out.push(line);
            }
        }
    }
    out
}

/// Group `lines` by paragraph, consuming lines until their text — with the
/// inserted hyphens removed — is the paragraph's own. An assertion in
/// itself: if layout dropped, duplicated or mis-hyphenated anything, no
/// grouping exists and this panics.
fn by_paragraph(lines: &[Vec<String>], paragraphs: &[Paragraph]) -> Vec<Vec<Vec<String>>> {
    let mut out = Vec::new();
    let mut next = 0;
    for paragraph in paragraphs {
        assert!(
            !paragraph.text.contains('-'),
            "the fixture must contain no authored hyphen: {:?}",
            paragraph.text
        );
        let mut taken: Vec<Vec<String>> = Vec::new();
        let rebuilt = |taken: &[Vec<String>]| -> String {
            taken
                .iter()
                .map(|line| line.concat().replace('-', ""))
                .collect::<String>()
        };
        while rebuilt(&taken).chars().count() < paragraph.text.chars().count() {
            let line = lines.get(next).unwrap_or_else(|| {
                panic!(
                    "ran out of lines rebuilding {:?}; got {:?}",
                    paragraph.text, taken
                )
            });
            taken.push(line.clone());
            next += 1;
        }
        assert_eq!(
            rebuilt(&taken),
            paragraph.text,
            "a paragraph's lines, hyphens removed, are exactly its text"
        );
        out.push(taken);
    }
    assert_eq!(next, lines.len(), "every line belongs to a paragraph");
    out
}

/// The head piece of a line that ends in an inserted hyphen, without it.
fn hyphenated_head(line: &[String]) -> Option<&str> {
    line.last()?.strip_suffix('-')
}

/// The run of letters inside a token, as `layout::line::hyphenation_core`
/// defines a word.
fn letters(token: &str) -> &str {
    token.trim_matches(|c: char| !c.is_alphabetic())
}

/// Every word a paragraph's lines cut, as `(word, cut byte offset, lang)`.
fn cuts(paragraph: &Paragraph, lines: &[Vec<String>]) -> Vec<(String, usize, HyphenLanguage)> {
    let mut out = Vec::new();
    for (i, line) in lines.iter().enumerate() {
        let Some(head) = hyphenated_head(line) else {
            continue;
        };
        let tail = lines
            .get(i + 1)
            .and_then(|next| next.first())
            .unwrap_or_else(|| panic!("a hyphenated line has a line after it: {line:?}"));
        let tail_token = tail.split_whitespace().next().unwrap_or("");
        let head_letters = letters(head);
        let word = format!("{head_letters}{}", letters(tail_token));
        assert!(
            paragraph.text.contains(&word),
            "{head}- | {tail} does not rejoin to a word of the paragraph"
        );
        let lang = if paragraph.english.iter().any(|run| run.contains(&word)) {
            HyphenLanguage::EnUs
        } else {
            HyphenLanguage::De1996
        };
        out.push((word, head_letters.len(), lang));
    }
    out
}

fn hyphenated_line_count(lines: &[Vec<String>]) -> usize {
    lines
        .iter()
        .filter(|l| hyphenated_head(l).is_some())
        .count()
}

// ── §17.15.1.10: the setting hyphenates, at the run's language's seams ────

/// With `<w:autoHyphenation/>` the German prose is cut, and every cut is at
/// a point the `de-1996` patterns allow — while the words of the run tagged
/// `en-US` inside the fourth paragraph are cut at `en-us` points. The two
/// languages come from the run's cascaded `w:lang`, not the paragraph's.
#[test]
fn the_setting_cuts_words_at_the_pattern_seams_of_the_runs_language() {
    let (paragraphs, pages) = fixture(BASE);
    let grouped = by_paragraph(&lines(&pages, None), &paragraphs);

    let mut german = 0;
    let mut english = 0;
    for (paragraph, lines) in paragraphs.iter().zip(&grouped) {
        for (word, cut, lang) in cuts(paragraph, lines) {
            let allowed = opportunities(&word, lang);
            assert!(
                allowed.contains(&cut),
                "{word:?} was cut at byte {cut} ({:?}|{:?}), which the {lang:?} patterns \
                 do not allow ({allowed:?})",
                &word[..cut],
                &word[cut..]
            );
            match lang {
                HyphenLanguage::EnUs => english += 1,
                _ => german += 1,
            }
        }
    }
    assert!(
        german >= 5,
        "the German prose should be cut several times, got {german}"
    );
    assert!(
        english >= 1,
        "a word of the en-US run should be cut, got {english}"
    );
}

/// Hyphenation is a line-end event: the head is the last thing painted on
/// its line, never followed by its own tail or by another word.
#[test]
fn an_inserted_hyphen_is_only_ever_the_last_piece_of_its_line() {
    for path in [BASE, LIMIT_1, NOCAPS, COLUMNS] {
        let split = if path == COLUMNS { Some(240.0) } else { None };
        let (_, pages) = fixture(path);
        for line in lines(&pages, split) {
            for (idx, piece) in line.iter().enumerate() {
                assert!(
                    !piece.ends_with('-') || idx + 1 == line.len(),
                    "{path}: an inserted hyphen inside a line — {line:?}"
                );
            }
        }
    }
}

// ── §17.3.1.34: the paragraph's exemption ─────────────────────────────────

/// The second paragraph is the first with `<w:suppressAutoHyphens/>`; "the
/// paragraph shall be exempted from any hyphenation" — and the first one,
/// right above it in the same document, is the control that the document
/// setting is on.
#[test]
fn suppress_auto_hyphens_exempts_the_paragraph() {
    let (paragraphs, pages) = fixture(BASE);
    let grouped = by_paragraph(&lines(&pages, None), &paragraphs);
    assert!(
        hyphenated_line_count(&grouped[0]) >= 3,
        "the control hyphenates"
    );
    assert_eq!(
        hyphenated_line_count(&grouped[1]),
        0,
        "the exempted twin is not hyphenated: {:?}",
        grouped[1]
    );
    assert!(
        grouped[1].len() > grouped[0].len(),
        "without hyphenation the same text needs more lines"
    );
}

/// §17.7.5.2: `w:suppressAutoHyphens` in `docDefaults` reaches every
/// paragraph through the cascade, so a document that turns hyphenation on
/// and exempts everything at the root hyphenates nothing — every `-` it
/// paints is one its text already had.
#[test]
fn a_docdefaults_exemption_reaches_every_paragraph() {
    let bytes = std::fs::read(OUTLINE_CONTROL).unwrap();
    let doc = dxpdf::docx::parse(&bytes).unwrap();
    assert!(
        doc.settings.hyphenation.enabled,
        "the control declares the setting"
    );
    assert_eq!(doc.settings.hyphenation.zone.raw(), 0, "and a zone of 0");
    assert_eq!(
        doc.styles.doc_defaults_paragraph.suppress_auto_hyphens,
        Some(true),
        "and exempts everything in docDefaults"
    );
    let mut authored = String::new();
    all_text(&doc.body, &mut authored);
    let pages = dxpdf::render::resolve_and_layout(doc).1;
    for page in &pages {
        for command in &page.commands {
            if let DrawCommand::Text { text, .. } = command {
                if text.ends_with('-') {
                    assert!(
                        authored.contains(&**text),
                        "{text:?} was painted but is not in the document's text"
                    );
                }
            }
        }
    }
}

// ── §17.15.1.53: the zone ─────────────────────────────────────────────────

/// A 7200-twip zone is wider than the 2 in column, so no line can ever leave
/// a gap that exceeds it, and "the amount of whitespace which can be left at
/// the end of a line" is all of it: nothing is hyphenated.
#[test]
fn a_zone_wider_than_the_column_hyphenates_nothing() {
    let (paragraphs, pages) = fixture(ZONE_WIDE);
    let grouped = by_paragraph(&lines(&pages, None), &paragraphs);
    for lines in &grouped {
        assert_eq!(hyphenated_line_count(lines), 0, "{lines:?}");
    }
}

// ── §17.15.1.22: the consecutive limit ────────────────────────────────────

/// With `w:consecutiveHyphenLimit w:val="1"` "the following line shall not be
/// hyphenated regardless of whether or not it meets the criteria", so no two
/// consecutive lines of a paragraph end in an inserted hyphen — at zone 0,
/// where without the limit the same prose ends line after line in one.
#[test]
fn the_consecutive_limit_holds_across_every_paragraph() {
    let (paragraphs, pages) = fixture(LIMIT_1);
    let grouped = by_paragraph(&lines(&pages, None), &paragraphs);
    let mut total = 0;
    for lines in &grouped {
        for pair in lines.windows(2) {
            assert!(
                !(hyphenated_head(&pair[0]).is_some() && hyphenated_head(&pair[1]).is_some()),
                "two consecutive hyphenated lines: {pair:?}"
            );
        }
        total += hyphenated_line_count(lines);
    }
    assert!(
        total >= 3,
        "the limit is exercised, not vacuous: {total} cuts"
    );

    // The control: the same document at zone 0 without the limit does
    // produce consecutive hyphenated lines (the nocaps fixture is that).
    let (paragraphs, pages) = fixture(NOCAPS);
    let grouped = by_paragraph(&lines(&pages, None), &paragraphs);
    assert!(
        grouped.iter().any(|lines| {
            lines
                .windows(2)
                .any(|p| hyphenated_head(&p[0]).is_some() && hyphenated_head(&p[1]).is_some())
        }),
        "without a limit the same prose does end consecutive lines in hyphens"
    );
}

// ── §17.15.1.27: capitals ─────────────────────────────────────────────────

fn is_all_caps(head: &str) -> bool {
    let word = letters(head);
    !word.is_empty() && word.chars().all(char::is_uppercase)
}

/// Under `w:doNotHyphenateCaps` every `KRAFTFAHRZEUG` moves down whole while
/// the `Kraftfahrzeug` beside it is still cut; "If omitted, words in CAPITAL
/// LETTERS shall be hyphenated" — the base document, same paragraph, is the
/// control and does cut one.
#[test]
fn capitals_stay_whole_under_do_not_hyphenate_caps() {
    let (paragraphs, pages) = fixture(NOCAPS);
    let grouped = by_paragraph(&lines(&pages, None), &paragraphs);
    let caps_paragraph = &grouped[2];
    let heads: Vec<&str> = caps_paragraph
        .iter()
        .filter_map(|l| hyphenated_head(l))
        .collect();
    assert!(
        heads.iter().all(|h| !is_all_caps(h)),
        "a capital word was cut: {heads:?}"
    );
    assert!(
        heads.iter().any(|h| letters(h).starts_with("Kraft")),
        "the lowercase compound beside it is still cut: {heads:?}"
    );

    let (paragraphs, pages) = fixture(BASE);
    let grouped = by_paragraph(&lines(&pages, None), &paragraphs);
    let heads: Vec<&str> = grouped[2]
        .iter()
        .filter_map(|l| hyphenated_head(l))
        .collect();
    assert!(
        heads.iter().any(|h| is_all_caps(h)),
        "with the element omitted a capital word is cut: {heads:?}"
    );
}

// ── §17.15.1.10 off ───────────────────────────────────────────────────────

/// `<w:autoHyphenation w:val="0"/>` — the document says no.
#[test]
fn the_document_setting_off_hyphenates_nothing() {
    let (paragraphs, pages) = fixture(OFF);
    let grouped = by_paragraph(&lines(&pages, None), &paragraphs);
    for lines in &grouped {
        assert_eq!(hyphenated_line_count(lines), 0, "{lines:?}");
    }
}

// ── the continuation re-fit ───────────────────────────────────────────────

/// One paragraph flows from a 2 in column (x from 72 pt) into a 4 in one
/// (x from 252 pt). Words cut for the narrow column are re-fitted whole in
/// the wide one: no line anywhere carries an inserted hyphen before its last
/// piece, the narrow column does hyphenate, and the paragraph does reach the
/// wide column — the three facts that together say the remainder was
/// re-fitted from the paragraph's words and not from this placement's cut
/// fragments.
#[test]
fn a_word_cut_for_the_narrow_column_is_whole_again_in_the_wide_one() {
    let (paragraphs, pages) = fixture(COLUMNS);
    let lines = lines(&pages, Some(240.0));
    let (narrow, wide): (Vec<_>, Vec<_>) = pages
        .iter()
        .flat_map(|p| p.commands.iter())
        .filter_map(|c| match c {
            DrawCommand::Text { position, .. } => Some(position.x.raw()),
            _ => None,
        })
        .partition(|x| *x < 240.0);
    assert!(
        !narrow.is_empty() && !wide.is_empty(),
        "both columns hold text"
    );
    assert!(
        hyphenated_line_count(&lines) >= 3,
        "the narrow column hyphenates"
    );
    for line in &lines {
        for (idx, piece) in line.iter().enumerate() {
            assert!(
                !piece.ends_with('-') || idx + 1 == line.len(),
                "an inserted hyphen inside a line — {line:?}"
            );
        }
    }
    // And every cut, in either column, is at a German pattern seam.
    let paragraph = &paragraphs[0];
    for (word, cut, lang) in cuts(paragraph, &lines) {
        assert!(
            opportunities(&word, lang).contains(&cut),
            "{word:?} cut at {cut} is not a {lang:?} pattern point"
        );
    }
}
