//! Numbering resolution — flatten abstract + instance + overrides into lookup table.

use std::collections::HashMap;

use crate::model::{
    AbstractNumId, Alignment, Indentation, LevelSuffix, NumId, NumPicBulletId, NumberFormat,
    NumberingDefinitions, NumberingLevelDefinition, RunProperties,
};
use crate::render::resolve::counting;
use crate::render::resolve::locale::Locale;
use crate::render::resolve::spellout;

/// §17.9.2 / §17.9.15: an abstract numbering definition with the level
/// overrides of one `w:num` instance applied.
///
/// The abstract id travels with the levels because the *counter* belongs to the
/// abstract definition, not to the instance: several `w:num` entries may point
/// at one `w:abstractNum`, and Word runs one sequence across all of them
/// (issue #230). Layout keys `BuildState::list_counters` on
/// `abstract_num_id`, so it has to arrive here.
#[derive(Clone, Debug)]
pub struct ResolvedNumbering {
    pub abstract_num_id: AbstractNumId,
    pub levels: Vec<ResolvedNumberingLevel>,
}

/// A resolved numbering level — ready for label generation.
#[derive(Clone, Debug)]
pub struct ResolvedNumberingLevel {
    pub format: NumberFormat,
    pub level_text: String,
    pub start: u32,
    /// §17.9.28: `w:startOverride` on this instance's level — a one-shot
    /// restart of the shared sequence, fired when the instance is first used,
    /// and *not* the level's start value. Collapsing the two erases the
    /// difference: a later item on the same instance would restart again, and
    /// the other instances of the abstract definition would never see it.
    pub start_override: Option<u32>,
    /// §17.9.3: run properties for the numbering symbol (font, color, etc.).
    pub run_properties: Option<RunProperties>,
    /// §17.9.3: paragraph indentation from the numbering level definition.
    /// A fallback only: any indentation the paragraph already carries — from
    /// its style cascade or direct formatting — overrides this field-by-field
    /// (`build::list_label::effective_indentation`), matching Word rather
    /// than a reading where the numbering level wins.
    pub indentation: Option<Indentation>,
    /// §17.9.7: justification of the numbering symbol (left, center, right).
    pub justification: Option<Alignment>,
    /// §17.9.10: reference to a picture bullet definition.
    pub lvl_pic_bullet_id: Option<NumPicBulletId>,
    /// §17.9.29: separator between the label and the paragraph text.
    pub suffix: LevelSuffix,
    /// §17.9.8: render all level numbers as decimal (legal numbering).
    pub is_legal: bool,
}

/// Resolve numbering definitions into a flat lookup: `NumId` →
/// `Vec<ResolvedNumberingLevel>`.
/// Each instance's abstract definition is looked up and level overrides applied.
pub fn resolve_numbering(defs: &NumberingDefinitions) -> HashMap<NumId, ResolvedNumbering> {
    let mut result = HashMap::new();

    for (num_id, instance) in &defs.numbering_instances {
        let abstract_levels = defs
            .abstract_nums
            .get(&instance.abstract_num_id)
            .map(|a| a.levels.as_slice())
            .unwrap_or(&[]);

        let mut levels: Vec<ResolvedNumberingLevel> =
            abstract_levels.iter().map(resolve_level).collect();

        // Apply level overrides (§17.9.9). A `<w:lvlOverride>` may supply a full
        // replacement `<w:lvl>` and/or a `<w:startOverride>` that restarts the
        // level's counter.
        for ovr in &instance.level_overrides {
            let idx = ovr.level as usize;
            if idx >= levels.len() {
                continue; // override references a level beyond the abstract def
            }
            if let Some(def) = &ovr.definition {
                levels[idx] = resolve_level(def);
            }
            if let Some(start) = ovr.start_override {
                levels[idx].start_override = Some(start);
            }
        }

        result.insert(
            *num_id,
            ResolvedNumbering {
                abstract_num_id: instance.abstract_num_id,
                levels,
            },
        );
    }

    result
}

fn resolve_level(def: &NumberingLevelDefinition) -> ResolvedNumberingLevel {
    ResolvedNumberingLevel {
        format: def.format.unwrap_or(NumberFormat::None),
        level_text: def.level_text.clone(),
        start_override: None,
        start: def.start.unwrap_or(1),
        run_properties: def.run_properties.clone(),
        indentation: def.indentation,
        justification: def.justification,
        lvl_pic_bullet_id: def.lvl_pic_bullet_id,
        suffix: def.suffix,
        is_legal: def.is_legal,
    }
}

/// §17.9.11: format a list label by expanding the level_text template.
/// `%1` is replaced with the formatted counter for level 0, `%2` for level 1, etc.
/// Returns `None` for `NumberFormat::None`.
pub fn format_list_label(
    levels: &[ResolvedNumberingLevel],
    level: u8,
    counters: &HashMap<(AbstractNumId, u8), u32>,
    abstract_num_id: AbstractNumId,
    locale: Locale,
) -> Option<String> {
    let lvl = levels.get(level as usize)?;
    if lvl.format == NumberFormat::None {
        return None;
    }
    if lvl.format == NumberFormat::Bullet {
        return Some(lvl.level_text.clone());
    }

    // Expand template: %1 → level 0 counter, %2 → level 1 counter, etc.
    // §17.9.8: when this level is "legal", every referenced counter is rendered
    // as decimal regardless of the individual levels' own formats.
    let mut result = lvl.level_text.clone();
    for i in (0..=level).rev() {
        // Widen before adding: `i` is a raw `w:ilvl` u8, so a crafted
        // ilvl=255 would overflow `i + 1` (§17.9.9 placeholders are 1-based).
        let placeholder = format!("%{}", u32::from(i) + 1);
        if result.contains(&placeholder) {
            let count = counters.get(&(abstract_num_id, i)).copied().unwrap_or(1);
            let fmt = if lvl.is_legal {
                NumberFormat::Decimal
            } else {
                levels
                    .get(i as usize)
                    .map(|l| l.format)
                    .unwrap_or(NumberFormat::Decimal)
            };
            let formatted = format_number(count, fmt, locale);
            result = result.replace(&placeholder, &formatted);
        }
    }
    Some(result)
}

/// §17.18.59 `ST_NumberFormat`: render one counter.
///
/// Total over `NumberFormat` — the `_ =>` arm this replaced answered
/// `n.to_string()` for every format it did not implement, which is right for
/// none of them: it printed a digit where `none` asks for nothing, and it hid
/// `cardinalText` and `ordinalText` behind an answer that looked deliberate.
///
/// `locale` decides only the three language-dependent formats; the rest are
/// the same in every language, which is why they take it without using it.
///
/// Those three used to be spelled here, in English only. They now live whole
/// in [`super::spellout`], which is also where issue #132's decision is
/// recorded — why the words are hand-written rather than taken from CLDR, and
/// what the one crate that would have supplied them cost when it was measured.
fn format_number(n: u32, fmt: NumberFormat, locale: Locale) -> String {
    match fmt {
        // §17.18.59 `decimalHalfWidth` *is* decimal: "half-width Arabic
        // numerals" names the ASCII digits, which is what `decimal` already
        // writes. It is a separate spec value because `decimalFullWidth` is
        // the one it contrasts with, not because it renders differently.
        NumberFormat::Decimal | NumberFormat::DecimalHalfWidth => n.to_string(),
        NumberFormat::LowerLetter => to_letter_lower(n),
        NumberFormat::UpperLetter => to_letter_upper(n),
        NumberFormat::LowerRoman => to_roman_lower(n),
        NumberFormat::UpperRoman => to_roman_upper(n),
        // Same in every locale: the Cyrillic sequence *is* the format.
        NumberFormat::RussianLower => to_russian_lower(n),
        NumberFormat::RussianUpper => to_russian_upper(n),

        // ── §17.18.59 digit substitution ──────────────────────────────────
        // Positional decimal, ten other characters. `decimalFullWidth2` is
        // Word's second full-width value and writes the same ten as the
        // first; the pair differs in font selection, not in numbering.
        NumberFormat::DecimalFullWidth | NumberFormat::DecimalFullWidth2 => {
            to_digit_set(n, &FULL_WIDTH_DIGITS)
        }
        NumberFormat::HindiNumbers => to_digit_set(n, &DEVANAGARI_DIGITS),
        NumberFormat::ThaiNumbers => to_digit_set(n, &THAI_DIGITS),
        NumberFormat::IdeographDigital => to_digit_set(n, &IDEOGRAPH_DIGITS),

        // ── §17.18.59 decorated decimal ───────────────────────────────────
        NumberFormat::DecimalZero => format!("{n:02}"),
        NumberFormat::Hex => format!("{n:X}"),
        NumberFormat::NumberInDash => format!("-{n}-"),
        // `decimalEnclosedCircleChinese` is the same U+2460 series as
        // `decimalEnclosedCircle`; Word tells them apart by the font it draws
        // them with, which is not something a counter can carry.
        NumberFormat::DecimalEnclosedCircle | NumberFormat::DecimalEnclosedCircleChinese => {
            to_enclosed(n, '\u{2460}', 20)
        }
        NumberFormat::DecimalEnclosedParen => to_enclosed(n, '\u{2474}', 20),
        NumberFormat::DecimalEnclosedFullstop => to_enclosed(n, '\u{2488}', 20),
        NumberFormat::IdeographEnclosedCircle => to_enclosed(n, '\u{3280}', 10),

        // ── §17.18.59 fixed alphabets ─────────────────────────────────────
        NumberFormat::Aiueo => from_alphabet(n, &KATAKANA_GOJUON_HALF),
        NumberFormat::AiueoFullWidth => from_alphabet(n, &KATAKANA_GOJUON),
        NumberFormat::Iroha => from_alphabet(n, &KATAKANA_IROHA_HALF),
        NumberFormat::IrohaFullWidth => from_alphabet(n, &KATAKANA_IROHA),
        NumberFormat::Ganada => from_alphabet(n, &HANGUL_GANADA),
        NumberFormat::Chosung => from_alphabet(n, &HANGUL_CHOSUNG),
        NumberFormat::Hebrew2 => from_alphabet(n, &HEBREW_ALPHABET),
        NumberFormat::ArabicAlpha => from_alphabet(n, &ARABIC_ALPHABET),
        NumberFormat::HindiVowels => from_alphabet(n, &DEVANAGARI_VOWELS),
        NumberFormat::HindiConsonants => from_alphabet(n, &DEVANAGARI_CONSONANTS),
        NumberFormat::ThaiLetters => from_alphabet(n, &THAI_CONSONANTS),
        NumberFormat::Chicago => from_alphabet(n, &CHICAGO_SYMBOLS),
        NumberFormat::IdeographTraditional => from_alphabet(n, &HEAVENLY_STEMS),
        NumberFormat::IdeographZodiac => from_alphabet(n, &EARTHLY_BRANCHES),
        NumberFormat::IdeographZodiacTraditional => to_sexagenary(n),

        // ── §17.18.59 additive numerals ───────────────────────────────────
        NumberFormat::Hebrew1 => to_hebrew_numeral(n),
        NumberFormat::ArabicAbjad => to_additive(n, &ABJAD_NUMERALS),

        // ── §17.18.59 counting systems (issue #152) ───────────────────────
        // The number read aloud in a language's own numerals. The format
        // names the language, so — unlike the three §17.9.27 formats above —
        // `locale` is not consulted. All rules live in `super::counting`.
        NumberFormat::ChineseCounting => counting::chinese_counting(n),
        NumberFormat::ChineseCountingThousand => counting::chinese_counting_thousand(n),
        NumberFormat::ChineseLegalSimplified => counting::chinese_legal_simplified(n),
        NumberFormat::TaiwaneseCounting => counting::taiwanese_counting(n),
        NumberFormat::TaiwaneseCountingThousand => counting::taiwanese_counting_thousand(n),
        NumberFormat::TaiwaneseDigital => counting::taiwanese_digital(n),
        NumberFormat::IdeographLegalTraditional => counting::ideograph_legal_traditional(n),
        NumberFormat::JapaneseCounting => counting::japanese_counting(n),
        NumberFormat::JapaneseLegal => counting::japanese_legal(n),
        NumberFormat::JapaneseDigitalTenThousand => counting::japanese_digital_ten_thousand(n),
        NumberFormat::KoreanCounting => counting::korean_counting(n),
        NumberFormat::KoreanLegal => counting::korean_legal(n),
        NumberFormat::KoreanDigital => counting::korean_digital(n),
        NumberFormat::KoreanDigital2 => counting::korean_digital2(n),
        NumberFormat::VietnameseCounting => counting::vietnamese_counting(n),
        NumberFormat::HindiCounting => counting::hindi_counting(n),
        NumberFormat::ThaiCounting => counting::thai_counting(n),
        NumberFormat::BahtText => counting::baht_text(n),
        NumberFormat::DollarText => counting::dollar_text(n),

        // §17.9.27: the three formats that are written differently in every
        // language, delegated whole to `spellout`. A language it cannot spell
        // answers `None` and gets the digits — not a degrade for its own sake:
        // writing `1st` onto a Polish list is not an approximation of Polish,
        // it is English text in a Polish document, and the digits Word itself
        // falls back to are closer than another language's words.
        NumberFormat::Ordinal => {
            spellout::ordinal_numeric(n, locale).unwrap_or_else(|| n.to_string())
        }
        NumberFormat::CardinalText => {
            spellout::cardinal(n, locale).unwrap_or_else(|| n.to_string())
        }
        NumberFormat::OrdinalText => {
            spellout::ordinal_words(n, locale).unwrap_or_else(|| n.to_string())
        }

        // §17.18.59: `bullet` renders the level text, `none` renders nothing —
        // neither renders the counter. `format_list_label` returns before
        // reaching either, so these are unreachable in practice; answering with
        // the digit would be wrong if a future caller did reach them.
        NumberFormat::Bullet | NumberFormat::None => String::new(),
    }
}

/// Alphabetic numbering shared by the letter formats: Word repeats the
/// letter on overflow — a…z, then aa, bb, …, zz, aaa (a *repeating* scheme,
/// not bijective base-N). Item `len + 1` is the first letter doubled, not the
/// first letter again.
fn alphabetic_repeat(n: u32, len: u32, letter: impl Fn(u32) -> char) -> String {
    if n == 0 {
        return String::new();
    }
    let idx = (n - 1) % len;
    let count = ((n - 1) / len) as usize + 1;
    std::iter::repeat_n(letter(idx), count).collect()
}

fn to_letter_lower(n: u32) -> String {
    alphabetic_repeat(n, 26, |i| (b'a' + i as u8) as char)
}

fn to_letter_upper(n: u32) -> String {
    to_letter_lower(n).to_uppercase()
}

/// Word's Russian alphabetic numbering sequence (§17.18.59 russianLower):
/// 28 of the 33 Cyrillic letters — Ё, Й, Ъ, Ы, Ь are skipped.
const RUSSIAN_LETTERS: [char; 28] = [
    'а', 'б', 'в', 'г', 'д', 'е', 'ж', 'з', 'и', 'к', 'л', 'м', 'н', 'о', 'п', 'р', 'с', 'т', 'у',
    'ф', 'х', 'ц', 'ч', 'ш', 'щ', 'э', 'ю', 'я',
];

fn to_russian_lower(n: u32) -> String {
    alphabetic_repeat(n, RUSSIAN_LETTERS.len() as u32, |i| {
        RUSSIAN_LETTERS[i as usize]
    })
}

fn to_russian_upper(n: u32) -> String {
    to_russian_lower(n).to_uppercase()
}

fn to_roman_lower(mut n: u32) -> String {
    const VALS: [(u32, &str); 13] = [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ];
    let mut s = String::new();
    for &(val, sym) in &VALS {
        while n >= val {
            s.push_str(sym);
            n -= val;
        }
    }
    s
}

fn to_roman_upper(n: u32) -> String {
    to_roman_lower(n).to_uppercase()
}

// ── §17.18.59 sequences that need no language data (issue #132) ─────────────
//
// Thirty-one of the spec's values are a digit set, a wrapper, an ordered
// alphabet or a value table — all of which are *in the source below*, not in
// CLDR. What separates them from the values still degrading to decimal is
// stated on `StNumberFormat::Other`: those are spellout in a script that looks
// like digits (`chineseCounting` writes 12 as 十二, twelve read aloud), and
// spellout is a language question, which is `spellout.rs`'s.
//
// **Word reference render**: this environment has no Word to compare against,
// so where the spec names a sequence without listing it — which is every one
// of these — the tables are built from the sequence's own definition (the
// gojūon order, the Iroha poem, the hijāʾī order, the sexagenary cycle) rather
// than from an observed Word render. Each table below says which definition it
// is. What would overturn one: a Word render of a list using that `w:numFmt`.
//
// **A label is only as visible as its font.** Producing `①` or `ア` is this
// module's whole job, but painting it needs a typeface that covers it, and
// this engine has no per-glyph fallback: a codepoint the resolved face lacks
// is dropped, not substituted. That is general — a body run of `ASCII ① ア` in
// the spec fallback face loses both non-ASCII characters the same way — so it
// is not a numbering defect and is not fixed here (issue #139). It surfaces
// for these formats because they are the first to emit anything outside a
// Latin face's coverage on their own.
//
// It is also not the common case: §17.9.3 gives every level its own
// `<w:rPr>`, and Word writes a covering font into it when it writes one of
// these formats. Supplying that font is enough today — a level whose `rPr`
// names one renders `① ② ③`, `ア イ ウ`, `甲子 乙丑` correctly, verified by
// rendering exactly that. What is missing is only the fallback for a document
// that names *no* covering font, which is a font-resolution question rather
// than a numbering one.

/// Positional decimal in another set of ten digits.
pub(super) fn to_digit_set(n: u32, digits: &[char; 10]) -> String {
    // `n.to_string()` is ASCII, so a byte is a digit.
    n.to_string()
        .bytes()
        .map(|b| digits[usize::from(b - b'0')])
        .collect()
}

const FULL_WIDTH_DIGITS: [char; 10] = ['０', '１', '２', '３', '４', '５', '６', '７', '８', '９'];
const DEVANAGARI_DIGITS: [char; 10] = ['०', '१', '२', '३', '४', '५', '६', '७', '८', '९'];
const THAI_DIGITS: [char; 10] = ['๐', '๑', '๒', '๓', '๔', '๕', '๖', '๗', '๘', '๙'];
/// Ideographic digits used *positionally*: 12 is 一二, not 十二.
const IDEOGRAPH_DIGITS: [char; 10] = ['〇', '一', '二', '三', '四', '五', '六', '七', '八', '九'];

/// One character from a contiguous Unicode series of enclosed numbers —
/// `first` is the glyph for 1, `len` how many the series has.
///
/// **Every such series runs out**: circled digits stop at ⑳, circled
/// ideographs at ㊉. §17.18.59 does not say what a 21st item renders as, and
/// the choice is made once here rather than per format: **plain decimal**. It
/// is what Word falls back to for a format it cannot render at all, it is
/// unambiguously a number, and the alternatives — wrapping back to ① (two
/// items with the same label) or composing ②⓪ (not a number in any reading) —
/// are both worse than a digit.
fn to_enclosed(n: u32, first: char, len: u32) -> String {
    if n == 0 || n > len {
        return n.to_string();
    }
    char::from_u32(first as u32 + n - 1).map_or_else(|| n.to_string(), String::from)
}

/// The repeating scheme of [`alphabetic_repeat`], for alphabets whose items
/// are not single `char`s — Devanagari's अं is a letter plus a combining mark.
fn from_alphabet(n: u32, alphabet: &[&str]) -> String {
    if n == 0 || alphabet.is_empty() {
        return String::new();
    }
    let len = alphabet.len() as u32;
    let idx = ((n - 1) % len) as usize;
    let count = ((n - 1) / len) as usize + 1;
    alphabet[idx].repeat(count)
}

/// Katakana in gojūon (a-i-u-e-o) order — the 46 base syllables, no voiced
/// forms and no small kana, which is the order a Japanese dictionary uses.
///
/// ECMA-376 calls `aiueo` "hiragana"; Word writes katakana, and the
/// half-/full-width pairing the two spec values exist to distinguish only
/// exists for katakana — hiragana has no half-width form. Katakana is
/// therefore the reading taken here.
#[rustfmt::skip]
const KATAKANA_GOJUON: [&str; 46] = [
    "ア", "イ", "ウ", "エ", "オ", "カ", "キ", "ク", "ケ", "コ",
    "サ", "シ", "ス", "セ", "ソ", "タ", "チ", "ツ", "テ", "ト",
    "ナ", "ニ", "ヌ", "ネ", "ノ", "ハ", "ヒ", "フ", "ヘ", "ホ",
    "マ", "ミ", "ム", "メ", "モ", "ヤ", "ユ", "ヨ",
    "ラ", "リ", "ル", "レ", "ロ", "ワ", "ヲ", "ン",
];

#[rustfmt::skip]
const KATAKANA_GOJUON_HALF: [&str; 46] = [
    "ｱ", "ｲ", "ｳ", "ｴ", "ｵ", "ｶ", "ｷ", "ｸ", "ｹ", "ｺ",
    "ｻ", "ｼ", "ｽ", "ｾ", "ｿ", "ﾀ", "ﾁ", "ﾂ", "ﾃ", "ﾄ",
    "ﾅ", "ﾆ", "ﾇ", "ﾈ", "ﾉ", "ﾊ", "ﾋ", "ﾌ", "ﾍ", "ﾎ",
    "ﾏ", "ﾐ", "ﾑ", "ﾒ", "ﾓ", "ﾔ", "ﾕ", "ﾖ",
    "ﾗ", "ﾘ", "ﾙ", "ﾚ", "ﾛ", "ﾜ", "ｦ", "ﾝ",
];

/// Katakana in the order of the Iroha, the pangram poem that uses each of the
/// 47 classical syllables once. ン is *not* in it — the poem predates the
/// syllable — which is why this table is 47 long and the gojūon one 46.
#[rustfmt::skip]
const KATAKANA_IROHA: [&str; 47] = [
    "イ", "ロ", "ハ", "ニ", "ホ", "ヘ", "ト", "チ", "リ", "ヌ",
    "ル", "ヲ", "ワ", "カ", "ヨ", "タ", "レ", "ソ", "ツ", "ネ",
    "ナ", "ラ", "ム", "ウ", "ヰ", "ノ", "オ", "ク", "ヤ", "マ",
    "ケ", "フ", "コ", "エ", "テ", "ア", "サ", "キ", "ユ", "メ",
    "ミ", "シ", "ヱ", "ヒ", "モ", "セ", "ス",
];

/// The Iroha half-width, except at positions 25 and 43: ヰ and ヱ are archaic
/// and Unicode gives them **no** half-width form (U+FF66..U+FF9D has none), so
/// they stay full-width. A blank or a substitute there would silently change
/// which item a label names.
#[rustfmt::skip]
const KATAKANA_IROHA_HALF: [&str; 47] = [
    "ｲ", "ﾛ", "ﾊ", "ﾆ", "ﾎ", "ﾍ", "ﾄ", "ﾁ", "ﾘ", "ﾇ",
    "ﾙ", "ｦ", "ﾜ", "ｶ", "ﾖ", "ﾀ", "ﾚ", "ｿ", "ﾂ", "ﾈ",
    "ﾅ", "ﾗ", "ﾑ", "ｳ", "ヰ", "ﾉ", "ｵ", "ｸ", "ﾔ", "ﾏ",
    "ｹ", "ﾌ", "ｺ", "ｴ", "ﾃ", "ｱ", "ｻ", "ｷ", "ﾕ", "ﾒ",
    "ﾐ", "ｼ", "ヱ", "ﾋ", "ﾓ", "ｾ", "ｽ",
];

/// Hangul **syllables** in ganada order — each consonant with the vowel ㅏ.
#[rustfmt::skip]
const HANGUL_GANADA: [&str; 14] = [
    "가", "나", "다", "라", "마", "바", "사",
    "아", "자", "차", "카", "타", "파", "하",
];

/// Hangul **leading jamo** — the same fourteen consonants, bare. `ganada` and
/// `chosung` are the two spec values that differ in exactly this.
#[rustfmt::skip]
const HANGUL_CHOSUNG: [&str; 14] = [
    "ㄱ", "ㄴ", "ㄷ", "ㄹ", "ㅁ", "ㅂ", "ㅅ",
    "ㅇ", "ㅈ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ",
];

/// The 22 Hebrew letters in alphabetical order, final forms excluded — this is
/// `hebrew2`, the alphabet. `hebrew1` is the numeral system below.
#[rustfmt::skip]
const HEBREW_ALPHABET: [&str; 22] = [
    "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט", "י", "כ",
    "ל", "מ", "נ", "ס", "ע", "פ", "צ", "ק", "ר", "ש", "ת",
];

/// The 28 Arabic letters in modern **hijāʾī** order (ا ب ت ث …), which is the
/// alphabetical one. The abjadī order is a different sequence and belongs to
/// `arabicAbjad`, which uses it for its numeral *values* rather than as a
/// list.
#[rustfmt::skip]
const ARABIC_ALPHABET: [&str; 28] = [
    "ا", "ب", "ت", "ث", "ج", "ح", "خ", "د", "ذ", "ر",
    "ز", "س", "ش", "ص", "ض", "ط", "ظ", "ع", "غ", "ف",
    "ق", "ك", "ل", "م", "ن", "ه", "و", "ي",
];

/// The independent Devanagari vowels, followed by anusvāra and visarga — the
/// order Hindi teaching material calls the *svar*. The last two are a letter
/// plus a combining mark, which is why this table is `&str` and not `char`.
#[rustfmt::skip]
const DEVANAGARI_VOWELS: [&str; 13] = [
    "अ", "आ", "इ", "ई", "उ", "ऊ", "ऋ",
    "ए", "ऐ", "ओ", "औ", "अं", "अः",
];

/// The 33 Devanagari consonants in varga order, ending with the four
/// semivowels, three sibilants and ह.
#[rustfmt::skip]
const DEVANAGARI_CONSONANTS: [&str; 33] = [
    "क", "ख", "ग", "घ", "ङ", "च", "छ", "ज", "झ", "ञ", "ट",
    "ठ", "ड", "ढ", "ण", "त", "थ", "द", "ध", "न", "प", "फ",
    "ब", "भ", "म", "य", "र", "ल", "व", "श", "ष", "स", "ह",
];

/// The 44 Thai consonants in alphabetical order, including the two obsolete
/// letters ฃ and ฅ, which the alphabet still counts.
#[rustfmt::skip]
const THAI_CONSONANTS: [&str; 44] = [
    "ก", "ข", "ฃ", "ค", "ฅ", "ฆ", "ง", "จ", "ฉ", "ช", "ซ",
    "ฌ", "ญ", "ฎ", "ฏ", "ฐ", "ฑ", "ฒ", "ณ", "ด", "ต", "ถ",
    "ท", "ธ", "น", "บ", "ป", "ผ", "ฝ", "พ", "ฟ", "ภ", "ม",
    "ย", "ร", "ล", "ว", "ศ", "ษ", "ส", "ห", "ฬ", "อ", "ฮ",
];

/// The Chicago Manual of Style footnote sequence. Doubling on overflow is not
/// this engine's invention — it is the manual's own rule, and it happens to be
/// exactly [`alphabetic_repeat`]'s scheme: `*`, `†`, `‡`, `§`, `**`, `††`, …
const CHICAGO_SYMBOLS: [&str; 4] = ["*", "†", "‡", "§"];

/// The ten Heavenly Stems (天干) — `ideographTraditional`.
#[rustfmt::skip]
const HEAVENLY_STEMS: [&str; 10] = [
    "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸",
];

/// The twelve Earthly Branches (地支) — `ideographZodiac`.
#[rustfmt::skip]
const EARTHLY_BRANCHES: [&str; 12] = [
    "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥",
];

/// The sexagenary cycle (干支) — `ideographZodiacTraditional`: stem *paired*
/// with branch, 甲子, 乙丑, …, 癸亥, repeating every 60.
///
/// Not [`from_alphabet`]'s scheme, and the difference is the point: the cycle
/// does not repeat a symbol on overflow, it advances both wheels
/// independently, so item 61 is 甲子 again rather than 甲甲子子. Taking both
/// indices modulo their own lengths gives that for free — 10 and 12 have
/// lowest common multiple 60.
fn to_sexagenary(n: u32) -> String {
    if n == 0 {
        return String::new();
    }
    let i = n - 1;
    let stem = HEAVENLY_STEMS[(i % HEAVENLY_STEMS.len() as u32) as usize];
    let branch = EARTHLY_BRANCHES[(i % EARTHLY_BRANCHES.len() as u32) as usize];
    format!("{stem}{branch}")
}

/// Additive numerals: emit the largest symbol that fits, repeatedly. The shape
/// [`to_roman_lower`] already has, with the table supplied.
///
/// **Bounded at 9999.** These are closed systems — Hebrew's largest symbol is
/// ת (400) and the abjad's is غ (1000) — and neither has a notation above its
/// top symbol other than repeating it. Word repeats; so does this, up to the
/// point where repetition stops carrying information (25 ת's), past which the
/// decimal is the honest answer rather than a wall of one letter. A list
/// counter reaching five figures is not a numeral in any of these systems.
fn to_additive(mut n: u32, table: &[(u32, char)]) -> String {
    if n == 0 || n > 9_999 {
        return n.to_string();
    }
    let mut s = String::new();
    for &(value, symbol) in table {
        while n >= value {
            s.push(symbol);
            n -= value;
        }
    }
    s
}

#[rustfmt::skip]
const HEBREW_NUMERALS: [(u32, char); 22] = [
    (400, 'ת'), (300, 'ש'), (200, 'ר'), (100, 'ק'),
    (90, 'צ'), (80, 'פ'), (70, 'ע'), (60, 'ס'), (50, 'נ'),
    (40, 'מ'), (30, 'ל'), (20, 'כ'), (10, 'י'),
    (9, 'ט'), (8, 'ח'), (7, 'ז'), (6, 'ו'), (5, 'ה'),
    (4, 'ד'), (3, 'ג'), (2, 'ב'), (1, 'א'),
];

/// Abjad numerals in abjadī order — ا=1 … ي=10, ك=20 … ق=100, ر=200 … غ=1000.
#[rustfmt::skip]
const ABJAD_NUMERALS: [(u32, char); 28] = [
    (1000, 'غ'), (900, 'ظ'), (800, 'ض'), (700, 'ذ'), (600, 'خ'), (500, 'ث'),
    (400, 'ت'), (300, 'ش'), (200, 'ر'), (100, 'ق'),
    (90, 'ص'), (80, 'ف'), (70, 'ع'), (60, 'س'), (50, 'ن'),
    (40, 'م'), (30, 'ل'), (20, 'ك'), (10, 'ي'),
    (9, 'ط'), (8, 'ح'), (7, 'ز'), (6, 'و'), (5, 'ه'),
    (4, 'د'), (3, 'ج'), (2, 'ب'), (1, 'ا'),
];

/// §17.18.59 `hebrew1`: the Hebrew numeral system.
///
/// Plain addition, with the one substitution every Hebrew numeral has: 15 and
/// 16 would spell יה and יו, which are forms of the Tetragrammaton, so they
/// are written טו (9+6) and טז (9+7) instead. The rule applies to the last two
/// digits at any magnitude — 115 is קטו, not קיה — which a suffix rewrite
/// gets exactly right, because greedy addition emits at most one י and it is
/// always immediately before the units.
fn to_hebrew_numeral(n: u32) -> String {
    let s = to_additive(n, &HEBREW_NUMERALS);
    if let Some(rest) = s.strip_suffix("יה") {
        format!("{rest}טו")
    } else if let Some(rest) = s.strip_suffix("יו") {
        format!("{rest}טז")
    } else {
        s
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::*;

    fn make_defs(
        abstracts: Vec<(AbstractNumId, Vec<NumberingLevelDefinition>)>,
        instances: Vec<(NumId, AbstractNumId, Vec<NumberingLevelDefinition>)>,
    ) -> NumberingDefinitions {
        NumberingDefinitions {
            abstract_nums: abstracts
                .into_iter()
                .map(|(id, levels)| (id, AbstractNumbering { levels }))
                .collect(),
            numbering_instances: instances
                .into_iter()
                .map(|(num_id, abstract_id, overrides)| {
                    (
                        num_id,
                        NumberingInstance {
                            abstract_num_id: abstract_id,
                            level_overrides: overrides
                                .into_iter()
                                .map(|def| crate::model::LevelOverride {
                                    level: def.level,
                                    start_override: None,
                                    definition: Some(def),
                                })
                                .collect(),
                        },
                    )
                })
                .collect(),
            pic_bullets: HashMap::new(),
        }
    }

    fn level(lvl: u8, fmt: NumberFormat, text: &str, start: u32) -> NumberingLevelDefinition {
        NumberingLevelDefinition {
            level: lvl,
            format: Some(fmt),
            level_text: text.to_string(),
            start: Some(start),
            justification: None,
            indentation: None,
            run_properties: None,
            lvl_pic_bullet_id: None,
            suffix: LevelSuffix::default(),
            is_legal: false,
        }
    }

    #[test]
    fn russian_letters_follow_word_sequence() {
        assert_eq!(to_russian_lower(1), "а");
        assert_eq!(to_russian_lower(2), "б");
        // Ё (7th of the full alphabet) is skipped: 6 → е, 7 → ж.
        assert_eq!(to_russian_lower(6), "е");
        assert_eq!(to_russian_lower(7), "ж");
        // Й is skipped: 9 → и, 10 → к.
        assert_eq!(to_russian_lower(10), "к");
        assert_eq!(to_russian_lower(28), "я");
        // Overflow repeats the letter, like latin `aa`.
        assert_eq!(to_russian_lower(29), "аа");
        assert_eq!(to_russian_upper(1), "А");
        assert_eq!(to_russian_upper(28), "Я");
        assert_eq!(to_russian_lower(0), "");
    }

    #[test]
    fn russian_format_expands_in_label_template() {
        let locale = Locale::default();
        assert_eq!(format_number(3, NumberFormat::RussianUpper, locale), "В");
        assert_eq!(format_number(3, NumberFormat::RussianLower, locale), "в");
    }

    #[test]
    fn single_instance_resolves_from_abstract() {
        let defs = make_defs(
            vec![(
                AbstractNumId::new(0),
                vec![level(0, NumberFormat::Decimal, "%1.", 1)],
            )],
            vec![(NumId::new(1), AbstractNumId::new(0), vec![])],
        );

        let resolved = resolve_numbering(&defs);
        let levels = &resolved.get(&NumId::new(1)).unwrap().levels;

        assert_eq!(levels.len(), 1);
        assert_eq!(levels[0].format, NumberFormat::Decimal);
        assert_eq!(levels[0].level_text, "%1.");
        assert_eq!(levels[0].start, 1);
    }

    #[test]
    fn level_override_replaces_abstract_level() {
        let defs = make_defs(
            vec![(
                AbstractNumId::new(0),
                vec![
                    level(0, NumberFormat::Decimal, "%1.", 1),
                    level(1, NumberFormat::LowerLetter, "%2)", 1),
                ],
            )],
            vec![(
                NumId::new(1),
                AbstractNumId::new(0),
                // Override level 0 to bullet
                vec![level(0, NumberFormat::Bullet, "•", 1)],
            )],
        );

        let resolved = resolve_numbering(&defs);
        let levels = &resolved.get(&NumId::new(1)).unwrap().levels;

        assert_eq!(levels.len(), 2);
        assert_eq!(levels[0].format, NumberFormat::Bullet, "overridden");
        assert_eq!(levels[0].level_text, "•");
        assert_eq!(levels[1].format, NumberFormat::LowerLetter, "from abstract");
    }

    #[test]
    fn missing_abstract_produces_empty_levels() {
        let defs = make_defs(
            vec![],
            vec![(NumId::new(1), AbstractNumId::new(99), vec![])],
        );

        let resolved = resolve_numbering(&defs);
        let levels = &resolved.get(&NumId::new(1)).unwrap().levels;
        assert!(levels.is_empty());
    }

    #[test]
    fn multiple_instances_same_abstract() {
        let defs = make_defs(
            vec![(
                AbstractNumId::new(0),
                vec![level(0, NumberFormat::Decimal, "%1.", 1)],
            )],
            vec![
                (NumId::new(1), AbstractNumId::new(0), vec![]),
                (
                    NumId::new(2),
                    AbstractNumId::new(0),
                    vec![level(0, NumberFormat::Decimal, "%1)", 10)],
                ),
            ],
        );

        let resolved = resolve_numbering(&defs);

        let l1 = &resolved.get(&NumId::new(1)).unwrap().levels;
        assert_eq!(l1[0].level_text, "%1.");
        assert_eq!(l1[0].start, 1);

        let l2 = &resolved.get(&NumId::new(2)).unwrap().levels;
        assert_eq!(l2[0].level_text, "%1)");
        assert_eq!(l2[0].start, 10);
    }

    /// §17.9.28: the override is carried beside the level's own `w:start`, not
    /// written over it. Layout fires it once, at the instance's first use, and
    /// then the shared sequence continues — behaviour a single `start` field
    /// cannot express (issue #230).
    #[test]
    fn start_override_is_carried_separately_from_start() {
        let mut abstract_nums = HashMap::new();
        abstract_nums.insert(
            AbstractNumId::new(0),
            AbstractNumbering {
                levels: vec![level(0, NumberFormat::Decimal, "%1.", 1)],
            },
        );
        let mut numbering_instances = HashMap::new();
        numbering_instances.insert(
            NumId::new(1),
            NumberingInstance {
                abstract_num_id: AbstractNumId::new(0),
                level_overrides: vec![crate::model::LevelOverride {
                    level: 0,
                    start_override: Some(5),
                    definition: None,
                }],
            },
        );
        let defs = NumberingDefinitions {
            abstract_nums,
            numbering_instances,
            pic_bullets: HashMap::new(),
        };
        let resolved = resolve_numbering(&defs);
        let lvl = &resolved[&NumId::new(1)].levels[0];
        assert_eq!(lvl.start, 1, "the abstract level's own start is untouched");
        assert_eq!(lvl.start_override, Some(5));
    }

    #[test]
    fn legal_numbering_renders_all_levels_decimal() {
        let levels = vec![
            ResolvedNumberingLevel {
                format: NumberFormat::UpperRoman,
                level_text: "%1".to_string(),
                start: 1,
                start_override: None,
                run_properties: None,
                indentation: None,
                justification: None,
                lvl_pic_bullet_id: None,
                suffix: LevelSuffix::default(),
                is_legal: false,
            },
            ResolvedNumberingLevel {
                format: NumberFormat::LowerLetter,
                level_text: "%1.%2".to_string(),
                start: 1,
                start_override: None,
                run_properties: None,
                indentation: None,
                justification: None,
                lvl_pic_bullet_id: None,
                suffix: LevelSuffix::default(),
                is_legal: true,
            },
        ];
        let mut counters = HashMap::new();
        counters.insert((AbstractNumId::new(1), 0u8), 3u32); // would be "III" un-legal
        counters.insert((AbstractNumId::new(1), 1u8), 2u32); // would be "b" un-legal
        let label = format_list_label(
            &levels,
            1,
            &counters,
            AbstractNumId::new(1),
            Locale::English,
        )
        .unwrap();
        assert_eq!(
            label, "3.2",
            "isLgl forces decimal for every referenced level"
        );
    }

    #[test]
    fn level_with_no_format_defaults_to_none() {
        let defs = make_defs(
            vec![(
                AbstractNumId::new(0),
                vec![NumberingLevelDefinition {
                    level: 0,
                    format: None,
                    level_text: String::new(),
                    start: None,
                    justification: None,
                    indentation: None,
                    run_properties: None,
                    lvl_pic_bullet_id: None,
                    suffix: LevelSuffix::default(),
                    is_legal: false,
                }],
            )],
            vec![(NumId::new(1), AbstractNumId::new(0), vec![])],
        );

        let resolved = resolve_numbering(&defs);
        let levels = &resolved.get(&NumId::new(1)).unwrap().levels;
        assert_eq!(levels[0].format, NumberFormat::None);
        assert_eq!(levels[0].start, 1);
    }

    #[test]
    fn lower_letter_repeats_on_overflow() {
        // §17.9 lowerLetter: a…z then aa, bb, … (repeating, not bijective).
        assert_eq!(
            format_number(1, NumberFormat::LowerLetter, Locale::English),
            "a"
        );
        assert_eq!(
            format_number(26, NumberFormat::LowerLetter, Locale::English),
            "z"
        );
        assert_eq!(
            format_number(27, NumberFormat::LowerLetter, Locale::English),
            "aa"
        );
        assert_eq!(
            format_number(28, NumberFormat::LowerLetter, Locale::English),
            "bb"
        );
        assert_eq!(
            format_number(52, NumberFormat::LowerLetter, Locale::English),
            "zz"
        );
        assert_eq!(
            format_number(53, NumberFormat::LowerLetter, Locale::English),
            "aaa"
        );
    }

    #[test]
    fn upper_letter_matches_lower_uppercased() {
        assert_eq!(
            format_number(27, NumberFormat::UpperLetter, Locale::English),
            "AA"
        );
    }

    #[test]
    fn roman_and_ordinal_formats() {
        assert_eq!(
            format_number(4, NumberFormat::LowerRoman, Locale::English),
            "iv"
        );
        assert_eq!(
            format_number(2026, NumberFormat::UpperRoman, Locale::English),
            "MMXXVI"
        );
        assert_eq!(
            format_number(1, NumberFormat::Ordinal, Locale::English),
            "1st"
        );
        assert_eq!(
            format_number(2, NumberFormat::Ordinal, Locale::English),
            "2nd"
        );
        assert_eq!(
            format_number(11, NumberFormat::Ordinal, Locale::English),
            "11th"
        );
        assert_eq!(
            format_number(23, NumberFormat::Ordinal, Locale::English),
            "23rd"
        );
        assert_eq!(
            format_number(111, NumberFormat::Ordinal, Locale::English),
            "111th"
        );
    }

    /// §17.18.59: neither `bullet` nor `none` renders the counter. The `_ =>`
    /// arm this replaced printed the digit for both.
    #[test]
    fn formats_that_render_no_counter_render_nothing() {
        assert_eq!(format_number(7, NumberFormat::Bullet, Locale::English), "");
        assert_eq!(format_number(7, NumberFormat::None, Locale::English), "");
    }

    /// A language whose number words this engine cannot spell gets the digits —
    /// for all three text formats, not just the one that was implemented.
    #[test]
    fn a_non_spelling_locale_gets_digits_for_every_text_format() {
        for fmt in [
            NumberFormat::Ordinal,
            NumberFormat::CardinalText,
            NumberFormat::OrdinalText,
        ] {
            assert_eq!(format_number(3, fmt, Locale::CommaDecimal), "3", "{fmt:?}");
            assert_eq!(format_number(3, fmt, Locale::PointDecimal), "3", "{fmt:?}");
        }
    }

    // ── §17.18.59 sequences (issue #132) ─────────────────────────────────

    fn fmt(n: u32, f: NumberFormat) -> String {
        format_number(n, f, Locale::English)
    }

    /// Every §17.18.59 value closed by issue #132 — the list the claim "these
    /// need no language data" is made about, kept next to the test that
    /// checks it.
    const SEQUENCE_FORMATS: [NumberFormat; 31] = [
        NumberFormat::DecimalFullWidth,
        NumberFormat::DecimalFullWidth2,
        NumberFormat::DecimalHalfWidth,
        NumberFormat::HindiNumbers,
        NumberFormat::ThaiNumbers,
        NumberFormat::IdeographDigital,
        NumberFormat::DecimalZero,
        NumberFormat::Hex,
        NumberFormat::NumberInDash,
        NumberFormat::DecimalEnclosedFullstop,
        NumberFormat::DecimalEnclosedParen,
        NumberFormat::DecimalEnclosedCircle,
        NumberFormat::DecimalEnclosedCircleChinese,
        NumberFormat::IdeographEnclosedCircle,
        NumberFormat::Aiueo,
        NumberFormat::AiueoFullWidth,
        NumberFormat::Iroha,
        NumberFormat::IrohaFullWidth,
        NumberFormat::Ganada,
        NumberFormat::Chosung,
        NumberFormat::Hebrew2,
        NumberFormat::ArabicAlpha,
        NumberFormat::HindiVowels,
        NumberFormat::HindiConsonants,
        NumberFormat::ThaiLetters,
        NumberFormat::Chicago,
        NumberFormat::IdeographTraditional,
        NumberFormat::IdeographZodiac,
        NumberFormat::IdeographZodiacTraditional,
        NumberFormat::Hebrew1,
        NumberFormat::ArabicAbjad,
    ];

    /// The claim issue #132's classification rests on, as a test: not one of
    /// these formats reads the locale, so none of them needs CLDR.
    #[test]
    fn every_sequence_format_renders_the_same_in_every_language() {
        for f in SEQUENCE_FORMATS {
            let english = fmt(7, f);
            assert!(!english.is_empty(), "{f:?} renders nothing for 7");
            for locale in [
                Locale::CommaDecimal,
                Locale::PointDecimal,
                Locale::Unrecognised,
            ] {
                assert_eq!(format_number(7, f, locale), english, "{f:?}/{locale:?}");
            }
        }
    }

    /// Digit substitution is positional — each decimal place is replaced on
    /// its own, so 12 is two characters and not the word for twelve. That is
    /// what separates `ideographDigital` (一二) from `chineseCounting` (十二),
    /// which is still unsupported.
    #[test]
    fn digit_substitution_replaces_each_place_independently() {
        assert_eq!(fmt(12, NumberFormat::DecimalFullWidth), "１２");
        assert_eq!(fmt(12, NumberFormat::DecimalFullWidth2), "１２");
        assert_eq!(fmt(12, NumberFormat::DecimalHalfWidth), "12");
        assert_eq!(fmt(120, NumberFormat::HindiNumbers), "१२०");
        assert_eq!(fmt(45, NumberFormat::ThaiNumbers), "๔๕");
        assert_eq!(fmt(12, NumberFormat::IdeographDigital), "一二");
        assert_eq!(fmt(10, NumberFormat::IdeographDigital), "一〇");
    }

    #[test]
    fn decorated_decimals_pad_wrap_and_enclose() {
        assert_eq!(fmt(1, NumberFormat::DecimalZero), "01");
        assert_eq!(fmt(9, NumberFormat::DecimalZero), "09");
        assert_eq!(fmt(10, NumberFormat::DecimalZero), "10");
        assert_eq!(fmt(100, NumberFormat::DecimalZero), "100");

        assert_eq!(fmt(9, NumberFormat::Hex), "9");
        assert_eq!(fmt(10, NumberFormat::Hex), "A");
        assert_eq!(fmt(255, NumberFormat::Hex), "FF");

        assert_eq!(fmt(3, NumberFormat::NumberInDash), "-3-");

        assert_eq!(fmt(1, NumberFormat::DecimalEnclosedCircle), "①");
        assert_eq!(fmt(20, NumberFormat::DecimalEnclosedCircle), "⑳");
        assert_eq!(fmt(1, NumberFormat::DecimalEnclosedParen), "⑴");
        assert_eq!(fmt(20, NumberFormat::DecimalEnclosedParen), "⒇");
        assert_eq!(fmt(1, NumberFormat::DecimalEnclosedFullstop), "⒈");
        assert_eq!(fmt(20, NumberFormat::DecimalEnclosedFullstop), "⒛");
        assert_eq!(fmt(1, NumberFormat::IdeographEnclosedCircle), "㊀");
        assert_eq!(fmt(10, NumberFormat::IdeographEnclosedCircle), "㊉");

        // The Chinese-locale circled form is the same series.
        assert_eq!(
            fmt(7, NumberFormat::DecimalEnclosedCircleChinese),
            fmt(7, NumberFormat::DecimalEnclosedCircle),
        );
    }

    /// The documented choice at the end of every enclosed series: a plain
    /// decimal, not a wrap-around that would give two items the same label.
    #[test]
    fn an_enclosed_series_that_runs_out_falls_back_to_decimal() {
        for f in [
            NumberFormat::DecimalEnclosedCircle,
            NumberFormat::DecimalEnclosedParen,
            NumberFormat::DecimalEnclosedFullstop,
        ] {
            assert_eq!(fmt(21, f), "21", "{f:?}");
        }
        assert_eq!(fmt(11, NumberFormat::IdeographEnclosedCircle), "11");
    }

    /// Fixed alphabets cycle by repetition, exactly as `lowerLetter` does —
    /// item `len + 1` is the first item doubled.
    #[test]
    fn fixed_alphabets_cycle_by_repeating_the_item() {
        for (f, first, last, len) in [
            (NumberFormat::Aiueo, "ｱ", "ﾝ", 46),
            (NumberFormat::AiueoFullWidth, "ア", "ン", 46),
            (NumberFormat::Iroha, "ｲ", "ｽ", 47),
            (NumberFormat::IrohaFullWidth, "イ", "ス", 47),
            (NumberFormat::Ganada, "가", "하", 14),
            (NumberFormat::Chosung, "ㄱ", "ㅎ", 14),
            (NumberFormat::Hebrew2, "א", "ת", 22),
            (NumberFormat::ArabicAlpha, "ا", "ي", 28),
            (NumberFormat::HindiVowels, "अ", "अः", 13),
            (NumberFormat::HindiConsonants, "क", "ह", 33),
            (NumberFormat::ThaiLetters, "ก", "ฮ", 44),
            (NumberFormat::Chicago, "*", "§", 4),
            (NumberFormat::IdeographTraditional, "甲", "癸", 10),
            (NumberFormat::IdeographZodiac, "子", "亥", 12),
        ] {
            assert_eq!(fmt(1, f), first, "{f:?} first");
            assert_eq!(fmt(len, f), last, "{f:?} at {len}");
            assert_eq!(fmt(len + 1, f), first.repeat(2), "{f:?} overflow");
        }
    }

    /// ヰ and ヱ have no half-width form in Unicode, so the half-width Iroha
    /// keeps them full-width rather than dropping or substituting them — a
    /// blank there would silently shift every later label.
    #[test]
    fn the_half_width_iroha_keeps_its_two_archaic_syllables() {
        assert_eq!(fmt(25, NumberFormat::Iroha), "ヰ");
        assert_eq!(fmt(43, NumberFormat::Iroha), "ヱ");
        assert_eq!(fmt(25, NumberFormat::IrohaFullWidth), "ヰ");
    }

    /// The sexagenary cycle advances stem and branch independently, so it
    /// repeats at 61 rather than doubling a symbol.
    #[test]
    fn the_sexagenary_cycle_advances_both_wheels() {
        let f = NumberFormat::IdeographZodiacTraditional;
        assert_eq!(fmt(1, f), "甲子");
        assert_eq!(fmt(2, f), "乙丑");
        assert_eq!(fmt(11, f), "甲戌");
        assert_eq!(fmt(13, f), "丙子");
        assert_eq!(fmt(60, f), "癸亥");
        assert_eq!(fmt(61, f), "甲子", "the cycle repeats, it does not double");
    }

    /// §17.18.59 `hebrew1`, including the substitution every Hebrew numeral
    /// makes: 15 and 16 are written 9+6 and 9+7 so they do not spell a form of
    /// the divine name. The rule follows the last two digits at any magnitude.
    #[test]
    fn hebrew_numerals_add_and_avoid_the_divine_name() {
        for (n, want) in [
            (1, "א"),
            (5, "ה"),
            (6, "ו"),
            (10, "י"),
            (15, "טו"),
            (16, "טז"),
            (17, "יז"),
            (26, "כו"),
            (106, "קו"),
            (115, "קטו"),
            (116, "קטז"),
            (400, "ת"),
            (999, "תתקצט"),
        ] {
            assert_eq!(fmt(n, NumberFormat::Hebrew1), want, "{n}");
        }
    }

    #[test]
    fn abjad_numerals_add_their_symbols() {
        for (n, want) in [
            (1, "ا"),
            (10, "ي"),
            (11, "يا"),
            (28, "كح"),
            (100, "ق"),
            (1000, "غ"),
            (1999, "غظصط"),
        ] {
            assert_eq!(fmt(n, NumberFormat::ArabicAbjad), want, "{n}");
        }
    }

    /// Both numeral systems are closed, so past the point where repeating the
    /// largest symbol carries information they answer with the decimal.
    #[test]
    fn additive_numerals_past_their_range_render_as_digits() {
        for f in [NumberFormat::Hebrew1, NumberFormat::ArabicAbjad] {
            assert!(fmt(9_999, f).chars().count() > 1, "{f:?} in range");
            assert_eq!(fmt(10_000, f), "10000", "{f:?}");
            assert_eq!(fmt(u32::MAX, f), u32::MAX.to_string(), "{f:?}");
        }
    }

    /// …and the formats that are the same in every language are untouched by it.
    #[test]
    fn language_independent_formats_ignore_the_locale() {
        for locale in [
            Locale::English,
            Locale::CommaDecimal,
            Locale::PointDecimal,
            Locale::Unrecognised,
        ] {
            assert_eq!(format_number(4, NumberFormat::Decimal, locale), "4");
            assert_eq!(format_number(4, NumberFormat::LowerRoman, locale), "iv");
            assert_eq!(format_number(4, NumberFormat::UpperLetter, locale), "D");
        }
    }
}
