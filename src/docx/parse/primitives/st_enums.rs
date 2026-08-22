//! OOXML `ST_*` simple-type enums with strict `Deserialize` impls.
//!
//! Each schema enum mirrors an OOXML spec-defined simple type. `From<St…>`
//! implementations convert the schema enum into the matching model type.
//! Unknown string values fail deserialization (plan §Decisions: strict) — the
//! one exception is `StNumberFormat`, a large extensible value space (~60 spec
//! values) whose unsupported members must degrade rather than fail the parse
//! (see its `#[serde(other)]` variant).
//!
//! Alphabetically ordered by schema type name. Layered as:
//!
//! 1. Schema enum with `#[derive(Deserialize)]` + `#[serde(rename_all)]`
//!    (or explicit per-variant rename where the value doesn't match `camelCase`).
//! 2. `impl From<StXxx> for ModelXxx` — identity mapping in most cases,
//!    re-naming/re-mapping where the model is coarser or uses different names.
//!
//! Tests live at the bottom of the file and cover every variant plus a
//! known-bad value for each enum.

use serde::Deserialize;

use crate::docx::model::{
    Alignment, BorderStyle, BreakClear, CellVerticalAlign, FieldCharType, FrameWrap, HeightRule,
    HighlightColor, NumberFormat, PTabAlignment, PTabLeader, PTabRelativeTo, PageOrientation,
    SectionType, ShadingPattern, TabAlignment, TabLeader, TableAnchor, TableLayout, TableOverlap,
    TableXAlign, TableYAlign, TextAlignment, TextDirection, ThemeFontRef, UnderlineStyle,
    VerticalAlign,
};

// ── StBorderType (§17.18.2) ───────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StBorderType {
    /// §17.18.2: "no border" — distinct from `none`, and carried through as
    /// such. See `BorderStyle` for why the two must not be merged.
    Nil,
    None,
    Single,
    Thick,
    Double,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Triple,
    ThinThickSmallGap,
    ThickThinSmallGap,
    ThinThickThinSmallGap,
    ThinThickMediumGap,
    ThickThinMediumGap,
    ThinThickThinMediumGap,
    ThinThickLargeGap,
    ThickThinLargeGap,
    ThinThickThinLargeGap,
    Wave,
    DoubleWave,
    DashSmallGap,
    DashDotStroked,
    ThreeDEmboss,
    ThreeDEngrave,
    Outset,
    Inset,
}

impl From<StBorderType> for BorderStyle {
    fn from(s: StBorderType) -> Self {
        match s {
            // §17.18.2: kept distinct. Both draw nothing, but [MS-OI29500]
            // §17.4.66 gives them opposite behaviour in table border conflict
            // resolution — `nil` suppresses the shared edge, `none` yields.
            StBorderType::Nil => Self::Nil,
            StBorderType::None => Self::None,
            StBorderType::Single => Self::Single,
            StBorderType::Thick => Self::Thick,
            StBorderType::Double => Self::Double,
            StBorderType::Dotted => Self::Dotted,
            StBorderType::Dashed => Self::Dashed,
            StBorderType::DotDash => Self::DotDash,
            StBorderType::DotDotDash => Self::DotDotDash,
            StBorderType::Triple => Self::Triple,
            StBorderType::ThinThickSmallGap => Self::ThinThickSmallGap,
            StBorderType::ThickThinSmallGap => Self::ThickThinSmallGap,
            StBorderType::ThinThickThinSmallGap => Self::ThinThickThinSmallGap,
            StBorderType::ThinThickMediumGap => Self::ThinThickMediumGap,
            StBorderType::ThickThinMediumGap => Self::ThickThinMediumGap,
            StBorderType::ThinThickThinMediumGap => Self::ThinThickThinMediumGap,
            StBorderType::ThinThickLargeGap => Self::ThinThickLargeGap,
            StBorderType::ThickThinLargeGap => Self::ThickThinLargeGap,
            StBorderType::ThinThickThinLargeGap => Self::ThinThickThinLargeGap,
            StBorderType::Wave => Self::Wave,
            StBorderType::DoubleWave => Self::DoubleWave,
            StBorderType::DashSmallGap => Self::DashSmallGap,
            StBorderType::DashDotStroked => Self::DashDotStroked,
            StBorderType::ThreeDEmboss => Self::ThreeDEmboss,
            StBorderType::ThreeDEngrave => Self::ThreeDEngrave,
            StBorderType::Outset => Self::Outset,
            StBorderType::Inset => Self::Inset,
        }
    }
}

// ── StBrClear (§17.18.4) ──────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StBrClear {
    None,
    Left,
    Right,
    All,
}

impl From<StBrClear> for BreakClear {
    fn from(s: StBrClear) -> Self {
        match s {
            StBrClear::None => Self::None,
            StBrClear::Left => Self::Left,
            StBrClear::Right => Self::Right,
            StBrClear::All => Self::All,
        }
    }
}

// ── StFldCharType (§17.18.29) ─────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StFldCharType {
    Begin,
    Separate,
    End,
}

impl From<StFldCharType> for FieldCharType {
    fn from(s: StFldCharType) -> Self {
        match s {
            StFldCharType::Begin => Self::Begin,
            StFldCharType::Separate => Self::Separate,
            StFldCharType::End => Self::End,
        }
    }
}

// ── StFrameWrap (§17.18.104 ST_Wrap) ──────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StFrameWrap {
    Auto,
    NotBeside,
    Around,
    Tight,
    Through,
    None,
}

impl From<StFrameWrap> for FrameWrap {
    fn from(s: StFrameWrap) -> Self {
        match s {
            StFrameWrap::Auto => Self::Auto,
            StFrameWrap::NotBeside => Self::NotBeside,
            StFrameWrap::Around => Self::Around,
            StFrameWrap::Tight => Self::Tight,
            StFrameWrap::Through => Self::Through,
            StFrameWrap::None => Self::None,
        }
    }
}

// ── StHeightRule (§17.18.38) ──────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StHeightRule {
    Auto,
    Exact,
    AtLeast,
}

impl From<StHeightRule> for HeightRule {
    fn from(s: StHeightRule) -> Self {
        match s {
            StHeightRule::Auto => Self::Auto,
            StHeightRule::Exact => Self::Exact,
            StHeightRule::AtLeast => Self::AtLeast,
        }
    }
}

// ── StHighlightColor (§17.18.40) ──────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StHighlightColor {
    /// §17.18.40: explicit "no highlight" — overrides any inherited
    /// highlight. Distinct from absence of the element, which inherits.
    None,
    Black,
    Blue,
    Cyan,
    DarkBlue,
    DarkCyan,
    DarkGray,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    Green,
    LightGray,
    Magenta,
    Red,
    White,
    Yellow,
}

impl From<StHighlightColor> for HighlightColor {
    fn from(s: StHighlightColor) -> Self {
        match s {
            StHighlightColor::None => Self::None,
            StHighlightColor::Black => Self::Black,
            StHighlightColor::Blue => Self::Blue,
            StHighlightColor::Cyan => Self::Cyan,
            StHighlightColor::DarkBlue => Self::DarkBlue,
            StHighlightColor::DarkCyan => Self::DarkCyan,
            StHighlightColor::DarkGray => Self::DarkGray,
            StHighlightColor::DarkGreen => Self::DarkGreen,
            StHighlightColor::DarkMagenta => Self::DarkMagenta,
            StHighlightColor::DarkRed => Self::DarkRed,
            StHighlightColor::DarkYellow => Self::DarkYellow,
            StHighlightColor::Green => Self::Green,
            StHighlightColor::LightGray => Self::LightGray,
            StHighlightColor::Magenta => Self::Magenta,
            StHighlightColor::Red => Self::Red,
            StHighlightColor::White => Self::White,
            StHighlightColor::Yellow => Self::Yellow,
        }
    }
}

// ── StJc (§17.18.44) ──────────────────────────────────────────────────────
//
// OOXML `both` and `justify` are synonyms per the spec; both produce
// `Alignment::Both`. Rust variant names follow the model's directional
// naming (Start/End) rather than OOXML's presentation naming (Left/Right),
// preserving fidelity to the schema side.

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
pub enum StJc {
    #[serde(rename = "left", alias = "start")]
    Left,
    #[serde(rename = "center")]
    Center,
    #[serde(rename = "right", alias = "end")]
    Right,
    #[serde(rename = "both", alias = "justify")]
    Both,
    #[serde(rename = "distribute")]
    Distribute,
    #[serde(rename = "thaiDistribute")]
    ThaiDistribute,
    // §17.18.44: Arabic kashida justification and the legacy numbering-tab
    // alignment are spec-legal but not modelled distinctly. They must still
    // parse (else an Arabic document fails outright); each degrades to the
    // nearest modelled alignment. Kept as named variants rather than a
    // catch-all so a genuine typo still fails deserialization.
    #[serde(rename = "mediumKashida")]
    MediumKashida,
    #[serde(rename = "highKashida")]
    HighKashida,
    #[serde(rename = "lowKashida")]
    LowKashida,
    #[serde(rename = "numTab")]
    NumTab,
}

impl From<StJc> for Alignment {
    fn from(s: StJc) -> Self {
        match s {
            StJc::Left => Self::Start,
            StJc::Center => Self::Center,
            StJc::Right => Self::End,
            StJc::Both => Self::Both,
            StJc::Distribute => Self::Distribute,
            StJc::ThaiDistribute => Self::Thai,
            // Kashida is a form of full justification; numTab aligns to the
            // start of the numbering area.
            StJc::MediumKashida | StJc::HighKashida | StJc::LowKashida => Self::Both,
            StJc::NumTab => Self::Start,
        }
    }
}

// ── StLineSpacingRule (§17.18.48) ─────────────────────────────────────────
//
// No `From` impl: the model's `LineSpacing` is a discriminated union that
// combines the rule with the value. The conversion happens at the owning
// schema (see `parse::properties::paragraph::SpacingXml`).

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StLineSpacingRule {
    Auto,
    Exact,
    AtLeast,
}

// ── StNumberFormat (§17.18.59) ────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StNumberFormat {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
    Bullet,
    Ordinal,
    CardinalText,
    OrdinalText,
    RussianUpper,
    RussianLower,
    None,

    // ── Digit substitution ────────────────────────────────────────────────
    // Positional decimal in another digit set: same arithmetic, ten different
    // characters. Nothing about the language is involved beyond which ten.
    /// Full-width (double-byte) Arabic numerals — `１`, `２`, `３`.
    DecimalFullWidth,
    /// §17.18.59's second full-width form. Word writes the same ten
    /// characters as `decimalFullWidth`; the two differ only in which
    /// East-Asian font Word picks to draw them with, which is not a
    /// numbering question.
    DecimalFullWidth2,
    /// Half-width Arabic numerals — the ASCII digits, i.e. `decimal`.
    DecimalHalfWidth,
    /// Devanagari digits — `१`, `२`, `३`.
    HindiNumbers,
    /// Thai digits — `๑`, `๒`, `๓`.
    ThaiNumbers,
    /// Ideographic digits used *positionally* — `一`, `二`, … `一二` for 12.
    /// Not `chineseCounting`, which writes 12 as `十二`; that one is spellout
    /// and stays in [`Other`](Self::Other).
    IdeographDigital,

    // ── Decorated decimal ─────────────────────────────────────────────────
    // Ordinary arithmetic plus a wrapper, a pad, or an enclosed-glyph series.
    /// Decimal padded to two digits — `01`, `02`, … `10`.
    DecimalZero,
    /// Upper-case hexadecimal — `1`, … `9`, `A`, `B`.
    Hex,
    /// The number between dashes — `-1-`, `-2-`.
    NumberInDash,
    /// Enclosed by a full stop — `⒈`, `⒉` (U+2488…, 1–20).
    DecimalEnclosedFullstop,
    /// Parenthesised — `⑴`, `⑵` (U+2474…, 1–20).
    DecimalEnclosedParen,
    /// Circled — `①`, `②` (U+2460…, 1–20).
    DecimalEnclosedCircle,
    /// §17.18.59's Chinese-locale circled form. The same U+2460 series;
    /// Word distinguishes it by font, not by character.
    DecimalEnclosedCircleChinese,
    /// Circled ideographic digits — `㊀`, `㊁` (U+3280…, 1–10).
    IdeographEnclosedCircle,

    // ── Fixed alphabet ────────────────────────────────────────────────────
    // A finite ordered list of characters, cycled with repetition exactly as
    // `lowerLetter` cycles a…z, aa, bb — see `numbering::alphabetic_repeat`.
    /// Katakana in gojūon (a-i-u-e-o) order, half-width — `ｱ`, `ｲ`, `ｳ`.
    Aiueo,
    /// Katakana in gojūon order, full-width — `ア`, `イ`, `ウ`.
    AiueoFullWidth,
    /// Katakana in iroha order, half-width — `ｲ`, `ﾛ`, `ﾊ`.
    Iroha,
    /// Katakana in iroha order, full-width — `イ`, `ロ`, `ハ`.
    IrohaFullWidth,
    /// Hangul syllables in ganada order — `가`, `나`, `다`.
    Ganada,
    /// Hangul leading jamo (chosung) — `ㄱ`, `ㄴ`, `ㄷ`.
    Chosung,
    /// The Hebrew *alphabet* — `א`, `ב`, `ג`. Distinct from
    /// [`Hebrew1`](Self::Hebrew1), which is the numeral system.
    Hebrew2,
    /// The Arabic alphabet in modern hijāʾī order — `ا`, `ب`, `ت`.
    ArabicAlpha,
    /// Devanagari vowels — `अ`, `आ`, `इ`.
    HindiVowels,
    /// Devanagari consonants — `क`, `ख`, `ग`.
    HindiConsonants,
    /// Thai consonants — `ก`, `ข`, `ฃ`.
    ThaiLetters,
    /// The Chicago Manual of Style footnote symbols — `*`, `†`, `‡`, `§`,
    /// then doubled.
    Chicago,
    /// The ten Heavenly Stems — `甲`, `乙`, `丙`.
    IdeographTraditional,
    /// The twelve Earthly Branches — `子`, `丑`, `寅`.
    IdeographZodiac,
    /// The sexagenary cycle, stem paired with branch — `甲子`, `乙丑`.
    IdeographZodiacTraditional,

    // ── Closed numeral algorithm ──────────────────────────────────────────
    // Additive numerals with a fixed value table, the shape `lowerRoman`
    // already has.
    /// Hebrew numerals (gematria) — `א`=1, `י`=10, `טו`=15.
    Hebrew1,
    /// Arabic abjad numerals — `ا`=1, `ي`=10, `ق`=100.
    ArabicAbjad,

    /// §17.18.59's remaining values, which this engine renders as decimal.
    ///
    /// This is the **exception to the strict-enum rule**: a large, extensible
    /// value space where a legal-but-unsupported value must degrade, not fail
    /// — otherwise one exotic `<w:numFmt>` would fail the *whole* document
    /// parse (`parse_numbering(..)?` propagates the error). Word itself falls
    /// back to decimal for formats it can't render.
    ///
    /// # What is left here, and why
    ///
    /// Issue #132 classified all 63 §17.18.59 values by *what data* rendering
    /// them needs. The 31 above need none — a digit set, a wrapper, an
    /// alphabet or a value table, each of which is in the source rather than
    /// in CLDR. These need language data this engine does not carry:
    ///
    /// * **Counting systems** — `japaneseCounting`, `japaneseLegal`,
    ///   `japaneseDigitalTenThousand`, `chineseCounting`,
    ///   `chineseCountingThousand`, `chineseLegalSimplified`,
    ///   `taiwaneseCounting`, `taiwaneseCountingThousand`, `taiwaneseDigital`,
    ///   `ideographLegalTraditional`, `koreanCounting`, `koreanLegal`,
    ///   `koreanDigital`, `koreanDigital2`, `vietnameseCounting`,
    ///   `hindiCounting`, `thaiCounting`. These *look* like digits and are
    ///   spellout: `chineseCounting` writes 12 as `十二`, twelve read aloud,
    ///   not two positional digits. Each needs its own language's rules,
    ///   which is the same data `cardinalText` needs — see
    ///   `crate::render::resolve::spellout` for why that data is hand-written
    ///   here and what the alternative cost.
    /// * **Spellout with a currency** — `bahtText`, `dollarText`.
    /// * **`custom`** — §17.9.30's picture string. Not a format at all: a
    ///   template the consumer evaluates, and a separate feature.
    #[serde(other)]
    Other,
}

impl From<StNumberFormat> for NumberFormat {
    fn from(s: StNumberFormat) -> Self {
        match s {
            StNumberFormat::Decimal => Self::Decimal,
            StNumberFormat::UpperRoman => Self::UpperRoman,
            StNumberFormat::LowerRoman => Self::LowerRoman,
            StNumberFormat::UpperLetter => Self::UpperLetter,
            StNumberFormat::LowerLetter => Self::LowerLetter,
            StNumberFormat::Bullet => Self::Bullet,
            StNumberFormat::Ordinal => Self::Ordinal,
            StNumberFormat::CardinalText => Self::CardinalText,
            StNumberFormat::OrdinalText => Self::OrdinalText,
            StNumberFormat::RussianUpper => Self::RussianUpper,
            StNumberFormat::RussianLower => Self::RussianLower,
            StNumberFormat::None => Self::None,

            StNumberFormat::DecimalFullWidth => Self::DecimalFullWidth,
            StNumberFormat::DecimalFullWidth2 => Self::DecimalFullWidth2,
            StNumberFormat::DecimalHalfWidth => Self::DecimalHalfWidth,
            StNumberFormat::HindiNumbers => Self::HindiNumbers,
            StNumberFormat::ThaiNumbers => Self::ThaiNumbers,
            StNumberFormat::IdeographDigital => Self::IdeographDigital,

            StNumberFormat::DecimalZero => Self::DecimalZero,
            StNumberFormat::Hex => Self::Hex,
            StNumberFormat::NumberInDash => Self::NumberInDash,
            StNumberFormat::DecimalEnclosedFullstop => Self::DecimalEnclosedFullstop,
            StNumberFormat::DecimalEnclosedParen => Self::DecimalEnclosedParen,
            StNumberFormat::DecimalEnclosedCircle => Self::DecimalEnclosedCircle,
            StNumberFormat::DecimalEnclosedCircleChinese => Self::DecimalEnclosedCircleChinese,
            StNumberFormat::IdeographEnclosedCircle => Self::IdeographEnclosedCircle,

            StNumberFormat::Aiueo => Self::Aiueo,
            StNumberFormat::AiueoFullWidth => Self::AiueoFullWidth,
            StNumberFormat::Iroha => Self::Iroha,
            StNumberFormat::IrohaFullWidth => Self::IrohaFullWidth,
            StNumberFormat::Ganada => Self::Ganada,
            StNumberFormat::Chosung => Self::Chosung,
            StNumberFormat::Hebrew2 => Self::Hebrew2,
            StNumberFormat::ArabicAlpha => Self::ArabicAlpha,
            StNumberFormat::HindiVowels => Self::HindiVowels,
            StNumberFormat::HindiConsonants => Self::HindiConsonants,
            StNumberFormat::ThaiLetters => Self::ThaiLetters,
            StNumberFormat::Chicago => Self::Chicago,
            StNumberFormat::IdeographTraditional => Self::IdeographTraditional,
            StNumberFormat::IdeographZodiac => Self::IdeographZodiac,
            StNumberFormat::IdeographZodiacTraditional => Self::IdeographZodiacTraditional,

            StNumberFormat::Hebrew1 => Self::Hebrew1,
            StNumberFormat::ArabicAbjad => Self::ArabicAbjad,

            StNumberFormat::Other => Self::Decimal,
        }
    }
}

// ── StPageOrientation (§17.18.65) ─────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StPageOrientation {
    Portrait,
    Landscape,
}

impl From<StPageOrientation> for PageOrientation {
    fn from(s: StPageOrientation) -> Self {
        match s {
            StPageOrientation::Portrait => Self::Portrait,
            StPageOrientation::Landscape => Self::Landscape,
        }
    }
}

// ── StSectionMark (§17.18.77) ─────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StSectionMark {
    NextPage,
    Continuous,
    EvenPage,
    OddPage,
    NextColumn,
}

impl From<StSectionMark> for SectionType {
    fn from(s: StSectionMark) -> Self {
        match s {
            StSectionMark::NextPage => Self::NextPage,
            StSectionMark::Continuous => Self::Continuous,
            StSectionMark::EvenPage => Self::EvenPage,
            StSectionMark::OddPage => Self::OddPage,
            StSectionMark::NextColumn => Self::NextColumn,
        }
    }
}

// ── StShd (§17.18.78) ─────────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StShd {
    /// §17.18.78 ST_Shd: `nil` — no shading whatsoever (distinct from
    /// `clear`, which means "transparent shading with a defined color").
    Nil,
    Clear,
    Solid,
    HorzStripe,
    VertStripe,
    ReverseDiagStripe,
    DiagStripe,
    HorzCross,
    DiagCross,
    ThinHorzStripe,
    ThinVertStripe,
    ThinReverseDiagStripe,
    ThinDiagStripe,
    ThinHorzCross,
    ThinDiagCross,
    Pct5,
    Pct10,
    Pct12,
    Pct15,
    Pct20,
    Pct25,
    Pct30,
    Pct35,
    Pct37,
    Pct40,
    Pct45,
    Pct50,
    Pct55,
    Pct60,
    Pct62,
    Pct65,
    Pct70,
    Pct75,
    Pct80,
    Pct85,
    Pct87,
    Pct90,
    Pct95,
}

impl From<StShd> for ShadingPattern {
    fn from(s: StShd) -> Self {
        match s {
            StShd::Nil => Self::Nil,
            StShd::Clear => Self::Clear,
            StShd::Solid => Self::Solid,
            StShd::HorzStripe => Self::HorzStripe,
            StShd::VertStripe => Self::VertStripe,
            StShd::ReverseDiagStripe => Self::ReverseDiagStripe,
            StShd::DiagStripe => Self::DiagStripe,
            StShd::HorzCross => Self::HorzCross,
            StShd::DiagCross => Self::DiagCross,
            StShd::ThinHorzStripe => Self::ThinHorzStripe,
            StShd::ThinVertStripe => Self::ThinVertStripe,
            StShd::ThinReverseDiagStripe => Self::ThinReverseDiagStripe,
            StShd::ThinDiagStripe => Self::ThinDiagStripe,
            StShd::ThinHorzCross => Self::ThinHorzCross,
            StShd::ThinDiagCross => Self::ThinDiagCross,
            StShd::Pct5 => Self::Pct5,
            StShd::Pct10 => Self::Pct10,
            StShd::Pct12 => Self::Pct12,
            StShd::Pct15 => Self::Pct15,
            StShd::Pct20 => Self::Pct20,
            StShd::Pct25 => Self::Pct25,
            StShd::Pct30 => Self::Pct30,
            StShd::Pct35 => Self::Pct35,
            StShd::Pct37 => Self::Pct37,
            StShd::Pct40 => Self::Pct40,
            StShd::Pct45 => Self::Pct45,
            StShd::Pct50 => Self::Pct50,
            StShd::Pct55 => Self::Pct55,
            StShd::Pct60 => Self::Pct60,
            StShd::Pct62 => Self::Pct62,
            StShd::Pct65 => Self::Pct65,
            StShd::Pct70 => Self::Pct70,
            StShd::Pct75 => Self::Pct75,
            StShd::Pct80 => Self::Pct80,
            StShd::Pct85 => Self::Pct85,
            StShd::Pct87 => Self::Pct87,
            StShd::Pct90 => Self::Pct90,
            StShd::Pct95 => Self::Pct95,
        }
    }
}

// ── StHAnchor/StVAnchor (§17.18.35/106) ─────────────────────────────────
//
// Shared by table positioning (`<w:tblpPr>`) and frame positioning
// (`<w:framePr>`). OOXML uses the same tag set in both spots.

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StAnchor {
    Text,
    Margin,
    Page,
}

impl From<StAnchor> for TableAnchor {
    fn from(s: StAnchor) -> Self {
        match s {
            StAnchor::Text => Self::Text,
            StAnchor::Margin => Self::Margin,
            StAnchor::Page => Self::Page,
        }
    }
}

// ── StXAlign (§17.18.108) ────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StXAlign {
    Left,
    Center,
    Right,
    Inside,
    Outside,
}

impl From<StXAlign> for TableXAlign {
    fn from(s: StXAlign) -> Self {
        match s {
            StXAlign::Left => Self::Left,
            StXAlign::Center => Self::Center,
            StXAlign::Right => Self::Right,
            StXAlign::Inside => Self::Inside,
            StXAlign::Outside => Self::Outside,
        }
    }
}

// ── StYAlign (§17.18.109) ────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StYAlign {
    Top,
    Center,
    Bottom,
    Inside,
    Outside,
    Inline,
}

impl From<StYAlign> for TableYAlign {
    fn from(s: StYAlign) -> Self {
        match s {
            StYAlign::Top => Self::Top,
            StYAlign::Center => Self::Center,
            StYAlign::Bottom => Self::Bottom,
            StYAlign::Inside => Self::Inside,
            StYAlign::Outside => Self::Outside,
            StYAlign::Inline => Self::Inline,
        }
    }
}

// ── StTabJc (§17.18.85 tab alignment) ─────────────────────────────────────

/// ISO/IEC 29500 Strict spells the directional values `start`/`end`; the
/// Transitional→Strict migration maps `left`↔`start` and `right`↔`end`
/// losslessly, the same aliasing [`StJc`] carries — and losing it here was
/// worse than a wrong stop: an unknown variant fails deserialization, so one
/// Strict `w:tab` made the whole document unreadable.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTabJc {
    #[serde(alias = "start")]
    Left,
    Center,
    #[serde(alias = "end")]
    Right,
    Decimal,
    Bar,
    Clear,
    /// Legacy — treated as `Left`.
    Num,
}

impl From<StTabJc> for TabAlignment {
    fn from(s: StTabJc) -> Self {
        match s {
            StTabJc::Left | StTabJc::Num => Self::Left,
            StTabJc::Center => Self::Center,
            StTabJc::Right => Self::Right,
            StTabJc::Decimal => Self::Decimal,
            StTabJc::Bar => Self::Bar,
            StTabJc::Clear => Self::Clear,
        }
    }
}

// ── StTabTlc (§17.18.86 tab leader character) ─────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTabTlc {
    None,
    Dot,
    Hyphen,
    Underscore,
    Heavy,
    MiddleDot,
}

impl From<StTabTlc> for TabLeader {
    fn from(s: StTabTlc) -> Self {
        match s {
            StTabTlc::None => Self::None,
            StTabTlc::Dot => Self::Dot,
            StTabTlc::Hyphen => Self::Hyphen,
            StTabTlc::Underscore => Self::Underscore,
            StTabTlc::Heavy => Self::Heavy,
            StTabTlc::MiddleDot => Self::MiddleDot,
        }
    }
}

// ── StPTabAlignment (§17.18.59 absolute position tab alignment) ────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StPTabAlignment {
    Left,
    Center,
    Right,
}

impl From<StPTabAlignment> for PTabAlignment {
    fn from(s: StPTabAlignment) -> Self {
        match s {
            StPTabAlignment::Left => Self::Left,
            StPTabAlignment::Center => Self::Center,
            StPTabAlignment::Right => Self::Right,
        }
    }
}

// ── StPTabRelativeTo (§17.18.61 absolute position tab base) ────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StPTabRelativeTo {
    Margin,
    Indent,
}

impl From<StPTabRelativeTo> for PTabRelativeTo {
    fn from(s: StPTabRelativeTo) -> Self {
        match s {
            StPTabRelativeTo::Margin => Self::Margin,
            StPTabRelativeTo::Indent => Self::Indent,
        }
    }
}

// ── StPTabLeader (§17.18.60 absolute position tab leader) ──────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StPTabLeader {
    None,
    Dot,
    Hyphen,
    Underscore,
    MiddleDot,
}

impl From<StPTabLeader> for PTabLeader {
    fn from(s: StPTabLeader) -> Self {
        match s {
            StPTabLeader::None => Self::None,
            StPTabLeader::Dot => Self::Dot,
            StPTabLeader::Hyphen => Self::Hyphen,
            StPTabLeader::Underscore => Self::Underscore,
            StPTabLeader::MiddleDot => Self::MiddleDot,
        }
    }
}

// ── StTblLayoutType (§17.18.87) ───────────────────────────────────────────

/// §17.18.87 `ST_TblLayoutType` — `fixed` | **`autofit`**. Not `auto`.
///
/// # The spec says `auto` twice, and means neither time
///
/// §17.4.52 and §17.4.53 both close with "If this element is omitted, then the
/// value of this element shall be assumed to be **auto**" — a value the
/// enumeration does not contain. Annex A settles it both ways over: the XSD
/// restriction lists `<xsd:enumeration value="fixed"/>` and
/// `<xsd:enumeration value="autofit"/>`, the RELAX NG grammar reads
/// `w_ST_TblLayoutType = string "fixed" | string "autofit"`, and §17.18.87's
/// own value table names `autofit` and `fixed`. So the prose default is a
/// typo for `autofit`; [MS-OI29500] Part 1 §2.1.158(b) reads it that way too.
/// Nothing in the spec permits a producer to write `auto`, and this enum
/// accepted only that — so `<w:tblLayout w:type="autofit"/>`, the sole
/// spelling that can name the mode, failed deserialization, and because the
/// `ST_*` catalogue is strict by design (module doc) that rejected the entire
/// document rather than the one attribute.
///
/// The same subclauses say the algorithms are "discussed in the simple type
/// referenced by the **val** attribute" while the attribute is `@type` —
/// `docx::parse::properties::schema::table::TblLayoutXml` already carries a
/// note about that one.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTblLayoutType {
    Autofit,
    Fixed,
}

impl From<StTblLayoutType> for TableLayout {
    fn from(s: StTblLayoutType) -> Self {
        match s {
            StTblLayoutType::Autofit => Self::Autofit,
            StTblLayoutType::Fixed => Self::Fixed,
        }
    }
}

// ── StTblOverlap (§17.4.57) ───────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTblOverlap {
    Overlap,
    Never,
}

impl From<StTblOverlap> for TableOverlap {
    fn from(s: StTblOverlap) -> Self {
        match s {
            StTblOverlap::Overlap => Self::Overlap,
            StTblOverlap::Never => Self::Never,
        }
    }
}

// ── StTextAlignment (§17.18.91) ───────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTextAlignment {
    Auto,
    Top,
    Center,
    Baseline,
    Bottom,
}

impl From<StTextAlignment> for TextAlignment {
    fn from(s: StTextAlignment) -> Self {
        match s {
            StTextAlignment::Auto => Self::Auto,
            StTextAlignment::Top => Self::Top,
            StTextAlignment::Center => Self::Center,
            StTextAlignment::Baseline => Self::Baseline,
            StTextAlignment::Bottom => Self::Bottom,
        }
    }
}

// ── StTextDirection (§17.18.93) ───────────────────────────────────────────
//
// OOXML uses opaque two/three-letter codes; the model spells the full
// directional order. Mapping is static.

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
pub enum StTextDirection {
    #[serde(rename = "lrTb")]
    LrTb,
    #[serde(rename = "tbRl")]
    TbRl,
    #[serde(rename = "btLr")]
    BtLr,
    #[serde(rename = "lrTbV")]
    LrTbV,
    #[serde(rename = "tbRlV")]
    TbRlV,
    #[serde(rename = "tbLrV")]
    TbLrV,
}

impl From<StTextDirection> for TextDirection {
    fn from(s: StTextDirection) -> Self {
        match s {
            StTextDirection::LrTb => Self::LeftToRightTopToBottom,
            StTextDirection::TbRl => Self::TopToBottomRightToLeft,
            StTextDirection::BtLr => Self::BottomToTopLeftToRight,
            StTextDirection::LrTbV => Self::LeftToRightTopToBottomRotated,
            StTextDirection::TbRlV => Self::TopToBottomRightToLeftRotated,
            StTextDirection::TbLrV => Self::TopToBottomLeftToRightRotated,
        }
    }
}

// ── StTheme (§17.18.95) ───────────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StTheme {
    MajorHAnsi,
    MajorEastAsia,
    MajorBidi,
    MinorHAnsi,
    MinorEastAsia,
    MinorBidi,
}

impl From<StTheme> for ThemeFontRef {
    fn from(s: StTheme) -> Self {
        match s {
            StTheme::MajorHAnsi => Self::MajorHAnsi,
            StTheme::MajorEastAsia => Self::MajorEastAsia,
            StTheme::MajorBidi => Self::MajorBidi,
            StTheme::MinorHAnsi => Self::MinorHAnsi,
            StTheme::MinorEastAsia => Self::MinorEastAsia,
            StTheme::MinorBidi => Self::MinorBidi,
        }
    }
}

// ── StUnderline (§17.18.99) ───────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StUnderline {
    None,
    Single,
    Words,
    Double,
    Thick,
    Dotted,
    DottedHeavy,
    Dash,
    DashedHeavy,
    DashLong,
    DashLongHeavy,
    DotDash,
    DashDotHeavy,
    DotDotDash,
    DashDotDotHeavy,
    Wave,
    WavyHeavy,
    WavyDouble,
}

impl From<StUnderline> for UnderlineStyle {
    fn from(s: StUnderline) -> Self {
        match s {
            StUnderline::None => Self::None,
            StUnderline::Single => Self::Single,
            StUnderline::Words => Self::Words,
            StUnderline::Double => Self::Double,
            StUnderline::Thick => Self::Thick,
            StUnderline::Dotted => Self::Dotted,
            StUnderline::DottedHeavy => Self::DottedHeavy,
            StUnderline::Dash => Self::Dash,
            StUnderline::DashedHeavy => Self::DashedHeavy,
            StUnderline::DashLong => Self::DashLong,
            StUnderline::DashLongHeavy => Self::DashLongHeavy,
            StUnderline::DotDash => Self::DotDash,
            StUnderline::DashDotHeavy => Self::DashDotHeavy,
            StUnderline::DotDotDash => Self::DotDotDash,
            StUnderline::DashDotDotHeavy => Self::DashDotDotHeavy,
            StUnderline::Wave => Self::Wave,
            StUnderline::WavyHeavy => Self::WavyHeavy,
            StUnderline::WavyDouble => Self::WavyDouble,
        }
    }
}

// ── StVerticalAlignRun (§17.18.100) ───────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StVerticalAlignRun {
    Baseline,
    Superscript,
    Subscript,
}

impl From<StVerticalAlignRun> for VerticalAlign {
    fn from(s: StVerticalAlignRun) -> Self {
        match s {
            StVerticalAlignRun::Baseline => Self::Baseline,
            StVerticalAlignRun::Superscript => Self::Superscript,
            StVerticalAlignRun::Subscript => Self::Subscript,
        }
    }
}

// ── StVerticalJc (§17.18.101) ─────────────────────────────────────────────

#[derive(Clone, Copy, Debug, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum StVerticalJc {
    Top,
    Center,
    Bottom,
    Both,
}

impl From<StVerticalJc> for CellVerticalAlign {
    fn from(s: StVerticalJc) -> Self {
        match s {
            StVerticalJc::Top => Self::Top,
            StVerticalJc::Center => Self::Center,
            StVerticalJc::Bottom => Self::Bottom,
            StVerticalJc::Both => Self::Both,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde::de::DeserializeOwned;

    fn de<T: DeserializeOwned>(v: &str) -> Result<T, quick_xml::DeError> {
        #[derive(Deserialize)]
        struct Wrap<X> {
            #[serde(rename = "@v")]
            v: X,
        }
        quick_xml::de::from_str::<Wrap<T>>(&format!(r#"<x v="{v}"/>"#)).map(|w| w.v)
    }

    fn assert_bad<T: DeserializeOwned + std::fmt::Debug>(v: &str) {
        let r: Result<T, _> = de(v);
        assert!(r.is_err(), "expected error for {v:?}, got {r:?}");
    }

    // ── StBorderType ──
    #[test]
    fn border_type_all_variants() {
        assert_eq!(de::<StBorderType>("none").unwrap(), StBorderType::None);
        assert_eq!(de::<StBorderType>("single").unwrap(), StBorderType::Single);
        assert_eq!(
            de::<StBorderType>("dotDash").unwrap(),
            StBorderType::DotDash
        );
        assert_eq!(
            de::<StBorderType>("threeDEmboss").unwrap(),
            StBorderType::ThreeDEmboss
        );
        assert_eq!(
            de::<StBorderType>("threeDEngrave").unwrap(),
            StBorderType::ThreeDEngrave
        );
        assert_eq!(de::<StBorderType>("inset").unwrap(), StBorderType::Inset);
    }
    #[test]
    fn border_type_strict() {
        assert_bad::<StBorderType>("bogus");
    }
    #[test]
    fn border_type_converts_to_model() {
        let m: BorderStyle = StBorderType::DashDotStroked.into();
        assert_eq!(m, BorderStyle::DashDotStroked);
    }

    // ── StBrClear ──
    #[test]
    fn br_clear_all_variants() {
        assert_eq!(de::<StBrClear>("none").unwrap(), StBrClear::None);
        assert_eq!(de::<StBrClear>("left").unwrap(), StBrClear::Left);
        assert_eq!(de::<StBrClear>("right").unwrap(), StBrClear::Right);
        assert_eq!(de::<StBrClear>("all").unwrap(), StBrClear::All);
    }
    #[test]
    fn br_clear_strict() {
        assert_bad::<StBrClear>("middle");
    }
    #[test]
    fn br_clear_converts_to_model() {
        let m: BreakClear = StBrClear::Left.into();
        assert_eq!(m, BreakClear::Left);
    }

    // ── StFldCharType ──
    #[test]
    fn fld_char_type_all_variants() {
        assert_eq!(de::<StFldCharType>("begin").unwrap(), StFldCharType::Begin);
        assert_eq!(
            de::<StFldCharType>("separate").unwrap(),
            StFldCharType::Separate
        );
        assert_eq!(de::<StFldCharType>("end").unwrap(), StFldCharType::End);
    }
    #[test]
    fn fld_char_type_strict() {
        assert_bad::<StFldCharType>("middle");
    }

    // ── StFrameWrap ──
    #[test]
    fn frame_wrap_all_variants() {
        assert_eq!(de::<StFrameWrap>("auto").unwrap(), StFrameWrap::Auto);
        assert_eq!(
            de::<StFrameWrap>("notBeside").unwrap(),
            StFrameWrap::NotBeside
        );
        assert_eq!(de::<StFrameWrap>("around").unwrap(), StFrameWrap::Around);
        assert_eq!(de::<StFrameWrap>("tight").unwrap(), StFrameWrap::Tight);
        assert_eq!(de::<StFrameWrap>("through").unwrap(), StFrameWrap::Through);
        assert_eq!(de::<StFrameWrap>("none").unwrap(), StFrameWrap::None);
    }
    #[test]
    fn frame_wrap_strict() {
        assert_bad::<StFrameWrap>("wrap");
    }

    // ── StHeightRule ──
    #[test]
    fn height_rule_all_variants() {
        assert_eq!(de::<StHeightRule>("auto").unwrap(), StHeightRule::Auto);
        assert_eq!(de::<StHeightRule>("exact").unwrap(), StHeightRule::Exact);
        assert_eq!(
            de::<StHeightRule>("atLeast").unwrap(),
            StHeightRule::AtLeast
        );
    }
    #[test]
    fn height_rule_strict() {
        assert_bad::<StHeightRule>("maximum");
    }

    // ── StHighlightColor ──
    #[test]
    fn highlight_color_sample_variants() {
        assert_eq!(
            de::<StHighlightColor>("black").unwrap(),
            StHighlightColor::Black
        );
        assert_eq!(
            de::<StHighlightColor>("darkMagenta").unwrap(),
            StHighlightColor::DarkMagenta
        );
        assert_eq!(
            de::<StHighlightColor>("lightGray").unwrap(),
            StHighlightColor::LightGray
        );
        assert_eq!(
            de::<StHighlightColor>("yellow").unwrap(),
            StHighlightColor::Yellow
        );
    }
    #[test]
    fn highlight_color_strict() {
        assert_bad::<StHighlightColor>("chartreuse");
    }

    // ── StJc — includes the both/justify alias and Start/End rename ──
    #[test]
    fn jc_all_variants_and_aliases() {
        assert_eq!(de::<StJc>("left").unwrap(), StJc::Left);
        assert_eq!(de::<StJc>("start").unwrap(), StJc::Left); // alias
        assert_eq!(de::<StJc>("center").unwrap(), StJc::Center);
        assert_eq!(de::<StJc>("right").unwrap(), StJc::Right);
        assert_eq!(de::<StJc>("end").unwrap(), StJc::Right); // alias
        assert_eq!(de::<StJc>("both").unwrap(), StJc::Both);
        assert_eq!(de::<StJc>("justify").unwrap(), StJc::Both); // alias
        assert_eq!(de::<StJc>("distribute").unwrap(), StJc::Distribute);
        assert_eq!(de::<StJc>("thaiDistribute").unwrap(), StJc::ThaiDistribute);
    }
    #[test]
    fn jc_kashida_and_numtab_are_legal_and_degrade() {
        // §17.18.44: spec-legal values that would otherwise crash an Arabic /
        // legacy-numbered document's parse. They parse and degrade to the
        // nearest modelled alignment.
        assert_eq!(de::<StJc>("mediumKashida").unwrap(), StJc::MediumKashida);
        assert_eq!(de::<StJc>("highKashida").unwrap(), StJc::HighKashida);
        assert_eq!(de::<StJc>("lowKashida").unwrap(), StJc::LowKashida);
        assert_eq!(de::<StJc>("numTab").unwrap(), StJc::NumTab);
    }
    #[test]
    fn jc_strict() {
        // A genuine typo still fails — the kashida/numTab additions are named
        // variants, not a catch-all.
        assert_bad::<StJc>("middle");
    }
    #[test]
    fn jc_converts_to_model_with_rename() {
        assert_eq!(Alignment::from(StJc::Left), Alignment::Start);
        assert_eq!(Alignment::from(StJc::Right), Alignment::End);
        assert_eq!(Alignment::from(StJc::ThaiDistribute), Alignment::Thai);
        assert_eq!(Alignment::from(StJc::MediumKashida), Alignment::Both);
        assert_eq!(Alignment::from(StJc::NumTab), Alignment::Start);
    }

    // ── StNumberFormat ──
    #[test]
    fn number_format_all_variants() {
        assert_eq!(
            de::<StNumberFormat>("decimal").unwrap(),
            StNumberFormat::Decimal
        );
        assert_eq!(
            de::<StNumberFormat>("upperRoman").unwrap(),
            StNumberFormat::UpperRoman
        );
        assert_eq!(
            de::<StNumberFormat>("cardinalText").unwrap(),
            StNumberFormat::CardinalText
        );
        assert_eq!(
            de::<StNumberFormat>("bullet").unwrap(),
            StNumberFormat::Bullet
        );
        assert_eq!(
            de::<StNumberFormat>("russianUpper").unwrap(),
            StNumberFormat::RussianUpper
        );
        assert_eq!(
            de::<StNumberFormat>("russianLower").unwrap(),
            StNumberFormat::RussianLower
        );
        assert_eq!(
            NumberFormat::from(StNumberFormat::RussianUpper),
            NumberFormat::RussianUpper
        );
    }

    /// The §17.18.59 values issue #132 closed: each parses to its own variant
    /// and reaches the model as itself, rather than collapsing to decimal.
    /// One per classification group, plus the two spellings whose camelCase
    /// mapping is not obvious (`hebrew1`, `decimalFullWidth2`).
    #[test]
    fn number_format_sequence_values_parse_to_their_own_variants() {
        for (v, want) in [
            ("decimalFullWidth", StNumberFormat::DecimalFullWidth),
            ("decimalFullWidth2", StNumberFormat::DecimalFullWidth2),
            ("decimalHalfWidth", StNumberFormat::DecimalHalfWidth),
            ("thaiNumbers", StNumberFormat::ThaiNumbers),
            ("decimalZero", StNumberFormat::DecimalZero),
            ("hex", StNumberFormat::Hex),
            ("numberInDash", StNumberFormat::NumberInDash),
            (
                "decimalEnclosedCircle",
                StNumberFormat::DecimalEnclosedCircle,
            ),
            ("aiueoFullWidth", StNumberFormat::AiueoFullWidth),
            ("iroha", StNumberFormat::Iroha),
            ("chosung", StNumberFormat::Chosung),
            ("chicago", StNumberFormat::Chicago),
            (
                "ideographZodiacTraditional",
                StNumberFormat::IdeographZodiacTraditional,
            ),
            ("hebrew1", StNumberFormat::Hebrew1),
            ("hebrew2", StNumberFormat::Hebrew2),
            ("arabicAbjad", StNumberFormat::ArabicAbjad),
        ] {
            assert_eq!(de::<StNumberFormat>(v).unwrap(), want, "{v:?}");
        }
        assert_eq!(
            NumberFormat::from(StNumberFormat::Hebrew1),
            NumberFormat::Hebrew1,
            "and it reaches the model as itself, not as decimal",
        );
    }

    #[test]
    fn number_format_unsupported_legal_values_degrade_not_fail() {
        // §17.18.59's remaining values must parse (→ Other) and convert to
        // Decimal rather than failing the whole document parse. All of these
        // are spellout — a counting system, a currency, or §17.9.30's picture
        // string — which is what keeps them on this side of the boundary; see
        // `StNumberFormat::Other`.
        for v in [
            "japaneseCounting",
            "chineseCountingThousand",
            "koreanDigital2",
            "vietnameseCounting",
            "bahtText",
            "dollarText",
            "custom",
        ] {
            assert_eq!(
                de::<StNumberFormat>(v).unwrap(),
                StNumberFormat::Other,
                "{v:?} is spec-legal and must degrade to Other, not error"
            );
        }
        assert_eq!(
            NumberFormat::from(StNumberFormat::Other),
            NumberFormat::Decimal
        );
    }

    // ── StPageOrientation ──
    #[test]
    fn page_orientation_both() {
        assert_eq!(
            de::<StPageOrientation>("portrait").unwrap(),
            StPageOrientation::Portrait
        );
        assert_eq!(
            de::<StPageOrientation>("landscape").unwrap(),
            StPageOrientation::Landscape
        );
    }
    #[test]
    fn page_orientation_strict() {
        assert_bad::<StPageOrientation>("sideways");
    }

    // ── StSectionMark ──
    #[test]
    fn section_mark_all_variants() {
        assert_eq!(
            de::<StSectionMark>("nextPage").unwrap(),
            StSectionMark::NextPage
        );
        assert_eq!(
            de::<StSectionMark>("continuous").unwrap(),
            StSectionMark::Continuous
        );
        assert_eq!(
            de::<StSectionMark>("evenPage").unwrap(),
            StSectionMark::EvenPage
        );
        assert_eq!(
            de::<StSectionMark>("oddPage").unwrap(),
            StSectionMark::OddPage
        );
        assert_eq!(
            de::<StSectionMark>("nextColumn").unwrap(),
            StSectionMark::NextColumn
        );
    }
    #[test]
    fn section_mark_strict() {
        assert_bad::<StSectionMark>("previous");
    }

    // ── StShd ──
    #[test]
    fn shd_sample_variants() {
        assert_eq!(de::<StShd>("clear").unwrap(), StShd::Clear);
        assert_eq!(de::<StShd>("solid").unwrap(), StShd::Solid);
        assert_eq!(de::<StShd>("horzStripe").unwrap(), StShd::HorzStripe);
        assert_eq!(de::<StShd>("thinDiagCross").unwrap(), StShd::ThinDiagCross);
        assert_eq!(de::<StShd>("pct5").unwrap(), StShd::Pct5);
        assert_eq!(de::<StShd>("pct95").unwrap(), StShd::Pct95);
    }
    #[test]
    fn shd_strict() {
        assert_bad::<StShd>("pct100");
    }

    // ── StTblLayoutType ──
    #[test]
    fn tbl_layout_type_both() {
        assert_eq!(
            de::<StTblLayoutType>("autofit").unwrap(),
            StTblLayoutType::Autofit
        );
        assert_eq!(
            de::<StTblLayoutType>("fixed").unwrap(),
            StTblLayoutType::Fixed
        );
    }
    #[test]
    fn tbl_layout_type_strict() {
        assert_bad::<StTblLayoutType>("flex");
        // `auto` is the word §17.4.52/§17.4.53's *prose* uses for the default,
        // and it is not a value of the type — see `StTblLayoutType`. A parser
        // that accepts it accepts something no producer may write, and this
        // enum accepted it while rejecting `autofit`, which every producer may.
        assert_bad::<StTblLayoutType>("auto");
    }

    // ── StTblOverlap ──
    #[test]
    fn tbl_overlap_both() {
        assert_eq!(
            de::<StTblOverlap>("overlap").unwrap(),
            StTblOverlap::Overlap
        );
        assert_eq!(de::<StTblOverlap>("never").unwrap(), StTblOverlap::Never);
    }
    #[test]
    fn tbl_overlap_strict() {
        assert_bad::<StTblOverlap>("always");
    }

    // ── StTextAlignment ──
    #[test]
    fn text_alignment_all_variants() {
        assert_eq!(
            de::<StTextAlignment>("auto").unwrap(),
            StTextAlignment::Auto
        );
        assert_eq!(de::<StTextAlignment>("top").unwrap(), StTextAlignment::Top);
        assert_eq!(
            de::<StTextAlignment>("center").unwrap(),
            StTextAlignment::Center
        );
        assert_eq!(
            de::<StTextAlignment>("baseline").unwrap(),
            StTextAlignment::Baseline
        );
        assert_eq!(
            de::<StTextAlignment>("bottom").unwrap(),
            StTextAlignment::Bottom
        );
    }
    #[test]
    fn text_alignment_strict() {
        assert_bad::<StTextAlignment>("middle");
    }

    // ── StTextDirection ──
    #[test]
    fn text_direction_all_variants() {
        assert_eq!(
            de::<StTextDirection>("lrTb").unwrap(),
            StTextDirection::LrTb
        );
        assert_eq!(
            de::<StTextDirection>("tbRl").unwrap(),
            StTextDirection::TbRl
        );
        assert_eq!(
            de::<StTextDirection>("btLr").unwrap(),
            StTextDirection::BtLr
        );
        assert_eq!(
            de::<StTextDirection>("lrTbV").unwrap(),
            StTextDirection::LrTbV
        );
        assert_eq!(
            de::<StTextDirection>("tbRlV").unwrap(),
            StTextDirection::TbRlV
        );
        assert_eq!(
            de::<StTextDirection>("tbLrV").unwrap(),
            StTextDirection::TbLrV
        );
    }
    #[test]
    fn text_direction_strict() {
        assert_bad::<StTextDirection>("ltr");
    }
    #[test]
    fn text_direction_converts_to_model() {
        assert_eq!(
            TextDirection::from(StTextDirection::LrTb),
            TextDirection::LeftToRightTopToBottom
        );
        assert_eq!(
            TextDirection::from(StTextDirection::TbLrV),
            TextDirection::TopToBottomLeftToRightRotated
        );
    }

    // ── StTheme ──
    #[test]
    fn theme_all_variants() {
        assert_eq!(de::<StTheme>("majorHAnsi").unwrap(), StTheme::MajorHAnsi);
        assert_eq!(
            de::<StTheme>("majorEastAsia").unwrap(),
            StTheme::MajorEastAsia
        );
        assert_eq!(de::<StTheme>("majorBidi").unwrap(), StTheme::MajorBidi);
        assert_eq!(de::<StTheme>("minorHAnsi").unwrap(), StTheme::MinorHAnsi);
        assert_eq!(
            de::<StTheme>("minorEastAsia").unwrap(),
            StTheme::MinorEastAsia
        );
        assert_eq!(de::<StTheme>("minorBidi").unwrap(), StTheme::MinorBidi);
    }
    #[test]
    fn theme_strict() {
        assert_bad::<StTheme>("default");
    }

    // ── StUnderline ──
    #[test]
    fn underline_sample_variants() {
        assert_eq!(de::<StUnderline>("none").unwrap(), StUnderline::None);
        assert_eq!(de::<StUnderline>("single").unwrap(), StUnderline::Single);
        assert_eq!(de::<StUnderline>("dotted").unwrap(), StUnderline::Dotted);
        assert_eq!(
            de::<StUnderline>("dashDotHeavy").unwrap(),
            StUnderline::DashDotHeavy
        );
        assert_eq!(
            de::<StUnderline>("wavyDouble").unwrap(),
            StUnderline::WavyDouble
        );
    }
    #[test]
    fn underline_strict() {
        assert_bad::<StUnderline>("italic");
    }

    // ── StVerticalAlignRun ──
    #[test]
    fn vertical_align_run_all_variants() {
        assert_eq!(
            de::<StVerticalAlignRun>("baseline").unwrap(),
            StVerticalAlignRun::Baseline
        );
        assert_eq!(
            de::<StVerticalAlignRun>("superscript").unwrap(),
            StVerticalAlignRun::Superscript
        );
        assert_eq!(
            de::<StVerticalAlignRun>("subscript").unwrap(),
            StVerticalAlignRun::Subscript
        );
    }
    #[test]
    fn vertical_align_run_strict() {
        assert_bad::<StVerticalAlignRun>("middle");
    }

    // ── StVerticalJc ──
    #[test]
    fn vertical_jc_all_variants() {
        assert_eq!(de::<StVerticalJc>("top").unwrap(), StVerticalJc::Top);
        assert_eq!(de::<StVerticalJc>("center").unwrap(), StVerticalJc::Center);
        assert_eq!(de::<StVerticalJc>("bottom").unwrap(), StVerticalJc::Bottom);
        assert_eq!(de::<StVerticalJc>("both").unwrap(), StVerticalJc::Both);
    }
    #[test]
    fn vertical_jc_strict() {
        assert_bad::<StVerticalJc>("start");
    }

    // ── StPTabAlignment ──
    #[test]
    fn ptab_alignment_all_variants() {
        assert_eq!(
            de::<StPTabAlignment>("left").unwrap(),
            StPTabAlignment::Left
        );
        assert_eq!(
            de::<StPTabAlignment>("center").unwrap(),
            StPTabAlignment::Center
        );
        assert_eq!(
            de::<StPTabAlignment>("right").unwrap(),
            StPTabAlignment::Right
        );
    }
    #[test]
    fn ptab_alignment_strict() {
        assert_bad::<StPTabAlignment>("decimal");
    }
    #[test]
    fn ptab_alignment_converts_to_model() {
        let m: PTabAlignment = StPTabAlignment::Center.into();
        assert_eq!(m, PTabAlignment::Center);
    }

    // ── StPTabRelativeTo ──
    #[test]
    fn ptab_relative_to_all_variants() {
        assert_eq!(
            de::<StPTabRelativeTo>("margin").unwrap(),
            StPTabRelativeTo::Margin
        );
        assert_eq!(
            de::<StPTabRelativeTo>("indent").unwrap(),
            StPTabRelativeTo::Indent
        );
    }
    #[test]
    fn ptab_relative_to_strict() {
        assert_bad::<StPTabRelativeTo>("page");
    }

    // ── StPTabLeader ──
    #[test]
    fn ptab_leader_all_variants() {
        assert_eq!(de::<StPTabLeader>("none").unwrap(), StPTabLeader::None);
        assert_eq!(de::<StPTabLeader>("dot").unwrap(), StPTabLeader::Dot);
        assert_eq!(de::<StPTabLeader>("hyphen").unwrap(), StPTabLeader::Hyphen);
        assert_eq!(
            de::<StPTabLeader>("underscore").unwrap(),
            StPTabLeader::Underscore
        );
        assert_eq!(
            de::<StPTabLeader>("middleDot").unwrap(),
            StPTabLeader::MiddleDot
        );
    }
    #[test]
    fn ptab_leader_strict() {
        // `heavy` is a regular tab leader (§17.18.86) but not a ptab leader.
        assert_bad::<StPTabLeader>("heavy");
    }
    #[test]
    fn ptab_leader_converts_to_model() {
        let m: PTabLeader = StPTabLeader::MiddleDot.into();
        assert_eq!(m, PTabLeader::MiddleDot);
        // …and the model leader maps onto the shared TabLeader painter enum.
        let t: TabLeader = m.into();
        assert_eq!(t, TabLeader::MiddleDot);
    }
}
