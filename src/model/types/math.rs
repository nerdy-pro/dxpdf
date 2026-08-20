//! Office Math (OMML) — the minimal ADT this engine renders: math runs,
//! superscripts and fractions. Everything else OMML defines (n-ary
//! operators, radicals, delimiters, matrices, …) is dropped with a warning
//! at parse-conversion time; the containing paragraph still renders.

/// The face Word uses for math when the document does not override it via
/// `w:settings/m:mathPr/m:mathFont`. That override is not consumed yet —
/// when it is, this becomes the fallback, not the answer.
pub const DEFAULT_MATH_FONT: &str = "Cambria Math";

/// One `m:oMath` — an inline run of math content inside a paragraph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MathBlock {
    pub content: Vec<MathElement>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum MathElement {
    /// `m:r` — literal math text.
    Run(MathRun),
    /// `m:sSup` — base with a superscript. (A `Subscript` twin for `m:sSub`
    /// is the natural next variant; the layout path is shared.)
    Superscript {
        base: Vec<MathElement>,
        sup: Vec<MathElement>,
    },
    /// `m:f` — numerator over a fraction bar over a denominator.
    Fraction {
        num: Vec<MathElement>,
        den: Vec<MathElement>,
    },
}

/// `m:r` text content, `m:t` parts joined.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MathRun {
    pub text: String,
}
