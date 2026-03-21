use std::cell::RefCell;
use std::collections::HashMap;

use skia_safe::{Font, FontMgr, FontStyle, Typeface};

/// Open-source metric-compatible substitutes for proprietary fonts.
/// Each entry maps a proprietary font name to a list of alternatives
/// tried in order. These fonts have matching metrics so line wrapping
/// and spacing remain accurate.
const FONT_SUBSTITUTIONS: &[(&str, &[&str])] = &[
    // Microsoft fonts → metric-compatible open-source alternatives
    ("Calibri", &["Carlito", "Liberation Sans", "Noto Sans"]),
    ("Cambria", &["Caladea", "Liberation Serif", "Noto Serif"]),
    ("Arial", &["Liberation Sans", "Noto Sans", "Helvetica"]),
    (
        "Times New Roman",
        &["Liberation Serif", "Noto Serif", "Times"],
    ),
    (
        "Courier New",
        &["Liberation Mono", "Noto Sans Mono", "Courier"],
    ),
    ("Verdana", &["DejaVu Sans", "Noto Sans"]),
    ("Georgia", &["DejaVu Serif", "Noto Serif"]),
    ("Trebuchet MS", &["Ubuntu", "Noto Sans"]),
    (
        "Consolas",
        &["Inconsolata", "Liberation Mono", "Noto Sans Mono"],
    ),
    ("Segoe UI", &["Noto Sans", "Liberation Sans"]),
];

/// Cache key for resolved typefaces.
#[derive(Hash, Eq, PartialEq)]
struct TypefaceKey {
    family: String,
    weight: i32,
    slant: skia_safe::font_style::Slant,
}

// Thread-local cache of resolved typefaces to avoid repeated fontconfig lookups.
thread_local! {
    static TYPEFACE_CACHE: RefCell<HashMap<TypefaceKey, Typeface>> = RefCell::new(HashMap::new());
}

/// Try to match a font by family name, returning it only if the system
/// actually has that exact font (not a Skia-substituted fallback).
fn match_exact(font_mgr: &FontMgr, family: &str, style: FontStyle) -> Option<Typeface> {
    let tf = font_mgr.match_family_style(family, style)?;
    // Skia may silently return a fallback font instead of None.
    // Verify the returned typeface actually matches the requested family.
    let actual = tf.family_name();
    if actual.eq_ignore_ascii_case(family) {
        Some(tf)
    } else {
        None
    }
}

/// Resolve a font family name, trying the requested font first, then
/// metric-compatible substitutes from the table, then Helvetica as a final fallback.
fn resolve_typeface_uncached(font_mgr: &FontMgr, font_family: &str, style: FontStyle) -> Typeface {
    // Try the requested font first (exact match only)
    if let Some(tf) = match_exact(font_mgr, font_family, style) {
        return tf;
    }

    // Try metric-compatible substitutes
    if let Some((_, subs)) = FONT_SUBSTITUTIONS
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(font_family))
    {
        for sub in *subs {
            if let Some(tf) = match_exact(font_mgr, sub, style) {
                return tf;
            }
        }
    }

    // Final fallbacks
    font_mgr
        .match_family_style("Helvetica", style)
        .or_else(|| font_mgr.legacy_make_typeface(None::<&str>, style))
        .expect("no fallback typeface available")
}

/// Resolve a typeface with caching. Lookups are cached per (family, style)
/// to avoid repeated fontconfig queries.
pub fn resolve_typeface(font_mgr: &FontMgr, font_family: &str, style: FontStyle) -> Typeface {
    let key = TypefaceKey {
        family: font_family.to_lowercase(),
        weight: *style.weight(),
        slant: style.slant(),
    };

    TYPEFACE_CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        if let Some(tf) = cache.get(&key) {
            return tf.clone();
        }
        let tf = resolve_typeface_uncached(font_mgr, font_family, style);
        cache.insert(key, tf.clone());
        tf
    })
}

/// Pre-resolve all font families for all style variants (normal, bold, italic, bold-italic).
/// Call this before layout/paint to move all fontconfig lookups into a dedicated pipeline step.
pub fn preload_fonts(font_mgr: &FontMgr, families: &[std::rc::Rc<str>]) {
    let styles = [
        FontStyle::normal(),
        FontStyle::bold(),
        FontStyle::italic(),
        FontStyle::bold_italic(),
    ];
    for family in families {
        for &style in &styles {
            resolve_typeface(font_mgr, family, style);
        }
    }
}

/// Create a Skia Font for the given properties with substitution support.
pub fn make_font(
    font_mgr: &FontMgr,
    font_family: &str,
    font_size: f32,
    bold: bool,
    italic: bool,
) -> Font {
    let style = match (bold, italic) {
        (true, true) => FontStyle::bold_italic(),
        (true, false) => FontStyle::bold(),
        (false, true) => FontStyle::italic(),
        (false, false) => FontStyle::normal(),
    };
    let typeface = resolve_typeface(font_mgr, font_family, style);
    Font::from_typeface(typeface, font_size)
}
