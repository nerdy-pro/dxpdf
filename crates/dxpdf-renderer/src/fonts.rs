//! Font resolution and caching — Skia font substitution with metric-compatible fallbacks.

use std::cell::RefCell;
use std::collections::HashMap;

use skia_safe::{Font, FontMgr, FontStyle, Typeface};

use crate::dimension::Pt;

/// Open-source metric-compatible substitutes for proprietary fonts.
const FONT_SUBSTITUTIONS: &[(&str, &[&str])] = &[
    ("Calibri", &["Carlito", "Liberation Sans", "Noto Sans"]),
    ("Cambria", &["Caladea", "Liberation Serif", "Noto Serif"]),
    ("Arial", &["Liberation Sans", "Noto Sans", "Helvetica"]),
    ("Times New Roman", &["Liberation Serif", "Noto Serif", "Times"]),
    ("Courier New", &["Liberation Mono", "Noto Sans Mono", "Courier"]),
    ("Verdana", &["DejaVu Sans", "Noto Sans"]),
    ("Georgia", &["DejaVu Serif", "Noto Serif"]),
    ("Trebuchet MS", &["Ubuntu", "Noto Sans"]),
    ("Consolas", &["Inconsolata", "Liberation Mono", "Noto Sans Mono"]),
    ("Segoe UI", &["Noto Sans", "Liberation Sans"]),
];

#[derive(Hash, Eq, PartialEq)]
struct TypefaceKey {
    family: String,
    weight: i32,
    slant: skia_safe::font_style::Slant,
}

thread_local! {
    static TYPEFACE_CACHE: RefCell<HashMap<TypefaceKey, Typeface>> = RefCell::new(HashMap::new());
}

fn match_exact(font_mgr: &FontMgr, family: &str, style: FontStyle) -> Option<Typeface> {
    let tf = font_mgr.match_family_style(family, style)?;
    let actual = tf.family_name();
    if actual.eq_ignore_ascii_case(family) {
        Some(tf)
    } else {
        None
    }
}

fn resolve_typeface_uncached(font_mgr: &FontMgr, font_family: &str, style: FontStyle) -> Typeface {
    if let Some(tf) = match_exact(font_mgr, font_family, style) {
        return tf;
    }

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

    font_mgr
        .match_family_style("Helvetica", style)
        .or_else(|| font_mgr.legacy_make_typeface(None::<&str>, style))
        .expect("no fallback typeface available")
}

/// Resolve a typeface with caching.
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

/// Pre-resolve all font families for all style variants.
pub fn preload_fonts(font_mgr: &FontMgr, families: &[String]) {
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

/// Create a Skia Font for the given properties.
pub fn make_font(font_mgr: &FontMgr, font_family: &str, font_size: Pt, bold: bool, italic: bool) -> Font {
    let style = match (bold, italic) {
        (true, true) => FontStyle::bold_italic(),
        (true, false) => FontStyle::bold(),
        (false, true) => FontStyle::italic(),
        (false, false) => FontStyle::normal(),
    };
    let typeface = resolve_typeface(font_mgr, font_family, style);
    Font::from_typeface(typeface, f32::from(font_size))
}
