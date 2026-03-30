//! Font resolution and caching — Skia font substitution with metric-compatible fallbacks.

use std::cell::RefCell;
use std::collections::HashMap;

use dxpdf_docx_model::model::{EmbeddedFont, EmbeddedFontVariant};
use skia_safe::{Data, Font, FontMgr, FontStyle, Typeface};

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
        log::debug!("[font] '{}' {:?} → exact match", font_family, style);
        return tf;
    }

    if let Some((_, subs)) = FONT_SUBSTITUTIONS
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(font_family))
    {
        for sub in *subs {
            if let Some(tf) = match_exact(font_mgr, sub, style) {
                log::debug!("[font] '{}' {:?} → substitute '{}'", font_family, style, sub);
                return tf;
            }
        }
    }

    // No exact match or known substitute — use the system's default sans-serif.
    let tf = font_mgr.legacy_make_typeface(None::<&str>, style)
        .expect("no fallback typeface available");
    log::debug!("[font] '{}' {:?} → system default '{}'", font_family, style, tf.family_name());
    tf
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

/// Register embedded fonts from the document into the typeface cache.
/// These take priority over system fonts since the document bundled them.
pub fn register_embedded_fonts(font_mgr: &FontMgr, embedded: &[EmbeddedFont]) {
    for font in embedded {
        let skia_data = Data::new_copy(&font.data);
        let typeface = match font_mgr.new_from_data(&skia_data, 0) {
            Some(tf) => tf,
            None => {
                log::warn!(
                    "failed to load embedded font '{}' ({:?}): invalid font data",
                    font.family,
                    font.variant
                );
                continue;
            }
        };

        let style = match font.variant {
            EmbeddedFontVariant::Regular => FontStyle::normal(),
            EmbeddedFontVariant::Bold => FontStyle::bold(),
            EmbeddedFontVariant::Italic => FontStyle::italic(),
            EmbeddedFontVariant::BoldItalic => FontStyle::bold_italic(),
        };

        let key = TypefaceKey {
            family: font.family.to_lowercase(),
            weight: *style.weight(),
            slant: style.slant(),
        };

        TYPEFACE_CACHE.with(|cache| {
            cache.borrow_mut().insert(key, typeface);
        });

        log::debug!(
            "registered embedded font '{}' {:?}",
            font.family,
            font.variant
        );
    }
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
