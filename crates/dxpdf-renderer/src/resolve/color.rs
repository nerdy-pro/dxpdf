//! Color resolution — Color::Auto to RGB, theme color index to RGB.

use dxpdf_docx_model::model::{Color, ThemeColorIndex, ThemeColorScheme};

/// Resolved RGB color (0xRRGGBB).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct RgbColor {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl RgbColor {
    pub const BLACK: Self = Self { r: 0, g: 0, b: 0 };
    pub const WHITE: Self = Self {
        r: 255,
        g: 255,
        b: 255,
    };
}

/// Context for resolving Color::Auto.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ColorContext {
    /// Text color — Auto means black.
    Text,
    /// Background/fill color — Auto means white.
    Background,
}

/// Resolve a model Color to concrete RGB.
pub fn resolve_color(color: Color, context: ColorContext) -> RgbColor {
    match color {
        Color::Rgb(v) => rgb_from_u32(v),
        Color::Auto => match context {
            ColorContext::Text => RgbColor::BLACK,
            ColorContext::Background => RgbColor::WHITE,
        },
    }
}

/// Resolve a theme color index to RGB via the color scheme.
pub fn resolve_theme_color(index: ThemeColorIndex, scheme: &ThemeColorScheme) -> RgbColor {
    rgb_from_u32(scheme.resolve(index))
}

/// Convert a packed u32 (0xRRGGBB) to RgbColor.
pub fn rgb_from_u32(v: u32) -> RgbColor {
    RgbColor {
        r: ((v >> 16) & 0xFF) as u8,
        g: ((v >> 8) & 0xFF) as u8,
        b: (v & 0xFF) as u8,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rgb_from_u32_red() {
        let c = rgb_from_u32(0xFF0000);
        assert_eq!(c, RgbColor { r: 255, g: 0, b: 0 });
    }

    #[test]
    fn rgb_from_u32_white() {
        let c = rgb_from_u32(0xFFFFFF);
        assert_eq!(c, RgbColor::WHITE);
    }

    #[test]
    fn rgb_from_u32_black() {
        let c = rgb_from_u32(0x000000);
        assert_eq!(c, RgbColor::BLACK);
    }

    #[test]
    fn rgb_from_u32_mixed() {
        let c = rgb_from_u32(0x1A2B3C);
        assert_eq!(
            c,
            RgbColor {
                r: 0x1A,
                g: 0x2B,
                b: 0x3C
            }
        );
    }

    #[test]
    fn resolve_color_rgb_passes_through() {
        let c = resolve_color(Color::Rgb(0x336699), ColorContext::Text);
        assert_eq!(
            c,
            RgbColor {
                r: 0x33,
                g: 0x66,
                b: 0x99
            }
        );
    }

    #[test]
    fn resolve_color_auto_text_is_black() {
        let c = resolve_color(Color::Auto, ColorContext::Text);
        assert_eq!(c, RgbColor::BLACK);
    }

    #[test]
    fn resolve_color_auto_background_is_white() {
        let c = resolve_color(Color::Auto, ColorContext::Background);
        assert_eq!(c, RgbColor::WHITE);
    }

    #[test]
    fn resolve_theme_color_accent1() {
        let scheme = ThemeColorScheme {
            accent1: 0x4472C4,
            ..Default::default()
        };
        let c = resolve_theme_color(ThemeColorIndex::Accent1, &scheme);
        assert_eq!(
            c,
            RgbColor {
                r: 0x44,
                g: 0x72,
                b: 0xC4
            }
        );
    }

    #[test]
    fn resolve_theme_color_dark1() {
        let scheme = ThemeColorScheme {
            dark1: 0x000000,
            ..Default::default()
        };
        let c = resolve_theme_color(ThemeColorIndex::Dark1, &scheme);
        assert_eq!(c, RgbColor::BLACK);
    }

    #[test]
    fn resolve_theme_color_hyperlink() {
        let scheme = ThemeColorScheme {
            hyperlink: 0x0563C1,
            ..Default::default()
        };
        let c = resolve_theme_color(ThemeColorIndex::Hyperlink, &scheme);
        assert_eq!(
            c,
            RgbColor {
                r: 0x05,
                g: 0x63,
                b: 0xC1
            }
        );
    }
}
