//! DrawingML color resolver (§20.1.2.3).
//!
//! Takes a parsed `DrawingColor` and a `ColorContext` (theme + bg/tx
//! convention) and produces a concrete `Rgba`. Color transforms are applied
//! in document order via a pure left-fold.
//!
//! ## Spec references for transforms
//!
//! Most transforms can be described in one of two spaces:
//!
//! * **RGB-direct**: `red`/`green`/`blue` and their `Off`/`Mod` siblings
//!   act directly on sRGB channels. Absolute forms overwrite the channel;
//!   offsets add a signed percentage; multipliers scale (clamped to [0,1]).
//! * **HSL-space**: `hue`/`sat`/`lum` and their siblings convert the color
//!   to HSL, modify the named component, and convert back.
//!
//! Intensity transforms:
//!
//! * `tint(p)`    — per-channel: `c' = c + (1 - c) * p` (§20.1.2.3.34).
//! * `shade(p)`   — per-channel: `c' = c * (1 - p)`     (§20.1.2.3.31).
//! * `comp`       — HSL hue rotated by 180° (§20.1.2.3.7).
//! * `inv`        — per-channel: `c' = 1 - c`          (§20.1.2.3.17).
//! * `gray`       — luminance-preserving desaturation (§20.1.2.3.8). The
//!   resulting value sets all three sRGB channels to the original luminance.
//! * `gamma`      — sRGB → linear (§20.1.2.3.9).
//! * `invGamma`   — linear → sRGB (§20.1.2.3.18).
//!
//! Alpha transforms act on the alpha channel only (§20.1.2.3.1/2/3).

use crate::model::dimension::{Dimension, SixtieThousandthDeg, ThousandthPercent};
use crate::model::{ColorTransform, DrawingColor, SchemeColorVal, Theme};

/// RGBA with per-channel `f32` in `[0, 1]`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Rgba {
    pub r: f32,
    pub g: f32,
    pub b: f32,
    pub a: f32,
}

impl Rgba {
    pub const BLACK: Self = Self {
        r: 0.0,
        g: 0.0,
        b: 0.0,
        a: 1.0,
    };
    pub const WHITE: Self = Self {
        r: 1.0,
        g: 1.0,
        b: 1.0,
        a: 1.0,
    };

    /// Construct from a packed `0xRRGGBB` sRGB value with full alpha.
    pub fn from_rgb24(v: u32) -> Self {
        Self {
            r: ((v >> 16) & 0xFF) as f32 / 255.0,
            g: ((v >> 8) & 0xFF) as f32 / 255.0,
            b: (v & 0xFF) as f32 / 255.0,
            a: 1.0,
        }
    }

    /// Convert to a packed `0xRRGGBB` sRGB value (alpha discarded).
    pub fn to_rgb24(self) -> u32 {
        let c = |v: f32| -> u32 { ((v.clamp(0.0, 1.0) * 255.0).round() as u32) & 0xFF };
        (c(self.r) << 16) | (c(self.g) << 8) | c(self.b)
    }
}

/// §17.3.1.31: Word documents treat `tx1`/`bg1`/`tx2`/`bg2` as references
/// against an assumed background brightness. For light-backgrounded bodies
/// (Word's default) `tx1 → dk1`, `bg1 → lt1`, etc. Dark themes invert.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub enum BgTxConvention {
    #[default]
    LightBackground,
    DarkBackground,
}

/// Resolver context. `theme` is optional because documents can omit theme
/// parts; when missing, scheme references fall back to black.
#[derive(Clone, Copy, Debug)]
pub struct DrawingColorContext<'a> {
    pub theme: Option<&'a Theme>,
    pub bg_tx_convention: BgTxConvention,
}

impl<'a> DrawingColorContext<'a> {
    pub fn new(theme: Option<&'a Theme>) -> Self {
        Self {
            theme,
            bg_tx_convention: BgTxConvention::LightBackground,
        }
    }
}

/// Resolve a `DrawingColor` to its final RGBA.
pub fn resolve_drawing_color(color: &DrawingColor, ctx: &DrawingColorContext<'_>) -> Rgba {
    let base = resolve_base(color, ctx);
    color.transforms().iter().fold(base, apply_transform)
}

// ── Base color resolution ───────────────────────────────────────────────────

fn resolve_base(color: &DrawingColor, ctx: &DrawingColorContext<'_>) -> Rgba {
    match color {
        DrawingColor::Srgb { rgb, .. } => Rgba::from_rgb24(*rgb),
        DrawingColor::ScRgb { r, g, b, .. } => Rgba {
            r: pct_f32(*r),
            g: pct_f32(*g),
            b: pct_f32(*b),
            a: 1.0,
        },
        DrawingColor::Hsl { hue, sat, lum, .. } => {
            hsl_to_rgba(angle_f32(*hue), pct_f32(*sat), pct_f32(*lum))
        }
        DrawingColor::Sys { name, last_clr, .. } => {
            Rgba::from_rgb24(last_clr.unwrap_or_else(|| name.default_rgb()))
        }
        DrawingColor::Scheme { name, .. } => resolve_scheme(*name, ctx),
        DrawingColor::Prst { name, .. } => Rgba::from_rgb24(name.rgb()),
    }
}

fn resolve_scheme(name: SchemeColorVal, ctx: &DrawingColorContext<'_>) -> Rgba {
    let theme = match ctx.theme {
        Some(t) => t,
        None => return Rgba::BLACK,
    };
    let scheme = &theme.color_scheme;
    let rgb = match (name, ctx.bg_tx_convention) {
        (SchemeColorVal::Dk1, _) => scheme.dark1,
        (SchemeColorVal::Lt1, _) => scheme.light1,
        (SchemeColorVal::Dk2, _) => scheme.dark2,
        (SchemeColorVal::Lt2, _) => scheme.light2,

        // §17.3.1.31: Word light-background convention maps tx↔dk and bg↔lt.
        (SchemeColorVal::Tx1, BgTxConvention::LightBackground) => scheme.dark1,
        (SchemeColorVal::Bg1, BgTxConvention::LightBackground) => scheme.light1,
        (SchemeColorVal::Tx2, BgTxConvention::LightBackground) => scheme.dark2,
        (SchemeColorVal::Bg2, BgTxConvention::LightBackground) => scheme.light2,
        (SchemeColorVal::Tx1, BgTxConvention::DarkBackground) => scheme.light1,
        (SchemeColorVal::Bg1, BgTxConvention::DarkBackground) => scheme.dark1,
        (SchemeColorVal::Tx2, BgTxConvention::DarkBackground) => scheme.light2,
        (SchemeColorVal::Bg2, BgTxConvention::DarkBackground) => scheme.dark2,

        (SchemeColorVal::Accent1, _) => scheme.accent1,
        (SchemeColorVal::Accent2, _) => scheme.accent2,
        (SchemeColorVal::Accent3, _) => scheme.accent3,
        (SchemeColorVal::Accent4, _) => scheme.accent4,
        (SchemeColorVal::Accent5, _) => scheme.accent5,
        (SchemeColorVal::Accent6, _) => scheme.accent6,
        (SchemeColorVal::Hlink, _) => scheme.hyperlink,
        (SchemeColorVal::FolHlink, _) => scheme.followed_hyperlink,

        (SchemeColorVal::PhClr, _) => {
            // §20.1.2.3.29 phClr is a placeholder used by master/layout parts;
            // for Word body content it has no concrete RGB. Fall back to black
            // rather than producing a panic.
            0x000000
        }
    };
    Rgba::from_rgb24(rgb)
}

// ── Transform application ───────────────────────────────────────────────────

fn apply_transform(c: Rgba, t: &ColorTransform) -> Rgba {
    match t {
        // Intensity
        ColorTransform::Tint(p) => apply_tint(c, pct_f32(*p)),
        ColorTransform::Shade(p) => apply_shade(c, pct_f32(*p)),
        ColorTransform::Comp => apply_comp(c),
        ColorTransform::Inv => Rgba {
            r: 1.0 - c.r,
            g: 1.0 - c.g,
            b: 1.0 - c.b,
            a: c.a,
        },
        ColorTransform::Gray => apply_gray(c),
        ColorTransform::Gamma => Rgba {
            r: srgb_to_linear(c.r),
            g: srgb_to_linear(c.g),
            b: srgb_to_linear(c.b),
            a: c.a,
        },
        ColorTransform::InvGamma => Rgba {
            r: linear_to_srgb(c.r),
            g: linear_to_srgb(c.g),
            b: linear_to_srgb(c.b),
            a: c.a,
        },

        // Alpha
        ColorTransform::Alpha(p) => Rgba {
            a: pct_f32(*p).clamp(0.0, 1.0),
            ..c
        },
        ColorTransform::AlphaOff(p) => Rgba {
            a: (c.a + pct_f32(*p)).clamp(0.0, 1.0),
            ..c
        },
        ColorTransform::AlphaMod(p) => Rgba {
            a: (c.a * pct_f32(*p)).clamp(0.0, 1.0),
            ..c
        },

        // HSL
        ColorTransform::Hue(a) => hsl_adjust(c, |_h, s, l| (angle_f32(*a), s, l)),
        ColorTransform::HueOff(a) => {
            hsl_adjust(c, |h, s, l| ((h + angle_f32(*a)).rem_euclid(1.0), s, l))
        }
        ColorTransform::HueMod(p) => {
            hsl_adjust(c, |h, s, l| ((h * pct_f32(*p)).rem_euclid(1.0), s, l))
        }
        ColorTransform::Sat(p) => hsl_adjust(c, |h, _s, l| (h, pct_f32(*p).clamp(0.0, 1.0), l)),
        ColorTransform::SatOff(p) => {
            hsl_adjust(c, |h, s, l| (h, (s + pct_f32(*p)).clamp(0.0, 1.0), l))
        }
        ColorTransform::SatMod(p) => {
            hsl_adjust(c, |h, s, l| (h, (s * pct_f32(*p)).clamp(0.0, 1.0), l))
        }
        ColorTransform::Lum(p) => hsl_adjust(c, |h, s, _l| (h, s, pct_f32(*p).clamp(0.0, 1.0))),
        ColorTransform::LumOff(p) => {
            hsl_adjust(c, |h, s, l| (h, s, (l + pct_f32(*p)).clamp(0.0, 1.0)))
        }
        ColorTransform::LumMod(p) => {
            hsl_adjust(c, |h, s, l| (h, s, (l * pct_f32(*p)).clamp(0.0, 1.0)))
        }

        // Per-RGB-channel
        ColorTransform::Red(p) => Rgba {
            r: pct_f32(*p).clamp(0.0, 1.0),
            ..c
        },
        ColorTransform::RedOff(p) => Rgba {
            r: (c.r + pct_f32(*p)).clamp(0.0, 1.0),
            ..c
        },
        ColorTransform::RedMod(p) => Rgba {
            r: (c.r * pct_f32(*p)).clamp(0.0, 1.0),
            ..c
        },
        ColorTransform::Green(p) => Rgba {
            g: pct_f32(*p).clamp(0.0, 1.0),
            ..c
        },
        ColorTransform::GreenOff(p) => Rgba {
            g: (c.g + pct_f32(*p)).clamp(0.0, 1.0),
            ..c
        },
        ColorTransform::GreenMod(p) => Rgba {
            g: (c.g * pct_f32(*p)).clamp(0.0, 1.0),
            ..c
        },
        ColorTransform::Blue(p) => Rgba {
            b: pct_f32(*p).clamp(0.0, 1.0),
            ..c
        },
        ColorTransform::BlueOff(p) => Rgba {
            b: (c.b + pct_f32(*p)).clamp(0.0, 1.0),
            ..c
        },
        ColorTransform::BlueMod(p) => Rgba {
            b: (c.b * pct_f32(*p)).clamp(0.0, 1.0),
            ..c
        },
    }
}

fn apply_tint(c: Rgba, p: f32) -> Rgba {
    let p = p.clamp(0.0, 1.0);
    Rgba {
        r: c.r + (1.0 - c.r) * p,
        g: c.g + (1.0 - c.g) * p,
        b: c.b + (1.0 - c.b) * p,
        a: c.a,
    }
}

fn apply_shade(c: Rgba, p: f32) -> Rgba {
    let p = p.clamp(0.0, 1.0);
    Rgba {
        r: c.r * (1.0 - p),
        g: c.g * (1.0 - p),
        b: c.b * (1.0 - p),
        a: c.a,
    }
}

fn apply_comp(c: Rgba) -> Rgba {
    let (h, s, l) = rgba_to_hsl(c);
    let h2 = (h + 0.5).rem_euclid(1.0);
    let (r, g, b) = hsl_to_rgb(h2, s, l);
    Rgba { r, g, b, a: c.a }
}

fn apply_gray(c: Rgba) -> Rgba {
    // ITU-R BT.601 luminance.
    let y = 0.299 * c.r + 0.587 * c.g + 0.114 * c.b;
    Rgba {
        r: y,
        g: y,
        b: y,
        a: c.a,
    }
}

fn hsl_adjust<F: FnOnce(f32, f32, f32) -> (f32, f32, f32)>(c: Rgba, f: F) -> Rgba {
    let (h, s, l) = rgba_to_hsl(c);
    let (h, s, l) = f(h, s, l);
    let (r, g, b) = hsl_to_rgb(h, s, l);
    Rgba { r, g, b, a: c.a }
}

// ── Color space conversions ────────────────────────────────────────────────

fn pct_f32(p: Dimension<ThousandthPercent>) -> f32 {
    p.raw() as f32 / 100_000.0
}

fn angle_f32(a: Dimension<SixtieThousandthDeg>) -> f32 {
    // Normalize any angle to [0, 1) expressed as a fraction of a full turn.
    let degrees = a.raw() as f32 / 60_000.0;
    (degrees / 360.0).rem_euclid(1.0)
}

fn hsl_to_rgba(h: f32, s: f32, l: f32) -> Rgba {
    let (r, g, b) = hsl_to_rgb(h, s, l);
    Rgba { r, g, b, a: 1.0 }
}

/// HSL → RGB. All inputs and outputs in [0, 1].
fn hsl_to_rgb(h: f32, s: f32, l: f32) -> (f32, f32, f32) {
    if s <= 0.0 {
        return (l, l, l);
    }
    let q = if l < 0.5 {
        l * (1.0 + s)
    } else {
        l + s - l * s
    };
    let p = 2.0 * l - q;
    let r = hue_to_channel(p, q, h + 1.0 / 3.0);
    let g = hue_to_channel(p, q, h);
    let b = hue_to_channel(p, q, h - 1.0 / 3.0);
    (r, g, b)
}

fn hue_to_channel(p: f32, q: f32, mut t: f32) -> f32 {
    if t < 0.0 {
        t += 1.0;
    }
    if t > 1.0 {
        t -= 1.0;
    }
    if t < 1.0 / 6.0 {
        return p + (q - p) * 6.0 * t;
    }
    if t < 0.5 {
        return q;
    }
    if t < 2.0 / 3.0 {
        return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
    }
    p
}

/// RGB → HSL. All inputs and outputs in [0, 1].
fn rgba_to_hsl(c: Rgba) -> (f32, f32, f32) {
    let max = c.r.max(c.g).max(c.b);
    let min = c.r.min(c.g).min(c.b);
    let l = (max + min) / 2.0;
    if (max - min).abs() < f32::EPSILON {
        return (0.0, 0.0, l);
    }
    let d = max - min;
    let s = if l > 0.5 {
        d / (2.0 - max - min)
    } else {
        d / (max + min)
    };
    let h = if (max - c.r).abs() < f32::EPSILON {
        (c.g - c.b) / d + if c.g < c.b { 6.0 } else { 0.0 }
    } else if (max - c.g).abs() < f32::EPSILON {
        (c.b - c.r) / d + 2.0
    } else {
        (c.r - c.g) / d + 4.0
    } / 6.0;
    (h, s, l)
}

/// §20.1.2.3.9 gamma — sRGB EOTF (sRGB → linear).
fn srgb_to_linear(c: f32) -> f32 {
    if c <= 0.04045 {
        c / 12.92
    } else {
        ((c + 0.055) / 1.055).powf(2.4)
    }
}

/// §20.1.2.3.18 invGamma — inverse sRGB EOTF (linear → sRGB).
fn linear_to_srgb(c: f32) -> f32 {
    if c <= 0.0031308 {
        c * 12.92
    } else {
        1.055 * c.powf(1.0 / 2.4) - 0.055
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{PresetColorVal, SystemColorVal, ThemeColorScheme, ThemeFontScheme};

    fn approx(a: f32, b: f32) -> bool {
        (a - b).abs() < 5e-3
    }

    fn default_ctx() -> DrawingColorContext<'static> {
        DrawingColorContext {
            theme: None,
            bg_tx_convention: BgTxConvention::LightBackground,
        }
    }

    fn theme_ctx(theme: &Theme) -> DrawingColorContext<'_> {
        DrawingColorContext {
            theme: Some(theme),
            bg_tx_convention: BgTxConvention::LightBackground,
        }
    }

    fn sample_theme() -> Theme {
        Theme {
            color_scheme: ThemeColorScheme {
                dark1: 0x000000,
                light1: 0xFFFFFF,
                dark2: 0x44546A,
                light2: 0xE7E6E6,
                accent1: 0x4472C4,
                accent2: 0xED7D31,
                accent3: 0xA5A5A5,
                accent4: 0xFFC000,
                accent5: 0x5B9BD5,
                accent6: 0x70AD47,
                hyperlink: 0x0563C1,
                followed_hyperlink: 0x954F72,
            },
            major_font: ThemeFontScheme::default(),
            minor_font: ThemeFontScheme::default(),
        }
    }

    // ── Base color resolution ──────────────────────────────────────────

    #[test]
    fn srgb_passes_through() {
        let c = DrawingColor::Srgb {
            rgb: 0xAABBCC,
            transforms: vec![],
        };
        let r = resolve_drawing_color(&c, &default_ctx());
        assert_eq!(r.to_rgb24(), 0xAABBCC);
    }

    #[test]
    fn scheme_accent1_resolves_via_theme() {
        let theme = sample_theme();
        let c = DrawingColor::Scheme {
            name: SchemeColorVal::Accent1,
            transforms: vec![],
        };
        let r = resolve_drawing_color(&c, &theme_ctx(&theme));
        assert_eq!(r.to_rgb24(), 0x4472C4);
    }

    #[test]
    fn scheme_tx1_maps_to_dk1_on_light_bg() {
        let theme = sample_theme();
        let c = DrawingColor::Scheme {
            name: SchemeColorVal::Tx1,
            transforms: vec![],
        };
        let r = resolve_drawing_color(&c, &theme_ctx(&theme));
        assert_eq!(r.to_rgb24(), 0x000000);
    }

    #[test]
    fn scheme_bg1_maps_to_lt1_on_light_bg() {
        let theme = sample_theme();
        let c = DrawingColor::Scheme {
            name: SchemeColorVal::Bg1,
            transforms: vec![],
        };
        let r = resolve_drawing_color(&c, &theme_ctx(&theme));
        assert_eq!(r.to_rgb24(), 0xFFFFFF);
    }

    #[test]
    fn scheme_tx1_maps_to_lt1_on_dark_bg() {
        let theme = sample_theme();
        let ctx = DrawingColorContext {
            theme: Some(&theme),
            bg_tx_convention: BgTxConvention::DarkBackground,
        };
        let c = DrawingColor::Scheme {
            name: SchemeColorVal::Tx1,
            transforms: vec![],
        };
        let r = resolve_drawing_color(&c, &ctx);
        assert_eq!(r.to_rgb24(), 0xFFFFFF);
    }

    #[test]
    fn sys_clr_uses_last_clr_when_present() {
        let c = DrawingColor::Sys {
            name: SystemColorVal::Window,
            last_clr: Some(0xAABBCC),
            transforms: vec![],
        };
        assert_eq!(
            resolve_drawing_color(&c, &default_ctx()).to_rgb24(),
            0xAABBCC
        );
    }

    #[test]
    fn sys_clr_falls_back_to_default_rgb() {
        let c = DrawingColor::Sys {
            name: SystemColorVal::Window,
            last_clr: None,
            transforms: vec![],
        };
        assert_eq!(
            resolve_drawing_color(&c, &default_ctx()).to_rgb24(),
            SystemColorVal::Window.default_rgb()
        );
    }

    #[test]
    fn prst_clr_uses_preset_rgb() {
        let c = DrawingColor::Prst {
            name: PresetColorVal::Red,
            transforms: vec![],
        };
        assert_eq!(
            resolve_drawing_color(&c, &default_ctx()).to_rgb24(),
            0xFF0000
        );
    }

    #[test]
    fn scheme_without_theme_falls_back_to_black() {
        let c = DrawingColor::Scheme {
            name: SchemeColorVal::Accent1,
            transforms: vec![],
        };
        assert_eq!(
            resolve_drawing_color(&c, &default_ctx()).to_rgb24(),
            0x000000
        );
    }

    // ── Transforms ─────────────────────────────────────────────────────

    fn srgb(rgb: u32, transforms: Vec<ColorTransform>) -> DrawingColor {
        DrawingColor::Srgb { rgb, transforms }
    }

    fn pct(thousandth: i64) -> Dimension<ThousandthPercent> {
        Dimension::new(thousandth)
    }

    fn angle(sixtie: i64) -> Dimension<SixtieThousandthDeg> {
        Dimension::new(sixtie)
    }

    #[test]
    fn tint_toward_white() {
        let c = srgb(0x000000, vec![ColorTransform::Tint(pct(50_000))]);
        let r = resolve_drawing_color(&c, &default_ctx());
        assert!(approx(r.r, 0.5));
        assert!(approx(r.g, 0.5));
        assert!(approx(r.b, 0.5));
    }

    #[test]
    fn tint_100pct_produces_white() {
        let c = srgb(0x000000, vec![ColorTransform::Tint(pct(100_000))]);
        let r = resolve_drawing_color(&c, &default_ctx());
        assert_eq!(r.to_rgb24(), 0xFFFFFF);
    }

    #[test]
    fn shade_toward_black() {
        let c = srgb(0xFFFFFF, vec![ColorTransform::Shade(pct(50_000))]);
        let r = resolve_drawing_color(&c, &default_ctx());
        assert!(approx(r.r, 0.5));
    }

    #[test]
    fn shade_100pct_produces_black() {
        let c = srgb(0xFFFFFF, vec![ColorTransform::Shade(pct(100_000))]);
        let r = resolve_drawing_color(&c, &default_ctx());
        assert_eq!(r.to_rgb24(), 0x000000);
    }

    #[test]
    fn inv_flips_channels() {
        let c = srgb(0x112233, vec![ColorTransform::Inv]);
        let r = resolve_drawing_color(&c, &default_ctx());
        assert_eq!(r.to_rgb24(), 0xEEDDCC);
    }

    #[test]
    fn alpha_sets_absolute() {
        let c = srgb(0xFF0000, vec![ColorTransform::Alpha(pct(50_000))]);
        let r = resolve_drawing_color(&c, &default_ctx());
        assert!(approx(r.a, 0.5));
    }

    #[test]
    fn alpha_mod_multiplies() {
        let c = srgb(
            0xFF0000,
            vec![
                ColorTransform::Alpha(pct(80_000)),
                ColorTransform::AlphaMod(pct(50_000)),
            ],
        );
        let r = resolve_drawing_color(&c, &default_ctx());
        assert!(approx(r.a, 0.4));
    }

    #[test]
    fn alpha_off_shifts() {
        let c = srgb(0xFF0000, vec![ColorTransform::AlphaOff(pct(-25_000))]);
        let r = resolve_drawing_color(&c, &default_ctx());
        assert!(approx(r.a, 0.75));
    }

    #[test]
    fn lum_mod_half_darkens() {
        // Accent1 with lumMod 75000 + lumOff 25000 is the canonical Office
        // recipe for a lightened accent. Just sanity-check monotonicity.
        let theme = sample_theme();
        let c = DrawingColor::Scheme {
            name: SchemeColorVal::Accent1,
            transforms: vec![
                ColorTransform::LumMod(pct(75_000)),
                ColorTransform::LumOff(pct(25_000)),
            ],
        };
        let r = resolve_drawing_color(&c, &theme_ctx(&theme));
        // Accent1 is 0x4472C4 (bluish). LumMod+LumOff lightens it; the
        // resulting color should have L > input L.
        let input = Rgba::from_rgb24(0x4472C4);
        let (_, _, l_in) = rgba_to_hsl(input);
        let (_, _, l_out) = rgba_to_hsl(r);
        assert!(l_out > l_in);
    }

    #[test]
    fn hue_offset_rotates() {
        // Pure red → hue-offset 120° → pure green.
        let c = srgb(0xFF0000, vec![ColorTransform::HueOff(angle(120 * 60_000))]);
        let r = resolve_drawing_color(&c, &default_ctx());
        assert!(approx(r.r, 0.0));
        assert!(approx(r.g, 1.0));
        assert!(approx(r.b, 0.0));
    }

    #[test]
    fn gray_desaturates_to_luminance() {
        let c = srgb(0xFF0000, vec![ColorTransform::Gray]);
        let r = resolve_drawing_color(&c, &default_ctx());
        // BT.601 luminance of pure red = 0.299.
        assert!(approx(r.r, 0.299));
        assert!(approx(r.g, 0.299));
        assert!(approx(r.b, 0.299));
    }

    #[test]
    fn comp_rotates_hue_180() {
        // Pure red hue → cyan-ish (opposite hue).
        let c = srgb(0xFF0000, vec![ColorTransform::Comp]);
        let r = resolve_drawing_color(&c, &default_ctx());
        assert!(approx(r.r, 0.0));
        assert!(approx(r.g, 1.0));
        assert!(approx(r.b, 1.0));
    }

    #[test]
    fn red_channel_absolute() {
        let c = srgb(0x000000, vec![ColorTransform::Red(pct(100_000))]);
        let r = resolve_drawing_color(&c, &default_ctx());
        assert_eq!(r.to_rgb24(), 0xFF0000);
    }

    #[test]
    fn red_offset_adds() {
        let c = srgb(0x800000, vec![ColorTransform::RedOff(pct(-50_000))]);
        let r = resolve_drawing_color(&c, &default_ctx());
        assert!(approx(r.r, 0.0));
    }

    #[test]
    fn red_mod_multiplies() {
        let c = srgb(0x800000, vec![ColorTransform::RedMod(pct(50_000))]);
        let r = resolve_drawing_color(&c, &default_ctx());
        assert!(approx(r.r, 0.25));
    }

    #[test]
    fn transforms_apply_in_order() {
        // shade 50% then tint 50% is NOT the same as tint 50% then shade 50%.
        let c1 = srgb(
            0xFF0000,
            vec![
                ColorTransform::Shade(pct(50_000)),
                ColorTransform::Tint(pct(50_000)),
            ],
        );
        let c2 = srgb(
            0xFF0000,
            vec![
                ColorTransform::Tint(pct(50_000)),
                ColorTransform::Shade(pct(50_000)),
            ],
        );
        let r1 = resolve_drawing_color(&c1, &default_ctx());
        let r2 = resolve_drawing_color(&c2, &default_ctx());
        assert_ne!(r1.to_rgb24(), r2.to_rgb24());
    }

    #[test]
    fn gamma_then_inv_gamma_roundtrips() {
        let c = srgb(
            0x808080,
            vec![ColorTransform::Gamma, ColorTransform::InvGamma],
        );
        let r = resolve_drawing_color(&c, &default_ctx());
        // Allow small floating-point drift.
        let orig = Rgba::from_rgb24(0x808080);
        assert!(approx(r.r, orig.r));
        assert!(approx(r.g, orig.g));
        assert!(approx(r.b, orig.b));
    }

    // ── Rgba roundtrip ─────────────────────────────────────────────────

    #[test]
    fn rgba_from_to_rgb24_roundtrip() {
        for v in [0x000000, 0xFFFFFF, 0x808080, 0x123456, 0xAABBCC] {
            assert_eq!(Rgba::from_rgb24(v).to_rgb24(), v);
        }
    }

    #[test]
    fn rgba_clamp_in_to_rgb24() {
        let out = Rgba {
            r: 2.0,
            g: -1.0,
            b: 0.5,
            a: 1.0,
        }
        .to_rgb24();
        assert_eq!(out, 0xFF0080);
    }
}
