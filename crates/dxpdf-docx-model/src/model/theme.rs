//! Theme types — color schemes, font schemes, and script tags.

/// Resolved theme data from `theme1.xml`.
#[derive(Clone, Debug, Default)]
pub struct Theme {
    pub color_scheme: ThemeColorScheme,
    pub major_font: ThemeFontScheme,
    pub minor_font: ThemeFontScheme,
}

#[derive(Clone, Copy, Debug, Default)]
pub struct ThemeColorScheme {
    pub dark1: u32,
    pub light1: u32,
    pub dark2: u32,
    pub light2: u32,
    pub accent1: u32,
    pub accent2: u32,
    pub accent3: u32,
    pub accent4: u32,
    pub accent5: u32,
    pub accent6: u32,
    pub hyperlink: u32,
    pub followed_hyperlink: u32,
}

impl ThemeColorScheme {
    /// Resolve a theme color index to an RGB value.
    pub fn resolve(&self, idx: ThemeColorIndex) -> u32 {
        match idx {
            ThemeColorIndex::Dark1 => self.dark1,
            ThemeColorIndex::Light1 => self.light1,
            ThemeColorIndex::Dark2 => self.dark2,
            ThemeColorIndex::Light2 => self.light2,
            ThemeColorIndex::Accent1 => self.accent1,
            ThemeColorIndex::Accent2 => self.accent2,
            ThemeColorIndex::Accent3 => self.accent3,
            ThemeColorIndex::Accent4 => self.accent4,
            ThemeColorIndex::Accent5 => self.accent5,
            ThemeColorIndex::Accent6 => self.accent6,
            ThemeColorIndex::Hyperlink => self.hyperlink,
            ThemeColorIndex::FollowedHyperlink => self.followed_hyperlink,
        }
    }
}

/// Index into the theme color scheme (ST_ThemeColor).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ThemeColorIndex {
    Dark1,
    Light1,
    Dark2,
    Light2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

#[derive(Clone, Debug, Default)]
pub struct ThemeFontScheme {
    pub latin: String,
    pub east_asian: String,
    pub complex_script: String,
    /// §20.1.4.1.16: per-script font overrides.
    pub script_fonts: Vec<ThemeScriptFont>,
}

/// §20.1.4.1.16: a per-script font mapping in a theme font scheme.
#[derive(Clone, Debug)]
pub struct ThemeScriptFont {
    /// ISO 15924 script code.
    pub script: ScriptTag,
    /// Typeface name for this script.
    pub typeface: String,
}

/// ISO 15924 script codes used in OOXML theme font schemes (§20.1.4.1.16).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ScriptTag {
    Arab,
    Armn,
    Beng,
    Bopo,
    Bugi,
    Cans,
    Cher,
    Deva,
    Ethi,
    Geor,
    Gujr,
    Guru,
    Hang,
    Hans,
    Hant,
    Hebr,
    Java,
    Jpan,
    Khmr,
    Knda,
    Laoo,
    Lisu,
    Mlym,
    Mong,
    Mymr,
    Nkoo,
    Olck,
    Orya,
    Osma,
    Phag,
    Sinh,
    Sora,
    Syre,
    Syrj,
    Syrn,
    Syrc,
    Tale,
    Talu,
    Taml,
    Telu,
    Tfng,
    Thaa,
    Thai,
    Tibt,
    Uigh,
    Viet,
    Yiii,
    /// Unrecognized script code — preserved as-is.
    Other(u32),
}
