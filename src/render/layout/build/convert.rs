use std::collections::HashMap;
use std::rc::Rc;

use crate::model::{self, FirstLineIndent, LineSpacing};
use crate::render::dimension::Pt;
use crate::render::geometry::PtSize;
use crate::render::layout::fragment::Fragment;
use crate::render::layout::measurer::TextMeasurer;
use crate::render::layout::paragraph::TabStopDef;
use crate::render::layout::paragraph::{
    BorderLine, LineSpacingRule, ParagraphBorderStyle, ParagraphStyle,
};
use crate::render::layout::table::{
    CellBorderOverride, TableBorderConfig, TableBorderLine, TableBorderStyle,
};
use crate::render::resolve::color::{resolve_color, ColorContext, RgbColor};
use crate::render::resolve::fonts::effective_font;
use crate::render::resolve::properties::merge_paragraph_properties;
use crate::render::resolve::ResolvedDocument;

use super::{BuildContext, SPEC_DEFAULT_FONT_SIZE, SPEC_FALLBACK_FONT};

/// Resolve a paragraph's effective defaults.
/// Cascade: direct → style → doc defaults.
///
/// Returns (font_family, font_size, color, merged_paragraph_props, run_defaults).
///
/// When `defer_doc_defaults` is true, doc defaults are NOT merged into the
/// paragraph properties — the caller is responsible for merging them after
/// inserting table style / conditional formatting in the cascade.
pub(super) fn resolve_paragraph_defaults(
    para: &model::Paragraph,
    resolved: &ResolvedDocument,
    defer_doc_defaults: bool,
) -> (
    String,
    Pt,
    RgbColor,
    model::ParagraphProperties,
    model::RunProperties,
) {
    let mut para_props = para.properties.clone();
    let mut run_defaults = resolved.doc_defaults_run.clone();

    // Derive default font from: doc defaults → theme minor font → spec fallback.
    let mut default_family = resolved
        .theme
        .as_ref()
        .map(|t| t.minor_font.latin.as_str())
        .filter(|s| !s.is_empty())
        .unwrap_or(SPEC_FALLBACK_FONT)
        .to_string();
    let mut default_size = resolved
        .doc_defaults_run
        .font_size
        .map(Pt::from)
        .unwrap_or(SPEC_DEFAULT_FONT_SIZE);
    let mut default_color = RgbColor::BLACK;

    // §17.7.4.17: if no style is specified, use the default paragraph style.
    let effective_style_id = para
        .style_id
        .as_ref()
        .or(resolved.default_paragraph_style_id.as_ref());

    if let Some(style_id) = effective_style_id {
        if let Some(resolved_style) = resolved.styles.get(style_id) {
            merge_paragraph_properties(&mut para_props, &resolved_style.paragraph);
            run_defaults = resolved_style.run.clone();
        }
    }

    // Merge doc defaults as lowest-priority fallback (unless deferred for table cascade).
    if !defer_doc_defaults {
        merge_paragraph_properties(&mut para_props, &resolved.doc_defaults_paragraph);
    }

    // Style's run font overrides document-level default.
    if let Some(f) = effective_font(&run_defaults.fonts) {
        default_family = f.to_string();
    }
    if let Some(fs) = run_defaults.font_size {
        default_size = Pt::from(fs);
    }
    if let Some(c) = run_defaults.color {
        default_color = resolve_color(c, ColorContext::Text);
    }

    (
        default_family,
        default_size,
        default_color,
        para_props,
        run_defaults,
    )
}

pub(super) fn doc_font_family(ctx: &BuildContext) -> String {
    ctx.resolved
        .theme
        .as_ref()
        .map(|t| t.minor_font.latin.as_str())
        .filter(|s| !s.is_empty())
        .unwrap_or(SPEC_FALLBACK_FONT)
        .to_string()
}

pub(super) fn doc_font_size(ctx: &BuildContext) -> Pt {
    ctx.resolved
        .doc_defaults_run
        .font_size
        .map(Pt::from)
        .unwrap_or(SPEC_DEFAULT_FONT_SIZE)
}

/// Convert a model paragraph properties into a layout ParagraphStyle.
pub(super) fn paragraph_style_from_props(props: &model::ParagraphProperties) -> ParagraphStyle {
    let indent_left = props
        .indentation
        .and_then(|i| i.start)
        .map(Pt::from)
        .unwrap_or(Pt::ZERO);
    let indent_right = props
        .indentation
        .and_then(|i| i.end)
        .map(Pt::from)
        .unwrap_or(Pt::ZERO);
    let indent_first_line = props
        .indentation
        .and_then(|i| i.first_line)
        .map(|fl| match fl {
            FirstLineIndent::FirstLine(v) => Pt::from(v),
            FirstLineIndent::Hanging(v) => -Pt::from(v),
            FirstLineIndent::None => Pt::ZERO,
        })
        .unwrap_or(Pt::ZERO);

    // §17.3.1.33: when autoSpacing is true, use 14pt instead of explicit value.
    let space_before = if props.spacing.and_then(|s| s.before_auto_spacing) == Some(true) {
        Pt::new(14.0)
    } else {
        props
            .spacing
            .and_then(|s| s.before)
            .map(Pt::from)
            .unwrap_or(Pt::ZERO)
    };
    let space_after = if props.spacing.and_then(|s| s.after_auto_spacing) == Some(true) {
        Pt::new(14.0)
    } else {
        props
            .spacing
            .and_then(|s| s.after)
            .map(Pt::from)
            .unwrap_or(Pt::ZERO)
    };

    // §17.3.1.33: line spacing defaults to single (auto, 240 twips = 1.0x).
    let line_spacing = props
        .spacing
        .and_then(|s| s.line)
        .map(|ls| match ls {
            LineSpacing::Auto(v) => LineSpacingRule::Auto(Pt::from(v).raw() / 12.0),
            LineSpacing::Exact(v) => LineSpacingRule::Exact(Pt::from(v)),
            LineSpacing::AtLeast(v) => LineSpacingRule::AtLeast(Pt::from(v)),
        })
        .unwrap_or(LineSpacingRule::Auto(1.0));

    // §17.3.1.38: convert tab stops to layout format.
    // Clear entries are directives consumed during style merging, not layout stops.
    let tabs: Vec<TabStopDef> = props
        .tabs
        .iter()
        .filter(|t| t.alignment != model::TabAlignment::Clear)
        .map(|t| TabStopDef {
            position: Pt::from(t.position),
            alignment: t.alignment,
            leader: t.leader,
        })
        .collect();

    ParagraphStyle {
        alignment: props.alignment.unwrap_or(model::Alignment::Start),
        space_before,
        space_after,
        indent_left,
        indent_right,
        indent_first_line,
        line_spacing,
        tabs,
        drop_cap: None,
        borders: resolve_paragraph_borders(props),
        shading: props
            .shading
            .as_ref()
            .map(|s| resolve_color(s.fill, ColorContext::Background)),
        keep_next: props.keep_next.unwrap_or(false),
        contextual_spacing: props.contextual_spacing.unwrap_or(false),
        style_id: None, // set by caller when available
        page_floats: Vec::new(),
        page_y: crate::render::dimension::Pt::ZERO,
        page_x: crate::render::dimension::Pt::ZERO,
        page_content_width: crate::render::dimension::Pt::ZERO,
    }
}

/// §17.3.1.24: resolve paragraph borders.
fn resolve_paragraph_borders(props: &model::ParagraphProperties) -> Option<ParagraphBorderStyle> {
    let pbdr = props.borders.as_ref()?;

    let convert = |b: &model::Border| -> BorderLine {
        BorderLine {
            width: Pt::from(b.width),
            color: resolve_color(b.color, ColorContext::Text),
            space: Pt::from(b.space),
        }
    };

    let style = ParagraphBorderStyle {
        top: pbdr.top.as_ref().map(convert),
        bottom: pbdr.bottom.as_ref().map(convert),
        left: pbdr.left.as_ref().map(convert),
        right: pbdr.right.as_ref().map(convert),
    };

    if style.top.is_some()
        || style.bottom.is_some()
        || style.left.is_some()
        || style.right.is_some()
    {
        Some(style)
    } else {
        None
    }
}

/// Convert a model `Border` to a layout `TableBorderLine`.
fn convert_model_border(b: &model::Border) -> TableBorderLine {
    TableBorderLine {
        width: Pt::from(b.width),
        color: resolve_color(b.color, ColorContext::Text),
        style: match b.style {
            model::BorderStyle::Double => TableBorderStyle::Double,
            _ => TableBorderStyle::Single,
        },
    }
}

/// Convert a model cell border to a `CellBorderOverride`.
pub(super) fn convert_cell_border_override(b: &Option<model::Border>) -> Option<CellBorderOverride> {
    b.as_ref().map(|b| {
        if b.style == model::BorderStyle::None {
            CellBorderOverride::Nil
        } else {
            CellBorderOverride::Border(convert_model_border(b))
        }
    })
}

/// §17.4.38: merge direct table borders over style borders.
/// Each edge in `direct` overrides the corresponding edge in `style`;
/// unspecified edges (`None`) inherit from the style.
pub(super) fn merge_table_borders(
    direct: &model::TableBorders,
    style: &model::TableBorders,
) -> model::TableBorders {
    model::TableBorders {
        top: direct.top.or(style.top),
        bottom: direct.bottom.or(style.bottom),
        left: direct.left.or(style.left),
        right: direct.right.or(style.right),
        inside_h: direct.inside_h.or(style.inside_h),
        inside_v: direct.inside_v.or(style.inside_v),
    }
}

/// Convert model `TableBorders` to a layout `TableBorderConfig`.
/// §17.4.38: borders with `val="none"` or `val="nil"` are suppressed.
pub(super) fn convert_table_border_config(b: &model::TableBorders) -> TableBorderConfig {
    let convert = |border: &Option<model::Border>| -> Option<TableBorderLine> {
        border.as_ref().and_then(|b| {
            if b.style == model::BorderStyle::None {
                None
            } else {
                Some(convert_model_border(b))
            }
        })
    };
    TableBorderConfig {
        top: convert(&b.top),
        bottom: convert(&b.bottom),
        left: convert(&b.left),
        right: convert(&b.right),
        inside_h: convert(&b.inside_h),
        inside_v: convert(&b.inside_v),
    }
}

/// Split text fragments wider than `max_width` into per-character fragments
/// with individually measured widths. Used in narrow table cells for
/// character-level line breaking.
pub(super) fn split_oversized_fragments(
    fragments: Vec<Fragment>,
    max_width: Pt,
    ctx: &BuildContext,
) -> Vec<Fragment> {
    if max_width <= Pt::ZERO {
        return fragments;
    }
    let mut result = Vec::with_capacity(fragments.len());
    for frag in fragments {
        match &frag {
            Fragment::Text {
                text, width, font, ..
            } if *width > max_width && text.chars().count() > 1 => {
                // Re-measure each character individually.
                for ch in text.chars() {
                    let ch_str = ch.to_string();
                    let (w, m) = ctx.measurer.measure(&ch_str, font);
                    if let Fragment::Text {
                        color,
                        shading,
                        border,
                        hyperlink_url,
                        baseline_offset,
                        ..
                    } = &frag
                    {
                        result.push(Fragment::Text {
                            text: Rc::from(ch_str.as_str()),
                            font: font.clone(),
                            color: *color,
                            shading: *shading,
                            border: *border,
                            width: w,
                            trimmed_width: w,
                            metrics: m,
                            hyperlink_url: hyperlink_url.clone(),
                            baseline_offset: *baseline_offset,
                            text_offset: Pt::ZERO,
                        });
                    }
                }
            }
            _ => result.push(frag),
        }
    }
    result
}

/// Populate image data on Fragment::Image fragments from the media map.
pub(super) fn populate_image_data(fragments: &mut [Fragment], media: &HashMap<model::RelId, Vec<u8>>) {
    for frag in fragments.iter_mut() {
        if let Fragment::Image {
            rel_id, image_data, ..
        } = frag
        {
            if image_data.is_none() {
                if let Some(bytes) = media.get(&model::RelId::new(rel_id.as_str())) {
                    *image_data = Some(bytes.as_slice().into());
                }
            }
        }
    }
}

/// Remap PUA codepoints (0xF0xx) from legacy Symbol/Wingdings encoding
/// to standard Unicode, and return a portable font to render them with.
///
/// OOXML stores Symbol/Wingdings characters as PUA codepoints. These are
/// not portable across platforms — different OS font versions have different
/// cmap coverage. The standard approach (used by LibreOffice, Google Docs)
/// is to remap to Unicode equivalents from the official mapping tables:
/// - Symbol: unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt
/// - Wingdings: standard Microsoft Wingdings-to-Unicode mapping
///
/// Returns (remapped_text, font_family).
pub(super) fn remap_legacy_font_chars(
    text: &str,
    font_family: &str,
    fallback_family: &str,
) -> (String, String) {
    let is_symbol = font_family.eq_ignore_ascii_case("Symbol");
    let is_wingdings = font_family.eq_ignore_ascii_case("Wingdings");

    if !is_symbol && !is_wingdings {
        let family = if font_family.is_empty() {
            fallback_family
        } else {
            font_family
        };
        return (text.to_string(), family.to_string());
    }

    let remapped: String = text
        .chars()
        .map(|ch| {
            let code = ch as u32;
            if is_symbol && (0xF020..=0xF0FF).contains(&code) {
                // Symbol font PUA mapping per unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt
                match code {
                    0xF020 => '\u{0020}',                                           // SPACE
                    0xF021 => '\u{0021}', // EXCLAMATION MARK
                    0xF025 => '\u{0025}', // PERCENT SIGN
                    0xF028 => '\u{0028}', // LEFT PARENTHESIS
                    0xF029 => '\u{0029}', // RIGHT PARENTHESIS
                    0xF02B => '\u{002B}', // PLUS SIGN
                    0xF02E => '\u{002E}', // FULL STOP
                    0xF030..=0xF039 => char::from_u32(code - 0xF000).unwrap_or(ch), // DIGITS
                    0xF03C => '\u{003C}', // LESS-THAN SIGN
                    0xF03D => '\u{003D}', // EQUALS SIGN
                    0xF03E => '\u{003E}', // GREATER-THAN SIGN
                    0xF05B => '\u{005B}', // LEFT SQUARE BRACKET
                    0xF05D => '\u{005D}', // RIGHT SQUARE BRACKET
                    0xF07B => '\u{007B}', // LEFT CURLY BRACKET
                    0xF07C => '\u{007C}', // VERTICAL LINE
                    0xF07D => '\u{007D}', // RIGHT CURLY BRACKET
                    0xF07E => '\u{223C}', // TILDE OPERATOR
                    0xF0A0 => '\u{20AC}', // EURO SIGN
                    0xF0A5 => '\u{221E}', // INFINITY
                    0xF0A7 => '\u{2663}', // BLACK CLUB SUIT
                    0xF0A8 => '\u{2666}', // BLACK DIAMOND SUIT
                    0xF0A9 => '\u{2665}', // BLACK HEART SUIT
                    0xF0AA => '\u{2660}', // BLACK SPADE SUIT
                    0xF0AB => '\u{2194}', // LEFT RIGHT ARROW
                    0xF0AC => '\u{2190}', // LEFTWARDS ARROW
                    0xF0AD => '\u{2191}', // UPWARDS ARROW
                    0xF0AE => '\u{2192}', // RIGHTWARDS ARROW
                    0xF0AF => '\u{2193}', // DOWNWARDS ARROW
                    0xF0B0 => '\u{00B0}', // DEGREE SIGN
                    0xF0B1 => '\u{00B1}', // PLUS-MINUS SIGN
                    0xF0B2 => '\u{2033}', // DOUBLE PRIME
                    0xF0B3 => '\u{2265}', // GREATER-THAN OR EQUAL TO
                    0xF0B4 => '\u{00D7}', // MULTIPLICATION SIGN
                    0xF0B5 => '\u{221D}', // PROPORTIONAL TO
                    0xF0B7 => '\u{2022}', // BULLET
                    0xF0B8 => '\u{00F7}', // DIVISION SIGN
                    0xF0B9 => '\u{2260}', // NOT EQUAL TO
                    0xF0BA => '\u{2261}', // IDENTICAL TO
                    0xF0BB => '\u{2248}', // ALMOST EQUAL TO
                    0xF0BC => '\u{2026}', // HORIZONTAL ELLIPSIS
                    0xF0C0 => '\u{2135}', // ALEF SYMBOL
                    0xF0C1 => '\u{2111}', // BLACK-LETTER CAPITAL I
                    0xF0C2 => '\u{211C}', // BLACK-LETTER CAPITAL R
                    0xF0C3 => '\u{2118}', // SCRIPT CAPITAL P
                    0xF0C5 => '\u{2297}', // CIRCLED TIMES
                    0xF0C6 => '\u{2295}', // CIRCLED PLUS
                    0xF0C7 => '\u{2205}', // EMPTY SET
                    0xF0C8 => '\u{2229}', // INTERSECTION
                    0xF0C9 => '\u{222A}', // UNION
                    0xF0CB => '\u{2283}', // SUPERSET OF
                    0xF0CC => '\u{2287}', // SUPERSET OF OR EQUAL TO
                    0xF0CD => '\u{2284}', // NOT A SUBSET OF
                    0xF0CE => '\u{2282}', // SUBSET OF
                    0xF0CF => '\u{2286}', // SUBSET OF OR EQUAL TO
                    0xF0D0 => '\u{2208}', // ELEMENT OF
                    0xF0D1 => '\u{2209}', // NOT AN ELEMENT OF
                    0xF0D5 => '\u{220F}', // N-ARY PRODUCT
                    0xF0D6 => '\u{221A}', // SQUARE ROOT
                    0xF0D7 => '\u{22C5}', // DOT OPERATOR
                    0xF0D8 => '\u{00AC}', // NOT SIGN
                    0xF0D9 => '\u{2227}', // LOGICAL AND
                    0xF0DA => '\u{2228}', // LOGICAL OR
                    0xF0E0 => '\u{21D0}', // LEFTWARDS DOUBLE ARROW
                    0xF0E1 => '\u{21D1}', // UPWARDS DOUBLE ARROW
                    0xF0E2 => '\u{21D2}', // RIGHTWARDS DOUBLE ARROW
                    0xF0E3 => '\u{21D3}', // DOWNWARDS DOUBLE ARROW
                    0xF0E4 => '\u{21D4}', // LEFT RIGHT DOUBLE ARROW
                    0xF0E5 => '\u{2329}', // LEFT-POINTING ANGLE BRACKET
                    0xF0F1 => '\u{232A}', // RIGHT-POINTING ANGLE BRACKET
                    0xF0F2 => '\u{222B}', // INTEGRAL
                    _ => ch,
                }
            } else if is_wingdings && (0xF020..=0xF0FF).contains(&code) {
                // Wingdings PUA mapping per Microsoft Wingdings-to-Unicode table
                match code {
                    0xF021 => '\u{270E}',  // LOWER RIGHT PENCIL
                    0xF022 => '\u{2702}',  // BLACK SCISSORS
                    0xF023 => '\u{2701}',  // UPPER BLADE SCISSORS
                    0xF028 => '\u{1F4CB}', // CLIPBOARD (may need supplementary)
                    0xF029 => '\u{1F4CB}', // CLIPBOARD
                    0xF041 => '\u{FE4E}',  // WAVY LOW LINE (approximate)
                    0xF046 => '\u{1F44D}', // THUMBS UP SIGN
                    0xF04C => '\u{2639}',  // WHITE FROWNING FACE
                    0xF04A => '\u{263A}',  // WHITE SMILING FACE
                    0xF06C => '\u{25CF}',  // BLACK CIRCLE
                    0xF06D => '\u{274D}',  // SHADOWED WHITE CIRCLE
                    0xF06E => '\u{25A0}',  // BLACK SQUARE
                    0xF06F => '\u{25A1}',  // WHITE SQUARE
                    0xF070 => '\u{25A1}',  // WHITE SQUARE (alt)
                    0xF071 => '\u{2751}',  // LOWER RIGHT SHADOWED WHITE SQUARE
                    0xF072 => '\u{2752}',  // UPPER RIGHT SHADOWED WHITE SQUARE
                    0xF073 => '\u{25C6}',  // BLACK DIAMOND
                    0xF074 => '\u{2756}',  // BLACK DIAMOND MINUS WHITE X
                    0xF076 => '\u{2756}',  // BLACK DIAMOND MINUS WHITE X
                    0xF09F => '\u{2708}',  // AIRPLANE
                    0xF0A1 => '\u{270C}',  // VICTORY HAND
                    0xF0A4 => '\u{261C}',  // WHITE LEFT POINTING INDEX
                    0xF0A5 => '\u{261E}',  // WHITE RIGHT POINTING INDEX
                    0xF0A7 => '\u{25AA}',  // BLACK SMALL SQUARE
                    0xF0A8 => '\u{25FB}',  // WHITE MEDIUM SQUARE
                    0xF0D5 => '\u{232B}',  // ERASE TO THE LEFT
                    0xF0D8 => '\u{27A2}',  // THREE-D TOP-LIGHTED RIGHTWARDS ARROWHEAD
                    0xF0E8 => '\u{2B22}',  // BLACK HEXAGON (approximate)
                    0xF0F0 => '\u{2B1A}',  // DOTTED SQUARE (approximate)
                    0xF0FC => '\u{2714}',  // HEAVY CHECK MARK
                    0xF0FB => '\u{2718}',  // HEAVY BALLOT X
                    0xF0FE => '\u{2612}',  // BALLOT BOX WITH X (approximate)
                    _ => ch,
                }
            } else {
                ch
            }
        })
        .collect();

    // After PUA→Unicode remapping, the original Symbol/Wingdings font
    // cannot render the standard Unicode codepoints (legacy fonts lack
    // Unicode cmaps). Use the document's fallback font for the remapped
    // glyphs — standard Unicode bullets (U+2022, U+25AA) are in most
    // text fonts.
    (remapped, fallback_family.to_string())
}

/// Populate underline position/thickness from Skia font metrics.
pub(super) fn populate_underline_metrics(fragments: &mut [Fragment], measurer: &TextMeasurer) {
    for frag in fragments.iter_mut() {
        if let Fragment::Text { font, .. } = frag {
            if font.underline {
                let (pos, thickness) = measurer.underline_metrics(font);
                font.underline_position = pos;
                font.underline_thickness = thickness;
            }
        }
    }
}

/// Extract the display size for a picture bullet from its VML shape style.
/// Falls back to 9pt × 9pt (common Word default for picture bullets).
pub(super) fn pic_bullet_size(bullet: &model::NumPicBullet) -> PtSize {
    use crate::model::VmlLengthUnit;

    let default = PtSize::new(Pt::new(9.0), Pt::new(9.0));
    let shape = match bullet.pict.as_ref().and_then(|p| p.shapes.first()) {
        Some(s) => s,
        None => return default,
    };

    let to_pt = |len: &crate::model::VmlLength| -> Pt {
        let val = len.value as f32;
        match len.unit {
            VmlLengthUnit::Pt => Pt::new(val),
            VmlLengthUnit::In => Pt::new(val * 72.0),
            VmlLengthUnit::Cm => Pt::new(val * 28.3465),
            VmlLengthUnit::Mm => Pt::new(val * 2.83465),
            VmlLengthUnit::Px => Pt::new(val * 0.75),
            _ => Pt::new(val),
        }
    };

    let w = shape
        .style
        .width
        .as_ref()
        .map(to_pt)
        .unwrap_or(default.width);
    let h = shape
        .style
        .height
        .as_ref()
        .map(to_pt)
        .unwrap_or(default.height);
    PtSize::new(w, h)
}
