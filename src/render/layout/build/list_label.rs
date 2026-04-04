//! List label injection — prepend bullet/number labels to paragraph fragments.
//!
//! §17.9.22: when a paragraph carries a numbering reference, resolve the label
//! text (or picture bullet) and inject it as the first fragment, followed by a
//! tab that advances to the body text indent position.

use std::rc::Rc;

use crate::model::{self, ParagraphProperties};
use crate::render::dimension::Pt;
use crate::render::layout::fragment::{FontProps, Fragment};

use super::convert::{pic_bullet_size, remap_legacy_font_chars, resolve_paragraph_defaults};
use super::{BuildContext, BuildState};

/// Inject list label fragments into a paragraph if it has a numbering reference.
///
/// Updates `fragments` (prepends label + tab), `merged_props` (overrides indentation
/// from the numbering level), and `state.list_counters` (increments/resets counters).
pub(super) fn inject_list_label(
    para: &model::Paragraph,
    fragments: &mut Vec<Fragment>,
    merged_props: &mut ParagraphProperties,
    ctx: &BuildContext,
    state: &mut BuildState,
) {
    let num_ref = match merged_props.numbering {
        Some(ref nr) => nr,
        None => return,
    };

    let num_id = model::NumId::new(num_ref.num_id);
    let level = num_ref.level;

    let levels = match ctx.resolved.numbering.get(&num_id) {
        Some(levels) => levels,
        None => return,
    };

    // Update counters: increment this level, reset deeper levels.
    {
        let counters = &mut state.list_counters;
        let count = counters
            .entry((num_id, level))
            .or_insert_with(|| levels.get(level as usize).map(|l| l.start).unwrap_or(1) - 1);
        *count += 1;
        // Reset deeper levels.
        let max_level = levels.len() as u8;
        for deeper in (level + 1)..max_level {
            counters.remove(&(num_id, deeper));
        }
    }

    let level_def = levels.get(level as usize);

    // §17.9.10: check for picture bullet before text label.
    let pic_bullet_injected = level_def
        .and_then(|l| l.lvl_pic_bullet_id)
        .and_then(|pic_id| ctx.resolved.pic_bullets.get(&pic_id))
        .and_then(|bullet| {
            let rel_id = bullet
                .pict
                .as_ref()?
                .shapes
                .first()?
                .image_data
                .as_ref()?
                .rel_id
                .as_ref()?;
            let image_bytes = ctx.media().get(rel_id)?;
            // Size from VML shape style (width/height), default 9pt.
            let size = pic_bullet_size(bullet);
            let label_frag = Fragment::Image {
                size,
                rel_id: rel_id.as_str().to_string(),
                image_data: Some(image_bytes.clone()),
            };
            Some((label_frag, size.height))
        });

    if let Some((label_frag, label_height)) = pic_bullet_injected {
        let hanging = extract_hanging(level_def);
        let tab_frag = Fragment::Tab {
            line_height: label_height,
            fitting_width: Some(hanging),
        };
        fragments.insert(0, tab_frag);
        fragments.insert(0, label_frag);

        if let Some(lvl_left) = level_def
            .and_then(|l| l.indentation.as_ref())
            .and_then(|ind| ind.start)
        {
            merged_props.tabs.insert(
                0,
                crate::model::TabStop {
                    position: lvl_left,
                    alignment: crate::model::TabAlignment::Left,
                    leader: crate::model::TabLeader::None,
                },
            );
        }
    } else {
        inject_text_label(
            para,
            fragments,
            merged_props,
            ctx,
            &state.list_counters,
            levels,
            level,
            level_def,
        );
    }

    // §17.9.23: numbering level pPr overrides the paragraph style.
    // Only the paragraph's direct ind overrides the numbering level.
    if let Some(lvl_ind) = levels
        .get(level as usize)
        .and_then(|l| l.indentation.as_ref())
    {
        let mut ind = *lvl_ind;
        if let Some(direct) = para.properties.indentation {
            if let Some(start) = direct.start {
                ind.start = Some(start);
            }
            if let Some(end) = direct.end {
                ind.end = Some(end);
            }
            if let Some(first_line) = direct.first_line {
                ind.first_line = Some(first_line);
            }
        }
        merged_props.indentation = Some(ind);
    }
}

/// Inject a text label (non-picture bullet) into the paragraph fragments.
#[allow(clippy::too_many_arguments)]
fn inject_text_label(
    para: &model::Paragraph,
    fragments: &mut Vec<Fragment>,
    merged_props: &mut ParagraphProperties,
    ctx: &BuildContext,
    counters: &std::collections::HashMap<(model::NumId, u8), u32>,
    levels: &[crate::render::resolve::numbering::ResolvedNumberingLevel],
    level: u8,
    level_def: Option<&crate::render::resolve::numbering::ResolvedNumberingLevel>,
) {
    let num_id = model::NumId::new(merged_props.numbering.as_ref().unwrap().num_id);
    let label_text =
        match crate::render::resolve::numbering::format_list_label(levels, level, counters, num_id)
        {
            Some(t) => t,
            None => return,
        };

    // Resolve label font from level run_properties or paragraph defaults.
    let (default_family, default_size, default_color, _, _) =
        resolve_paragraph_defaults(para, ctx.resolved, false);
    let level_font_family = level_def
        .and_then(|l| l.run_properties.as_ref())
        .and_then(|rp| crate::render::resolve::fonts::effective_font(&rp.fonts))
        .unwrap_or("");

    // Remap PUA codepoints from legacy Symbol/Wingdings encoding.
    let (label_text, label_family) =
        remap_legacy_font_chars(&label_text, level_font_family, &default_family);
    let label_family: Rc<str> = Rc::from(label_family.as_str());

    let label_size = level_def
        .and_then(|l| l.run_properties.as_ref())
        .and_then(|rp| rp.font_size)
        .map(Pt::from)
        .unwrap_or(default_size);
    let label_bold = level_def
        .and_then(|l| l.run_properties.as_ref())
        .and_then(|rp| rp.bold)
        .unwrap_or(false);
    let label_italic = level_def
        .and_then(|l| l.run_properties.as_ref())
        .and_then(|rp| rp.italic)
        .unwrap_or(false);

    let label_font = FontProps {
        family: label_family,
        size: label_size,
        bold: label_bold,
        italic: label_italic,
        underline: false,
        char_spacing: Pt::ZERO,
        underline_position: Pt::ZERO,
        underline_thickness: Pt::ZERO,
    };
    let (w, m) = ctx.measurer.measure(&label_text, &label_font);
    let h = m.height();

    let hanging = extract_hanging(level_def);
    // §17.9.7: lvlJc controls label justification within the hanging indent area.
    let jc = level_def.and_then(|l| l.justification);
    let text_offset = match jc {
        Some(crate::model::Alignment::End) => -w,
        Some(crate::model::Alignment::Center) => w * -0.5,
        _ => Pt::ZERO,
    };
    let label_width = w;
    let label_frag = Fragment::Text {
        text: Rc::from(label_text.as_str()),
        font: label_font.clone(),
        color: default_color,
        shading: None,
        border: None,
        width: label_width,
        trimmed_width: label_width,
        metrics: m,
        hyperlink_url: None,
        baseline_offset: Pt::ZERO,
        text_offset,
    };
    let tab_fitting = (hanging - label_width).max(Pt::ZERO);
    let tab_frag = Fragment::Tab {
        line_height: h,
        fitting_width: Some(tab_fitting),
    };
    fragments.insert(0, tab_frag);
    fragments.insert(0, label_frag);

    // Add implicit tab stop at numLvl.left so the tab lands at the body text position.
    let lvl_left = level_def
        .and_then(|l| l.indentation.as_ref())
        .and_then(|ind| ind.start);
    if let Some(lvl_left) = lvl_left {
        merged_props.tabs.insert(
            0,
            crate::model::TabStop {
                position: lvl_left,
                alignment: crate::model::TabAlignment::Left,
                leader: crate::model::TabLeader::None,
            },
        );
    }
}

/// Extract the hanging indent from a numbering level definition.
fn extract_hanging(
    level_def: Option<&crate::render::resolve::numbering::ResolvedNumberingLevel>,
) -> Pt {
    level_def
        .and_then(|l| l.indentation.as_ref())
        .and_then(|ind| ind.first_line)
        .map(|fl| match fl {
            model::FirstLineIndent::Hanging(v) => Pt::from(v),
            _ => Pt::ZERO,
        })
        .unwrap_or(Pt::ZERO)
}
