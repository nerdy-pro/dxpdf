//! Numbering resolution — flatten abstract + instance + overrides into lookup table.

use std::collections::HashMap;

use dxpdf_docx_model::model::{
    Alignment, NumberFormat, NumberingDefinitions, NumberingLevelDefinition, NumId, NumPicBulletId,
    RunProperties, Indentation,
};

/// A resolved numbering level — ready for label generation.
#[derive(Clone, Debug)]
pub struct ResolvedNumberingLevel {
    pub format: NumberFormat,
    pub level_text: String,
    pub start: u32,
    /// §17.9.3: run properties for the numbering symbol (font, color, etc.).
    pub run_properties: Option<RunProperties>,
    /// §17.9.3: paragraph indentation from the numbering level definition.
    /// When present, overrides the paragraph style's indentation.
    pub indentation: Option<Indentation>,
    /// §17.9.7: justification of the numbering symbol (left, center, right).
    pub justification: Option<Alignment>,
    /// §17.9.10: reference to a picture bullet definition.
    pub lvl_pic_bullet_id: Option<NumPicBulletId>,
}

/// Resolve numbering definitions into a flat lookup: NumId → Vec<ResolvedNumberingLevel>.
/// Each instance's abstract definition is looked up and level overrides applied.
pub fn resolve_numbering(
    defs: &NumberingDefinitions,
) -> HashMap<NumId, Vec<ResolvedNumberingLevel>> {
    let mut result = HashMap::new();

    for (num_id, instance) in &defs.numbering_instances {
        let abstract_levels = defs
            .abstract_nums
            .get(&instance.abstract_num_id)
            .map(|a| a.levels.as_slice())
            .unwrap_or(&[]);

        let mut levels: Vec<ResolvedNumberingLevel> = abstract_levels
            .iter()
            .map(resolve_level)
            .collect();

        // Apply level overrides
        for ovr in &instance.level_overrides {
            let idx = ovr.level as usize;
            let resolved = resolve_level(ovr);
            if idx < levels.len() {
                levels[idx] = resolved;
            }
            // If override references a level beyond abstract, we skip it
        }

        result.insert(*num_id, levels);
    }

    result
}

fn resolve_level(def: &NumberingLevelDefinition) -> ResolvedNumberingLevel {
    ResolvedNumberingLevel {
        format: def.format.unwrap_or(NumberFormat::None),
        level_text: def.level_text.clone(),
        start: def.start.unwrap_or(1),
        run_properties: def.run_properties.clone(),
        indentation: def.indentation,
        justification: def.justification,
        lvl_pic_bullet_id: def.lvl_pic_bullet_id,
    }
}

/// §17.9.11: format a list label by expanding the level_text template.
/// `%1` is replaced with the formatted counter for level 0, `%2` for level 1, etc.
/// Returns `None` for `NumberFormat::None`.
pub fn format_list_label(
    levels: &[ResolvedNumberingLevel],
    level: u8,
    counters: &HashMap<(NumId, u8), u32>,
    num_id: NumId,
) -> Option<String> {
    let lvl = levels.get(level as usize)?;
    if lvl.format == NumberFormat::None {
        return None;
    }
    if lvl.format == NumberFormat::Bullet {
        return Some(lvl.level_text.clone());
    }

    // Expand template: %1 → level 0 counter, %2 → level 1 counter, etc.
    let mut result = lvl.level_text.clone();
    for i in (0..=level).rev() {
        let placeholder = format!("%{}", i + 1);
        if result.contains(&placeholder) {
            let count = counters.get(&(num_id, i)).copied().unwrap_or(1);
            let fmt = levels.get(i as usize).map(|l| l.format).unwrap_or(NumberFormat::Decimal);
            let formatted = format_number(count, fmt);
            result = result.replace(&placeholder, &formatted);
        }
    }
    Some(result)
}

/// Format a number according to the OOXML number format.
fn format_number(n: u32, fmt: NumberFormat) -> String {
    match fmt {
        NumberFormat::Decimal => n.to_string(),
        NumberFormat::LowerLetter => to_letter_lower(n),
        NumberFormat::UpperLetter => to_letter_upper(n),
        NumberFormat::LowerRoman => to_roman_lower(n),
        NumberFormat::UpperRoman => to_roman_upper(n),
        NumberFormat::Ordinal => format_ordinal(n),
        _ => n.to_string(),
    }
}

fn to_letter_lower(n: u32) -> String {
    if n == 0 { return String::new(); }
    let idx = ((n - 1) % 26) as u8;
    String::from((b'a' + idx) as char)
}

fn to_letter_upper(n: u32) -> String {
    to_letter_lower(n).to_uppercase()
}

fn to_roman_lower(mut n: u32) -> String {
    const VALS: [(u32, &str); 13] = [
        (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"),
        (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
        (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"),
    ];
    let mut s = String::new();
    for &(val, sym) in &VALS {
        while n >= val { s.push_str(sym); n -= val; }
    }
    s
}

fn to_roman_upper(n: u32) -> String {
    to_roman_lower(n).to_uppercase()
}

fn format_ordinal(n: u32) -> String {
    let suffix = match n % 100 {
        11..=13 => "th",
        _ => match n % 10 {
            1 => "st",
            2 => "nd",
            3 => "rd",
            _ => "th",
        },
    };
    format!("{n}{suffix}")
}

#[cfg(test)]
mod tests {
    use super::*;
    use dxpdf_docx_model::model::*;

    fn make_defs(
        abstracts: Vec<(AbstractNumId, Vec<NumberingLevelDefinition>)>,
        instances: Vec<(NumId, AbstractNumId, Vec<NumberingLevelDefinition>)>,
    ) -> NumberingDefinitions {
        NumberingDefinitions {
            abstract_nums: abstracts
                .into_iter()
                .map(|(id, levels)| (id, AbstractNumbering { levels }))
                .collect(),
            numbering_instances: instances
                .into_iter()
                .map(|(num_id, abstract_id, overrides)| {
                    (
                        num_id,
                        NumberingInstance {
                            abstract_num_id: abstract_id,
                            level_overrides: overrides,
                        },
                    )
                })
                .collect(),
            pic_bullets: HashMap::new(),
        }
    }

    fn level(lvl: u8, fmt: NumberFormat, text: &str, start: u32) -> NumberingLevelDefinition {
        NumberingLevelDefinition {
            level: lvl,
            format: Some(fmt),
            level_text: text.to_string(),
            start: Some(start),
            justification: None,
            indentation: None,
            run_properties: None,
            lvl_pic_bullet_id: None,
        }
    }

    #[test]
    fn single_instance_resolves_from_abstract() {
        let defs = make_defs(
            vec![(
                AbstractNumId::new(0),
                vec![level(0, NumberFormat::Decimal, "%1.", 1)],
            )],
            vec![(NumId::new(1), AbstractNumId::new(0), vec![])],
        );

        let resolved = resolve_numbering(&defs);
        let levels = resolved.get(&NumId::new(1)).unwrap();

        assert_eq!(levels.len(), 1);
        assert_eq!(levels[0].format, NumberFormat::Decimal);
        assert_eq!(levels[0].level_text, "%1.");
        assert_eq!(levels[0].start, 1);
    }

    #[test]
    fn level_override_replaces_abstract_level() {
        let defs = make_defs(
            vec![(
                AbstractNumId::new(0),
                vec![
                    level(0, NumberFormat::Decimal, "%1.", 1),
                    level(1, NumberFormat::LowerLetter, "%2)", 1),
                ],
            )],
            vec![(
                NumId::new(1),
                AbstractNumId::new(0),
                // Override level 0 to bullet
                vec![level(0, NumberFormat::Bullet, "•", 1)],
            )],
        );

        let resolved = resolve_numbering(&defs);
        let levels = resolved.get(&NumId::new(1)).unwrap();

        assert_eq!(levels.len(), 2);
        assert_eq!(levels[0].format, NumberFormat::Bullet, "overridden");
        assert_eq!(levels[0].level_text, "•");
        assert_eq!(levels[1].format, NumberFormat::LowerLetter, "from abstract");
    }

    #[test]
    fn missing_abstract_produces_empty_levels() {
        let defs = make_defs(
            vec![],
            vec![(NumId::new(1), AbstractNumId::new(99), vec![])],
        );

        let resolved = resolve_numbering(&defs);
        let levels = resolved.get(&NumId::new(1)).unwrap();
        assert!(levels.is_empty());
    }

    #[test]
    fn multiple_instances_same_abstract() {
        let defs = make_defs(
            vec![(
                AbstractNumId::new(0),
                vec![level(0, NumberFormat::Decimal, "%1.", 1)],
            )],
            vec![
                (NumId::new(1), AbstractNumId::new(0), vec![]),
                (
                    NumId::new(2),
                    AbstractNumId::new(0),
                    vec![level(0, NumberFormat::Decimal, "%1)", 10)],
                ),
            ],
        );

        let resolved = resolve_numbering(&defs);

        let l1 = resolved.get(&NumId::new(1)).unwrap();
        assert_eq!(l1[0].level_text, "%1.");
        assert_eq!(l1[0].start, 1);

        let l2 = resolved.get(&NumId::new(2)).unwrap();
        assert_eq!(l2[0].level_text, "%1)");
        assert_eq!(l2[0].start, 10);
    }

    #[test]
    fn level_with_no_format_defaults_to_none() {
        let defs = make_defs(
            vec![(
                AbstractNumId::new(0),
                vec![NumberingLevelDefinition {
                    level: 0,
                    format: None,
                    level_text: String::new(),
                    start: None,
                    justification: None,
                    indentation: None,
                    run_properties: None,
                    lvl_pic_bullet_id: None,
                }],
            )],
            vec![(NumId::new(1), AbstractNumId::new(0), vec![])],
        );

        let resolved = resolve_numbering(&defs);
        let levels = resolved.get(&NumId::new(1)).unwrap();
        assert_eq!(levels[0].format, NumberFormat::None);
        assert_eq!(levels[0].start, 1);
    }
}
