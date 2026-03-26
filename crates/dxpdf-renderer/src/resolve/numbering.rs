//! Numbering resolution — flatten abstract + instance + overrides into lookup table.

use std::collections::HashMap;

use dxpdf_docx_model::model::{
    NumberFormat, NumberingDefinitions, NumberingLevelDefinition, NumId,
};

/// A resolved numbering level — ready for label generation.
#[derive(Clone, Debug)]
pub struct ResolvedNumberingLevel {
    pub format: NumberFormat,
    pub level_text: String,
    pub start: u32,
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
    }
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
