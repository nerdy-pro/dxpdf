//! Numbering definitions — abstract numbering, instances, picture bullets.

use std::collections::HashMap;

use super::formatting::{Alignment, NumberFormat};
use super::identifiers::{AbstractNumId, NumId, NumPicBulletId};
use super::paragraph::Indentation;
use super::run_properties::RunProperties;
use super::vml::Pict;

/// Raw numbering definitions as parsed from `word/numbering.xml`.
#[derive(Clone, Debug, Default)]
pub struct NumberingDefinitions {
    /// Abstract numbering definitions keyed by abstract numbering ID.
    pub abstract_nums: HashMap<AbstractNumId, AbstractNumbering>,
    /// Numbering instances keyed by numbering ID.
    pub numbering_instances: HashMap<NumId, NumberingInstance>,
    /// §17.9.21: picture bullet definitions keyed by numPicBulletId.
    pub pic_bullets: HashMap<NumPicBulletId, NumPicBullet>,
}

/// §17.9.21: a picture bullet definition.
#[derive(Clone, Debug)]
pub struct NumPicBullet {
    /// §17.9.21 @numPicBulletId: unique identifier.
    pub id: NumPicBulletId,
    /// §17.3.3.19: VML picture content.
    pub pict: Option<Pict>,
}

/// An abstract numbering definition.
#[derive(Clone, Debug)]
pub struct AbstractNumbering {
    pub levels: Vec<NumberingLevelDefinition>,
}

/// A single level within an abstract numbering definition.
#[derive(Clone, Debug)]
pub struct NumberingLevelDefinition {
    pub level: u8,
    pub format: Option<NumberFormat>,
    pub level_text: String,
    pub start: Option<u32>,
    /// §17.9.7: justification of the numbering symbol (uses ST_Jc).
    pub justification: Option<Alignment>,
    pub indentation: Option<Indentation>,
    pub run_properties: Option<RunProperties>,
    /// §17.9.10: reference to a picture bullet definition.
    pub lvl_pic_bullet_id: Option<NumPicBulletId>,
}

/// A numbering instance — maps to an abstract numbering, with optional level overrides.
#[derive(Clone, Debug)]
pub struct NumberingInstance {
    pub abstract_num_id: AbstractNumId,
    pub level_overrides: Vec<NumberingLevelDefinition>,
}
