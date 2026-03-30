//! Style sheet — style definitions, latent styles, table style overrides.

use std::collections::HashMap;

use super::identifiers::StyleId;
use super::paragraph::ParagraphProperties;
use super::run_properties::RunProperties;
use super::table::{TableCellProperties, TableProperties, TableRowProperties};

/// Raw style definitions as parsed from `word/styles.xml`.
/// No inheritance resolution — `basedOn` references are preserved as-is.
#[derive(Clone, Debug, Default)]
pub struct StyleSheet {
    /// Document-level default paragraph properties (from `w:docDefaults/w:pPrDefault`).
    pub doc_defaults_paragraph: ParagraphProperties,
    /// Document-level default run properties (from `w:docDefaults/w:rPrDefault`).
    pub doc_defaults_run: RunProperties,
    /// All style definitions keyed by style ID.
    pub styles: HashMap<StyleId, Style>,
    /// §17.7.4.5: default properties for styles not explicitly defined.
    pub latent_styles: Option<LatentStyles>,
}

/// §17.7.4.5: default properties for latent (not explicitly defined) styles.
#[derive(Clone, Debug)]
pub struct LatentStyles {
    pub default_locked_state: Option<bool>,
    pub default_ui_priority: Option<u32>,
    pub default_semi_hidden: Option<bool>,
    pub default_unhide_when_used: Option<bool>,
    pub default_q_format: Option<bool>,
    pub count: Option<u32>,
    /// §17.7.4.8: per-style exceptions to the defaults.
    pub exceptions: Vec<LatentStyleException>,
}

/// §17.7.4.8: exception to latent style defaults for a specific style name.
#[derive(Clone, Debug)]
pub struct LatentStyleException {
    pub name: Option<String>,
    pub locked: Option<bool>,
    pub ui_priority: Option<u32>,
    pub semi_hidden: Option<bool>,
    pub unhide_when_used: Option<bool>,
    pub q_format: Option<bool>,
}

/// A single style definition.
#[derive(Clone, Debug)]
pub struct Style {
    pub name: Option<String>,
    pub style_type: StyleType,
    /// Parent style ID. Properties not specified here should be inherited from this style.
    pub based_on: Option<StyleId>,
    pub is_default: bool,
    pub paragraph_properties: Option<ParagraphProperties>,
    pub run_properties: Option<RunProperties>,
    pub table_properties: Option<TableProperties>,
    /// §17.7.6.6: conditional formatting overrides for table regions.
    pub table_style_overrides: Vec<TableStyleOverride>,
}

/// §17.7.6.6: formatting override for a specific table region.
#[derive(Clone, Debug)]
pub struct TableStyleOverride {
    /// Which region this override applies to.
    pub override_type: TableStyleOverrideType,
    pub paragraph_properties: Option<ParagraphProperties>,
    pub run_properties: Option<RunProperties>,
    pub table_properties: Option<TableProperties>,
    pub table_row_properties: Option<TableRowProperties>,
    pub table_cell_properties: Option<TableCellProperties>,
}

/// §17.18.89 ST_TblStyleOverrideType — table region for conditional formatting.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableStyleOverrideType {
    FirstRow,
    LastRow,
    FirstCol,
    LastCol,
    Band1Vert,
    Band2Vert,
    Band1Horz,
    Band2Horz,
    NeCell,
    NwCell,
    SeCell,
    SwCell,
    WholeTable,
}

/// The type of a style definition (§17.7.4.17).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StyleType {
    Paragraph,
    Character,
    Table,
    Numbering,
}
