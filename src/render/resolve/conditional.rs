//! §17.7.6: Table conditional formatting resolution.
//!
//! Determines which conditional formatting regions apply to each cell
//! based on its position, tblLook flags, and band sizes. Overlays
//! applicable tblStylePr overrides in priority order.

use crate::model::model::{
    ParagraphProperties, RunProperties, TableCellProperties, TableLook, TableStyleOverride,
    TableStyleOverrideType,
};

/// Resolved conditional formatting for a single cell.
#[derive(Clone, Debug, Default)]
pub struct CellConditionalFormatting {
    pub cell_properties: Option<TableCellProperties>,
    pub run_properties: Option<RunProperties>,
    pub paragraph_properties: Option<ParagraphProperties>,
}

/// §17.7.6: resolve conditional formatting for a cell at (row, col).
///
/// Overlays applicable `tblStylePr` overrides in priority order per §17.7.6:
/// 1. Whole table (lowest)
/// 2. Band1/Band2 Horizontal
/// 3. Band1/Band2 Vertical
/// 4. First/Last Column
/// 5. First/Last Row
/// 6. Corner cells (highest)
#[allow(clippy::too_many_arguments)]
pub fn resolve_cell_conditional(
    row_idx: usize,
    col_idx: usize,
    num_rows: usize,
    num_cols: usize,
    look: Option<&TableLook>,
    overrides: &[TableStyleOverride],
    row_band_size: u32,
    col_band_size: u32,
) -> CellConditionalFormatting {
    let regions = applicable_regions(
        row_idx,
        col_idx,
        num_rows,
        num_cols,
        look,
        row_band_size,
        col_band_size,
    );

    let mut result = CellConditionalFormatting::default();

    // §17.7.6: apply overrides in priority order (lowest first, highest last).
    // Later overlays take precedence.
    for region in &regions {
        if let Some(ovr) = overrides.iter().find(|o| o.override_type == *region) {
            if let Some(ref tcp) = ovr.table_cell_properties {
                overlay_cell_properties(&mut result, tcp);
            }
            if let Some(ref rp) = ovr.run_properties {
                overlay_run_properties(&mut result, rp);
            }
            if let Some(ref pp) = ovr.paragraph_properties {
                overlay_paragraph_properties(&mut result, pp);
            }
        }
    }

    result
}

/// §17.7.6: determine which regions apply to a cell, in priority order
/// (lowest priority first). §17.4.56 tblLook controls which regions are active.
fn applicable_regions(
    row_idx: usize,
    col_idx: usize,
    num_rows: usize,
    num_cols: usize,
    look: Option<&TableLook>,
    row_band_size: u32,
    col_band_size: u32,
) -> Vec<TableStyleOverrideType> {
    let mut regions = Vec::new();

    // §17.4.56: tblLook flags. When absent, all regions are active.
    let first_row_active = look.and_then(|l| l.first_row).unwrap_or(true);
    let last_row_active = look.and_then(|l| l.last_row).unwrap_or(true);
    let first_col_active = look.and_then(|l| l.first_column).unwrap_or(true);
    let last_col_active = look.and_then(|l| l.last_column).unwrap_or(true);
    let h_band_active = !look.and_then(|l| l.no_h_band).unwrap_or(false);
    let v_band_active = !look.and_then(|l| l.no_v_band).unwrap_or(false);

    // §17.7.6 priority 2: horizontal banding.
    if h_band_active {
        // When firstRow is active, banding starts from row 1.
        let band_row = if first_row_active && row_idx > 0 {
            row_idx - 1
        } else {
            row_idx
        };
        let band_size = row_band_size.max(1) as usize;
        let in_first_band = (band_row / band_size).is_multiple_of(2);

        // Don't apply banding to first/last row if those regions are active.
        let is_first = first_row_active && row_idx == 0;
        let is_last = last_row_active && row_idx == num_rows - 1;
        if !is_first && !is_last {
            if in_first_band {
                regions.push(TableStyleOverrideType::Band1Horz);
            } else {
                regions.push(TableStyleOverrideType::Band2Horz);
            }
        }
    }

    // §17.7.6 priority 3: vertical banding.
    if v_band_active {
        let band_col = if first_col_active && col_idx > 0 {
            col_idx - 1
        } else {
            col_idx
        };
        let band_size = col_band_size.max(1) as usize;
        let in_first_band = (band_col / band_size).is_multiple_of(2);

        let is_first = first_col_active && col_idx == 0;
        let is_last = last_col_active && col_idx == num_cols - 1;
        if !is_first && !is_last {
            if in_first_band {
                regions.push(TableStyleOverrideType::Band1Vert);
            } else {
                regions.push(TableStyleOverrideType::Band2Vert);
            }
        }
    }

    // §17.7.6 priority 4: first/last column.
    if first_col_active && col_idx == 0 {
        regions.push(TableStyleOverrideType::FirstCol);
    }
    if last_col_active && col_idx == num_cols - 1 {
        regions.push(TableStyleOverrideType::LastCol);
    }

    // §17.7.6 priority 5: first/last row.
    if first_row_active && row_idx == 0 {
        regions.push(TableStyleOverrideType::FirstRow);
    }
    if last_row_active && row_idx == num_rows - 1 {
        regions.push(TableStyleOverrideType::LastRow);
    }

    // §17.7.6 priority 6: corner cells (highest priority).
    if first_row_active && first_col_active && row_idx == 0 && col_idx == 0 {
        regions.push(TableStyleOverrideType::NwCell);
    }
    if first_row_active && last_col_active && row_idx == 0 && col_idx == num_cols - 1 {
        regions.push(TableStyleOverrideType::NeCell);
    }
    if last_row_active && first_col_active && row_idx == num_rows - 1 && col_idx == 0 {
        regions.push(TableStyleOverrideType::SwCell);
    }
    if last_row_active && last_col_active && row_idx == num_rows - 1 && col_idx == num_cols - 1 {
        regions.push(TableStyleOverrideType::SeCell);
    }

    regions
}

/// Overlay cell properties (higher priority replaces existing values).
fn overlay_cell_properties(result: &mut CellConditionalFormatting, tcp: &TableCellProperties) {
    let target = result
        .cell_properties
        .get_or_insert_with(TableCellProperties::default);

    // §17.7.6: each non-None field from the overlay replaces the target.
    if tcp.shading.is_some() {
        target.shading = tcp.shading;
    }
    if let Some(ref src) = tcp.borders {
        // §17.7.6: when a tblStylePr has tcBorders, it REPLACES all cell
        // borders for that region. Sides not mentioned are implicitly nil.
        target.borders = Some(*src);
    }
    if tcp.vertical_align.is_some() {
        target.vertical_align = tcp.vertical_align;
    }
}

/// Overlay run properties (higher priority replaces existing values).
fn overlay_run_properties(result: &mut CellConditionalFormatting, rp: &RunProperties) {
    let target = result
        .run_properties
        .get_or_insert_with(RunProperties::default);
    // Higher priority: overlay's values replace target's.
    // Use merge in reverse: merge target into a clone of overlay.
    let mut merged = rp.clone();
    crate::render::resolve::properties::merge_run_properties(&mut merged, target);
    *target = merged;
}

/// Overlay paragraph properties (higher priority replaces existing values).
fn overlay_paragraph_properties(result: &mut CellConditionalFormatting, pp: &ParagraphProperties) {
    let target = result
        .paragraph_properties
        .get_or_insert_with(ParagraphProperties::default);
    let mut merged = pp.clone();
    crate::render::resolve::properties::merge_paragraph_properties(&mut merged, target);
    *target = merged;
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::model::*;

    fn make_override(
        override_type: TableStyleOverrideType,
        shading: Option<Shading>,
        bold: Option<bool>,
    ) -> TableStyleOverride {
        TableStyleOverride {
            override_type,
            paragraph_properties: None,
            run_properties: bold.map(|b| RunProperties {
                bold: Some(b),
                ..Default::default()
            }),
            table_properties: None,
            table_row_properties: None,
            table_cell_properties: shading.map(|s| TableCellProperties {
                shading: Some(s),
                ..Default::default()
            }),
        }
    }

    fn green_shading() -> Shading {
        Shading {
            fill: Color::Rgb(0x9BBB59),
            pattern: ShadingPattern::Clear,
            color: Color::Auto,
        }
    }

    fn blue_shading() -> Shading {
        Shading {
            fill: Color::Rgb(0xD3DFEE),
            pattern: ShadingPattern::Clear,
            color: Color::Auto,
        }
    }

    // ── Region detection ─────────────────────────────────────────────

    #[test]
    fn first_row_detected() {
        let regions = applicable_regions(0, 2, 6, 6, None, 1, 1);
        assert!(regions.contains(&TableStyleOverrideType::FirstRow));
        assert!(!regions.contains(&TableStyleOverrideType::LastRow));
    }

    #[test]
    fn last_row_detected() {
        let regions = applicable_regions(5, 2, 6, 6, None, 1, 1);
        assert!(regions.contains(&TableStyleOverrideType::LastRow));
        assert!(!regions.contains(&TableStyleOverrideType::FirstRow));
    }

    #[test]
    fn first_col_detected() {
        let regions = applicable_regions(2, 0, 6, 6, None, 1, 1);
        assert!(regions.contains(&TableStyleOverrideType::FirstCol));
    }

    #[test]
    fn last_col_detected() {
        let regions = applicable_regions(2, 5, 6, 6, None, 1, 1);
        assert!(regions.contains(&TableStyleOverrideType::LastCol));
    }

    #[test]
    fn nw_corner_detected() {
        let regions = applicable_regions(0, 0, 6, 6, None, 1, 1);
        assert!(regions.contains(&TableStyleOverrideType::NwCell));
        assert!(regions.contains(&TableStyleOverrideType::FirstRow));
        assert!(regions.contains(&TableStyleOverrideType::FirstCol));
    }

    #[test]
    fn se_corner_detected() {
        let regions = applicable_regions(5, 5, 6, 6, None, 1, 1);
        assert!(regions.contains(&TableStyleOverrideType::SeCell));
    }

    #[test]
    fn interior_cell_gets_banding() {
        let regions = applicable_regions(1, 1, 6, 6, None, 1, 1);
        // Row 1 with firstRow active: band_row = 0 → band1
        assert!(regions.contains(&TableStyleOverrideType::Band1Horz));
    }

    #[test]
    fn banding_alternates() {
        // Row 2 with firstRow active: band_row = 1 → band2
        let regions = applicable_regions(2, 1, 6, 6, None, 1, 1);
        assert!(regions.contains(&TableStyleOverrideType::Band2Horz));
    }

    #[test]
    fn no_h_band_disables_banding() {
        let look = TableLook {
            first_row: Some(true),
            last_row: Some(true),
            first_column: Some(true),
            last_column: Some(true),
            no_h_band: Some(true),
            no_v_band: None,
        };
        let regions = applicable_regions(1, 1, 6, 6, Some(&look), 1, 1);
        assert!(!regions.contains(&TableStyleOverrideType::Band1Horz));
        assert!(!regions.contains(&TableStyleOverrideType::Band2Horz));
    }

    #[test]
    fn first_row_disabled_by_look() {
        let look = TableLook {
            first_row: Some(false),
            last_row: None,
            first_column: None,
            last_column: None,
            no_h_band: None,
            no_v_band: None,
        };
        let regions = applicable_regions(0, 2, 6, 6, Some(&look), 1, 1);
        assert!(!regions.contains(&TableStyleOverrideType::FirstRow));
    }

    #[test]
    fn band_size_2() {
        // Row 1 with firstRow: band_row=0, band_size=2 → 0/2=0 → band1
        let r1 = applicable_regions(1, 1, 10, 6, None, 2, 1);
        assert!(r1.contains(&TableStyleOverrideType::Band1Horz));

        // Row 2: band_row=1, 1/2=0 → band1
        let r2 = applicable_regions(2, 1, 10, 6, None, 2, 1);
        assert!(r2.contains(&TableStyleOverrideType::Band1Horz));

        // Row 3: band_row=2, 2/2=1 → band2
        let r3 = applicable_regions(3, 1, 10, 6, None, 2, 1);
        assert!(r3.contains(&TableStyleOverrideType::Band2Horz));
    }

    // ── Priority overlay ─────────────────────────────────────────────

    #[test]
    fn first_row_shading_applied() {
        let overrides = vec![make_override(
            TableStyleOverrideType::FirstRow,
            Some(green_shading()),
            Some(true),
        )];
        let result = resolve_cell_conditional(0, 2, 6, 6, None, &overrides, 1, 1);
        assert!(result.cell_properties.is_some());
        assert!(result.cell_properties.as_ref().unwrap().shading.is_some());
        assert!(result.run_properties.as_ref().unwrap().bold == Some(true));
    }

    #[test]
    fn banding_shading_for_interior() {
        let overrides = vec![make_override(
            TableStyleOverrideType::Band1Horz,
            Some(blue_shading()),
            None,
        )];
        let result = resolve_cell_conditional(1, 2, 6, 6, None, &overrides, 1, 1);
        let shading = result
            .cell_properties
            .as_ref()
            .unwrap()
            .shading
            .as_ref()
            .unwrap();
        assert_eq!(shading.fill, Color::Rgb(0xD3DFEE));
    }

    #[test]
    fn corner_overrides_first_row() {
        let overrides = vec![
            make_override(
                TableStyleOverrideType::FirstRow,
                Some(green_shading()),
                None,
            ),
            make_override(TableStyleOverrideType::NwCell, Some(blue_shading()), None),
        ];
        let result = resolve_cell_conditional(0, 0, 6, 6, None, &overrides, 1, 1);
        // NW corner has higher priority than FirstRow.
        let shading = result
            .cell_properties
            .as_ref()
            .unwrap()
            .shading
            .as_ref()
            .unwrap();
        assert_eq!(shading.fill, Color::Rgb(0xD3DFEE));
    }

    #[test]
    fn no_overrides_returns_empty() {
        let result = resolve_cell_conditional(2, 2, 6, 6, None, &[], 1, 1);
        assert!(result.cell_properties.is_none());
        assert!(result.run_properties.is_none());
    }
}
