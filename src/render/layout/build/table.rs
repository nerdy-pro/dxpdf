use crate::model::dimension::{Dimension, Twips};
use crate::model::{self, Block, Table, TableCell};
use crate::render::dimension::Pt;
use crate::render::geometry;
use crate::render::layout::paragraph::DropCapInfo;
use crate::render::layout::section::LayoutBlock;
use crate::render::layout::table::{
    compute_column_widths, CellBorderConfig, CellBorderOverride, TableCellInput, TableRowInput,
};
use crate::render::resolve::color::{resolve_color, ColorContext};
use crate::render::resolve::conditional::{
    resolve_cell_conditional, CellConditionalFormatting, CellGridPosition,
};
use crate::render::resolve::styles::ResolvedStyle;

use super::block::build_paragraph_block;
use crate::render::layout::fragment::split_oversized_fragments;

use super::convert::{
    convert_cell_border_override, convert_table_border_config, merge_table_borders,
};
use super::{BuildContext, BuildState};

/// Result of building a table from the model.
pub(super) struct BuiltTable {
    pub(super) rows: Vec<TableRowInput>,
    /// Grid slots, already shrunk by `cell_spacing`.
    pub(super) col_widths: Vec<Pt>,
    /// §17.4.45 `tblCellSpacing` in points; zero when unset.
    pub(super) cell_spacing: Pt,
    pub(super) border_config: Option<crate::render::layout::table::TableBorderConfig>,
    /// §17.4.50: table indentation from left margin.
    pub(super) indent: Pt,
    /// §17.4.28: table horizontal alignment (left/center/right).
    pub(super) alignment: Option<model::Alignment>,
    pub(super) float_info: Option<super::super::section::TableFloatInfo>,
}

/// §17.4.80: turn a parsed `trHeight` into a layout constraint.
///
/// `exact` pins the height, `atLeast` is a minimum, and `auto` **ignores `val`**
/// — the row sizes to its content and carries no constraint at all, hence the
/// `None`.
///
/// An omitted `hRule` never reaches here as `Auto`: the parse seam defaults it
/// to `AtLeast`, matching Word ([MS-OI29500] §17.4.80(a)) rather than the
/// standard's `auto`. That default is what makes returning `None` for `auto`
/// safe — otherwise every Word row with a `trHeight` would lose its minimum.
fn row_height_rule(
    h: model::TableRowHeight,
) -> Option<crate::render::layout::table::RowHeightRule> {
    use crate::model::HeightRule;
    use crate::render::layout::table::RowHeightRule;
    match h.rule {
        HeightRule::Exact => Some(RowHeightRule::Exact(Pt::from(h.value))),
        HeightRule::AtLeast => Some(RowHeightRule::AtLeast(Pt::from(h.value))),
        HeightRule::Auto => None,
    }
}

/// §17.4.45: resolve `tblCellSpacing` to the gap it actually renders, in points.
///
/// `CT_TblWidth` allows `pct` and `auto`, and the spec says both **are ignored**
/// for this element — only `dxa` carries a usable value. `nil` and an omitted
/// element are zero, which is every table in the test corpora.
///
/// # No factor: the declared value *is* the gap, everywhere
///
/// Word renders `test-files/issue-165-cellspacing.docx` with gaps about twice
/// this width (issue #165), which twice prompted the obvious conclusion —
/// double the value — and it is wrong. The spec states no factor: §17.4.45 and
/// [MS-OI29500] describe the value only as "the minimum amount of space which
/// shall be left between all cells in the table including the width of the table
/// borders in the calculation".
///
/// **ONLYOFFICE settles it**, being an independent implementation that both
/// renders and targets Word compatibility. `sdkjs`, in
/// `word/Editor/Table/TableRecalculate.js`, insets each cell within its grid
/// slot: `+= CellSpacing` on the first cell's outer side and on the last cell's,
/// `+= CellSpacing / 2` on every interior side. So an interior gap is two halves
/// and an edge gap is one whole — **every gap is exactly the declared value**,
/// which is what this function returns and always did. It also carves them out
/// of the grid rather than growing the table, confirming `reserve_cell_spacing`
/// below.
///
/// So where does Word's doubling come from? The probe is the suspect, not the
/// factor: `issue-165-cellspacing.docx` declares `tblCellSpacing` **twice**, 400
/// in `tblPr` and 400 again in `trPr`. If Word sums them the effective spacing is
/// 800 twips — exactly the doubling observed, with no factor anywhere. §17.4.45
/// says the row value *supersedes* the table one and ONLYOFFICE implements
/// override (`RowPr.TableCellSpacing = TablePr.TablePr.TableCellSpacing`, then a
/// merge), but the spec saying "supersede" is not evidence about what Word does
/// — [MS-OI29500] exists because Word deviates. That probe cannot tell the two
/// apart, which is why it must not be used to justify a factor.
///
/// **Word reference render needed** (issue #165):
/// `test-files/issue-165-cellspacing-scale.docx` declares the spacing at table
/// level *only* in two of its four tables, so it separates "Word applies a
/// factor" from "Word sums the two declarations" — and its fourth table, with
/// 400 at table level and 800 on the row, measures the sum directly.
///
/// # Why the row-level value is still not applied
///
/// Not because the precedence is unclear. The primary text gives it exactly:
/// §17.4.45 (`tblPr`) is "superseded by a table-level exception (§17.4.44) or
/// the row cell spacing value (§17.4.43) in that order", and §17.4.44
/// (`tblPrEx`) is "superseded by the row cell spacing value (§17.4.43)". Row
/// beats exception beats table, and §17.4.43 adds two sentences no other level
/// has: row-level spacing is added *inside* the text margins where table-level
/// is added *outside*, and "Row-level cell spacing shall not increase the width
/// of the overall table."
///
/// **The one Word render on record contradicts "supersede".**
/// `issue-165-cellspacing.docx` declares 400 at table level *and* 400 on the
/// row. Superseding gives 400; Word draws about 800. Summing gives 800.
/// A factor is ruled out by ONLYOFFICE above, so of the two readings left it is
/// the spec's own that the measurement disagrees with — which is what
/// [MS-OI29500] exists to record. Implementing §17.4.43 as written would move
/// that fixture from matching the only Word measurement available to missing it
/// by half, on the strength of text already known not to bind Word here.
///
/// **And a differing row is not a local change.** The gap is uniform because
/// one spacing is carved out of the grid by `reserve_cell_spacing` and the same
/// value insets every cell, so slots plus one spacing equal the table width. A
/// row using a different value breaks that, and repairing it needs three
/// answers nothing supplies:
///
/// * does the row re-carve the grid from its own spacing — misaligning its
///   column boundaries with the rows above and below — or inset within the
///   slots the table-level value produced, leaving the row short of the table
///   edge? §17.4.43's "shall not increase the width of the overall table" rules
///   out *growing* the table and does not choose between these.
/// * what is the table's reported width when rows disagree? It is
///   `sum(slots) + cell_spacing` today, and it drives pagination, floating-table
///   registration and `jc` placement — so a wrong answer moves the whole table,
///   not one row.
/// * is border collapsing still table-wide? `collapse_borders` is
///   `cell_spacing <= 0` for the entire table. With spacing on some rows only,
///   a per-row answer would collapse one row's borders and not its neighbour's,
///   and no section says whether Word does that.
///
/// **What would settle it**: a Word render of
/// `issue-165-cellspacing-scale.docx`. Its two table-level-only tables give the
/// factor, and its fourth table — 400 on the table, 800 on every row —
/// distinguishes supersede (gap 800), sum (gap 1200) and table-wins (gap 400)
/// in a single measurement, because all three predict a different number. A
/// fifth table with the row value on *one* row only would answer the alignment
/// and border-collapse questions above; the fixture does not have one yet.
fn resolve_cell_spacing(m: Option<model::TableMeasure>) -> Pt {
    match m {
        Some(model::TableMeasure::Twips(tw)) => Pt::from(tw).max(Pt::ZERO),
        // Auto / Pct / Nil / absent.
        _ => Pt::ZERO,
    }
}

/// Carve one `cell_spacing` out of the grid so the slots plus one spacing add
/// up to the table's own width, scaling the columns proportionally.
///
/// Clamped: a spacing at least as large as the table would leave nothing to
/// scale, so the columns collapse to zero rather than going negative. Word
/// caps the usable spacing well below that, but the file format does not.
fn reserve_cell_spacing(col_widths: Vec<Pt>, cell_spacing: Pt) -> Vec<Pt> {
    if cell_spacing <= Pt::ZERO || col_widths.is_empty() {
        return col_widths;
    }
    let total: Pt = col_widths.iter().copied().sum();
    if total <= cell_spacing {
        return vec![Pt::ZERO; col_widths.len()];
    }
    let scale = (total - cell_spacing).raw() / total.raw();
    col_widths.into_iter().map(|w| w * scale).collect()
}

/// §17.4.63: an auto-width table keeps its declared `<w:tblGrid>` verbatim,
/// unless that grid is wide enough to carry the table off the **paper**, in
/// which case it is scaled down until it is not.
///
/// # This is a containment guard, not an autofit implementation
///
/// A `<w:tblW w:type="auto"/>` table states no preferred width, so there is
/// nothing for §17.4.63's scaling to scale *to*. Word would run its autofit
/// algorithm here; ECMA-376 defines no such algorithm, this engine does not
/// implement one (see `table::grid::compute_column_widths`), and inventing one
/// under the name of a bug fix is how a renderer acquires behaviour nobody
/// chose. So the declared grid is still used as declared. The only thing added
/// is a floor under it.
///
/// # Why the paper edge, and not the text column
///
/// Clamping to `available_width` is the obvious guard and the corpus refutes
/// it. Three of the 52 documents ship an auto table whose grid exceeds the text
/// column — `sample-docx-files-sample3.docx` by 101 twips, `ELH_2025-12-18` and
/// `KAB_2026-03-25` by ~172 twips across 39 tables — and every one of those 39
/// also declares `<w:tblLayout w:type="fixed"/>`, §17.4.52's instruction to use
/// the declared column widths rather than compute any. Those tables reach a few
/// points into the right margin, on the paper and fully legible. Normalising
/// them to the text column would restyle 40 tables of real Word and LibreOffice
/// output on no authority at all.
///
/// What survives that is narrower and needs no algorithm: content drawn past
/// the sheet is *gone* from the PDF, under every reading of every layout
/// property. So the limit is the width from the left margin to the right paper
/// edge — a table placed at the left margin, which is where a top-level table
/// with no `tblInd` sits, then ends exactly at the sheet's edge in the worst
/// case. Two known imprecisions, both deliberate, since a guard that has to be
/// exact is a layout algorithm again:
///
/// * a table displaced rightward (`w:tblInd`, `w:jc`, `w:tblpX`) can still
///   reach past the edge by that displacement;
/// * a *nested* table is measured against the page rather than against its
///   cell, so it may still overflow the cell — but the enclosing table's own
///   width already keeps it on the sheet.
///
/// **Word reference render needed**: what width Word actually gives an
/// overflowing `w:type="auto"` table, at `tblLayout` both `fixed` and
/// `autofit`. That measurement is what would replace this guard with a rule.
fn clamp_auto_grid_to_page(
    grid_cols: &[Pt],
    num_cols: usize,
    available_width: Pt,
    page: &crate::render::layout::page::PageConfig,
) -> Vec<Pt> {
    let declared: Pt = grid_cols.iter().copied().sum();
    // `max(available_width)` so a pathological page — margins wider than the
    // sheet, which `PageConfig::content_width` also has to defend against —
    // cannot make the limit narrower than the space the caller just offered.
    let limit = (page.page_size.width - page.margins.left).max(available_width);
    if declared <= limit {
        return grid_cols.to_vec();
    }
    compute_column_widths(grid_cols, num_cols, limit)
}

/// §17.4.48 / §17.4.71: give every declared cell a grid column, appending
/// columns to `<w:tblGrid>` when it is too short to seat one.
///
/// # What the spec settles, and what it does not
///
/// §17.4.63 (`tblW`) and §17.4.71 (`tcW`) carry the same paragraph verbatim:
/// *"All widths in a table are considered preferred because: The table **shall**
/// satisfy the shared columns as specified by the `tblGrid` element … Two or
/// more widths can have conflicting values for the width of the same grid column
/// … The table layout algorithm can require a preference to be overridden."*
///
/// So the grid is the invariant and both width elements are preferences that may
/// lose to it — and the spec states the conflict without resolving it. That is
/// why nothing here lets a `tcW` override the grid column it disagrees with:
/// which one Word believes is open, and settling it needs a **Word reference
/// render**, not a rule invented at this line. `w:tcW` is read here and nowhere
/// else in `src/render/`.
///
/// What is *not* open is the seating. A grid with fewer columns than a row has
/// cells cannot satisfy the shared columns for that row under any reading of the
/// sentence above, because the last cell has no column to be satisfied in. The
/// file contradicts itself and something has to give.
///
/// # Why the cell may not be what gives
///
/// It used to be. Both this module and `table/measure.rs` clamp the grid walk to
/// the declared column count so a short grid cannot panic on a slice index —
/// and the clamp gives the overrun cell an empty slice, so it laid out at zero
/// width, stacked at the table's right edge under the previous cell's border.
/// Its text, borders and shading were all emitted and none of them were legible.
///
/// Content drawn at zero width is gone from the PDF exactly as content drawn off
/// the paper is, which is the line `clamp_auto_grid_to_page` above is already
/// drawn at, and for the same reason: it is the one consequence every reading of
/// every layout property agrees is wrong. So the grid grows to fit the cells
/// rather than the cells shrinking to fit the grid.
///
/// # The repair is minimal, and that is a choice
///
/// The declared columns are kept exactly as declared and the missing ones are
/// appended. Every row the grid *could* seat therefore lays out unchanged, which
/// is the property that keeps this off all 398 tables in the corpus.
///
/// LibreOffice resolves the same input differently — `DomainMapperTableManager`
/// discards the grid entirely once the cell count disagrees with it and rebuilds
/// the separators from the declared cell widths — and that is the reading with
/// 13 years of third-party-producer pressure behind it. It is also the larger
/// change: it moves rows that were already seated. Neither is derivable from
/// §17.4.48, and a **Word reference render** of a table whose grid is one column
/// short is what would choose between them.
///
/// # Sizing an appended column
///
/// From the unseated cell's own `w:tcW`, which is the only width the file offers
/// for a column the grid never mentioned. Where a cell straddles the end of the
/// declared grid, the columns it already has are subtracted first and the
/// remainder is split evenly across the new ones. Rows may disagree about the
/// same appended column, and the widest claim wins: a narrower one would put
/// some row back into less width than it declared, and taking the first or last
/// row's claim instead would make the result depend on row order, which nothing
/// in §17.4.48 supports.
///
/// Only `w:type="dxa"` counts. §17.4.71's `pct` is a fraction *of the table's
/// own width*, and the table's width is computed from the grid this function
/// returns — so reading it here would be circular. Such a cell falls to the
/// same default as one that declares nothing: the mean of the declared columns,
/// or an arbitrary equal unit when there are none, which reproduces the
/// equal-distribution fallback in `compute_column_widths` exactly.
///
/// # Only cells count toward the demand
///
/// §17.4.15 `gridBefore` is counted because it displaces real cells rightward.
/// §17.4.14 `gridAfter` is **not**: it declares trailing columns that hold no
/// cell, so a grid too short for them loses no content, and appending for it
/// would narrow a cell that renders perfectly well today on nothing but
/// speculation. A row with no cells at all asks for nothing for the same
/// reason.
///
/// # The demand is bounded by the cells, not by the attributes
///
/// `w:gridBefore` and `w:gridSpan` are both `ST_DecimalNumber`, so a file may
/// state four billion of either, and this is the one function that would turn
/// such a number into an allocation. The count of appended columns is therefore
/// capped at the table's own cell count — enough for every cell to get a column,
/// which is the whole invariant, and bounded by the size of the parsed document
/// rather than by a value inside it. A pathological `gridSpan` degrades to a
/// cell narrower than it asked for; it does not degrade to a cell that is not
/// drawn, and it cannot size a 32 GB `Vec`.
fn seat_every_cell(
    declared: &[Dimension<Twips>],
    rows: &[model::TableRow],
) -> Vec<Dimension<Twips>> {
    // §17.4.17: a `gridSpan` of 0 is well-formed and meaningless, so a cell
    // that exists covers at least one column — `max(1)`, as every grid walk in
    // this engine spells it.
    let span_of =
        |c: &model::TableCell| (*c.properties.grid_span.get().unwrap_or(&1)).max(1) as usize;
    let demand = |r: &model::TableRow| {
        (r.properties.grid_before as usize)
            .saturating_add(r.cells.iter().map(span_of).sum::<usize>())
    };

    let cells: usize = rows.iter().map(|r| r.cells.len()).sum();
    let needed = rows
        .iter()
        .filter(|r| !r.cells.is_empty())
        .map(demand)
        .max()
        .unwrap_or(0)
        .min(declared.len().saturating_add(cells));
    if needed <= declared.len() {
        // Every cell has a column: the grid is the invariant and it holds.
        // This is every table in `test-files/` + `test-cases/`.
        return declared.to_vec();
    }

    // What each appended column is claimed to be, by the widest claimant.
    let mut claimed = vec![Dimension::<Twips>::ZERO; needed - declared.len()];
    for row in rows {
        let mut col = row.properties.grid_before as usize;
        for cell in &row.cells {
            let span = span_of(cell);
            let end = col.saturating_add(span);
            // `needed` is capped, so on malformed input a cell can reach past
            // even the repaired grid. Only the claim is clamped, never the walk
            // — `col` keeps advancing by the declared span so the cells that do
            // fit stay in the columns they asked for.
            let claim_to = end.min(needed);
            let from = col.max(declared.len()).min(claim_to);
            // The cell's declared width covers its whole span; the part of that
            // span already inside the declared grid is subtracted before the
            // remainder is shared out.
            if from < claim_to {
                if let Some(Some(tcw)) = cell.properties.width.get().map(dxa_twips) {
                    let inside: i64 = declared[col.min(declared.len())..declared.len().min(end)]
                        .iter()
                        .map(|w| w.raw())
                        .sum();
                    let per = (tcw.raw() - inside).max(0) / (claim_to - from) as i64;
                    if per > 0 {
                        for slot in &mut claimed[from - declared.len()..claim_to - declared.len()] {
                            *slot = (*slot).max(Dimension::new(per));
                        }
                    }
                }
            }
            col = end;
        }
    }

    // A column nobody sized takes the mean of the declared ones. With no
    // declared columns either, any positive constant does: the grid is scaled
    // to the table's width downstream, so all-equal columns land on exactly the
    // equal distribution `compute_column_widths` would have produced.
    let default = if declared.is_empty() {
        Dimension::<Twips>::new(1)
    } else {
        Dimension::new(declared.iter().map(|w| w.raw()).sum::<i64>() / declared.len() as i64)
    };
    declared
        .iter()
        .copied()
        .chain(
            claimed
                .into_iter()
                .map(|w| if w > Dimension::ZERO { w } else { default }),
        )
        .collect()
}

/// A `CT_TblWidth` in twips, but only when it states one.
///
/// `pct` is deliberately excluded — see `seat_every_cell` for why its base is
/// circular here. `auto` states no width at all (§17.4.71 calls the value
/// ignored), and `nil` states zero, which is not a column width.
fn dxa_twips(m: &model::TableMeasure) -> Option<Dimension<Twips>> {
    match m {
        model::TableMeasure::Twips(tw) if *tw > Dimension::ZERO => Some(*tw),
        _ => None,
    }
}

/// Recursively build a table: resolve styles, conditional formatting, and
/// recurse into each cell's content blocks.
pub(super) fn build_table(
    t: &Table,
    available_width: Pt,
    ctx: &BuildContext,
    state: &mut BuildState,
) -> BuiltTable {
    // §17.4.48: grid column widths, after `seat_every_cell` has made sure the
    // grid has a column for every cell — for all 398 corpus tables it already
    // does, and the declared list comes back untouched.
    let declared_grid: Vec<Dimension<Twips>> = t.grid.iter().map(|g| g.width).collect();
    let grid = seat_every_cell(&declared_grid, &t.rows);
    let num_cols = if grid.is_empty() {
        // This arm cannot produce a non-zero answer, which is worth stating
        // because it reads like the handler for a missing `<w:tblGrid>` and is
        // not. `seat_every_cell` returns an empty grid only when **no row has
        // a cell**: one cell demands one column, which is already more than an
        // empty grid declares, so a table with any cell at all comes back with
        // appended columns instead. Whenever this line runs, every row is
        // empty and the maximum below is 0 too.
        //
        // Hence it does not consult `w:gridSpan` — adding that would change
        // nothing — and it is kept rather than deleted so that a later change
        // to the seating rule finds an answer here instead of a hole. A
        // mutation to any other expression survives the whole suite; that is
        // this branch being dead, not a gap in the tests.
        t.rows.iter().map(|r| r.cells.len()).max().unwrap_or(0)
    } else {
        grid.len()
    };
    let grid_cols: Vec<Pt> = grid.iter().copied().map(Pt::from).collect();

    // §17.7.6: table style for conditional formatting, borders, cell margins.
    //
    // §17.7.4.17: a table that names no style takes the stylesheet's
    // `w:default="1"` table style. That is Word's `TableNormal`, whose
    // 108-twip left/right `tblCellMar` is why cell text in a Word table sits
    // in from the cell edge rather than against it — without this, such a
    // table drew its text 5.4 pt to the left of where Word puts it.
    //
    // Only when none is named. §17.7.4.17 scopes the default to objects that
    // do not reference a style, and a table style that wants `TableNormal`
    // underneath it says `basedOn` — which is exactly how Word's own built-in
    // table styles are written, so the two readings agree on Word's output.
    let raw_table_style = t
        .properties
        .style_id
        .as_ref()
        .or(ctx.resolved.default_table_style_id.as_ref())
        .and_then(|sid| ctx.resolved.styles.get(sid));

    // §17.7.2: the table style's own `<w:tblPr>`, already folded through
    // `basedOn` and `wholeTable` by style resolution. The properties below
    // cascade direct-then-style through it — but not all of `CT_TblPrBase`
    // does, and which do is not a property of the content model.
    //
    // # Six of the twelve do not cascade
    //
    // [MS-OI29500] §2.1.250(a), annotating §17.7.6.4 (a style's own `tblPr`):
    // the standard allows `bidiVisual`, `tblLayout`, `tblLook`, `tblOverlap`,
    // `tblpPr`, `tblStyle` and `tblW` as its children, and "Word does not allow
    // these elements to be child elements of the tblPr element". §2.1.249(a)
    // says the same of §17.7.6.3, the conditional `tblPr` that `wholeTable`
    // folds in here, so both sources feeding `style_table` are covered. A
    // document declaring one of them in a style either fails to open in Word or
    // renders without it; reading it would render a third-party document
    // differently from the Word its author checked it against.
    //
    // Four of those seven are modelled and read here — `tblW`, `tblLook`,
    // `tblpPr`, `tblOverlap` — so each is taken from the `<w:tbl>` alone below.
    // `tblLayout` has no read site at all. The remaining two need no code:
    // `TableProperties` has no `bidiVisual` field, and `tblStyle` is already
    // excluded from `merge_table_properties` for an independent reason (it
    // would be a second inheritance edge competing with `basedOn`).
    //
    // The five that do cascade — `jc`, `tblInd`, `tblBorders`, `tblCellMar`,
    // `tblCellSpacing` — appear on neither list. So do the two band sizes,
    // and their case runs the *other* way: §2.1.164(a) says Word does not
    // accept `tblStyleColBandSize`/`tblStyleRowBandSize` on a `<w:tbl>`'s own
    // `tblPr` (§17.4.59), which makes the style the only level it reads them
    // at. ECMA documents them under Table Styles to match (§17.7.6.5 and
    // §17.7.6.7, not §17.4), and the corpus agrees without exception: 693
    // band-size declarations across 8 documents, every one in a style's
    // `tblPr`, none on a table. The direct read below is kept anyway — it costs
    // nothing and honours a producer Word would not.
    //
    // `cell_margins` and `borders` merge *per edge*, since §17.4.42 and
    // §17.4.38 define those as edge-wise exceptions; every other property is a
    // single value, so the direct occurrence simply wins outright.
    //
    // Note the narrowing is of the **application** only. `merge_table_
    // properties` still folds all twelve fields through `basedOn`, because
    // style-to-style inheritance is a different question: §17.7.4.3 gives a
    // child "all of the properties of the base style" unconditionally, and
    // nothing in either erratum speaks about what one style inherits from
    // another. A field that both levels carry and no one applies is inert, and
    // making the merge selective would put a rule about Word's *reader* inside
    // a function that models the *stylesheet*.
    let style_table = raw_table_style.and_then(|s| s.table.as_ref());

    // §17.4.42: default cell margins from table style cascade.
    let style_cell_margins = style_table.and_then(|tp| tp.cell_margins.cloned());
    // Per-edge merge: direct tblCellMar overrides style per-edge, with
    // unspecified edges (value 0) falling back to the style's value.
    // Word merges per-edge rather than replacing the entire set.
    let default_cell_margins = match (t.properties.cell_margins.cloned(), style_cell_margins) {
        (Some(direct), Some(style)) => {
            use crate::model::geometry::EdgeInsets;
            Some(EdgeInsets {
                top: if direct.top.raw() != 0 {
                    direct.top
                } else {
                    style.top
                },
                bottom: if direct.bottom.raw() != 0 {
                    direct.bottom
                } else {
                    style.bottom
                },
                left: if direct.left.raw() != 0 {
                    direct.left
                } else {
                    style.left
                },
                right: if direct.right.raw() != 0 {
                    direct.right
                } else {
                    style.right
                },
            })
        }
        (Some(direct), None) => Some(direct),
        (None, Some(style)) => Some(style),
        (None, None) => None,
    };

    // Each property read once here, so the several sites below cannot disagree
    // about which level won. Which of them consult `style_table` is decided by
    // the errata cited above it, not by the content model.
    //
    // §17.4.28 `jc`, §17.4.50 `tblInd`, §17.7.6.5/§17.7.6.7 the band sizes:
    // direct, then style.
    let alignment = t
        .properties
        .alignment
        .get()
        .or_else(|| style_table.and_then(|tp| tp.alignment.get()))
        .copied();
    let table_indent = t
        .properties
        .indent
        .get()
        .or_else(|| style_table.and_then(|tp| tp.indent.get()));
    let row_band_size = t
        .properties
        .style_row_band_size
        .get()
        .or_else(|| style_table.and_then(|tp| tp.style_row_band_size.get()))
        .copied()
        .unwrap_or(1);
    let col_band_size = t
        .properties
        .style_col_band_size
        .get()
        .or_else(|| style_table.and_then(|tp| tp.style_col_band_size.get()))
        .copied()
        .unwrap_or(1);

    // §17.4.63 `tblW`, §17.4.55 `tblLook`, §17.4.57 `tblpPr`, §17.4.56
    // `tblOverlap`: the `<w:tbl>`'s own `<w:tblPr>` and nothing else, per
    // §2.1.250(a)/§2.1.249(a). Written as plain reads rather than as a cascade
    // with the style arm deleted, so no later edit can restore one by reflex.
    let width = t.properties.width.get();
    let tbl_look = t.properties.look.get();
    let positioning = t.properties.positioning.get();
    let overlap = t.properties.overlap.cloned();

    // §17.4.63: resolve table width from tblW.
    let is_auto_width = matches!(
        width,
        None | Some(model::TableMeasure::Auto) | Some(model::TableMeasure::Nil)
    );
    let cell_margins_h = default_cell_margins
        .map(|m| Pt::from(m.left) + Pt::from(m.right))
        .unwrap_or(Pt::ZERO);
    // §17.4.63 / Word heuristic: a full-width left-aligned table extends
    // beyond the body content area by its cell margins so cell content
    // aligns with surrounding paragraph text. Centered/right-aligned tables
    // are positioned as a unit — extending them would just shift them out
    // by half the margins, producing a width discrepancy when consecutive
    // centered tables have different cell margins (cf. the stacked
    // "Anhang: Sauberkeit" tables in the Volvo Annahme-Protokoll).
    let extends_for_alignment = !matches!(
        alignment,
        Some(model::Alignment::Center) | Some(model::Alignment::End)
    );
    let target_width = match width {
        Some(model::TableMeasure::Pct(pct)) => {
            // §17.4.63 / §17.18.90: the `pct` scale is fiftieths of a percent,
            // which `Dimension<FiftiethPercent>` carries and divides by.
            //
            // # The base is the caller's width, not the page's
            //
            // §17.4.63 says a `pct` table width is relative to "the text
            // extents of the page", i.e. the page width less its margins.
            // `available_width` is that only for a *top-level* table; for a
            // nested one it is the enclosing cell's inner width, so the two
            // bases diverge exactly when a table is nested.
            //
            // Left as the cell's width deliberately, and it is a choice rather
            // than an oversight. Read literally, §17.4.63 makes a nested table
            // declaring 100% as wide as the page's text column however narrow
            // the cell holding it — which would draw it straight out of its own
            // cell, and no implementation does that. §17.4.71 gives `tcW`'s
            // `pct` a different base again ("the overall width of the table"),
            // so the two sibling elements do not even agree on what a
            // percentage is measured against.
            //
            // The corpus reaches this once: of 398 tables, exactly 1 is nested
            // *and* `pct` — `sample-docx-files-sample1.docx` at `w="4000"`
            // (80%), in a cell narrower than the page. Changing the base would
            // move that table and nothing else, and there is no reference to
            // move it toward. **Word reference render needed**: a nested table
            // at `w:tblW w:type="pct" w:w="5000"` inside a half-width cell
            // separates the two readings in one measurement.
            let ratio = pct.to_fraction();
            let base = if *pct >= Dimension::FULL && extends_for_alignment {
                available_width + cell_margins_h
            } else {
                available_width
            };
            base * ratio
        }
        Some(model::TableMeasure::Twips(tw)) => Pt::from(*tw),
        _ => available_width, // auto/nil: use grid cols or available width
    };
    // §17.4.52: tblLayout controls whether columns may auto-resize to fit
    // content; it does not override the preferred table width from tblW.
    // Word scales grid column widths proportionally to match tblW in both
    // fixed and auto layouts. Only when tblW is auto/nil do we keep the raw
    // grid widths (no preferred width was specified).
    //
    // `tblLayout` has no read site here, and when auto-fit lands it must **not**
    // acquire a style level — §2.1.250(a) lists it, so a style's `tblLayout` is
    // one of the elements Word does not read.
    //
    // Producers do write it there, so this is a decision and not a shortage of
    // documents: `sample-docx-files-sample-5.docx` and `-6.docx` each ship a
    // `TableNormal` whose own `tblPr` declares `<w:tblLayout w:type="fixed"/>`,
    // and each marks it `w:default="1"` — the style §17.7.4.17 hands to a table
    // naming none, which is the path `raw_table_style` above already resolves.
    // Neither document happens to contain a table, so nothing in this corpus
    // reaches the value; the point is that Word, reading the same two files,
    // would ignore the element either way. The 226 `<w:tblLayout>` elements the
    // corpus does put on a `<w:tbl>` are the direct level, which §2.1.250(a)
    // says nothing about.
    let col_widths = if is_auto_width && !grid_cols.is_empty() {
        clamp_auto_grid_to_page(&grid_cols, num_cols, available_width, &state.page_config)
    } else {
        compute_column_widths(&grid_cols, num_cols, target_width)
    };
    // §17.4.45: cell spacing is carved out of the table's own width rather than
    // added to it — the spec calls it "the minimum amount of space which shall
    // be left between all cells", not extra width. Reserving one spacing here
    // and offsetting each cell by one in `measure_table_rows` yields exactly
    // `cell_spacing` between adjacent cells *and* at both table edges.
    let cell_spacing = resolve_cell_spacing(
        t.properties
            .cell_spacing
            .get()
            .or_else(|| style_table.and_then(|tp| tp.cell_spacing.get()))
            .copied(),
    );
    let col_widths = reserve_cell_spacing(col_widths, cell_spacing);
    let style_overrides = raw_table_style
        .map(|s| s.table_style_overrides.as_slice())
        .unwrap_or(&[]);
    let num_rows = t.rows.len();

    // §17.4.38: resolve table borders — merge direct properties over table style.
    // Direct tblBorders may specify only a subset of edges (e.g. insideH=none);
    // unspecified edges inherit from the table style. Computed up front so
    // per-row tblPrEx merges (§17.4.60) below have a stable basis.
    let style_borders = style_table.and_then(|tp| tp.borders.get());
    let tbl_borders = match (t.properties.borders.get(), style_borders) {
        (Some(direct), Some(style)) => Some(merge_table_borders(direct, style)),
        (Some(direct), None) => Some(*direct),
        (None, Some(style)) => Some(*style),
        (None, None) => None,
    };
    let border_config = tbl_borders
        .as_ref()
        .map(|b| convert_table_border_config(b, state));

    // Build rows by iterating cells and recursing into their content.
    let rows: Vec<TableRowInput> = t
        .rows
        .iter()
        .enumerate()
        .map(|(row_idx, row)| {
            let cells: Vec<TableCellInput> = row
                .cells
                .iter()
                .enumerate()
                .map(|(col_idx, cell)| {
                    // §17.4.15: gridBefore offsets the row's first cell to the
                    // right by that many grid columns; subsequent spans accumulate.
                    let span = grid_span(cell) as usize;
                    let mut grid_start = row.properties.grid_before as usize;
                    for ci in 0..col_idx {
                        grid_start += grid_span(&row.cells[ci]) as usize;
                    }

                    // §17.7.6: the cell's *grid* position, which is what decides
                    // its column regions — see `applicable_regions` for Word's
                    // own answer. `col_idx` and this row's cell count would be
                    // the same thing only while no row spans or skips a column,
                    // and a table with `gridSpan`/`gridBefore` is exactly the
                    // table a conditional style is worth getting right on.
                    let cond = resolve_cell_conditional(
                        &CellGridPosition {
                            row_idx,
                            grid_col: grid_start,
                            grid_span: span,
                            num_rows,
                            num_cols,
                            row_band_size,
                            col_band_size,
                        },
                        tbl_look,
                        style_overrides,
                    );

                    // Compute available width for nested content.
                    // A row may address more grid columns than `tblGrid` declares —
                    // a `gridBefore` past the end, or simply more `<w:tc>` than
                    // `<w:gridCol>`. Both occur in real producer output and Word
                    // recovers from them. Clamp *both* ends: clamping only the end
                    // inverts the range and panics on the slice.
                    let grid_start = grid_start.min(col_widths.len());
                    let grid_end = (grid_start + span).min(col_widths.len());
                    let cell_width: Pt = col_widths[grid_start..grid_end].iter().copied().sum();
                    // Per-side cascade against the table default (see
                    // `build_table_cell` for the spec rationale): the horizontal
                    // padding contribution is the resolved left+right insets.
                    let table_default =
                        default_cell_margins.unwrap_or(crate::model::geometry::EdgeInsets::ZERO);
                    let resolved_h = match cell.properties.margins.get() {
                        Some(partial) => partial.resolve_against(table_default),
                        None => table_default,
                    };
                    let cell_margins_h = Pt::from(resolved_h.left) + Pt::from(resolved_h.right);
                    let inner_width = (cell_width - cell_margins_h).max(Pt::ZERO);

                    build_table_cell(
                        cell,
                        raw_table_style,
                        default_cell_margins,
                        &cond,
                        inner_width,
                        ctx,
                        state,
                    )
                })
                .collect();

            // Word/LibreOffice row-uniform content-area quirk — see
            // `normalize_row_uniform_vertical_insets` for the spec gap and
            // empirical evidence motivating this pass.
            let mut cells = cells;
            normalize_row_uniform_vertical_insets(&mut cells);

            // A row may override the table's `tblCellSpacing`, and the spec is
            // unambiguous about the order. Verified against the primary text:
            // §17.4.45 (`tblPr`) "shall be superseded by a table-level
            // exception (§17.4.44) or the row cell spacing value (§17.4.43) in
            // that order", and §17.4.44 (`tblPrEx`) "shall be superseded by
            // the row cell spacing value (§17.4.43)". So trPr > tblPrEx > tblPr.
            //
            // It is still not applied, and the reason is not that the order is
            // unclear. See `resolve_cell_spacing` for why: implementing it
            // requires three answers the spec does not give and no render has
            // measured, and the one Word render on record points *away* from
            // "supersede". Report it rather than dropping it silently; the
            // parsed value stays on the model, so nothing has to be reparsed
            // when a render settles it.
            //
            // Note this fires only on *disagreement*. A row restating the
            // table's own value resolves to the same gap either way, which is
            // why `issue-165-cellspacing.docx` is silent here and
            // `issue-165-cellspacing-scale.docx` — table 400, row 800 — is not.
            let row_spacing = row.properties.cell_spacing.cloned().or_else(|| {
                row.property_exceptions
                    .as_ref()
                    .and_then(|e| e.cell_spacing)
            });
            if let Some(rs) = row_spacing {
                if resolve_cell_spacing(Some(rs)) != cell_spacing && !state.warned_row_cell_spacing
                {
                    state.warned_row_cell_spacing = true;
                    log::warn!(
                        "§17.4.43/§17.4.44: row-level tblCellSpacing overrides are not applied; \
                         using the table-level value ({cell_spacing:?})"
                    );
                }
            }

            TableRowInput {
                cells,
                // §17.4.80: `exact` pins the height, `atLeast` is a minimum,
                // and `auto` **ignores `val`** — the row sizes to its content,
                // so it carries no constraint at all. An omitted `hRule`
                // arrives as `AtLeast` (Word's default, set at the parse seam),
                // which is why folding `auto` in with it used to be invisible:
                // [MS-OE376] §2.4.77(c) notes Word requires `val = 0` whenever
                // `hRule="auto"`, making `AtLeast(0)` a no-op on Word output.
                // Other producers are not bound by that.
                height_rule: row.properties.height.cloned().and_then(row_height_rule),
                is_header: row.properties.is_header,
                cant_split: row.properties.cant_split,
                grid_before: row.properties.grid_before,
                // §17.4.60: row-level tblPrEx.tblBorders — per-side
                // override of the table's effective borders. We merge
                // *at the model layer* (Option<Border> with style=None
                // is preserved), then convert to layout — that keeps
                // the spec's "specified as none" vs "not specified"
                // distinction that converting first would erase.
                border_overrides: row
                    .property_exceptions
                    .as_ref()
                    .and_then(|ex| ex.borders.as_ref())
                    .map(|over| {
                        let merged = match tbl_borders.as_ref() {
                            Some(table) => merge_table_borders(over, table),
                            None => *over,
                        };
                        convert_table_border_config(&merged, state)
                    }),
            }
        })
        .collect();

    // §17.4.84: an unpaired `w:vMerge w:val="continue"` is malformed input, and
    // is repaired here — once, over the finished rows — so that measure, emit,
    // border resolution and the paginator cannot disagree about whether a given
    // cell is merged.
    let mut rows = rows;
    let orphans = promote_orphan_vmerge_continues(&mut rows);
    if orphans > 0 && !state.warned_orphan_vmerge {
        state.warned_orphan_vmerge = true;
        log::warn!(
            "§17.4.84: {orphans} <w:vMerge w:val=\"continue\"/> cell(s) have no \
             <w:vMerge w:val=\"restart\"/> above them in the same grid column; \
             rendering each as an ordinary cell"
        );
    }

    // §17.4.57: floating table positioning.
    //
    // Horizontal positioning is partial: only `tblpXSpec` (`x_align`, below)
    // and `rightFromText` (`right_gap`, which only widens the wrap-reservation
    // width, not the table's own position) currently reach layout. `tblpX`
    // (`pos.x`), `horzAnchor` and `leftFromText` are parsed into the model but
    // not carried into `TableFloatInfo` — a horizontally-offset or
    // margin/page-anchored-horizontally table places as if none of those were
    // set.
    let float_info = positioning.map(|pos| {
        super::super::section::TableFloatInfo {
            right_gap: pos.right_from_text.map(Pt::from).unwrap_or(Pt::ZERO),
            bottom_gap: pos.bottom_from_text.map(Pt::from).unwrap_or(Pt::ZERO),
            x_align: pos.x_align,
            // §17.4.57: tblpY — absolute Y offset from the vertical anchor.
            y_offset: pos.y.map(Pt::from).unwrap_or(Pt::ZERO),
            // §17.4.57: default vertical anchor is "text".
            vert_anchor: pos.vert_anchor.unwrap_or(crate::model::TableAnchor::Text),
            // §17.4.56: tblOverlap controls collision behavior with
            // other floats on the same page.
            overlap,
        }
    });

    // §17.4.50: table indentation from left margin.
    // For full-width left-aligned tables, MS Word shifts the table left
    // by the default cell margin so cell content aligns with paragraph text.
    let is_full_width = matches!(
        width,
        Some(model::TableMeasure::Pct(pct)) if *pct >= Dimension::FULL
    );
    let is_left_aligned = !matches!(
        alignment,
        Some(model::Alignment::Center) | Some(model::Alignment::End)
    );
    // The shift is keyed on the *resolved* indent being zero, not on the
    // absence of a `tblInd` element. Those were the same question only while
    // the style level was unread: every built-in Word table style declares
    // `<w:tblInd w:w="0" w:type="dxa"/>`, so reading the style would otherwise
    // suppress the shift on essentially every Word document — silently undoing
    // the corpus tuning the shift exists for, without anything having decided
    // to. An explicitly-zero indent and an absent one are the same indent, and
    // Word cannot tell them apart either.
    //
    // A *non-zero* indent is still taken literally. Whether Word measures
    // `tblInd` to the table's border box or to the first cell's text edge — the
    // reading under which the shift is not a special case at all but the same
    // rule at zero — is not settled here; a **Word reference render** of a
    // full-width left-aligned table at a non-zero `tblInd` would settle it.
    let indent = match table_indent {
        Some(model::TableMeasure::Twips(tw)) if tw.raw() != 0 => Pt::from(*tw),
        _ if is_full_width && is_left_aligned => -default_cell_margins
            .map(|m| Pt::from(m.left))
            .unwrap_or(Pt::ZERO),
        _ => Pt::ZERO,
    };

    BuiltTable {
        rows,
        col_widths,
        cell_spacing,
        border_config,
        indent,
        alignment,
        float_info,
    }
}

/// §17.4.17: how many grid columns a `<w:tc>` covers — at least one.
///
/// `gridSpan` is a `ST_DecimalNumber`, so `w:val="0"` is well-formed and the
/// spec gives it no meaning: §17.4.17 defines the element as "the number of
/// grid columns spanned by the current cell", and a cell that exists occupies
/// a column. Every pass that walks the grid already says exactly that with
/// `span.max(1)` — `measure_table_rows`, `emit_table_rows`,
/// `find_cell_at_grid_col`, `cell_index_at_grid_col`,
/// `promote_orphan_vmerge_continues`.
///
/// This is the seam those passes are fed from, and it spelled the same rule
/// `unwrap_or(1)`: identical on an *absent* `gridSpan`, and off by one column
/// for every later cell in the row on `w:val="0"`. No corpus table writes one,
/// and while the walk produced nothing but widths the divergence changed no
/// pixel — but the walk now also addresses §17.7.6 conditional formatting, and
/// there the shortfall costs the row's last cell its `lastCol` layer. One
/// spelling, in one place, is what keeps the build layer and the layout passes
/// looking at the same grid.
fn grid_span(cell: &TableCell) -> u32 {
    cell.properties.grid_span.cloned().unwrap_or(1).max(1)
}

/// §17.4.84: turn every **orphaned** `w:vMerge w:val="continue"` cell — one
/// with no `restart` above it in any grid column it covers — into an ordinary
/// cell, and report how many were found.
///
/// # The defect this exists for
///
/// Nothing downstream sizes an orphan. `measure_table_rows` gives every
/// `Continue` cell an empty `CellLayout` and leaves it out of the row's
/// `max_height`, because a real merge is sized once by
/// `expand_rows_for_vmerge` over the whole span — and *that* pass only ever
/// starts from a `Restart`. With no `Restart` to start from, the two passes
/// between them account for the cell nowhere: its text is dropped from the
/// output entirely and its row collapses to zero height. That is data loss,
/// and no reading of §17.4.84 produces it: the section describes `continue`
/// only as continuing a merge a `restart` began, and a merged region shows the
/// *restart* cell's content precisely because there is one to show.
///
/// # The choice, and why it is a choice
///
/// §17.4.84 does not say what an unpaired `continue` means, so "promote it to
/// an ordinary cell" is this engine's decision, not a spec reading. The
/// alternatives were to treat the orphan as an implicit `restart` (inventing a
/// merge the author did not write, and one whose extent is guesswork), or to
/// drop the cell but keep its height (which loses the text just as surely).
///
/// **LibreOffice does the same thing**, which is corroboration rather than
/// proof: in `sw/source/core/unocore/unotext.cxx`, `lcl_ApplyCellProperties`
/// looks for an open merge group at the cell's position, and when it finds
/// none leaves `bFound` false, logs `SAL_WARN("sw.uno", "couldn't find first
/// vertically merged cell")` and does nothing else — the cell is never added to
/// a merge group and keeps the default RowSpan of 1, so it renders full height
/// with its content intact. The binary filter agrees: `ww8par2.cxx` treats a
/// null `FindMergeGroup` as "not merged". LibreOffice flags this explicitly as
/// a malformed-input fallback, so it is evidence about what a second
/// implementation chose, not about what Word does.
///
/// **A Word render would settle it** — a document with a `continue` cell that
/// no `restart` precedes, carrying text, rendered in Word: whether that text
/// appears, and whether its row takes its natural height, decides between this
/// choice and the implicit-`restart` one. Until then the fix is scoped to the
/// part every candidate reading agrees on, which is that the content survives.
///
/// # Which grid column, and why only one
///
/// A merge is tracked here per **grid column**, because `gridBefore` and
/// `gridSpan` make "which column" and "which cell index" different questions.
/// *Which* column is not a free choice either: the passes this repair has to
/// agree with anchor a merge on the restart cell's **first** column and no
/// other. `expand_rows_for_vmerge` walks down from a restart while
/// `is_vmerge_continue(row_below, entry.grid_col)` holds, and `entry.grid_col`
/// is that first column; `merged_span_height` repeats the same walk when it
/// emits. A `gridSpan="2"` restart therefore joins whatever covers its leftmost
/// column, and nothing else — the second column it spans opens no merge.
///
/// So a `Restart` opens `cols.start` alone, and a paired `Continue` carries
/// forward only the columns it covers that were **already** open, never the
/// rest of its own span. Both halves are load-bearing, and each was its own
/// content-losing bug while this pass paired on "any column the cell covers":
///
/// * a `Restart` opening its whole span made a narrower `Continue` *beside*
///   the restart's column read as paired. The merge passes looked in the
///   restart's column, found that row's ordinary cell there and never merged,
///   so `measure_table_rows` left the `Continue` with an empty `CellLayout` and
///   its text never reached the page.
/// * a `Continue` re-opening its whole span let a **wider** continuation
///   propagate openness into a column no `Restart` had ever begun, and an
///   orphan below it in that column was dropped the same way.
///
/// Stating the rule this way makes the two passes agree exactly rather than
/// approximately, which is what the "repair" has to mean: a cell this pass
/// promotes is precisely one that no merge would have sized or drawn, so
/// promoting it can only ever turn nothing into the cell the author wrote.
///
/// Both shapes are asserted end-to-end in `tests/table_row_height.rs` over
/// **spanning** cells, because a rule stated in grid columns cannot be tested
/// by a fixture whose every cell is one column wide — which is why the leak
/// survived a suite of span-1 fixtures.
fn promote_orphan_vmerge_continues(rows: &mut [TableRowInput]) -> usize {
    use crate::render::layout::table::VerticalMergeState;

    fn open_column(flags: &mut Vec<bool>, col: usize) {
        if flags.len() <= col {
            flags.resize(col + 1, false);
        }
        flags[col] = true;
    }

    // Per grid column: is a merge group open at this row?  A group opens on a
    // `Restart` and stays open only through consecutive `Continue`s, which is
    // exactly the run `expand_rows_for_vmerge` walks.
    let mut open: Vec<bool> = Vec::new();
    let mut promoted = 0;

    for row in rows.iter_mut() {
        let mut still_open: Vec<bool> = Vec::new();
        // §17.4.15: the row's first cell starts `gridBefore` columns in.
        let mut grid_col = row.grid_before as usize;
        for cell in row.cells.iter_mut() {
            let cols = grid_col..grid_col + cell.grid_span.max(1) as usize;
            match cell.vertical_merge {
                Some(VerticalMergeState::Restart) => open_column(&mut still_open, cols.start),
                Some(VerticalMergeState::Continue) => {
                    let mut carried = false;
                    for c in cols.clone() {
                        if open.get(c) == Some(&true) {
                            open_column(&mut still_open, c);
                            carried = true;
                        }
                    }
                    if !carried {
                        cell.vertical_merge = None;
                        promoted += 1;
                    }
                }
                None => {}
            }
            grid_col = cols.end;
        }
        open = still_open;
    }

    promoted
}

/// Word/LibreOffice row layout quirk: top/bottom cell margins are normalized
/// to the row-wide maximum across all cells in the row, while left/right stay
/// per-cell.
///
/// # Spec relationship
///
/// ECMA-376 §17.4.68 (`tcMar`, "Single Table Cell Margins") defines the cell
/// margin as a per-side exception over §17.4.42 (`tblCellMar`, the table-level
/// default). Each side is a `CT_TblWidth` (§17.4.87), where `@type="dxa"
/// @w="N"` is an explicit `N`-twip value. The spec is silent on how
/// *neighbouring* cells in the same row interact when their per-cell margins
/// disagree — there is no row-level "content area" concept defined in
/// §17.4.78 (`tr`) or §17.4.80 (`trHeight`).
///
/// The de-facto behaviour of every mainstream renderer (Word and LibreOffice
/// Writer in particular) is to compute a row-uniform content inset:
///
/// ```text
/// row.uniform_top    = max(cell.tcMar.top    for cell in row)
/// row.uniform_bottom = max(cell.tcMar.bottom for cell in row)
/// ```
///
/// and to position every cell's content within that uniform inset regardless
/// of the cell's own per-cell override. Without this pass, a cell with an
/// explicit `tcMar.top=0` in a row whose siblings inherit a larger value sits
/// flush against the row's top border while its neighbours sit padded — a
/// positioning no mainstream editor produces.
///
/// # Empirical basis
///
/// Verified against the `Wohnungsübergabeprotokoll` sample (LibreOffice
/// origin) by editing the DOCX directly and observing Word's render:
///
/// 1. Rewriting all `<w:tcMar><w:top w:w="0"/></w:tcMar>` to
///    `<w:top w:w="1"/>` produced **byte-equivalent visual output** — Word
///    is invariant to the literal value at this magnitude, ruling out
///    "Word treats `w=0` as no-override" as the explanation.
/// 2. Rewriting just one cell's `<w:tcMar>` to `<w:top w:w="500"/>` (≈ 25pt)
///    pushed **every** cell in that row down by ~25pt, confirming the
///    row-wide max(...) discipline.
///
/// # Scope
///
/// Only `top` and `bottom` are normalized. `left` and `right` stay per-cell
/// because each column's content width is independent — there is no shared
/// row-wide horizontal area in the de-facto layout, and editors do honour
/// per-cell horizontal overrides.
///
/// Applied at the build layer (post per-side `<w:tcMar>`/`<w:tblCellMar>`
/// cascade resolution) so downstream measure/emit/split code sees uniform
/// vertical insets and needs no further changes.
fn normalize_row_uniform_vertical_insets(cells: &mut [TableCellInput]) {
    let max_top = cells.iter().fold(Pt::ZERO, |acc, c| acc.max(c.margins.top));
    let max_bottom = cells
        .iter()
        .fold(Pt::ZERO, |acc, c| acc.max(c.margins.bottom));
    for cell in cells {
        cell.margins.top = max_top;
        cell.margins.bottom = max_bottom;
    }
}

/// Build a single table cell: resolve content blocks, margins, shading, borders.
fn build_table_cell(
    cell: &TableCell,
    table_style: Option<&ResolvedStyle>,
    style_cell_margins: Option<crate::model::geometry::EdgeInsets<crate::model::dimension::Twips>>,
    cond: &CellConditionalFormatting,
    inner_width: Pt,
    ctx: &BuildContext,
    state: &mut BuildState,
) -> TableCellInput {
    // §17.4.42: cell margins cascade *per side* against the pre-merged
    // table default. A cell-level `<w:tcMar>` that specifies only some sides
    // (e.g. `top`/`bottom` only — common in LibreOffice output) must inherit
    // the remaining sides from `<w:tblCellMar>` rather than zeroing them out;
    // collapsing missing sides to 0 produces text that hugs the cell borders
    // instead of carrying the table's intended padding.
    let table_default = style_cell_margins.unwrap_or(crate::model::geometry::EdgeInsets::ZERO);
    let resolved_margins = match cell.properties.margins.get() {
        Some(partial) => partial.resolve_against(table_default),
        None => table_default,
    };
    let cell_margins = geometry::PtEdgeInsets::new(
        Pt::from(resolved_margins.top),
        Pt::from(resolved_margins.right),
        Pt::from(resolved_margins.bottom),
        Pt::from(resolved_margins.left),
    );

    // §17.7.6: resolve cell shading.  Priority: direct → conditional → none.
    let shading = cell
        .properties
        .shading
        .get()
        .map(|s| resolve_color(s.fill, ColorContext::Background))
        .or_else(|| {
            cond.cell_properties
                .as_ref()
                .and_then(|tcp| tcp.shading.get())
                .map(|s| resolve_color(s.fill, ColorContext::Background))
        });

    // §17.4.66: cell borders cascade — direct cell borders (highest priority)
    // → conditional formatting → table-level borders (resolved in layout).
    let cond_borders = cond
        .cell_properties
        .as_ref()
        .and_then(|tcp| tcp.borders.get());
    let direct_borders = cell.properties.borders.get();

    let cell_borders = match (direct_borders, cond_borders) {
        (Some(db), _) => {
            // Direct cell borders: highest priority.  Fall through to
            // conditional for edges not specified directly.
            Some(CellBorderConfig {
                top: convert_cell_border_override(&db.top, state).or_else(|| {
                    cond_borders.and_then(|cb| convert_cell_border_override(&cb.top, state))
                }),
                bottom: convert_cell_border_override(&db.bottom, state).or_else(|| {
                    cond_borders.and_then(|cb| convert_cell_border_override(&cb.bottom, state))
                }),
                left: convert_cell_border_override(&db.left, state).or_else(|| {
                    cond_borders.and_then(|cb| convert_cell_border_override(&cb.left, state))
                }),
                right: convert_cell_border_override(&db.right, state).or_else(|| {
                    cond_borders.and_then(|cb| convert_cell_border_override(&cb.right, state))
                }),
            })
        }
        (None, Some(cb)) => Some(CellBorderConfig {
            top: convert_cell_border_override(&cb.top, state),
            bottom: convert_cell_border_override(&cb.bottom, state),
            left: convert_cell_border_override(&cb.left, state),
            right: convert_cell_border_override(&cb.right, state),
        }),
        (None, None) => None,
    };

    // §17.4.83: vertical alignment — direct cell, conditional, or default top.
    let valign = cell
        .properties
        .vertical_align
        .cloned()
        .or_else(|| {
            cond.cell_properties
                .as_ref()
                .and_then(|tcp| tcp.vertical_align.cloned())
        })
        .map(|va| match va {
            model::CellVerticalAlign::Bottom => crate::render::layout::table::CellVAlign::Bottom,
            model::CellVerticalAlign::Center => crate::render::layout::table::CellVAlign::Center,
            // `ST_VerticalJc`'s fourth value, `both`, asks for the cell's lines
            // to be *distributed* over its height rather than placed at one end
            // of it. That is a layout this engine does not have — `CellVAlign`
            // offers three placements and no distribution — so it takes the
            // same placement as `top`. A capability boundary rather than a
            // reading of §17.4.83, and deliberately not "no request": falling
            // through to the conditional layer instead would silently give such
            // a cell some *other* alignment the author did not ask for.
            _ => crate::render::layout::table::CellVAlign::Top,
        })
        .unwrap_or(crate::render::layout::table::CellVAlign::Top);

    // Estimate border insets to compute effective content width for
    // character-level splitting of oversized fragments.
    let border_w = |ovr: &Option<CellBorderOverride>| -> Pt {
        match ovr {
            Some(CellBorderOverride::Border(b)) => b.width,
            _ => Pt::ZERO,
        }
    };
    let border_inset_h = cell_borders
        .as_ref()
        .map(|cb| {
            let bl = (border_w(&cb.left) - cell_margins.left).max(Pt::ZERO);
            let br = (border_w(&cb.right) - cell_margins.right).max(Pt::ZERO);
            bl + br
        })
        .unwrap_or(Pt::ZERO);
    let content_width = (inner_width - border_inset_h).max(Pt::ZERO);

    // Recurse into cell content blocks.
    let cell_blocks =
        build_cell_blocks(&cell.content, table_style, cond, content_width, ctx, state);

    TableCellInput {
        blocks: cell_blocks,
        margins: cell_margins,
        grid_span: grid_span(cell),
        shading,
        cell_borders,
        vertical_merge: cell.properties.vertical_merge.cloned().map(|vm| match vm {
            model::VerticalMerge::Restart => {
                crate::render::layout::table::VerticalMergeState::Restart
            }
            model::VerticalMerge::Continue => {
                crate::render::layout::table::VerticalMergeState::Continue
            }
        }),
        vertical_align: valign,
    }
}

/// Recursively build cell content blocks.
///
/// Paragraphs are resolved with table style + conditional overrides.
/// Nested tables recurse via `build_table()` → `layout_table()`.
fn build_cell_blocks(
    content: &[Block],
    table_style: Option<&ResolvedStyle>,
    cond: &CellConditionalFormatting,
    inner_width: Pt,
    ctx: &BuildContext,
    state: &mut BuildState,
) -> Vec<LayoutBlock> {
    let mut blocks = Vec::new();
    let mut pending_dropcap: Option<DropCapInfo> = None;

    for (i, block) in content.iter().enumerate() {
        match block {
            Block::Paragraph(p) => {
                // §17.4.66: every cell must end with a paragraph. When the
                // last block is an empty paragraph following a table, it is
                // structural — Word renders it with zero height.
                if p.content.is_empty()
                    && i > 0
                    && matches!(content[i - 1], Block::Table(_))
                    && i == content.len() - 1
                {
                    continue;
                }
                if let Some(lb) = build_paragraph_block(
                    p,
                    ctx,
                    state,
                    &mut pending_dropcap,
                    table_style,
                    Some(cond),
                ) {
                    // Split oversized text fragments for narrow cells.
                    let lb = if let LayoutBlock::Paragraph {
                        fragments,
                        style,
                        page_break_before,
                        footnotes,
                        floating_images,
                        floating_shapes,
                    } = lb
                    {
                        // Keep the owned vector when no split is needed.
                        let measure = |t: &str, f: &crate::render::layout::fragment::FontProps| {
                            ctx.measurer.measure(t, f)
                        };
                        let fragments =
                            split_oversized_fragments(&fragments, inner_width, Some(&measure))
                                .unwrap_or(fragments);
                        LayoutBlock::Paragraph {
                            fragments,
                            style,
                            page_break_before,
                            footnotes,
                            floating_images,
                            floating_shapes,
                        }
                    } else {
                        lb
                    };
                    blocks.push(lb);
                }
            }
            Block::Table(nested_t) => {
                let built = build_table(nested_t, inner_width, ctx, state);
                blocks.push(LayoutBlock::Table {
                    rows: built.rows,
                    col_widths: built.col_widths,
                    cell_spacing: built.cell_spacing,
                    border_config: built.border_config,
                    indent: built.indent,
                    alignment: built.alignment,
                    float_info: built.float_info,
                    style_id: nested_t.properties.style_id.clone(),
                });
            }
            _ => {}
        }
    }

    blocks
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::render::layout::table::{CellVAlign, TableCellInput};

    // ── §17.4.48 seat_every_cell ─────────────────────────────────────────
    //
    // The shapes here are the ones no `.docx` can express in a way a rendered
    // page could then be measured — a four-billion-column row has no geometry
    // to check — so they are asserted against the function directly.

    /// A model row of `spans.len()` cells with the given `gridSpan`s.
    fn model_row(grid_before: u32, spans: &[u32]) -> model::TableRow {
        model::TableRow {
            properties: model::TableRowProperties {
                grid_before,
                ..Default::default()
            },
            cells: spans
                .iter()
                .map(|s| model::TableCell {
                    properties: model::TableCellProperties {
                        grid_span: model::Dup::from(Some(*s)),
                        ..Default::default()
                    },
                    content: Vec::new(),
                })
                .collect(),
            rsids: Default::default(),
            property_exceptions: None,
        }
    }

    fn twips(v: &[i64]) -> Vec<Dimension<Twips>> {
        v.iter().copied().map(Dimension::new).collect()
    }

    /// The gate: a grid that already seats every cell comes back untouched.
    /// Every one of the 398 tables in `test-files/` + `test-cases/` is this
    /// case, so this is the assertion that the repair is invisible to them.
    ///
    /// The early return it exercises is an optimisation, not the behaviour:
    /// falling through with `needed == declared.len()` appends an empty list
    /// and returns the same values, which the mutation check confirmed by
    /// finding `<=` → `<` to be an equivalent mutant. What that mutation cannot
    /// reach, `repair disabled entirely` in the integration mutation run does.
    #[test]
    fn a_seated_grid_is_returned_verbatim() {
        let declared = twips(&[1000, 2000, 3000]);
        let rows = vec![model_row(0, &[1, 1, 1]), model_row(1, &[2])];
        assert_eq!(seat_every_cell(&declared, &rows), declared);
    }

    /// §17.4.15 `gridBefore` is `ST_DecimalNumber`, so a file may state four
    /// billion of them. The demand is capped at one appended column per cell,
    /// so the allocation is bounded by the document that was parsed rather than
    /// by a number inside it — without the cap this asks for a 32 GB `Vec`.
    #[test]
    fn an_absurd_grid_before_cannot_size_the_allocation() {
        let declared = twips(&[1000, 2000]);
        let rows = vec![model_row(4_000_000_000, &[1])];
        let seated = seat_every_cell(&declared, &rows);
        assert_eq!(
            seated.len(),
            3,
            "two declared columns plus one cell is the whole budget"
        );
    }

    /// §17.4.17 `gridSpan` is the same type and gets the same bound. The cell
    /// is still seated — it has a column and will be drawn — it is merely
    /// narrower than the four billion columns it asked for.
    #[test]
    fn an_absurd_grid_span_cannot_size_the_allocation() {
        let declared = twips(&[1000, 2000]);
        let rows = vec![model_row(0, &[1, 4_000_000_000])];
        let seated = seat_every_cell(&declared, &rows);
        assert_eq!(seated.len(), 4, "two declared columns plus two cells");
        assert!(
            seated.iter().all(|w| *w > Dimension::ZERO),
            "no column may be zero-width: {seated:?}"
        );
    }

    /// A row with no cells asks for nothing, whatever its `gridBefore` says —
    /// it is the same empty-column case as `gridAfter`, and there is no content
    /// in it to lose. Without the filter this row would append 6 columns and
    /// widen the table for every other row.
    #[test]
    fn a_row_with_no_cells_demands_no_columns() {
        let declared = twips(&[1000, 2000]);
        let rows = vec![model_row(6, &[]), model_row(0, &[1, 1])];
        assert_eq!(seat_every_cell(&declared, &rows), declared);
    }

    /// An empty table — no grid, no rows — stays empty rather than becoming a
    /// one-column table.
    ///
    /// Held by the cell budget rather than by `unwrap_or(0)`: with no cells the
    /// cap is zero, so the demand is clamped to zero whatever the `unwrap_or`
    /// says. The mutation check found `unwrap_or(1)` to be equivalent for
    /// exactly that reason, which is worth knowing — the two guards overlap
    /// here, and removing either still leaves this correct.
    #[test]
    fn a_table_with_no_rows_and_no_grid_stays_empty() {
        assert!(seat_every_cell(&[], &[]).is_empty());
    }

    fn cell_with_margins(top: f32, right: f32, bottom: f32, left: f32) -> TableCellInput {
        TableCellInput {
            blocks: vec![],
            margins: geometry::PtEdgeInsets::new(
                Pt::new(top),
                Pt::new(right),
                Pt::new(bottom),
                Pt::new(left),
            ),
            grid_span: 1,
            shading: None,
            cell_borders: None,
            vertical_merge: None,
            vertical_align: CellVAlign::Top,
        }
    }

    /// Word/Writer row-uniform content-area normalization: when cells in a row
    /// disagree on `tcMar.top` (or `bottom`), the row-wide *maximum* wins for
    /// every cell. Verified empirically against MS Word — see
    /// [`normalize_row_uniform_vertical_insets`] for the experimental
    /// reproducer and spec references.
    #[test]
    fn row_uniform_picks_max_top_and_bottom_across_row() {
        // Mirrors the Wohnungsübergabe "Keller" row: one cell with
        // explicit zero top/bottom (the Keller cell after per-side tcMar
        // cascade) and another with the inherited 57-twip table default
        // (≈ 2.85 pt) — siblings without a tcMar override.
        let mut cells = vec![
            cell_with_margins(0.0, 5.4, 0.0, 5.15),   // Keller-like
            cell_with_margins(2.85, 5.4, 2.85, 5.15), // sibling with table default
            cell_with_margins(0.0, 5.4, 0.0, 5.15),   // another Keller-like
        ];
        normalize_row_uniform_vertical_insets(&mut cells);

        for (i, cell) in cells.iter().enumerate() {
            assert_eq!(
                cell.margins.top.raw(),
                2.85,
                "cell #{i}: every cell's top inset must equal the row-wide max"
            );
            assert_eq!(
                cell.margins.bottom.raw(),
                2.85,
                "cell #{i}: every cell's bottom inset must equal the row-wide max"
            );
            // Horizontal stays per-cell.
            assert_eq!(cell.margins.left.raw(), 5.15, "left is per-cell");
            assert_eq!(cell.margins.right.raw(), 5.4, "right is per-cell");
        }
    }

    /// Mirrors the `_kellertop500` experiment: a single cell with a `tcMar.top`
    /// far larger than its siblings becomes the row-wide max, so every cell —
    /// including the ones with no override — picks up that large top inset.
    /// In Word, this is observable as the entire row's content shifting down.
    #[test]
    fn row_uniform_one_large_cell_top_pushes_whole_row_down() {
        let mut cells = vec![
            cell_with_margins(25.0, 5.4, 0.0, 5.15), // Keller with top=25pt
            cell_with_margins(2.85, 5.4, 2.85, 5.15), // small sibling
        ];
        normalize_row_uniform_vertical_insets(&mut cells);

        assert_eq!(
            cells[1].margins.top.raw(),
            25.0,
            "sibling cell inherits the large top from the dominating cell"
        );
        // Bottom max is the smaller cell's 2.85 (Keller bottom stayed 0).
        assert_eq!(cells[0].margins.bottom.raw(), 2.85);
    }

    #[test]
    fn row_uniform_single_cell_row_is_noop() {
        let mut cells = vec![cell_with_margins(2.85, 5.4, 2.85, 5.15)];
        let before = cells[0].margins;
        normalize_row_uniform_vertical_insets(&mut cells);
        assert_eq!(cells[0].margins, before);
    }

    #[test]
    fn row_uniform_handles_empty_row() {
        let mut cells: Vec<TableCellInput> = vec![];
        normalize_row_uniform_vertical_insets(&mut cells);
        // No panic, no work done.
        assert!(cells.is_empty());
    }

    /// §17.4.87 / §17.4.45: `CT_TblWidth` permits `pct` and `auto`, and the
    /// element's own section says both are **ignored** here — only `dxa`
    /// carries a usable value.
    /// Honouring a percentage here would scale the gap with the table width,
    /// which is exactly what the spec rules out.
    #[test]
    fn cell_spacing_only_honours_dxa() {
        use crate::model::dimension::Dimension;
        assert!(
            resolve_cell_spacing(Some(model::TableMeasure::Twips(Dimension::new(240)))) > Pt::ZERO,
            "a dxa measurement carries a usable spacing"
        );
        for ignored in [
            Some(model::TableMeasure::Pct(Dimension::new(2500))),
            Some(model::TableMeasure::Auto),
            Some(model::TableMeasure::Nil),
            None,
        ] {
            assert_eq!(resolve_cell_spacing(ignored), Pt::ZERO, "{ignored:?}");
        }
    }

    /// §17.4.45: the declared value passes through unscaled — it *is* the gap.
    ///
    /// Pinned because it was briefly changed to `2x` on the strength of a Word
    /// render, and reverted when ONLYOFFICE's renderer turned out to apply no
    /// factor at all. See `resolve_cell_spacing`'s doc comment for that evidence
    /// and for what the doubling in that render is more likely to have been.
    #[test]
    fn a_declared_spacing_is_the_whole_gap() {
        use crate::model::dimension::Dimension;
        assert_eq!(
            resolve_cell_spacing(Some(model::TableMeasure::Twips(Dimension::new(240)))),
            Pt::new(12.0),
            "240 twips = 12pt declared and 12pt rendered"
        );
    }

    /// The spacing is carved *out of* the table, so the slots shrink by exactly
    /// one spacing in total and keep their proportions.
    #[test]
    fn reserving_spacing_shrinks_the_grid_by_exactly_one_spacing() {
        let widths = vec![Pt::new(60.0), Pt::new(40.0)];
        let out = reserve_cell_spacing(widths, Pt::new(10.0));
        let total: Pt = out.iter().copied().sum();
        assert_eq!(total, Pt::new(90.0), "one spacing removed in total");
        // Proportions preserved: 60:40 becomes 54:36.
        assert_eq!(out[0], Pt::new(54.0));
        assert_eq!(out[1], Pt::new(36.0));
    }

    /// Zero spacing must be a byte-for-byte no-op — every table that does not
    /// set `tblCellSpacing` goes through this path.
    #[test]
    fn reserving_zero_spacing_leaves_the_grid_untouched() {
        let widths = vec![Pt::new(60.0), Pt::new(40.0)];
        assert_eq!(reserve_cell_spacing(widths.clone(), Pt::ZERO), widths);
    }

    /// A spacing at least as wide as the table leaves nothing to distribute.
    /// Word caps the value long before this, but the file format does not, and
    /// scaling by a negative factor would put cells at negative widths.
    #[test]
    fn spacing_wider_than_the_table_collapses_columns_rather_than_going_negative() {
        let widths = vec![Pt::new(30.0), Pt::new(30.0)];
        let out = reserve_cell_spacing(widths, Pt::new(60.0));
        assert_eq!(out, vec![Pt::ZERO, Pt::ZERO]);
        assert!(out.iter().all(|w| *w >= Pt::ZERO));
    }

    // ── §17.4.63 clamp_auto_grid_to_page ─────────────────────────────────
    //
    // `tests/table_auto_width.rs` pins the end-to-end consequence and the
    // reasoning; these pin the arithmetic and the two edges of the branch.

    /// Letter, 1-inch margins: text column 468 pt, paper 612 pt, so the guard's
    /// limit is 612 − 72 = 540 pt.
    fn letter() -> crate::render::layout::page::PageConfig {
        crate::render::layout::page::PageConfig::default()
    }

    #[test]
    fn a_fitting_auto_grid_is_returned_verbatim() {
        let grid = vec![Pt::new(100.0), Pt::new(200.0)];
        assert_eq!(
            clamp_auto_grid_to_page(&grid, 2, Pt::new(468.0), &letter()),
            grid
        );
    }

    /// The guard's line is the paper edge, not the text column: 500 pt overflows
    /// the 468 pt column by 32 and is still left alone. A clamp written against
    /// `available_width` fails here, which is the point — 40 tables of real Word
    /// output in the corpus have exactly this shape.
    #[test]
    fn a_grid_past_the_text_column_but_on_the_paper_is_left_alone() {
        let grid = vec![Pt::new(500.0)];
        assert_eq!(
            clamp_auto_grid_to_page(&grid, 1, Pt::new(468.0), &letter()),
            grid
        );
    }

    /// Past the paper, the grid is scaled — proportionally, so the columns keep
    /// their declared ratio rather than being equalised.
    #[test]
    fn a_grid_past_the_paper_is_scaled_to_the_paper() {
        let out = clamp_auto_grid_to_page(
            &[Pt::new(800.0), Pt::new(400.0)],
            2,
            Pt::new(468.0),
            &letter(),
        );
        let total: Pt = out.iter().copied().sum();
        assert_eq!(total, Pt::new(540.0), "scaled to the left margin → paper");
        assert_eq!(out, vec![Pt::new(360.0), Pt::new(180.0)], "ratio kept 2:1");
    }

    /// A page whose margins exceed its width is expressible in the file format
    /// (`PageConfig::content_width` defends against the same thing), and would
    /// otherwise give a negative limit and negative column widths. The offered
    /// width is the floor.
    #[test]
    fn a_page_narrower_than_its_own_margins_falls_back_to_the_offered_width() {
        let mut page = letter();
        page.margins.left = Pt::new(900.0);
        let out = clamp_auto_grid_to_page(&[Pt::new(1000.0)], 1, Pt::new(468.0), &page);
        assert_eq!(out, vec![Pt::new(468.0)]);
    }

    // ── §17.4.84 promote_orphan_vmerge_continues ─────────────────────────
    //
    // The end-to-end consequence — an orphan's text surviving into the PDF —
    // is pinned by `tests/table_row_height.rs`. These pin the *pairing rule*,
    // which is by grid column rather than by cell index and is not reachable
    // from a document whose rows all start at column 0.

    use crate::render::layout::table::{TableRowInput, VerticalMergeState};

    fn merge_cell(span: u32, vm: Option<VerticalMergeState>) -> TableCellInput {
        TableCellInput {
            grid_span: span,
            vertical_merge: vm,
            ..cell_with_margins(0.0, 0.0, 0.0, 0.0)
        }
    }

    fn merge_row(cells: Vec<TableCellInput>, grid_before: u32) -> TableRowInput {
        TableRowInput {
            cells,
            height_rule: None,
            is_header: None,
            cant_split: None,
            grid_before,
            border_overrides: None,
        }
    }

    /// The defect itself, at the level the repair happens: a `Continue` with
    /// nothing above it becomes an ordinary cell and is counted.
    #[test]
    fn an_unpaired_continue_is_promoted_to_an_ordinary_cell() {
        let mut rows = vec![merge_row(
            vec![merge_cell(1, Some(VerticalMergeState::Continue))],
            0,
        )];
        assert_eq!(promote_orphan_vmerge_continues(&mut rows), 1);
        assert_eq!(rows[0].cells[0].vertical_merge, None);
    }

    /// A real merge is untouched, however long its chain of `Continue`s.
    #[test]
    fn a_paired_continue_chain_is_left_alone() {
        let mut rows = vec![
            merge_row(vec![merge_cell(1, Some(VerticalMergeState::Restart))], 0),
            merge_row(vec![merge_cell(1, Some(VerticalMergeState::Continue))], 0),
            merge_row(vec![merge_cell(1, Some(VerticalMergeState::Continue))], 0),
        ];
        assert_eq!(promote_orphan_vmerge_continues(&mut rows), 0);
        assert_eq!(
            rows[2].cells[0].vertical_merge,
            Some(VerticalMergeState::Continue)
        );
    }

    /// A span is the run of **consecutive** rows below the restart — that is
    /// how `expand_rows_for_vmerge` walks it, stopping at the first row whose
    /// cell in the column is not a `Continue`. A `Continue` that resumes after
    /// an intervening plain row therefore continues nothing, and would
    /// otherwise be sized by nobody.
    #[test]
    fn a_continue_that_resumes_after_a_plain_row_is_an_orphan() {
        let mut rows = vec![
            merge_row(vec![merge_cell(1, Some(VerticalMergeState::Restart))], 0),
            merge_row(vec![merge_cell(1, None)], 0),
            merge_row(vec![merge_cell(1, Some(VerticalMergeState::Continue))], 0),
        ];
        assert_eq!(promote_orphan_vmerge_continues(&mut rows), 1);
        assert_eq!(rows[2].cells[0].vertical_merge, None);
    }

    /// §17.4.15: the pairing is by **grid column**, not by cell index. With
    /// `gridBefore=1` the lower row's first cell sits in column 1, so it does
    /// not continue a merge that started in column 0 — and a fixture where
    /// `grid_col == cell_index` cannot tell the two rules apart.
    #[test]
    fn pairing_follows_the_grid_column_not_the_cell_index() {
        let mut rows = vec![
            merge_row(
                vec![
                    merge_cell(1, Some(VerticalMergeState::Restart)), // column 0
                    merge_cell(1, None),                              // column 1
                ],
                0,
            ),
            merge_row(vec![merge_cell(1, Some(VerticalMergeState::Continue))], 1),
        ];
        assert_eq!(promote_orphan_vmerge_continues(&mut rows), 1);
        assert_eq!(rows[1].cells[0].vertical_merge, None);
    }

    /// A `gridSpan` restart opens its **first** column only, because that is
    /// the column `expand_rows_for_vmerge` and `merged_span_height` anchor on.
    /// A `Continue` beside it — covering the restart's *second* column, which
    /// no merge pass looks in — continues nothing and must be promoted.
    ///
    /// This asserted 0 promotions until the end-to-end reproduction showed the
    /// consequence: the cell was called paired here, joined by nobody
    /// downstream, and its text was dropped from the PDF with no warning. The
    /// two shapes that follow are the ones a span-1 fixture cannot express.
    #[test]
    fn a_continue_beside_a_wide_restarts_column_is_an_orphan() {
        let mut rows = vec![
            merge_row(vec![merge_cell(2, Some(VerticalMergeState::Restart))], 0),
            merge_row(
                vec![
                    merge_cell(1, None),                               // column 0
                    merge_cell(1, Some(VerticalMergeState::Continue)), // column 1
                ],
                0,
            ),
        ];
        assert_eq!(promote_orphan_vmerge_continues(&mut rows), 1);
        assert_eq!(rows[1].cells[1].vertical_merge, None);
    }

    /// The other side of the same rule, and the control a blanket "promote
    /// anything under a wide restart" would break: a `Continue` covering the
    /// restart's *first* column is the merge those passes do join, so it stays.
    #[test]
    fn a_continue_under_a_wide_restarts_first_column_is_paired() {
        let mut rows = vec![
            merge_row(vec![merge_cell(2, Some(VerticalMergeState::Restart))], 0),
            merge_row(
                vec![
                    merge_cell(1, Some(VerticalMergeState::Continue)), // column 0
                    merge_cell(1, None),                               // column 1
                ],
                0,
            ),
        ];
        assert_eq!(promote_orphan_vmerge_continues(&mut rows), 0);
        assert_eq!(
            rows[1].cells[0].vertical_merge,
            Some(VerticalMergeState::Continue)
        );
    }

    /// A paired `Continue` carries forward only the columns that were already
    /// open — widening the open set to its whole span would invent a merge
    /// group in a column no `Restart` began, and orphan the cell below it into
    /// silence. Row 2's `Continue` sits in column 1, which only the wide
    /// `Continue` above it ever touched.
    #[test]
    fn a_wide_continue_does_not_open_a_column_no_restart_began() {
        let mut rows = vec![
            merge_row(
                vec![
                    merge_cell(1, Some(VerticalMergeState::Restart)), // column 0
                    merge_cell(1, None),                              // column 1
                ],
                0,
            ),
            merge_row(
                vec![merge_cell(2, Some(VerticalMergeState::Continue))], // columns 0–1
                0,
            ),
            merge_row(
                vec![
                    merge_cell(1, None),                               // column 0
                    merge_cell(1, Some(VerticalMergeState::Continue)), // column 1
                ],
                0,
            ),
        ];
        assert_eq!(promote_orphan_vmerge_continues(&mut rows), 1);
        assert_eq!(
            rows[1].cells[0].vertical_merge,
            Some(VerticalMergeState::Continue),
            "the wide continuation covers column 0 and stays a merge"
        );
        assert_eq!(rows[2].cells[1].vertical_merge, None);
    }

    // ── `build_table`: the orchestrator's own output ──────────────────────
    //
    // `BuiltTable` carries seven fields and every one of them is several
    // passes away from a pixel — a `col_widths` entry becomes a cell box only
    // after measure, grid and emit have each had a say, and `float_info` only
    // after the paginator has. The integration files pin what reaches the
    // page; these pin the arithmetic that produces it, so a failure names the
    // term that moved rather than the document that came out wrong.
    //
    // Every expected number is derived from the inputs and nothing else: 20
    // twips to the point, 8 eighth-points to the point, `pct` in fiftieths of
    // a percent (§17.18.90), and a width the caller offers rather than
    // defaults.

    use crate::render::fonts::FontRegistry;
    use crate::render::layout::measurer::TextMeasurer;
    use crate::render::resolve::ResolvedDocument;
    use std::collections::HashMap;

    /// The width every case below is offered: the Letter text column, which is
    /// what a top-level table gets from the default page config (8.5 in − 2 ×
    /// 1 in = 6.5 in = 468 pt).
    const OFFERED: Pt = Pt::new(468.0);

    /// Every width here is a twips-to-points division times a scale factor, so
    /// the comparison carries a tolerance — well below the 0.01 pt that any of
    /// the assertions turn on.
    #[track_caller]
    fn assert_pt(actual: Pt, expected: f32, what: &str) {
        assert!(
            (actual.raw() - expected).abs() < 0.005,
            "{what}: expected {expected} pt, got {actual:?}"
        );
    }

    fn total(widths: &[Pt]) -> Pt {
        widths.iter().copied().sum()
    }

    /// A table style whose `<w:tblPr>` is `table` and which defines no
    /// conditional layers.
    fn plain_style(table: model::TableProperties) -> ResolvedStyle {
        layered_style(table, Vec::new())
    }

    /// A table style with `<w:tblStylePr>` layers (§17.7.6.6) as well.
    fn layered_style(
        table: model::TableProperties,
        overrides: Vec<model::TableStyleOverride>,
    ) -> ResolvedStyle {
        ResolvedStyle {
            paragraph: model::ParagraphProperties::default(),
            run: model::RunProperties::default(),
            table: Some(table),
            table_style_overrides: overrides,
            is_toc_entry: false,
        }
    }

    /// A document whose stylesheet holds exactly `styles`.
    fn resolved_with(styles: &[(&str, ResolvedStyle)]) -> ResolvedDocument {
        ResolvedDocument {
            sections: Vec::new(),
            styles: styles
                .iter()
                .map(|(id, style)| (model::StyleId::new(*id), style.clone()))
                .collect(),
            numbering: HashMap::new(),
            font_families: Vec::new(),
            media: HashMap::new(),
            diagrams: Default::default(),
            charts: Default::default(),
            embedded_fonts: Vec::new(),
            pic_bullets: HashMap::new(),
            theme: None,
            doc_defaults_paragraph: model::ParagraphProperties::default(),
            doc_defaults_run: model::RunProperties::default(),
            default_paragraph_style_id: None,
            default_table_style_id: None,
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
            even_and_odd_headers: false,
            default_tab_stop: Dimension::new(720),
        }
    }

    /// Build, and hand back the state as well — the warn-once flags are the
    /// only record of what the build *reported*, and nothing on the page shows
    /// them.
    fn build_reporting_offered(
        t: &Table,
        offered: Pt,
        styles: &[(&str, ResolvedStyle)],
    ) -> (BuiltTable, BuildState) {
        let resolved = resolved_with(styles);
        let registry = FontRegistry::new(skia_safe::FontMgr::new());
        let measurer = TextMeasurer::new(&registry);
        let ctx = BuildContext {
            measurer: &measurer,
            resolved: &resolved,
        };
        let mut state = BuildState::default();
        let built = build_table(t, offered, &ctx, &mut state);
        (built, state)
    }

    fn build_offered(t: &Table, offered: Pt, styles: &[(&str, ResolvedStyle)]) -> BuiltTable {
        build_reporting_offered(t, offered, styles).0
    }

    /// The common case: the Letter text column, and an empty stylesheet.
    fn build(t: &Table) -> BuiltTable {
        build_offered(t, OFFERED, &[])
    }

    /// [`build`], keeping the state.
    fn build_reporting(t: &Table) -> (BuiltTable, BuildState) {
        build_reporting_offered(t, OFFERED, &[])
    }

    /// A table with `grid` declared in twips and one row of `cells` empty
    /// cells, each spanning one column.
    fn table_of(properties: model::TableProperties, grid: &[i64], cells: usize) -> Table {
        Table {
            properties,
            grid: grid
                .iter()
                .map(|w| model::GridColumn {
                    width: Dimension::new(*w),
                })
                .collect(),
            rows: vec![model_row(0, &vec![1u32; cells])],
        }
    }

    fn measure(m: model::TableMeasure) -> model::Dup<model::TableMeasure> {
        model::Dup::from(Some(m))
    }

    fn dxa(twips: i64) -> model::TableMeasure {
        model::TableMeasure::Twips(Dimension::new(twips))
    }

    fn pct(fiftieths: i64) -> model::TableMeasure {
        model::TableMeasure::Pct(Dimension::new(fiftieths))
    }

    /// A `<w:tblCellMar>` stating all four sides, in twips.
    fn cell_mar(
        top: i64,
        right: i64,
        bottom: i64,
        left: i64,
    ) -> crate::model::geometry::EdgeInsets<Twips> {
        crate::model::geometry::EdgeInsets::new(
            Dimension::new(top),
            Dimension::new(right),
            Dimension::new(bottom),
            Dimension::new(left),
        )
    }

    /// One border edge: `w:sz` in eighths of a point, so 8 is 1 pt.
    fn border(eighths: i64, style: model::BorderStyle) -> Option<model::Border> {
        Some(model::Border {
            style,
            width: Dimension::new(eighths),
            space: Dimension::ZERO,
            color: model::Color::BLACK,
        })
    }

    /// `<w:tblBorders>` with all six edges the same `single` line.
    fn uniform_borders(eighths: i64) -> model::TableBorders {
        let e = || border(eighths, model::BorderStyle::Single);
        model::TableBorders {
            top: e(),
            bottom: e(),
            left: e(),
            right: e(),
            inside_h: e(),
            inside_v: e(),
        }
    }

    /// `<w:tcBorders>` with the four outer edges the same line.
    fn uniform_cell_borders(eighths: i64) -> model::TableCellBorders {
        let e = || border(eighths, model::BorderStyle::Single);
        model::TableCellBorders {
            top: e(),
            bottom: e(),
            left: e(),
            right: e(),
            inside_h: None,
            inside_v: None,
            tl2br: None,
            tr2bl: None,
        }
    }

    /// The width of a resolved cell edge, panicking on a suppressed or absent
    /// one — every call below expects a line.
    #[track_caller]
    fn override_width(o: &Option<CellBorderOverride>) -> Pt {
        match o {
            Some(CellBorderOverride::Border(b)) => b.width,
            other => panic!("expected a drawn border, got {other:?}"),
        }
    }

    fn para(content: Vec<model::Inline>) -> Block {
        Block::Paragraph(Box::new(model::Paragraph {
            style_id: None,
            properties: model::ParagraphProperties::default(),
            mark_run_properties: None,
            content,
            rsids: model::ParagraphRevisionIds::default(),
        }))
    }

    fn text_run(s: &str) -> model::Inline {
        model::Inline::TextRun(Box::new(model::TextRun {
            style_id: None,
            properties: model::RunProperties::default(),
            content: vec![model::RunElement::Text(s.to_string())],
            rsids: model::RevisionIds::default(),
        }))
    }

    /// §17.4.63 / §17.18.90: a `pct` table width is that fraction of the width
    /// offered, and the full-width cell-margin extension is keyed on the
    /// fraction **reaching** `FULL` — 5000 fiftieths, 100% — rather than on it
    /// being large.
    ///
    /// One fiftieth of a percent short of full the base is the plain 468 pt
    /// offered, so the table is 4999/5000 × 468 = 467.906 pt. At full the base
    /// is 468 plus both of the table's cell margins, 2 × 108 twips = 10.8 pt,
    /// for 478.8. Two widths one unit apart in the file come out 10.9 pt apart
    /// on the page, and that discontinuity *is* the heuristic — see
    /// `extends_for_alignment`'s comment for why a full-width table is the one
    /// case Word widens.
    ///
    /// `tests/table_geometry_sizing.rs` pins each side of this on the page, but
    /// only ever with cell margins **or** a sub-full percentage, never both:
    /// with the comparison written `<= FULL` instead of `>= FULL` every one of
    /// those cases still passes, and so does every table in the corpus.
    #[test]
    fn the_cell_margin_extension_begins_exactly_at_a_full_width_pct() {
        let width = |w: model::TableMeasure| {
            let props = model::TableProperties {
                width: measure(w),
                cell_margins: model::Dup::from(Some(cell_mar(0, 108, 0, 108))),
                ..Default::default()
            };
            total(&build(&table_of(props, &[2000, 2000], 2)).col_widths)
        };

        assert_pt(width(pct(4999)), 467.9064, "one fiftieth short of full");
        assert_pt(width(pct(5000)), 478.8, "at full, plus both cell margins");
    }

    /// §17.4.63: `auto` and `nil` both state *no* preferred width, so the
    /// declared grid is kept as declared. `nil` in particular is not a
    /// zero-width table — `CT_TblWidth`'s `nil` is the absence of a value, and
    /// a table 0 pt wide is not a reading any section supports.
    ///
    /// The `dxa` control is what makes that a claim rather than an
    /// observation: the same grid under a declared width is scaled to it.
    #[test]
    fn an_auto_or_nil_table_width_keeps_the_declared_grid() {
        let grid = [1440, 2880]; // 72 pt and 144 pt, summing to 216
        let cols = |w: Option<model::TableMeasure>| {
            let props = model::TableProperties {
                width: model::Dup::from(w),
                ..Default::default()
            };
            build(&table_of(props, &grid, 2)).col_widths
        };

        for w in [
            None,
            Some(model::TableMeasure::Auto),
            Some(model::TableMeasure::Nil),
        ] {
            let out = cols(w);
            assert_pt(out[0], 72.0, &format!("{w:?}: column 0 as declared"));
            assert_pt(out[1], 144.0, &format!("{w:?}: column 1 as declared"));
        }

        let scaled = cols(Some(dxa(2160))); // 108 pt against a 216 pt grid
        assert_pt(scaled[0], 36.0, "a dxa width halves the same grid");
        assert_pt(scaled[1], 72.0, "…keeping its declared 1:2 ratio");
    }

    /// §17.4.50: a `dxa` `tblInd` is that many twips of indent, and the direct
    /// occurrence outranks the style's rather than adding to it.
    ///
    /// `tests/table_style_cascade.rs` pins that a style's `tblInd` reaches the
    /// table at all, written as parity against the same element on the
    /// `<w:tbl>`. What parity cannot show is the number, or which level wins
    /// when both levels declare one — the case that decides whether the read is
    /// a cascade or a sum.
    #[test]
    fn a_direct_tbl_ind_is_its_twips_in_points_and_outranks_the_styles() {
        let style = plain_style(model::TableProperties {
            indent: measure(dxa(1440)),
            ..Default::default()
        });
        let indent_of = |direct: Option<model::TableMeasure>| {
            let props = model::TableProperties {
                style_id: Some(model::StyleId::new("T")),
                indent: model::Dup::from(direct),
                width: measure(dxa(4000)),
                ..Default::default()
            };
            build_offered(
                &table_of(props, &[4000], 1),
                OFFERED,
                &[("T", style.clone())],
            )
            .indent
        };

        assert_pt(indent_of(None), 72.0, "1440 twips from the style");
        assert_pt(indent_of(Some(dxa(720))), 36.0, "the direct 720 twips wins");
    }

    /// §17.4.50 / §17.4.63: the full-width left-aligned shift is keyed on the
    /// **resolved indent being zero**, not on the `tblInd` element being
    /// absent — and "resolves to zero" covers every measure this engine does
    /// not read, not only `dxa 0`.
    ///
    /// `tblInd` is a `CT_TblWidth`, so a file may state `pct`, `auto` or `nil`
    /// there. Only `dxa` is read, which is the same reading
    /// `resolve_cell_spacing` applies to the same type — this engine has no
    /// base to measure a percentage indent against, the table's own width being
    /// the thing the indent helps place. **Word reference render needed**: a
    /// table at `<w:tblInd w:w="2500" w:type="pct"/>` says whether Word reads
    /// it and against what.
    ///
    /// What this pins regardless of that answer is the *rule*. Five spellings
    /// that resolve to no indent must all take the shift, so a guard rewritten
    /// as "the element is absent" fails on four of them; and the shift is
    /// −5.4 pt, one left cell margin, not some other multiple.
    /// `tests/table_style_cascade.rs` covers the two spellings the corpus has —
    /// absent, and `dxa 0` — on the page.
    #[test]
    fn every_tbl_ind_that_resolves_to_no_indent_takes_the_full_width_shift() {
        let indent_of = |ind: Option<model::TableMeasure>| {
            let props = model::TableProperties {
                width: measure(pct(5000)),
                indent: model::Dup::from(ind),
                cell_margins: model::Dup::from(Some(cell_mar(0, 108, 0, 108))),
                ..Default::default()
            };
            build(&table_of(props, &[2000, 2000], 2)).indent
        };

        for ind in [
            None,
            Some(model::TableMeasure::Auto),
            Some(model::TableMeasure::Nil),
            Some(pct(2500)),
            Some(dxa(0)),
        ] {
            assert_pt(indent_of(ind), -5.4, &format!("{ind:?} is no indent"));
        }
        assert_pt(
            indent_of(Some(dxa(720))),
            36.0,
            "…while a real indent is still taken literally",
        );
    }

    /// §17.4.28 / §17.4.63: the full-width extension and the negative shift are
    /// one rule with two halves, and both are withheld from a table placed as a
    /// unit — the `matches!(alignment, Center | End)` the two guards share.
    ///
    /// `tests/table_geometry_sizing.rs` pins absent-versus-`center` on the
    /// page. The two spellings it cannot reach are `end`, the other arm of that
    /// `matches!`, which no test at any level exercises; and `start`, the
    /// explicit spelling of the very case the heuristic is written for, which
    /// must behave exactly as an absent `w:jc` — the `matches!` is written as a
    /// negation, so a `start` that fell through it would lose the shift.
    #[test]
    fn only_a_start_aligned_full_width_table_extends_by_its_cell_margins() {
        let built = |jc: Option<model::Alignment>| {
            let props = model::TableProperties {
                width: measure(pct(5000)),
                alignment: model::Dup::from(jc),
                cell_margins: model::Dup::from(Some(cell_mar(0, 108, 0, 108))),
                ..Default::default()
            };
            build(&table_of(props, &[2000, 2000], 2))
        };

        for jc in [None, Some(model::Alignment::Start)] {
            let b = built(jc);
            assert_eq!(b.alignment, jc, "the alignment passes through as declared");
            assert_pt(total(&b.col_widths), 478.8, "468 + 2 × 5.4 pt of margin");
            assert_pt(b.indent, -5.4, "…pulled left by the left cell margin");
        }
        for jc in [Some(model::Alignment::Center), Some(model::Alignment::End)] {
            let b = built(jc);
            assert_eq!(b.alignment, jc);
            assert_pt(
                total(&b.col_widths),
                468.0,
                &format!("{jc:?}: the offered width verbatim"),
            );
            assert_pt(b.indent, 0.0, &format!("{jc:?}: and no shift"));
        }
    }

    /// §17.4.38: a direct `<w:tblBorders>` merges over the style's **per
    /// edge** — an edge the table does not mention keeps the style's — and a
    /// direct edge spelled `val="none"` still *wins* that merge, which is what
    /// makes it remove the style's border rather than inherit it.
    ///
    /// That is the asymmetry with §17.4.66, where the same `none` on a **cell**
    /// means inherit ([MS-OI29500] starts the cell cascade at "if the cell
    /// border is omitted or none"), and where `nil` is the spelling that
    /// suppresses. `tests/table_border_conflict.rs` pins the cell side on the
    /// page; the last two assertions here put both levels in one fixture, since
    /// "the same markup is read oppositely at two levels" is not a claim either
    /// level can make alone.
    #[test]
    fn direct_table_borders_merge_over_the_styles_edge_by_edge() {
        let style = plain_style(model::TableProperties {
            borders: model::Dup::from(Some(uniform_borders(8))),
            ..Default::default()
        });
        let props = model::TableProperties {
            style_id: Some(model::StyleId::new("T")),
            width: measure(dxa(4000)),
            borders: model::Dup::from(Some(model::TableBorders {
                top: border(16, model::BorderStyle::Single),
                inside_h: border(8, model::BorderStyle::None),
                bottom: None,
                left: None,
                right: None,
                inside_v: None,
            })),
            ..Default::default()
        };
        let mut t = table_of(props, &[2000, 2000], 2);
        t.rows[0].cells[0].properties.borders = model::Dup::from(Some(model::TableCellBorders {
            top: border(8, model::BorderStyle::None),
            bottom: border(8, model::BorderStyle::Nil),
            left: None,
            right: None,
            inside_h: None,
            inside_v: None,
            tl2br: None,
            tr2bl: None,
        }));

        let built = build_offered(&t, OFFERED, &[("T", style)]);
        let b = built.border_config.expect("the style declares six edges");
        assert_pt(
            b.top.expect("the direct top").width,
            2.0,
            "16 eighths of a point, from the table",
        );
        for (edge, name) in [
            (b.bottom, "bottom"),
            (b.left, "left"),
            (b.right, "right"),
            (b.inside_v, "insideV"),
        ] {
            assert_pt(
                edge.unwrap_or_else(|| panic!("{name} must survive from the style"))
                    .width,
                1.0,
                &format!("{name}: 8 eighths, from the style"),
            );
        }
        assert!(
            b.inside_h.is_none(),
            "a direct insideH of `none` must remove the style's, not inherit it"
        );

        let cell = built.rows[0].cells[0]
            .cell_borders
            .as_ref()
            .expect("the cell declares two edges");
        assert!(
            cell.top.is_none(),
            "the same `none` on a *cell* means inherit — §17.4.66, not §17.4.38"
        );
        assert!(
            matches!(cell.bottom, Some(CellBorderOverride::Suppress)),
            "…and `nil` is the cell spelling that does suppress, got {:?}",
            cell.bottom
        );
    }

    /// §17.4.42 / §17.4.68: two per-side cascades in series. The table's
    /// default is the direct `<w:tblCellMar>` over the style's, and the cell's
    /// margins are `<w:tcMar>` over *that*.
    ///
    /// One side stated at each level, with a different value at each, is what
    /// separates them: a merge that replaced either set wholesale, or that
    /// resolved the cell against the style rather than against the merged
    /// default, lands a different number on three of the four sides.
    ///
    /// The table level cannot tell an absent side from an explicit zero —
    /// `<w:tblCellMar>` parses to a whole `EdgeInsets`, missing sides collapsed
    /// to 0 because §17.4.42 gives a table default no further inheritance — so
    /// the merge reads 0 as "unstated". That is a limit of the model rather
    /// than a choice made here, and its one visible consequence is that
    /// `<w:left w:w="0"/>` on a `<w:tbl>` cannot override a style's non-zero
    /// left. Stated, not asserted.
    #[test]
    fn cell_margins_cascade_per_side_from_the_style_through_the_table_to_the_cell() {
        let style = plain_style(model::TableProperties {
            cell_margins: model::Dup::from(Some(cell_mar(100, 200, 300, 400))),
            ..Default::default()
        });
        let props = model::TableProperties {
            style_id: Some(model::StyleId::new("T")),
            width: measure(dxa(4000)),
            // Only `left` is stated; the other three are the parse seam's zero.
            cell_margins: model::Dup::from(Some(cell_mar(0, 0, 0, 800))),
            ..Default::default()
        };
        let mut t = table_of(props, &[4000], 1);
        t.rows[0].cells[0].properties.margins =
            model::Dup::from(Some(crate::model::geometry::PartialEdgeInsets::new(
                Some(Dimension::new(600)),
                None,
                None,
                None,
            )));

        let built = build_offered(&t, OFFERED, &[("T", style)]);
        let m = built.rows[0].cells[0].margins;
        assert_pt(m.top, 30.0, "the cell's own 600 twips");
        assert_pt(
            m.right,
            10.0,
            "the style's 200, restated at neither level below",
        );
        assert_pt(m.bottom, 15.0, "the style's 300, likewise");
        assert_pt(m.left, 40.0, "the table's 800 over the style's 400");
    }

    /// §17.4.57: which of `<w:tblpPr>`'s ten attributes reach layout.
    ///
    /// Five do — `rightFromText` and `bottomFromText` as the two wrap gaps,
    /// `tblpY` as the offset, `vertAnchor`, and `tblpXSpec`. Two more are
    /// dropped *structurally*: `TableFloatInfo` has no field for `horzAnchor`
    /// or `tblpYSpec` at all, so nothing here could carry them. The remaining
    /// three — `tblpX`, `leftFromText`, `topFromText` — are parsed and
    /// deliberately not carried, which `build_table` states at the site: a
    /// horizontally-offset table places as if none of them were set.
    ///
    /// Each dropped attribute is given a value no honoured field could be
    /// mistaken for, so a wire crossed between `left`/`right`, `top`/`bottom`
    /// or `x`/`y` shows up as a wrong number rather than as a missing one.
    #[test]
    fn five_of_tbl_p_prs_ten_attributes_reach_the_float_info() {
        let props = model::TableProperties {
            width: measure(dxa(4000)),
            positioning: model::Dup::from(Some(model::TablePositioning {
                left_from_text: Some(Dimension::new(9999)),
                right_from_text: Some(Dimension::new(100)),
                top_from_text: Some(Dimension::new(9999)),
                bottom_from_text: Some(Dimension::new(200)),
                vert_anchor: Some(model::TableAnchor::Page),
                horz_anchor: Some(model::TableAnchor::Margin),
                x_align: Some(model::TableXAlign::Right),
                y_align: Some(model::TableYAlign::Bottom),
                x: Some(Dimension::new(9999)),
                y: Some(Dimension::new(400)),
            })),
            overlap: model::Dup::from(Some(model::TableOverlap::Never)),
            ..Default::default()
        };
        let f = build(&table_of(props, &[4000], 1))
            .float_info
            .expect("a tblpPr floats the table");

        assert_pt(
            f.right_gap,
            5.0,
            "rightFromText's 100 twips, not leftFromText",
        );
        assert_pt(f.bottom_gap, 10.0, "bottomFromText's 200, not topFromText");
        assert_pt(f.y_offset, 20.0, "tblpY's 400, not tblpX");
        assert_eq!(f.vert_anchor, model::TableAnchor::Page);
        assert_eq!(f.x_align, Some(model::TableXAlign::Right));
        assert_eq!(
            f.overlap,
            Some(model::TableOverlap::Never),
            "§17.4.56 rides along, from its own element"
        );
    }

    /// §17.4.57: an empty `<w:tblpPr>` still floats the table, at the
    /// element's own defaults — no wrap gaps, no offset, and the `text`
    /// vertical anchor the section names as the default.
    ///
    /// And `<w:tblOverlap>` alone does **not** float one. §17.4.56 says how a
    /// float behaves when it meets another float, which is not a statement that
    /// this table is one — so a table carrying only `tblOverlap` must come back
    /// with no `float_info` at all, or the paginator would start treating an
    /// ordinary table as anchored.
    #[test]
    fn an_empty_tbl_p_pr_floats_at_the_defaults_and_tbl_overlap_alone_does_not() {
        let floated = model::TableProperties {
            width: measure(dxa(4000)),
            positioning: model::Dup::from(Some(model::TablePositioning {
                left_from_text: None,
                right_from_text: None,
                top_from_text: None,
                bottom_from_text: None,
                vert_anchor: None,
                horz_anchor: None,
                x_align: None,
                y_align: None,
                x: None,
                y: None,
            })),
            ..Default::default()
        };
        let f = build(&table_of(floated, &[4000], 1))
            .float_info
            .expect("an empty tblpPr still floats the table");
        assert_pt(f.right_gap, 0.0, "no rightFromText is no gap");
        assert_pt(f.bottom_gap, 0.0, "no bottomFromText is no gap");
        assert_pt(f.y_offset, 0.0, "no tblpY is no offset");
        assert_eq!(
            f.vert_anchor,
            model::TableAnchor::Text,
            "§17.4.57's default vertical anchor"
        );
        assert_eq!(f.x_align, None);

        let overlap_only = model::TableProperties {
            width: measure(dxa(4000)),
            overlap: model::Dup::from(Some(model::TableOverlap::Never)),
            ..Default::default()
        };
        assert!(
            build(&table_of(overlap_only, &[4000], 1))
                .float_info
                .is_none(),
            "tblOverlap describes a float; it does not create one"
        );
    }

    /// The width a nested table is built against is its cell's **inner**
    /// width: the grid slot less the cell margins that will be drawn inside it.
    ///
    /// 8000 twips of `tblW` is 400 pt, the one-column grid takes all of it, and
    /// the 108-twip left and right cell margins leave 389.2 pt of content. A
    /// nested table at `pct 2500` is half of that, 194.6 pt.
    ///
    /// Note what each half of that pins. **That the offered width is the
    /// cell's** is arithmetic, and is the assertion. **That the percentage is
    /// then taken against it** is a divergence from §17.4.63, which defines the
    /// `pct` base as the text extents of the *page* — `build_table` records why
    /// it is not read literally (a nested table at 100% would be drawn straight
    /// out of its cell, which no implementation does) and that a Word reference
    /// render is what would settle it. This test moves with that decision; it
    /// is not evidence for it.
    #[test]
    fn a_nested_table_is_offered_its_cells_inner_width() {
        let nested = table_of(
            model::TableProperties {
                width: measure(pct(2500)),
                ..Default::default()
            },
            &[1000],
            1,
        );
        let outer = model::TableProperties {
            width: measure(dxa(8000)),
            cell_margins: model::Dup::from(Some(cell_mar(0, 108, 0, 108))),
            ..Default::default()
        };
        let mut t = table_of(outer, &[8000], 1);
        t.rows[0].cells[0].content = vec![Block::Table(Box::new(nested))];

        let built = build(&t);
        assert_pt(built.col_widths[0], 400.0, "8000 twips of declared width");
        match &built.rows[0].cells[0].blocks[0] {
            LayoutBlock::Table { col_widths, .. } => assert_pt(
                total(col_widths),
                194.6,
                "half of 400 − 2 × 5.4 pt of cell margin",
            ),
            _ => panic!("the cell's first block is not the nested table"),
        }
    }

    /// §17.4.60: a row's `<w:tblPrEx><w:tblBorders>` is a per-edge exception
    /// over the table's *effective* borders, merged **at the model layer** —
    /// before `none` has been collapsed into "draws nothing".
    ///
    /// The order is the decision at the site. An `Option<Border>` whose style
    /// is `None` is a real element that wins the merge; converting each side to
    /// layout first would turn it into "this side was not specified" and let
    /// the table's border through underneath it. So an exception naming one
    /// edge leaves the other five alone, and one spelling an edge `none`
    /// removes it.
    ///
    /// Nothing under `tests/` writes a `<w:tblPrEx>` — the element appears in
    /// `table_border_conflict.rs`'s prose and in no fixture — so this path had
    /// no coverage at any level.
    #[test]
    fn a_rows_tbl_pr_ex_borders_are_a_per_edge_exception_over_the_tables() {
        let props = model::TableProperties {
            width: measure(dxa(4000)),
            borders: model::Dup::from(Some(uniform_borders(8))),
            ..Default::default()
        };
        let mut t = table_of(props, &[2000, 2000], 2);
        t.rows.push(model_row(0, &[1, 1]));
        t.rows[1].property_exceptions = Some(model::TableRowPropertyExceptions {
            borders: Some(model::TableBorders {
                top: border(16, model::BorderStyle::Single),
                inside_v: border(8, model::BorderStyle::None),
                bottom: None,
                left: None,
                right: None,
                inside_h: None,
            }),
            cell_spacing: None,
        });

        let built = build(&t);
        assert!(
            built.rows[0].border_overrides.is_none(),
            "a row with no exception carries none — it inherits the table's"
        );
        let o = built.rows[1]
            .border_overrides
            .as_ref()
            .expect("the exception must reach the row");
        assert_pt(
            o.top.expect("the exception's top").width,
            2.0,
            "16 eighths, from the exception",
        );
        assert_pt(
            o.bottom.expect("the table's bottom").width,
            1.0,
            "8 eighths, from the table the exception never mentions",
        );
        assert!(
            o.inside_v.is_none(),
            "an exception's `none` removes the table's edge rather than reading \
             as 'unspecified'"
        );
    }

    /// §17.4.43 / §17.4.44: a row-level `tblCellSpacing` is **not applied** —
    /// `resolve_cell_spacing` sets out the three answers that would be needed
    /// first — and it is *reported* rather than dropped in silence.
    ///
    /// What is reported is a disagreement, not the presence of a row-level
    /// element. A row restating the table's own spacing resolves to the same
    /// gap, so warning about it would report a difference that is not there:
    /// that is why `issue-165-cellspacing.docx` (400 twips at both levels) is
    /// silent while `issue-165-cellspacing-scale.docx` (400 and 800) is not.
    /// The two are indistinguishable on the page — the geometry is the table's
    /// value either way, which `tests/table_cell_spacing.rs` pins — so only
    /// `BuildState` can tell them apart.
    ///
    /// The `pct` case is the same rule from the other end. §17.4.45 leaves
    /// every measure but `dxa` without a usable value here, so a `pct` row
    /// spacing against a table that declares none resolves to the same zero and
    /// is not a disagreement either.
    #[test]
    fn a_row_cell_spacing_is_reported_only_when_it_resolves_to_a_different_gap() {
        let warned = |table: Option<model::TableMeasure>,
                      row: Option<model::TableMeasure>,
                      exception: Option<model::TableMeasure>| {
            let props = model::TableProperties {
                width: measure(dxa(4000)),
                cell_spacing: model::Dup::from(table),
                ..Default::default()
            };
            let mut t = table_of(props, &[2000, 2000], 2);
            t.rows[0].properties.cell_spacing = model::Dup::from(row);
            if exception.is_some() {
                t.rows[0].property_exceptions = Some(model::TableRowPropertyExceptions {
                    borders: None,
                    cell_spacing: exception,
                });
            }
            build_reporting(&t).1.warned_row_cell_spacing
        };

        assert!(
            !warned(Some(dxa(240)), None, None),
            "a row that says nothing is not a disagreement"
        );
        assert!(
            !warned(Some(dxa(240)), Some(dxa(240)), None),
            "…nor is a row restating the table's own value"
        );
        assert!(
            !warned(None, Some(pct(2500)), None),
            "…nor a row whose measure §17.4.45 leaves without a value, which \
             resolves to the same zero the table has"
        );
        assert!(
            warned(Some(dxa(240)), Some(dxa(480)), None),
            "a row asking for a different gap must be reported"
        );
        assert!(
            warned(Some(dxa(240)), None, Some(dxa(480))),
            "and §17.4.44's `tblPrEx` is read when the `trPr` is silent"
        );
    }

    /// §17.4.66: a table cell must end with a `<w:p>`, so a producer whose cell
    /// ends with a table writes an empty paragraph after it purely to satisfy
    /// the content model. Word gives that terminator no height, and it is
    /// dropped here rather than laid out — otherwise every nested table in a
    /// cell would be followed by a blank line the author never wrote.
    ///
    /// Four conjuncts decide it, and each gets a case that differs in exactly
    /// one of them: the paragraph must be empty, must follow a **table**, must
    /// be the cell's **last** block, and must not be its **first**. Drop any one
    /// and a paragraph that should have been laid out disappears — or, for the
    /// first-block guard, the index behind it goes negative.
    #[test]
    fn only_an_empty_last_paragraph_after_a_table_is_dropped_as_structural() {
        let blocks = |content: Vec<Block>| {
            let mut t = table_of(
                model::TableProperties {
                    width: measure(dxa(4000)),
                    ..Default::default()
                },
                &[4000],
                1,
            );
            t.rows[0].cells[0].content = content;
            build(&t).rows[0].cells[0].blocks.len()
        };
        let nested = || {
            Block::Table(Box::new(table_of(
                model::TableProperties::default(),
                &[1000],
                1,
            )))
        };

        assert_eq!(
            blocks(vec![nested(), para(vec![])]),
            1,
            "the terminator goes"
        );
        assert_eq!(
            blocks(vec![nested(), para(vec![text_run("x")])]),
            2,
            "a paragraph with content is not a terminator"
        );
        assert_eq!(
            blocks(vec![nested(), para(vec![]), para(vec![text_run("x")])]),
            3,
            "…and an empty one that is not last is not one either"
        );
        assert_eq!(
            blocks(vec![para(vec![]), nested()]),
            2,
            "the cell's first block is never a terminator — nothing precedes it"
        );
        assert_eq!(
            blocks(vec![para(vec![text_run("x")]), para(vec![])]),
            2,
            "an empty last paragraph after a *paragraph* is the author's"
        );
    }

    /// §17.7.6: a cell's own shading outranks the conditional layer's, and its
    /// borders outrank them **per edge** — a `<w:tcBorders>` naming only some
    /// sides takes the layer's for the rest rather than replacing it wholesale.
    ///
    /// `wholeTable` is the layer every cell gets whatever `tblLook` says, which
    /// is what makes it the right one to state a *priority* with: no region
    /// flag has to resolve correctly for the fixture to be about priority.
    #[test]
    fn a_cells_own_shading_and_borders_outrank_the_conditional_layer_per_edge() {
        let fill = |rgb: u32| {
            model::Dup::from(Some(model::Shading {
                fill: model::Color::Rgb(rgb),
                pattern: model::ShadingPattern::Clear,
                color: model::Color::Auto,
            }))
        };
        let layer = model::TableStyleOverride {
            override_type: model::TableStyleOverrideType::WholeTable,
            paragraph_properties: None,
            run_properties: None,
            table_properties: None,
            table_row_properties: None,
            table_cell_properties: Some(model::TableCellProperties {
                shading: fill(0x00FF00),
                borders: model::Dup::from(Some(uniform_cell_borders(8))),
                ..Default::default()
            }),
        };
        let props = model::TableProperties {
            style_id: Some(model::StyleId::new("T")),
            width: measure(dxa(4000)),
            ..Default::default()
        };
        let mut t = table_of(props, &[2000, 2000], 2);
        t.rows[0].cells[0].properties.shading = fill(0xFF0000);
        t.rows[0].cells[0].properties.borders = model::Dup::from(Some(model::TableCellBorders {
            top: border(16, model::BorderStyle::Single),
            bottom: None,
            left: None,
            right: None,
            inside_h: None,
            inside_v: None,
            tl2br: None,
            tr2bl: None,
        }));

        let built = build_offered(
            &t,
            OFFERED,
            &[(
                "T",
                layered_style(model::TableProperties::default(), vec![layer]),
            )],
        );
        let cells = &built.rows[0].cells;
        let rgb = |r, g, b| Some(crate::render::resolve::color::RgbColor { r, g, b });
        assert_eq!(cells[0].shading, rgb(255, 0, 0), "the cell's own fill wins");
        assert_eq!(
            cells[1].shading,
            rgb(0, 255, 0),
            "…and a cell that declares none takes the layer's"
        );

        let edges = |i: usize| cells[i].cell_borders.as_ref().expect("borders resolved");
        assert_pt(override_width(&edges(0).top), 2.0, "the cell's own top");
        assert_pt(
            override_width(&edges(0).bottom),
            1.0,
            "and the layer's bottom, which the cell never named",
        );
        assert_pt(
            override_width(&edges(1).top),
            1.0,
            "a cell with no borders of its own is the layer's throughout",
        );
    }

    /// §17.4.83: `w:vAlign` is direct, then the conditional layer, then `top` —
    /// and `both` lands on `top` as well, which is a capability boundary rather
    /// than a reading of the section. See the site for why.
    ///
    /// The third cell is what makes that observable: under a layer saying
    /// `center`, a cell asking for `both` must come out `top`. Were `both`
    /// treated as "no request" it would inherit the layer's `center` instead,
    /// and were it honoured it would be neither.
    #[test]
    fn v_align_is_direct_then_conditional_then_top_and_both_is_top() {
        let layer = model::TableStyleOverride {
            override_type: model::TableStyleOverrideType::WholeTable,
            paragraph_properties: None,
            run_properties: None,
            table_properties: None,
            table_row_properties: None,
            table_cell_properties: Some(model::TableCellProperties {
                vertical_align: model::Dup::from(Some(model::CellVerticalAlign::Center)),
                ..Default::default()
            }),
        };
        let props = model::TableProperties {
            style_id: Some(model::StyleId::new("T")),
            width: measure(dxa(6000)),
            ..Default::default()
        };
        let mut t = table_of(props, &[2000, 2000, 2000], 3);
        t.rows[0].cells[0].properties.vertical_align =
            model::Dup::from(Some(model::CellVerticalAlign::Bottom));
        t.rows[0].cells[2].properties.vertical_align =
            model::Dup::from(Some(model::CellVerticalAlign::Both));

        let built = build_offered(
            &t,
            OFFERED,
            &[(
                "T",
                layered_style(model::TableProperties::default(), vec![layer]),
            )],
        );
        let cells = &built.rows[0].cells;
        assert_eq!(
            cells[0].vertical_align,
            CellVAlign::Bottom,
            "the cell's own"
        );
        assert_eq!(cells[1].vertical_align, CellVAlign::Center, "the layer's");
        assert_eq!(
            cells[2].vertical_align,
            CellVAlign::Top,
            "`both` is vertical justification, which is not implemented — it \
             must take `top` rather than fall through to the layer"
        );

        let plain = build(&table_of(
            model::TableProperties {
                width: measure(dxa(4000)),
                ..Default::default()
            },
            &[4000],
            1,
        ));
        assert_eq!(
            plain.rows[0].cells[0].vertical_align,
            CellVAlign::Top,
            "§17.4.83's default, with nothing declared anywhere"
        );
    }

    /// §17.4.80: `hRule="auto"` ignores `val`, so the row carries **no** height
    /// constraint — it is not a zero-height minimum, and not a minimum of the
    /// stated value either.
    ///
    /// Word writes `val="0"` alongside `hRule="auto"` ([MS-OE376] §2.4.77(c)),
    /// which is why treating `auto` as `AtLeast(val)` was invisible on Word
    /// output: `AtLeast(0)` constrains nothing. Other producers are not bound by
    /// that, and a non-zero `val` under `auto` would have become a real minimum.
    #[test]
    fn auto_height_rule_carries_no_constraint() {
        use crate::model::{HeightRule, TableRowHeight};
        use crate::render::layout::table::RowHeightRule;
        let to_layout = |rule, twips: i64| -> Option<RowHeightRule> {
            row_height_rule(TableRowHeight {
                value: crate::model::dimension::Dimension::new(twips),
                rule,
            })
        };
        assert!(
            to_layout(HeightRule::Auto, 1000).is_none(),
            "auto ignores val entirely"
        );
        assert!(matches!(
            to_layout(HeightRule::AtLeast, 1000),
            Some(RowHeightRule::AtLeast(_))
        ));
        assert!(matches!(
            to_layout(HeightRule::Exact, 1000),
            Some(RowHeightRule::Exact(_))
        ));
    }
}
