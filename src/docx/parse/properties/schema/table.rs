//! Table property schemas: `<w:tblPr>`, `<w:trPr>`, `<w:tcPr>`.
//!
//! `TblPrXml::split` returns `(TableProperties, Option<StyleId>)` — the style
//! id travels separately for cascade reasons, matching the legacy parser's
//! signature.

use crate::model::Dup;
use serde::{Deserialize, Deserializer};

use crate::docx::model::dimension::Twips;
use crate::docx::model::{
    Alignment, CnfStyle, StyleId, TableCellProperties, TableLook, TablePositioning,
    TableProperties, TableRowHeight, TableRowProperties, VerticalMerge,
};
use crate::docx::parse::primitives::st_enums::{
    StAnchor, StHeightRule, StJc, StTblLayoutType, StTblOverlap, StTextDirection, StVerticalJc,
    StXAlign, StYAlign,
};
use crate::docx::parse::primitives::units::deserialize_optional_nonnegative_dimension;
use crate::docx::parse::primitives::{last_toggle, OnOff};
use crate::docx::parse::serde_xml::UnknownChildren;

use super::border::{TableBordersXml, TableCellBordersXml};
use super::cnf_style::CnfStyleXml;
use super::insets::EdgeInsetsTwipsXml;
use super::measure::{deserialize_vec_nonnegative_table_measure, TableMeasureXml};
use super::shading::ShdXml;

// ── tblPr ───────────────────────────────────────────────────────────────

#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct TblPrXml {
    #[serde(rename = "tblStyle", default)]
    tbl_style: Vec<ValString>,
    #[serde(rename = "tblBorders", default)]
    tbl_borders: Vec<TableBordersXml>,
    #[serde(rename = "tblCellMar", default)]
    tbl_cell_mar: Vec<EdgeInsetsTwipsXml>,
    #[serde(rename = "jc", default)]
    jc: Vec<ValAttr<StJc>>,
    #[serde(
        rename = "tblW",
        default,
        deserialize_with = "deserialize_vec_nonnegative_table_measure"
    )]
    tbl_w: Vec<TableMeasureXml>,
    #[serde(rename = "tblLayout", default)]
    tbl_layout: Vec<TblLayoutXml>,
    #[serde(rename = "tblInd", default)]
    tbl_ind: Vec<TableMeasureXml>,
    #[serde(
        rename = "tblCellSpacing",
        default,
        deserialize_with = "deserialize_vec_nonnegative_table_measure"
    )]
    tbl_cell_spacing: Vec<TableMeasureXml>,
    #[serde(rename = "tblLook", default)]
    tbl_look: Vec<TblLookXml>,
    #[serde(rename = "tblStyleRowBandSize", default)]
    tbl_style_row_band_size: Vec<ValAttr<u32>>,
    #[serde(rename = "tblStyleColBandSize", default)]
    tbl_style_col_band_size: Vec<ValAttr<u32>>,
    #[serde(rename = "tblpPr", default)]
    tblp_pr: Vec<TblpPrXml>,
    #[serde(rename = "tblOverlap", default)]
    tbl_overlap: Vec<ValAttr<StTblOverlap>>,
    /// §17.4.1 `<w:bidiVisual/>` — the table's columns run right to left, so
    /// the first cell of a row is the rightmost one. `CT_OnOff`, and
    /// `Vec<OnOff>` for the same last-wins reason as every other toggle here.
    #[serde(rename = "bidiVisual", default)]
    bidi_visual: Vec<OnOff>,
    /// Children this schema does not name — recorded so an unimplemented
    /// table property is visible under `RUST_LOG=warn` instead of vanishing.
    /// See [`UnknownChildren`].
    #[serde(rename = "$value", default)]
    unknown: UnknownChildren,
}

/// `<w:tblLayout w:type="fixed"/>` — note `@type` (not `@val`).
#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct TblLayoutXml {
    #[serde(rename = "@type")]
    ty: StTblLayoutType,
}

/// `<w:tblLook>` — supports both the modern explicit attributes (firstRow,
/// lastRow, ...) and the legacy hex bitfield on `@val`. Per
/// [MS-OI29500] §2.1.1583, when both are present the explicit attribute
/// wins per-flag; otherwise the bitfield supplies the value.
#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct TblLookXml {
    #[serde(rename = "@val", default)]
    val: Option<TblLookHex>,
    #[serde(rename = "@firstRow", default)]
    first_row: Option<AttrBool>,
    #[serde(rename = "@lastRow", default)]
    last_row: Option<AttrBool>,
    #[serde(rename = "@firstColumn", default)]
    first_column: Option<AttrBool>,
    #[serde(rename = "@lastColumn", default)]
    last_column: Option<AttrBool>,
    #[serde(rename = "@noHBand", default)]
    no_h_band: Option<AttrBool>,
    #[serde(rename = "@noVBand", default)]
    no_v_band: Option<AttrBool>,
}

/// Word's legacy `<w:tblLook val>` hex bitfield, per [MS-OI29500] §2.1.1583.
///
/// Bit positions:
/// | Mask     | Flag        |
/// |----------|-------------|
/// | `0x0020` | firstRow    |
/// | `0x0040` | lastRow     |
/// | `0x0080` | firstColumn |
/// | `0x0100` | lastColumn  |
/// | `0x0200` | noHBand     |
/// | `0x0400` | noVBand     |
///
/// Other bits are reserved/ignored. The Word default `04A0` =
/// firstRow + firstColumn + noVBand.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct TblLookHex(u16);

impl TblLookHex {
    const FIRST_ROW: u16 = 0x0020;
    const LAST_ROW: u16 = 0x0040;
    const FIRST_COLUMN: u16 = 0x0080;
    const LAST_COLUMN: u16 = 0x0100;
    const NO_H_BAND: u16 = 0x0200;
    const NO_V_BAND: u16 = 0x0400;

    fn first_row(self) -> bool {
        self.0 & Self::FIRST_ROW != 0
    }
    fn last_row(self) -> bool {
        self.0 & Self::LAST_ROW != 0
    }
    fn first_column(self) -> bool {
        self.0 & Self::FIRST_COLUMN != 0
    }
    fn last_column(self) -> bool {
        self.0 & Self::LAST_COLUMN != 0
    }
    fn no_h_band(self) -> bool {
        self.0 & Self::NO_H_BAND != 0
    }
    fn no_v_band(self) -> bool {
        self.0 & Self::NO_V_BAND != 0
    }
}

impl<'de> Deserialize<'de> for TblLookHex {
    fn deserialize<D: Deserializer<'de>>(d: D) -> Result<Self, D::Error> {
        let s = String::deserialize(d)?;
        u16::from_str_radix(s.trim_start_matches("0x"), 16)
            .map(TblLookHex)
            .map_err(serde::de::Error::custom)
    }
}

/// `<w:tblpPr>` — floating table positioning.
#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct TblpPrXml {
    #[serde(
        rename = "@leftFromText",
        default,
        deserialize_with = "deserialize_optional_nonnegative_dimension"
    )]
    left_from_text: Option<crate::docx::model::dimension::Dimension<Twips>>,
    #[serde(
        rename = "@rightFromText",
        default,
        deserialize_with = "deserialize_optional_nonnegative_dimension"
    )]
    right_from_text: Option<crate::docx::model::dimension::Dimension<Twips>>,
    #[serde(
        rename = "@topFromText",
        default,
        deserialize_with = "deserialize_optional_nonnegative_dimension"
    )]
    top_from_text: Option<crate::docx::model::dimension::Dimension<Twips>>,
    #[serde(
        rename = "@bottomFromText",
        default,
        deserialize_with = "deserialize_optional_nonnegative_dimension"
    )]
    bottom_from_text: Option<crate::docx::model::dimension::Dimension<Twips>>,
    #[serde(rename = "@vertAnchor", default)]
    vert_anchor: Option<StAnchor>,
    #[serde(rename = "@horzAnchor", default)]
    horz_anchor: Option<StAnchor>,
    #[serde(rename = "@tblpXSpec", default)]
    x_spec: Option<StXAlign>,
    #[serde(rename = "@tblpYSpec", default)]
    y_spec: Option<StYAlign>,
    #[serde(rename = "@tblpX", default)]
    x: Option<crate::docx::model::dimension::Dimension<Twips>>,
    #[serde(rename = "@tblpY", default)]
    y: Option<crate::docx::model::dimension::Dimension<Twips>>,
}

/// §17.4.60 `<w:tblPrEx>` — table-level property exceptions scoped to
/// a single row. Per the spec it accepts the same vocabulary as
/// `<w:tblPr>` minus `tblStyle` and `tblpPr`. We model only the slice
/// the layout currently honors (table borders); other fields can be
/// added incrementally.
#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct TblPrExXml {
    #[serde(rename = "tblBorders", default)]
    tbl_borders: Vec<TableBordersXml>,
    /// §17.4.44: per-row override of the table's `tblCellSpacing`.
    #[serde(
        rename = "tblCellSpacing",
        default,
        deserialize_with = "deserialize_vec_nonnegative_table_measure"
    )]
    tbl_cell_spacing: Vec<TableMeasureXml>,
    /// §17.4.1 on a row: this row's columns run right to left, independently of
    /// the table's own `w:bidiVisual`.
    ///
    /// Modelled so it is *visible* — an unnamed child vanishes into
    /// [`UnknownChildren`] and only ever surfaces under `RUST_LOG=warn` — but
    /// nothing acts on it yet. See `TableRowPropertyExceptions::bidi_visual`.
    #[serde(rename = "bidiVisual", default)]
    bidi_visual: Vec<OnOff>,
    /// Children this schema does not name — recorded so an unimplemented
    /// table property is visible under `RUST_LOG=warn` instead of vanishing.
    /// See [`UnknownChildren`].
    #[serde(rename = "$value", default)]
    unknown: UnknownChildren,
}

impl From<TblPrExXml> for crate::docx::model::TableRowPropertyExceptions {
    fn from(x: TblPrExXml) -> Self {
        x.unknown.warn_once("w:tblPrEx");
        Self {
            borders: Dup::from(x.tbl_borders).into_value().map(Into::into),
            cell_spacing: Dup::from(x.tbl_cell_spacing).into_value().map(Into::into),
            bidi_visual: last_toggle(x.bidi_visual),
        }
    }
}

impl TblPrXml {
    pub(crate) fn split(self) -> (TableProperties, Option<StyleId>) {
        self.unknown.warn_once("w:tblPr");
        let style_id = Dup::from(self.tbl_style)
            .into_value()
            .map(|v| StyleId::new(v.val));
        let props = TableProperties {
            style_id: style_id.clone(),
            alignment: Dup::from(self.jc).map(|v| Alignment::from(v.val)),
            width: Dup::from(self.tbl_w).map(Into::into),
            layout: Dup::from(self.tbl_layout).map(|v| crate::docx::model::TableLayout::from(v.ty)),
            indent: Dup::from(self.tbl_ind).map(Into::into),
            borders: Dup::from(self.tbl_borders).map(Into::into),
            cell_margins: Dup::from(self.tbl_cell_mar).map(Into::into),
            cell_spacing: Dup::from(self.tbl_cell_spacing).map(Into::into),
            look: Dup::from(self.tbl_look).filter_map(tbl_look),
            style_row_band_size: Dup::from(self.tbl_style_row_band_size).map(|v| v.val),
            style_col_band_size: Dup::from(self.tbl_style_col_band_size).map(|v| v.val),
            positioning: Dup::from(self.tblp_pr).map(Into::into),
            overlap: Dup::from(self.tbl_overlap)
                .map(|v| crate::docx::model::TableOverlap::from(v.val)),
            bidi_visual: Dup::from(self.bidi_visual).map(|OnOff(on)| on),
        };
        (props, style_id)
    }
}

/// §17.4.55 `<w:tblLook>` → the model, or `None` when the element states
/// nothing at all.
///
/// The legacy `@val` bitmask and the six modern attributes are **not** two
/// spellings of the same six flags to be merged flag by flag.
///
/// [MS-OI29500] Part 1 §17.4.55 note (c) is exact about it: "Word reads
/// the val attribute if, and only if, none of the attributes specified in
/// this subsection are present." So one modern attribute anywhere on the
/// element makes the whole bitmask unread — a per-flag fallback would make
/// that sentence say nothing, and would let `val`'s *cleared* bits switch
/// regions off that the document never mentioned.
///
/// # What an unmentioned sibling then means is a choice, not a rule
///
/// The erratum settles only which source is read. It does not say what an
/// attribute the element omits resolves to once `val` is out of the
/// picture, and neither does §17.4.55: `CT_TblLook`'s attributes are
/// optional `ST_OnOff` with no schema default.
///
/// **The choice taken here is `true` — every region on.** The evidence is
/// second-implementation: LibreOffice tested Word for tdf#167843 and
/// concluded that all unspecified attributes default to true, and its
/// regression fixture (`<w:tblLook w:val="04A0" w:firstRow="0"/>`, which
/// `tbl_pr_tbl_look_val_04a0_with_first_row_off` reproduces below) asserts
/// firstColumn, lastColumn and lastRow all on. That is testimony about
/// Word, not documentation of it.
///
/// It is applied to all six uniformly, including `noHBand`/`noVBand`,
/// where `true` switches banding *off* rather than on. Splitting the rule
/// — four flags one way, two the other — would be a second choice with no
/// evidence behind it at all, and a uniform "unset bits read as set" is
/// also the simplest thing an implementation holding a bitmask would do.
///
/// **What would settle it**: a Word render of a table whose style defines
/// a `band1Horz` layer, carrying `<w:tblLook w:firstRow="0"/>` and nothing
/// else. Banding paints iff Word's unspecified `noHBand` is false, which
/// is the one flag where the uniform reading and the "regions on" reading
/// disagree.
///
/// # Why an entirely silent element yields `None` rather than a silent value
///
/// `<w:tblLook/>` states no `@val` and no attribute, so there is nothing to
/// read from: it carries exactly what the *absent* element carries, which is
/// nothing. §17.4.55 note (a)'s absent-element default is the consumer's
/// question — `render::resolve::conditional::ActiveRegions::WORD_DEFAULT`
/// answers it — and the consumer can only answer it correctly if the two
/// spellings of "nothing" arrive here looking alike.
///
/// A `TableLook` with six `None` flags *looked* alike and was not, because the
/// carrier is a `crate::model::Dup`: any value at all reports the element as
/// **present**, and "this level did not set the property" (§17.7.2) is
/// precisely `Dup::is_absent`. So a silent `<w:tblLook/>` on a `<w:tbl>` beat
/// the table style's `tblLook` and replaced it with the absent-element
/// default, and a silent one in a child style beat its `basedOn` parent's the
/// same way — each switching off conditional layers the surviving level had
/// switched on. Dropping the occurrence is what makes the two spellings
/// interchangeable at *every* level of the cascade, rather than at whichever
/// read site remembered to ask.
fn tbl_look(x: TblLookXml) -> Option<TableLook> {
    let stated = [
        x.first_row,
        x.last_row,
        x.first_column,
        x.last_column,
        x.no_h_band,
        x.no_v_band,
    ];
    if stated.iter().all(Option::is_none) {
        // Nothing modern present, so `val` is read — and if it is absent too
        // the element states nothing at all.
        return x.val.map(|v| TableLook {
            first_row: Some(v.first_row()),
            last_row: Some(v.last_row()),
            first_column: Some(v.first_column()),
            last_column: Some(v.last_column()),
            no_h_band: Some(v.no_h_band()),
            no_v_band: Some(v.no_v_band()),
        });
    }
    // `is_none_or` *is* the rule: an unstated attribute reads as true.
    let attr = |a: Option<AttrBool>| Some(a.is_none_or(|b| b.0));
    Some(TableLook {
        first_row: attr(x.first_row),
        last_row: attr(x.last_row),
        first_column: attr(x.first_column),
        last_column: attr(x.last_column),
        no_h_band: attr(x.no_h_band),
        no_v_band: attr(x.no_v_band),
    })
}

impl From<TblpPrXml> for TablePositioning {
    fn from(x: TblpPrXml) -> Self {
        Self {
            left_from_text: x.left_from_text,
            right_from_text: x.right_from_text,
            top_from_text: x.top_from_text,
            bottom_from_text: x.bottom_from_text,
            vert_anchor: x.vert_anchor.map(Into::into),
            horz_anchor: x.horz_anchor.map(Into::into),
            x_align: x.x_spec.map(Into::into),
            y_align: x.y_spec.map(Into::into),
            x: x.x,
            y: x.y,
        }
    }
}

// ── trPr ───────────────────────────────────────────────────────────────

#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct TrPrXml {
    #[serde(rename = "trHeight", default)]
    tr_height: Vec<TrHeightXml>,
    // `Vec<OnOff>` (not `Option`) tolerates duplicated toggles per §17.7.2
    // last-wins — see `RPrXml` / `PPrXml` for the rationale.
    #[serde(rename = "tblHeader", default)]
    tbl_header: Vec<OnOff>,
    #[serde(rename = "cantSplit", default)]
    cant_split: Vec<OnOff>,
    #[serde(rename = "jc", default)]
    jc: Vec<ValAttr<StJc>>,
    #[serde(rename = "cnfStyle", default)]
    cnf_style: Vec<CnfStyleXml>,
    #[serde(rename = "gridBefore", default)]
    grid_before: Vec<ValAttr<u32>>,
    #[serde(
        rename = "wBefore",
        default,
        deserialize_with = "deserialize_vec_nonnegative_table_measure"
    )]
    w_before: Vec<TableMeasureXml>,
    #[serde(rename = "gridAfter", default)]
    grid_after: Vec<ValAttr<u32>>,
    #[serde(
        rename = "wAfter",
        default,
        deserialize_with = "deserialize_vec_nonnegative_table_measure"
    )]
    w_after: Vec<TableMeasureXml>,
    /// §17.4.43: row-level override of the table's `tblCellSpacing`.
    #[serde(
        rename = "tblCellSpacing",
        default,
        deserialize_with = "deserialize_vec_nonnegative_table_measure"
    )]
    tbl_cell_spacing: Vec<TableMeasureXml>,
    /// `<w:del>` — CT_TrPr's row-deletion marker, and the form **Word** writes
    /// when a row is deleted with change tracking on. Only its presence is
    /// read (see [`TrPrXml::marks_row_deleted`]); the `w:id`/`w:author`/
    /// `w:date` attributes describe the edit, not the document, and this
    /// parser renders the final view rather than the revision history.
    ///
    /// `Vec` for the same reason as the toggles above — a repeated child must
    /// not fail deserialization.
    #[serde(rename = "del", default)]
    del: Vec<RowRevisionMarkerXml>,
    /// Children this schema does not name — recorded so an unimplemented
    /// table property is visible under `RUST_LOG=warn` instead of vanishing.
    /// See [`UnknownChildren`].
    #[serde(rename = "$value", default)]
    unknown: UnknownChildren,
}

/// CT_TrackChange as it appears in `<w:trPr>` — presence-only, so no field is
/// modelled. Distinct from body-level `IgnoredXml` so this schema does not
/// depend on the body schema, which depends on it.
#[derive(Clone, Debug, Deserialize)]
pub(crate) struct RowRevisionMarkerXml {}

impl TrPrXml {
    /// Whether this row is marked deleted by `<w:trPr><w:del/>`.
    ///
    /// Read at the parse seam rather than carried on `TableRowProperties`,
    /// because that is where the same question is answered for runs: a
    /// `<w:del>`-wrapped run never reaches the model either. The model
    /// describes the document, not the edits that produced it.
    pub(crate) fn marks_row_deleted(&self) -> bool {
        !self.del.is_empty()
    }
}

#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct TrHeightXml {
    #[serde(
        rename = "@val",
        default,
        deserialize_with = "deserialize_optional_nonnegative_dimension"
    )]
    val: Option<crate::docx::model::dimension::Dimension<Twips>>,
    #[serde(rename = "@hRule", default)]
    rule: Option<StHeightRule>,
}

impl From<TrHeightXml> for TableRowHeight {
    fn from(x: TrHeightXml) -> Self {
        Self {
            value: x.val.unwrap_or_default(),
            // §17.4.80 says an omitted `hRule` means `auto`; [MS-OI29500]
            // §17.4.80(a) records that **Word assumes `atLeast`**, and Word is
            // what produced these files. Defaulting to `Auto` here would be
            // indistinguishable from an explicit `hRule="auto"`, where the
            // standard says the `val` is ignored — so the two must not collapse
            // to the same variant.
            rule: x
                .rule
                .map(Into::into)
                .unwrap_or(crate::docx::model::HeightRule::AtLeast),
        }
    }
}

impl From<TrPrXml> for TableRowProperties {
    fn from(x: TrPrXml) -> Self {
        x.unknown.warn_once("w:trPr");
        Self {
            height: Dup::from(x.tr_height).map(Into::into),
            is_header: last_toggle(x.tbl_header),
            cant_split: last_toggle(x.cant_split),
            justification: Dup::from(x.jc).map(|v| Alignment::from(v.val)),
            cnf_style: Dup::from(x.cnf_style).map(CnfStyle::from),
            grid_before: Dup::from(x.grid_before)
                .into_value()
                .map(|v| v.val)
                .unwrap_or(0),
            w_before: Dup::from(x.w_before).map(Into::into),
            grid_after: Dup::from(x.grid_after)
                .into_value()
                .map(|v| v.val)
                .unwrap_or(0),
            w_after: Dup::from(x.w_after).map(Into::into),
            cell_spacing: Dup::from(x.tbl_cell_spacing).map(Into::into),
        }
    }
}

// ── tcPr ───────────────────────────────────────────────────────────────

/// Table cell property bag (§17.4.69 `w:tcPr`).
///
/// Child properties are deserialized as `Vec<T>` to tolerate duplicate XML elements
/// (such as repeated `<w:tcMar>` or `<w:tcBorders>` emitted by Word/LibreOffice)
/// without failing deserialization. Per OOXML §17.7.2, the last occurrence wins.
#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct TcPrXml {
    #[serde(rename = "tcBorders", default)]
    tc_borders: Vec<TableCellBordersXml>,
    #[serde(rename = "tcMar", default)]
    tc_mar: Vec<EdgeInsetsTwipsXml>,
    #[serde(
        rename = "tcW",
        default,
        deserialize_with = "deserialize_vec_nonnegative_table_measure"
    )]
    tc_w: Vec<TableMeasureXml>,
    #[serde(rename = "shd", default)]
    shd: Vec<ShdXml>,
    #[serde(rename = "vAlign", default)]
    v_align: Vec<ValAttr<StVerticalJc>>,
    #[serde(rename = "vMerge", default)]
    v_merge: Vec<VMergeXml>,
    #[serde(rename = "gridSpan", default)]
    grid_span: Vec<ValAttr<u32>>,
    #[serde(rename = "textDirection", default)]
    text_direction: Vec<ValAttr<StTextDirection>>,
    #[serde(rename = "noWrap", default)]
    no_wrap: Vec<OnOff>,
    #[serde(rename = "cnfStyle", default)]
    cnf_style: Vec<CnfStyleXml>,
    /// Children this schema does not name — recorded so an unimplemented
    /// table property is visible under `RUST_LOG=warn` instead of vanishing.
    /// See [`UnknownChildren`].
    #[serde(rename = "$value", default)]
    unknown: UnknownChildren,
}

/// `<w:vMerge/>` — absent `@val` means "continue"; `@val="restart"` starts a
/// new vertical merge group; `@val="continue"` is explicit continue.
#[derive(Clone, Copy, Debug, Deserialize)]
pub(crate) struct VMergeXml {
    #[serde(rename = "@val", default)]
    val: Option<VMergeKind>,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
enum VMergeKind {
    Restart,
    Continue,
}

impl From<VMergeXml> for VerticalMerge {
    fn from(x: VMergeXml) -> Self {
        match x.val {
            Some(VMergeKind::Restart) => Self::Restart,
            Some(VMergeKind::Continue) | None => Self::Continue,
        }
    }
}

impl From<TcPrXml> for TableCellProperties {
    /// Every duplicable child is carried into the model whole; `Dup::get`
    /// applies last-wins where a consumer reads it. See `model::dup`.
    fn from(x: TcPrXml) -> Self {
        x.unknown.warn_once("w:tcPr");
        Self {
            width: x.tc_w.into_iter().map(Into::into).collect(),
            borders: x.tc_borders.into_iter().map(Into::into).collect(),
            shading: x.shd.into_iter().map(Into::into).collect(),
            margins: x.tc_mar.into_iter().map(Into::into).collect(),
            vertical_align: x
                .v_align
                .into_iter()
                .map(|v| crate::docx::model::CellVerticalAlign::from(v.val))
                .collect(),
            vertical_merge: x.v_merge.into_iter().map(Into::into).collect(),
            grid_span: x.grid_span.into_iter().map(|v| v.val).collect(),
            text_direction: x
                .text_direction
                .into_iter()
                .map(|v| crate::docx::model::TextDirection::from(v.val))
                .collect(),
            no_wrap: last_toggle(x.no_wrap),
            cnf_style: x.cnf_style.into_iter().map(CnfStyle::from).collect(),
        }
    }
}

// ── shared helpers ──────────────────────────────────────────────────────

#[derive(Clone, Debug, Deserialize)]
struct ValString {
    #[serde(rename = "@val")]
    val: String,
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(bound(deserialize = "T: serde::Deserialize<'de>"))]
pub(crate) struct ValAttr<T> {
    #[serde(rename = "@val")]
    val: T,
}

use crate::docx::parse::primitives::AttrBool;

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::model::{
        BorderStyle, CellVerticalAlign, HeightRule, TableLayout, TableMeasure, TableOverlap,
        TextDirection,
    };

    // ── tblPr ──

    fn parse_tbl_pr(xml: &str) -> (TableProperties, Option<StyleId>) {
        let x: TblPrXml = quick_xml::de::from_str(xml).unwrap();
        x.split()
    }

    #[test]
    fn tbl_pr_style_and_width() {
        let (tp, sid) = parse_tbl_pr(
            r#"<tblPr><tblStyle val="TableGrid"/><tblW w="5000" type="pct"/></tblPr>"#,
        );
        assert_eq!(
            sid.map(|s| s.as_str().to_string()),
            Some("TableGrid".into())
        );
        match tp.width.cloned().unwrap() {
            TableMeasure::Pct(d) => assert_eq!(d.raw(), 5000),
            other => panic!("expected Pct, got {other:?}"),
        }
    }

    #[test]
    fn negative_decimal_table_width_is_rejected() {
        let parsed: Result<TblPrXml, _> =
            quick_xml::de::from_str(r#"<tblPr><tblW w="-1.5" type="dxa"/></tblPr>"#);
        assert!(parsed.is_err(), "negative table widths must be rejected");
    }

    #[test]
    fn tbl_pr_layout_and_alignment() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><jc val="center"/><tblLayout type="fixed"/></tblPr>"#);
        assert_eq!(tp.layout, Dup::from(Some(TableLayout::Fixed)));
        assert_eq!(tp.alignment, Dup::from(Some(Alignment::Center)));
    }

    /// §17.18.87: the *other* value of `ST_TblLayoutType`, and the only one
    /// that names the auto-fit algorithm — `<w:tblLayout w:type="fixed"/>` was
    /// the sole spelling the corpus contained, so this half went untested and
    /// unparseable together. See `StTblLayoutType`.
    #[test]
    fn tbl_pr_layout_autofit() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLayout type="autofit"/></tblPr>"#);
        assert_eq!(tp.layout, Dup::from(Some(TableLayout::Autofit)));
    }

    #[test]
    fn tbl_pr_borders_and_margins() {
        let (tp, _) = parse_tbl_pr(
            r#"<tblPr>
                <tblBorders><top val="single"/><left val="double"/></tblBorders>
                <tblCellMar><top w="100"/><left w="80"/></tblCellMar>
            </tblPr>"#,
        );
        let b = tp.borders.cloned().unwrap();
        assert_eq!(b.top.unwrap().style, BorderStyle::Single);
        assert_eq!(b.left.unwrap().style, BorderStyle::Double);
        assert_eq!(tp.cell_margins.cloned().unwrap().top.raw(), 100);
    }

    #[test]
    fn tbl_pr_tbl_look_attrs() {
        let (tp, _) =
            parse_tbl_pr(r#"<tblPr><tblLook firstRow="1" lastRow="0" noHBand="true"/></tblPr>"#);
        let l = tp.look.cloned().unwrap();
        assert_eq!(l.first_row, Some(true));
        assert_eq!(l.last_row, Some(false));
        assert_eq!(l.no_h_band, Some(true));
        // The three the element never mentions are still answered — see
        // `tbl_look` for why, and on what evidence.
        assert_eq!(l.first_column, Some(true));
        assert_eq!(l.last_column, Some(true));
        assert_eq!(l.no_v_band, Some(true));
    }

    /// [MS-OI29500] §2.1.1583: legacy `@val` hex bitfield must decode to the
    /// same flags as the modern explicit attributes. Bit positions:
    /// 0x0020 firstRow, 0x0040 lastRow, 0x0080 firstColumn, 0x0100 lastColumn,
    /// 0x0200 noHBand, 0x0400 noVBand. The Word-default `04A0` =
    /// firstRow + firstColumn + noVBand.
    #[test]
    fn tbl_pr_tbl_look_legacy_val_default() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLook val="04A0"/></tblPr>"#);
        let l = tp.look.cloned().unwrap();
        assert_eq!(l.first_row, Some(true));
        assert_eq!(l.last_row, Some(false));
        assert_eq!(l.first_column, Some(true));
        assert_eq!(l.last_column, Some(false));
        assert_eq!(l.no_h_band, Some(false));
        assert_eq!(l.no_v_band, Some(true));
    }

    /// Sample1's ITEM/NEEDED table uses `val="0620"` =
    /// firstRow + noHBand + noVBand. Without legacy decoding, banding is
    /// erroneously enabled and `band1Horz` CF paints inner borders.
    #[test]
    fn tbl_pr_tbl_look_legacy_val_suppresses_banding() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLook val="0620"/></tblPr>"#);
        let l = tp.look.cloned().unwrap();
        assert_eq!(l.first_row, Some(true));
        assert_eq!(l.no_h_band, Some(true));
        assert_eq!(l.no_v_band, Some(true));
    }

    #[test]
    fn tbl_pr_tbl_look_legacy_val_zero_clears_all() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLook val="0000"/></tblPr>"#);
        let l = tp.look.cloned().unwrap();
        assert_eq!(l.first_row, Some(false));
        assert_eq!(l.last_row, Some(false));
        assert_eq!(l.first_column, Some(false));
        assert_eq!(l.last_column, Some(false));
        assert_eq!(l.no_h_band, Some(false));
        assert_eq!(l.no_v_band, Some(false));
    }

    /// [MS-OI29500] Part 1 §17.4.55 note (c): "Word reads the val attribute
    /// if, and only if, **none** of the attributes specified in this
    /// subsection are present." One modern attribute therefore suppresses
    /// `val` for *every* flag, not merely for its own — which is the whole
    /// content of the rule, since a per-flag reading would make the sentence
    /// say nothing.
    ///
    /// `val="0000"` clears all six bits, so under a per-flag fallback the
    /// four unmentioned flags would arrive as `Some(false)`.
    #[test]
    fn tbl_pr_tbl_look_one_explicit_attr_suppresses_val_for_every_flag() {
        let (tp, _) =
            parse_tbl_pr(r#"<tblPr><tblLook val="0000" firstRow="1" noVBand="1"/></tblPr>"#);
        let l = tp.look.cloned().unwrap();
        assert_eq!(l.first_row, Some(true), "explicit firstRow=1");
        assert_eq!(l.no_v_band, Some(true), "explicit noVBand=1");
        assert_eq!(l.last_row, Some(true), "not from val=0000");
        assert_eq!(l.first_column, Some(true), "not from val=0000");
        assert_eq!(l.last_column, Some(true), "not from val=0000");
        assert_eq!(l.no_h_band, Some(true), "not from val=0000");
    }

    /// LibreOffice's tdf#167843 regression fixture, verbatim: `val="04A0"`
    /// alongside a single `firstRow="0"`. 04A0 clears lastRow (0x040) and
    /// §17.4.1 on a *row* (`w:tblPrEx`) is a separate element from the one on
    /// the table, and this is the only thing that currently observes it.
    ///
    /// Nothing acts on it — `TableRowPropertyExceptions::bidi_visual` says why
    /// — so without this test the field could be renamed, misspelled or dropped
    /// and every render would agree. That is exactly the failure mode modelling
    /// it was meant to end: an unnamed child vanishes silently.
    #[test]
    fn tbl_pr_ex_bidi_visual_is_modelled_rather_than_unknown() {
        let parse = |xml: &str| -> (crate::docx::model::TableRowPropertyExceptions, Vec<String>) {
            let x: TblPrExXml = quick_xml::de::from_str(xml).unwrap();
            let unknown: Vec<String> = x.unknown.names().iter().map(|n| n.to_string()).collect();
            (x.into(), unknown)
        };

        let (ex, unknown) = parse("<tblPrEx><bidiVisual/></tblPrEx>");
        assert_eq!(ex.bidi_visual, Some(true));
        assert!(
            unknown.is_empty(),
            "a modelled child must not also be recorded as unknown: {unknown:?}"
        );

        let (off, _) = parse(r#"<tblPrEx><bidiVisual val="0"/></tblPrEx>"#);
        assert_eq!(off.bidi_visual, Some(false), "an explicit off is stated");

        let (absent, _) = parse("<tblPrEx/>");
        assert_eq!(absent.bidi_visual, None);
    }

    /// §17.4.1 `w:bidiVisual` is a `CT_OnOff`, so all four of its states have to
    /// be distinguishable at this seam — and the two that are *not* "on" are the
    /// ones the render tests cannot see.
    ///
    /// A render test can only tell a mirrored table from an unmirrored one, so
    /// it reads `Some(false)` and `None` alike as "did not mirror" and would
    /// pass with either wired to the other. Only the parse layer can say that an
    /// explicit `w:val="0"` is a *stated* off — which matters as soon as the
    /// value has anywhere to inherit from, and is the difference `Dup` exists to
    /// carry.
    #[test]
    fn tbl_pr_bidi_visual_is_a_toggle_with_four_states() {
        let read = |xml: &str| parse_tbl_pr(xml).0.bidi_visual.cloned();

        assert_eq!(read("<tblPr/>"), None, "absent states nothing");
        assert_eq!(
            read("<tblPr><bidiVisual/></tblPr>"),
            Some(true),
            "bare is on"
        );
        assert_eq!(
            read(r#"<tblPr><bidiVisual val="0"/></tblPr>"#),
            Some(false),
            "an explicit off is stated, not absent"
        );
        // §17.7.2 last-wins, which is why the field is a `Vec` and not an
        // `Option` — the same rule every other toggle in this file follows.
        assert_eq!(
            read(r#"<tblPr><bidiVisual/><bidiVisual val="0"/></tblPr>"#),
            Some(false),
            "the last occurrence wins"
        );
        assert_eq!(
            read(r#"<tblPr><bidiVisual val="0"/><bidiVisual/></tblPr>"#),
            Some(true),
            "…in both directions"
        );
    }

    /// lastColumn (0x100) and clears noHBand (0x200) — none of which may
    /// reach the model, because note (c) says the whole bitmask is unread.
    #[test]
    fn tbl_pr_tbl_look_val_04a0_with_first_row_off() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLook val="04A0" firstRow="0"/></tblPr>"#);
        let l = tp.look.cloned().unwrap();
        assert_eq!(l.first_row, Some(false));
        assert_eq!(l.last_row, Some(true), "04A0's cleared lastRow is not read");
        assert_eq!(l.first_column, Some(true));
        assert_eq!(l.last_column, Some(true), "nor its cleared lastColumn");
        assert_eq!(l.no_h_band, Some(true), "nor its cleared noHBand");
        assert_eq!(l.no_v_band, Some(true));
    }

    /// A `<w:tblLook/>` that states nothing at all states nothing: `val` is
    /// absent too, so there is no bitmask to read and no attribute to fill
    /// in from. What an entirely silent element means is §17.4.55 note (a)'s
    /// question, and it is answered by the resolver
    /// (`render::resolve::conditional::ActiveRegions::WORD_DEFAULT`), not here.
    ///
    /// "States nothing" has to mean **absent from the cascade**, not "present
    /// with six unstated flags": `Dup::is_absent` is what §17.7.2's "this
    /// level did not set the property" reads, so a value here — however empty
    /// — shadows every level below. This asserted the six `None`s until that
    /// consequence was measured; see `tbl_look`.
    #[test]
    fn tbl_pr_tbl_look_empty_element_states_nothing() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLook/></tblPr>"#);
        assert!(
            tp.look.is_absent(),
            "an element stating nothing must not occupy the cascade slot"
        );
    }

    /// The converse, and the guard on the line the check above draws: an
    /// element stating *one* attribute is a value, and every flag it does not
    /// mention is answered rather than left to the level below.
    #[test]
    fn tbl_pr_tbl_look_one_stated_attr_is_still_a_value() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLook lastRow="1"/></tblPr>"#);
        let l = tp.look.cloned().expect("one stated attribute is a value");
        assert_eq!(l.last_row, Some(true));
        assert_eq!(l.first_row, Some(true), "unstated reads as true");
    }

    /// …and so is a bare `@val`, whose bitmask answers all six.
    #[test]
    fn tbl_pr_tbl_look_a_bare_val_is_still_a_value() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLook val="0000"/></tblPr>"#);
        assert!(!tp.look.is_absent(), "a bitmask states all six flags");
    }

    #[test]
    fn tbl_pr_tbl_look_legacy_val_lowercase() {
        let (tp, _) = parse_tbl_pr(r#"<tblPr><tblLook val="04a0"/></tblPr>"#);
        let l = tp.look.cloned().unwrap();
        assert_eq!(l.first_row, Some(true));
        assert_eq!(l.no_v_band, Some(true));
    }

    #[test]
    fn tbl_pr_overlap_and_positioning() {
        let (tp, _) = parse_tbl_pr(
            r#"<tblPr>
                <tblOverlap val="never"/>
                <tblpPr tblpX="100" tblpY="200" vertAnchor="page"
                        horzAnchor="margin" tblpXSpec="center"/>
            </tblPr>"#,
        );
        assert_eq!(tp.overlap, Dup::from(Some(TableOverlap::Never)));
        let pos = tp.positioning.cloned().unwrap();
        assert_eq!(pos.x.unwrap().raw(), 100);
        assert_eq!(pos.y.unwrap().raw(), 200);
        assert_eq!(pos.vert_anchor, Some(crate::docx::model::TableAnchor::Page));
        assert_eq!(pos.x_align, Some(crate::docx::model::TableXAlign::Center));
    }

    // ── trPr ──

    fn parse_tr_pr(xml: &str) -> TableRowProperties {
        let x: TrPrXml = quick_xml::de::from_str(xml).unwrap();
        x.into()
    }

    #[test]
    fn tr_pr_height_with_rule() {
        let tr = parse_tr_pr(r#"<trPr><trHeight val="440" hRule="atLeast"/></trPr>"#);
        let h = tr.height.cloned().unwrap();
        assert_eq!(h.value.raw(), 440);
        assert_eq!(h.rule, HeightRule::AtLeast);
    }

    #[test]
    fn tr_pr_is_header_and_cant_split() {
        let tr = parse_tr_pr(r#"<trPr><tblHeader/><cantSplit/></trPr>"#);
        assert_eq!(tr.is_header, Some(true));
        assert_eq!(tr.cant_split, Some(true));
    }

    #[test]
    fn tr_pr_duplicate_toggles_tolerated_last_wins() {
        // Duplicated row toggles (as some writers emit) must not fail the parse.
        let tr = parse_tr_pr(r#"<trPr><cantSplit/><cantSplit/></trPr>"#);
        assert_eq!(tr.cant_split, Some(true));
        let tr = parse_tr_pr(r#"<trPr><tblHeader val="1"/><tblHeader val="0"/></trPr>"#);
        assert_eq!(tr.is_header, Some(false));
    }

    #[test]
    fn tc_pr_duplicate_no_wrap_tolerated() {
        let tc = parse_tc_pr(r#"<tcPr><noWrap/><noWrap/></tcPr>"#);
        assert_eq!(tc.no_wrap, Some(true));
    }

    #[test]
    fn tr_pr_grid_after_and_w_after() {
        let tr = parse_tr_pr(r#"<trPr><gridAfter val="2"/><wAfter w="500" type="dxa"/></trPr>"#);
        assert_eq!(tr.grid_after, 2);
        match tr.w_after.cloned().unwrap() {
            TableMeasure::Twips(d) => assert_eq!(d.raw(), 500),
            other => panic!("expected Twips, got {other:?}"),
        }
    }

    #[test]
    fn tr_pr_grid_before_and_w_before() {
        let tr = parse_tr_pr(r#"<trPr><gridBefore val="1"/><wBefore w="38" type="dxa"/></trPr>"#);
        assert_eq!(tr.grid_before, 1);
        match tr.w_before.cloned().unwrap() {
            TableMeasure::Twips(d) => assert_eq!(d.raw(), 38),
            other => panic!("expected Twips, got {other:?}"),
        }
    }

    #[test]
    fn tr_pr_grid_before_and_grid_after_default_zero() {
        let tr = parse_tr_pr(r#"<trPr/>"#);
        assert_eq!(tr.grid_before, 0);
        assert_eq!(tr.grid_after, 0);
        assert!(tr.w_before.is_absent());
        assert!(tr.w_after.is_absent());
    }

    // ── tcPr ──

    fn parse_tc_pr(xml: &str) -> TableCellProperties {
        let x: TcPrXml = quick_xml::de::from_str(xml).unwrap();
        x.into()
    }

    #[test]
    fn tc_pr_width_and_borders() {
        let tc = parse_tc_pr(
            r#"<tcPr>
                <tcW w="2500" type="dxa"/>
                <tcBorders><top val="single"/><tl2br val="dotted"/></tcBorders>
            </tcPr>"#,
        );
        match tc.width.cloned().unwrap() {
            TableMeasure::Twips(d) => assert_eq!(d.raw(), 2500),
            other => panic!("expected Twips, got {other:?}"),
        }
        assert!(tc.borders.cloned().unwrap().tl2br.is_some());
    }

    #[test]
    fn tc_pr_vertical_align() {
        let tc = parse_tc_pr(r#"<tcPr><vAlign val="center"/></tcPr>"#);
        assert_eq!(
            tc.vertical_align,
            Dup::from(Some(CellVerticalAlign::Center))
        );
    }

    #[test]
    fn tc_pr_v_merge_restart_and_continue() {
        let tc = parse_tc_pr(r#"<tcPr><vMerge val="restart"/></tcPr>"#);
        assert_eq!(tc.vertical_merge, Dup::from(Some(VerticalMerge::Restart)));

        let tc = parse_tc_pr(r#"<tcPr><vMerge/></tcPr>"#);
        assert_eq!(tc.vertical_merge, Dup::from(Some(VerticalMerge::Continue)));
    }

    #[test]
    fn tc_pr_grid_span_and_text_direction() {
        let tc = parse_tc_pr(r#"<tcPr><gridSpan val="3"/><textDirection val="tbRl"/></tcPr>"#);
        assert_eq!(tc.grid_span, Dup::from(Some(3)));
        assert_eq!(
            tc.text_direction,
            Dup::from(Some(TextDirection::TopToBottomRightToLeft))
        );
    }

    #[test]
    fn tc_pr_no_wrap_and_cnf_style() {
        let tc = parse_tc_pr(r#"<tcPr><noWrap/><cnfStyle val="100000000000"/></tcPr>"#);
        assert_eq!(tc.no_wrap, Some(true));
        assert_eq!(tc.cnf_style, Dup::from(Some(CnfStyle::FIRST_ROW)));
    }

    /// §17.4.80 vs [MS-OI29500] §17.4.80(a). The standard says an omitted
    /// `hRule` means `auto`; Word assumes `atLeast`, and Word wrote these files.
    ///
    /// The distinction is load-bearing rather than cosmetic: with `hRule="auto"`
    /// the standard says `val` is **ignored**, so collapsing "omitted" into
    /// `Auto` would make every Word row with a `trHeight` lose its minimum
    /// height — or, as before this change, force an explicit `auto` to be
    /// treated as a minimum it should not have.
    #[test]
    fn omitted_hrule_is_at_least_but_explicit_auto_is_auto() {
        use crate::docx::model::HeightRule;
        let parse = |xml: &str| -> crate::docx::model::TableRowHeight {
            quick_xml::de::from_str::<TrHeightXml>(xml).unwrap().into()
        };
        assert_eq!(parse(r#"<trHeight val="440"/>"#).rule, HeightRule::AtLeast);
        assert_eq!(
            parse(r#"<trHeight val="440" hRule="auto"/>"#).rule,
            HeightRule::Auto
        );
        assert_eq!(
            parse(r#"<trHeight val="440" hRule="atLeast"/>"#).rule,
            HeightRule::AtLeast
        );
        assert_eq!(
            parse(r#"<trHeight val="440" hRule="exact"/>"#).rule,
            HeightRule::Exact
        );
        // The value survives in every case; only the rule decides its meaning.
        assert_eq!(parse(r#"<trHeight val="440"/>"#).value.raw(), 440);
    }

    /// §17.4.44 / §17.4.43: both cell-spacing overrides now reach the model.
    /// Layout applies spacing per table and warns rather than honouring these,
    /// but dropping them at the parser would make that gap invisible.
    #[test]
    fn cell_spacing_overrides_reach_the_model() {
        let tr: crate::docx::model::TableRowProperties = quick_xml::de::from_str::<TrPrXml>(
            r#"<trPr><tblCellSpacing w="72" type="dxa"/></trPr>"#,
        )
        .unwrap()
        .into();
        assert!(tr.cell_spacing.cloned().is_some(), "row-level §17.4.43");

        let ex: crate::docx::model::TableRowPropertyExceptions =
            quick_xml::de::from_str::<TblPrExXml>(
                r#"<tblPrEx><tblCellSpacing w="72" type="dxa"/></tblPrEx>"#,
            )
            .unwrap()
            .into();
        assert!(ex.cell_spacing.is_some(), "tblPrEx §17.4.44");
    }

    /// A duplicated **non-toggle** child is schema-invalid and Word opens it
    /// anyway; see `primitives::duplicates` for why the last one wins.
    #[test]
    fn tbl_pr_duplicate_non_toggle_children_are_tolerated_last_wins() {
        let (tbl, _) = parse_tbl_pr(
            r#"<tblPr>
                 <jc val="left"/><jc val="center"/>
                 <tblW w="1000" type="dxa"/><tblW w="5000" type="dxa"/>
               </tblPr>"#,
        );
        assert_eq!(
            tbl.alignment.get(),
            Some(&Alignment::Center),
            "\u{a7}17.4.29"
        );
        assert!(
            matches!(tbl.width.get(), Some(TableMeasure::Twips(d)) if d.raw() == 5000),
            "\u{a7}17.4.63, got {:?}",
            tbl.width
        );
        // Both occurrences reach the model; only the read resolves.
        assert_eq!(tbl.alignment.all().len(), 2);
        assert_eq!(tbl.width.all().len(), 2);
    }

    #[test]
    fn tr_pr_duplicate_non_toggle_children_are_tolerated_last_wins() {
        let tr = parse_tr_pr(
            r#"<trPr>
                 <trHeight val="200"/><trHeight val="500"/>
               </trPr>"#,
        );
        assert_eq!(
            tr.height.get().map(|h| h.value.raw()),
            Some(500),
            "\u{a7}17.4.81"
        );
        assert_eq!(tr.height.all().len(), 2, "both occurrences reach the model");
    }

    #[test]
    fn tc_pr_duplicate_tc_mar_does_not_fail() {
        let tc = parse_tc_pr(
            r#"<tcPr>
                <tcMar><top w="100" type="dxa"/></tcMar>
                <tcMar><bottom w="200" type="dxa"/></tcMar>
            </tcPr>"#,
        );
        assert!(!tc.margins.is_absent());
        assert_eq!(
            tc.margins.get().unwrap().bottom.map(|d| d.raw()),
            Some(200),
            "last occurrence wins at the point of use"
        );
    }

    /// The point of `Dup`: the parser discards nothing, so a consumer that
    /// wants a different rule than last-wins can still have one.
    #[test]
    fn both_occurrences_reach_the_model() {
        let tc = parse_tc_pr(
            r#"<tcPr>
                <tcMar><top w="100" type="dxa"/></tcMar>
                <tcMar><bottom w="200" type="dxa"/></tcMar>
            </tcPr>"#,
        );
        assert!(tc.margins.is_duplicated(), "the document repeated w:tcMar");
        assert_eq!(tc.margins.all().len(), 2, "neither occurrence was dropped");
        // first-wins is reachable downstream without touching the parser
        assert_eq!(
            tc.margins.all().first().unwrap().top.map(|d| d.raw()),
            Some(100)
        );
        assert_eq!(tc.margins.get().unwrap().bottom.map(|d| d.raw()), Some(200));
    }

    #[test]
    fn an_unrepeated_child_is_not_flagged_as_duplicated() {
        let tc = parse_tc_pr(r#"<tcPr><tcMar><top w="100" type="dxa"/></tcMar></tcPr>"#);
        assert!(!tc.margins.is_duplicated());
        assert_eq!(tc.margins.all().len(), 1);
    }

    // ── unmodelled children are reported, not dropped ──
    //
    // A plain-struct property bag silently discards an element it does not
    // name, which is why `w:hMerge` and `w:bidiVisual` were invisible at
    // runtime rather than merely unimplemented. These pin both directions:
    // an unnamed child is captured by name, and a named one never is.

    #[test]
    fn tbl_pr_records_an_unmodelled_child() {
        // `w:tblCaption` (§17.4.60) stands in for `w:bidiVisual`, which used to
        // be the example here and is now modelled — the pairing this pins is
        // "unnamed is captured, named is not", so it needs a child the schema
        // still does not name.
        let x: TblPrXml = quick_xml::de::from_str(
            r#"<tblPr><tblStyle val="Grid"/><tblCaption val="c"/></tblPr>"#,
        )
        .unwrap();
        assert_eq!(x.unknown.names(), ["tblCaption"]);
        // The modelled sibling is unaffected by the catch-all.
        assert_eq!(
            x.split().1.map(|s| s.as_str().to_string()),
            Some("Grid".into())
        );
    }

    #[test]
    fn tc_pr_records_an_unmodelled_child() {
        let x: TcPrXml =
            quick_xml::de::from_str(r#"<tcPr><gridSpan val="2"/><hMerge val="restart"/></tcPr>"#)
                .unwrap();
        assert_eq!(x.unknown.names(), ["hMerge"]);
        assert_eq!(
            TableCellProperties::from(x).grid_span.get().copied(),
            Some(2)
        );
    }

    #[test]
    fn tr_pr_records_an_unmodelled_child() {
        let x: TrPrXml =
            quick_xml::de::from_str(r#"<trPr><cantSplit/><divId val="7"/></trPr>"#).unwrap();
        assert_eq!(x.unknown.names(), ["divId"]);
        assert_eq!(TableRowProperties::from(x).cant_split, Some(true));
    }

    #[test]
    fn tbl_pr_ex_records_an_unmodelled_child() {
        let x: TblPrExXml =
            quick_xml::de::from_str(r#"<tblPrEx><tblLayout type="fixed"/></tblPrEx>"#).unwrap();
        assert_eq!(
            x.unknown.names(),
            ["tblLayout"],
            "tblPrEx models only borders and cell spacing today, so its own \
             tblLayout is unmodelled and must say so"
        );
    }

    /// The trap detector for the three tests above: if a *modelled* child ever
    /// reached the catch-all, every real document would log a warning for a
    /// property that is in fact implemented, and the report would be noise.
    #[test]
    fn a_bag_of_only_modelled_children_reports_nothing() {
        let tbl: TblPrXml = quick_xml::de::from_str(
            r#"<tblPr><tblStyle val="G"/><tblW w="5000" type="pct"/><jc val="center"/>
               <tblLayout type="fixed"/><tblInd w="100" type="dxa"/>
               <tblCellSpacing w="40" type="dxa"/><tblLook val="04A0"/>
               <tblStyleRowBandSize val="1"/><tblStyleColBandSize val="2"/>
               <tblOverlap val="never"/></tblPr>"#,
        )
        .unwrap();
        assert!(
            tbl.unknown.names().is_empty(),
            "tblPr: {:?}",
            tbl.unknown.names()
        );

        let tr: TrPrXml = quick_xml::de::from_str(
            r#"<trPr><trHeight val="300" hRule="atLeast"/><tblHeader/><cantSplit/>
               <jc val="center"/><gridBefore val="1"/><gridAfter val="1"/>
               <wBefore w="10" type="dxa"/><wAfter w="10" type="dxa"/>
               <tblCellSpacing w="40" type="dxa"/></trPr>"#,
        )
        .unwrap();
        assert!(
            tr.unknown.names().is_empty(),
            "trPr: {:?}",
            tr.unknown.names()
        );

        let tc: TcPrXml = quick_xml::de::from_str(
            r#"<tcPr><tcW w="100" type="dxa"/><shd fill="FF0000"/><vAlign val="center"/>
               <vMerge val="restart"/><gridSpan val="2"/><noWrap/>
               <textDirection val="tbRl"/><tcMar><top w="1" type="dxa"/></tcMar></tcPr>"#,
        )
        .unwrap();
        assert!(
            tc.unknown.names().is_empty(),
            "tcPr: {:?}",
            tc.unknown.names()
        );

        let ex: TblPrExXml =
            quick_xml::de::from_str(r#"<tblPrEx><tblCellSpacing w="40" type="dxa"/></tblPrEx>"#)
                .unwrap();
        assert!(
            ex.unknown.names().is_empty(),
            "tblPrEx: {:?}",
            ex.unknown.names()
        );
    }
}
