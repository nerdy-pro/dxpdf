//! `<w:cnfStyle>` (§17.3.1.8) — conditional formatting flags for table cells
//! and paragraphs inside tables.
//!
//! Two wire encodings are supported per spec:
//!
//! 1. **Binary string on `@val`** — a 12-character `0`/`1` string, MSB-first,
//!    matching the `CnfStyle` bitflag order (first_row, last_row, …).
//! 2. **Individual boolean attributes** — `@firstRow="1"`, `@oddHBand="true"`,
//!    etc. These are merged on top of the `@val` bitmask (if both are present).

use serde::Deserialize;

use crate::docx::model::CnfStyle;

#[derive(Clone, Debug, Default, Deserialize)]
pub(crate) struct CnfStyleXml {
    #[serde(rename = "@val", default)]
    val: Option<String>,
    #[serde(rename = "@firstRow", default)]
    first_row: Option<OnOffAttr>,
    #[serde(rename = "@lastRow", default)]
    last_row: Option<OnOffAttr>,
    #[serde(rename = "@firstColumn", default)]
    first_column: Option<OnOffAttr>,
    #[serde(rename = "@lastColumn", default)]
    last_column: Option<OnOffAttr>,
    #[serde(rename = "@oddVBand", default)]
    odd_v_band: Option<OnOffAttr>,
    #[serde(rename = "@evenVBand", default)]
    even_v_band: Option<OnOffAttr>,
    #[serde(rename = "@oddHBand", default)]
    odd_h_band: Option<OnOffAttr>,
    #[serde(rename = "@evenHBand", default)]
    even_h_band: Option<OnOffAttr>,
    #[serde(rename = "@firstRowFirstColumn", default)]
    first_row_first_column: Option<OnOffAttr>,
    #[serde(rename = "@firstRowLastColumn", default)]
    first_row_last_column: Option<OnOffAttr>,
    #[serde(rename = "@lastRowFirstColumn", default)]
    last_row_first_column: Option<OnOffAttr>,
    #[serde(rename = "@lastRowLastColumn", default)]
    last_row_last_column: Option<OnOffAttr>,
}

use crate::docx::parse::primitives::AttrBool as OnOffAttr;

impl From<CnfStyleXml> for CnfStyle {
    fn from(x: CnfStyleXml) -> Self {
        let mut bits = match x.val.as_deref() {
            Some(s) => parse_cnf_val(s),
            None => CnfStyle::empty(),
        };
        set(&mut bits, CnfStyle::FIRST_ROW, x.first_row);
        set(&mut bits, CnfStyle::LAST_ROW, x.last_row);
        set(&mut bits, CnfStyle::FIRST_COLUMN, x.first_column);
        set(&mut bits, CnfStyle::LAST_COLUMN, x.last_column);
        set(&mut bits, CnfStyle::ODD_V_BAND, x.odd_v_band);
        set(&mut bits, CnfStyle::EVEN_V_BAND, x.even_v_band);
        set(&mut bits, CnfStyle::ODD_H_BAND, x.odd_h_band);
        set(&mut bits, CnfStyle::EVEN_H_BAND, x.even_h_band);
        set(
            &mut bits,
            CnfStyle::FIRST_ROW_FIRST_COLUMN,
            x.first_row_first_column,
        );
        set(
            &mut bits,
            CnfStyle::FIRST_ROW_LAST_COLUMN,
            x.first_row_last_column,
        );
        set(
            &mut bits,
            CnfStyle::LAST_ROW_FIRST_COLUMN,
            x.last_row_first_column,
        );
        set(
            &mut bits,
            CnfStyle::LAST_ROW_LAST_COLUMN,
            x.last_row_last_column,
        );
        bits
    }
}

fn set(bits: &mut CnfStyle, flag: CnfStyle, on: Option<OnOffAttr>) {
    if let Some(OnOffAttr(true)) = on {
        bits.insert(flag);
    } else if let Some(OnOffAttr(false)) = on {
        bits.remove(flag);
    }
}

/// Parse a 12-character 0/1 bitstring. Shorter strings pad with zeros;
/// longer strings are truncated. Non-0/1 chars clear the bit.
fn parse_cnf_val(s: &str) -> CnfStyle {
    const ORDER: [CnfStyle; 12] = [
        CnfStyle::FIRST_ROW,
        CnfStyle::LAST_ROW,
        CnfStyle::FIRST_COLUMN,
        CnfStyle::LAST_COLUMN,
        CnfStyle::ODD_V_BAND,
        CnfStyle::EVEN_V_BAND,
        CnfStyle::ODD_H_BAND,
        CnfStyle::EVEN_H_BAND,
        CnfStyle::FIRST_ROW_FIRST_COLUMN,
        CnfStyle::FIRST_ROW_LAST_COLUMN,
        CnfStyle::LAST_ROW_FIRST_COLUMN,
        CnfStyle::LAST_ROW_LAST_COLUMN,
    ];
    let mut out = CnfStyle::empty();
    for (ch, flag) in s.chars().zip(ORDER.iter()) {
        if ch == '1' {
            out.insert(*flag);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> CnfStyle {
        let x: CnfStyleXml = quick_xml::de::from_str(xml).unwrap();
        x.into()
    }

    #[test]
    fn empty_yields_empty_bits() {
        let c = parse(r#"<cnfStyle val="000000000000"/>"#);
        assert!(c.is_empty());
    }

    #[test]
    fn first_row_only() {
        let c = parse(r#"<cnfStyle val="100000000000"/>"#);
        assert_eq!(c, CnfStyle::FIRST_ROW);
    }

    #[test]
    fn multiple_bits_from_val() {
        let c = parse(r#"<cnfStyle val="101010000000"/>"#);
        assert_eq!(
            c,
            CnfStyle::FIRST_ROW | CnfStyle::FIRST_COLUMN | CnfStyle::ODD_V_BAND
        );
    }

    #[test]
    fn individual_attrs_override_val() {
        let c = parse(r#"<cnfStyle val="000000000000" firstRow="1" oddHBand="true"/>"#);
        assert_eq!(c, CnfStyle::FIRST_ROW | CnfStyle::ODD_H_BAND);
    }

    #[test]
    fn short_val_pads_with_zero() {
        let c = parse(r#"<cnfStyle val="11"/>"#);
        assert_eq!(c, CnfStyle::FIRST_ROW | CnfStyle::LAST_ROW);
    }

    #[test]
    fn no_val_no_attrs_is_empty() {
        let c = parse(r#"<cnfStyle/>"#);
        assert!(c.is_empty());
    }
}
