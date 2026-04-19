//! Parser for `word/numbering.xml` — single-pass serde over the whole file.
//! Picture bullets' `<w:pict>` contents are deserialized via the VML schema.

use serde::Deserialize;

use crate::docx::error::Result;
use crate::docx::model::{
    AbstractNumId, AbstractNumbering, Alignment, Indentation, NumId, NumPicBullet, NumPicBulletId,
    NumberFormat, NumberingDefinitions, NumberingInstance, NumberingLevelDefinition, RunProperties,
};
use crate::docx::parse::primitives::st_enums::{StJc, StNumberFormat};
use crate::docx::parse::properties::schema::paragraph::PPrXml;
use crate::docx::parse::properties::schema::run::RPrXml;
use crate::docx::parse::serde_xml::from_xml;

pub fn parse_numbering(data: &[u8]) -> Result<NumberingDefinitions> {
    if data.is_empty() {
        return Ok(NumberingDefinitions::default());
    }
    let schema: NumberingXml = from_xml(data)?;
    Ok(schema.into())
}

// ── serde schema ──────────────────────────────────────────────────────────

#[derive(Deserialize, Default)]
struct NumberingXml {
    #[serde(rename = "abstractNum", default)]
    abstract_nums: Vec<AbstractNumXml>,
    #[serde(rename = "num", default)]
    nums: Vec<NumXml>,
    /// Picture bullets are parsed structurally here (for id resolution)
    /// but their `<w:pict>` contents are filled in by the pre-pass.
    #[serde(rename = "numPicBullet", default)]
    num_pic_bullets: Vec<NumPicBulletXml>,
}

#[derive(Deserialize)]
struct AbstractNumXml {
    #[serde(rename = "@abstractNumId")]
    abstract_num_id: i64,
    #[serde(rename = "lvl", default)]
    levels: Vec<LvlXml>,
}

#[derive(Deserialize)]
struct LvlXml {
    #[serde(rename = "@ilvl")]
    ilvl: u8,
    #[serde(rename = "numFmt", default)]
    num_fmt: Option<ValAttr<StNumberFormat>>,
    #[serde(rename = "lvlText", default)]
    lvl_text: Option<ValString>,
    #[serde(rename = "start", default)]
    start: Option<ValAttr<u32>>,
    #[serde(rename = "lvlJc", default)]
    lvl_jc: Option<ValAttr<StJc>>,
    #[serde(rename = "pPr", default)]
    p_pr: Option<PPrXml>,
    #[serde(rename = "rPr", default)]
    r_pr: Option<RPrXml>,
    #[serde(rename = "lvlPicBulletId", default)]
    lvl_pic_bullet_id: Option<ValAttr<i64>>,
}

#[derive(Deserialize)]
struct NumXml {
    #[serde(rename = "@numId")]
    num_id: i64,
    #[serde(rename = "abstractNumId", default)]
    abstract_num_id: Option<ValAttr<i64>>,
    #[serde(rename = "lvlOverride", default)]
    overrides: Vec<LvlOverrideXml>,
}

#[derive(Deserialize)]
struct LvlOverrideXml {
    #[serde(rename = "@ilvl")]
    ilvl: u8,
    #[serde(rename = "lvl", default)]
    lvl: Option<LvlXml>,
}

#[derive(Deserialize)]
struct NumPicBulletXml {
    #[serde(rename = "@numPicBulletId")]
    num_pic_bullet_id: i64,
    #[serde(rename = "pict", default)]
    pict: Option<crate::docx::parse::vml::schema::PictXml>,
}

#[derive(Deserialize)]
struct ValString {
    #[serde(rename = "@val")]
    val: String,
}

#[derive(Deserialize)]
#[serde(bound(deserialize = "T: serde::Deserialize<'de>"))]
struct ValAttr<T> {
    #[serde(rename = "@val")]
    val: T,
}

// ── schema → model ────────────────────────────────────────────────────────

impl From<NumberingXml> for NumberingDefinitions {
    fn from(x: NumberingXml) -> Self {
        let mut defs = NumberingDefinitions::default();
        for a in x.abstract_nums {
            let id = AbstractNumId::new(a.abstract_num_id);
            defs.abstract_nums.insert(
                id,
                AbstractNumbering {
                    levels: a.levels.into_iter().map(Into::into).collect(),
                },
            );
        }
        for n in x.nums {
            defs.numbering_instances
                .insert(NumId::new(n.num_id), convert_num(n));
        }
        // Picture bullets may contain a VML `<w:pict>` (e.g., an imagedata
        // reference). Numbering has no body content, so no embeds crossing
        // into body convert — pass an empty ctx.
        let mut ctx = crate::docx::parse::body::ConvertCtx::new();
        for bullet in x.num_pic_bullets {
            let id = NumPicBulletId::new(bullet.num_pic_bullet_id);
            let pict = bullet.pict.map(|p| p.into_model(&mut ctx));
            defs.pic_bullets.insert(id, NumPicBullet { id, pict });
        }
        defs
    }
}

impl From<LvlXml> for NumberingLevelDefinition {
    fn from(x: LvlXml) -> Self {
        let (indentation, run_properties) = extract_level_properties(x.p_pr, x.r_pr);
        Self {
            level: x.ilvl,
            format: x.num_fmt.map(|v| NumberFormat::from(v.val)),
            level_text: x.lvl_text.map(|v| v.val).unwrap_or_default(),
            start: x.start.map(|v| v.val),
            justification: x.lvl_jc.map(|v| Alignment::from(v.val)),
            indentation,
            run_properties,
            lvl_pic_bullet_id: x.lvl_pic_bullet_id.map(|v| NumPicBulletId::new(v.val)),
        }
    }
}

fn extract_level_properties(
    p_pr: Option<PPrXml>,
    r_pr: Option<RPrXml>,
) -> (Option<Indentation>, Option<RunProperties>) {
    let indentation = p_pr.and_then(|p| p.split().properties.indentation);
    let run_properties = r_pr.map(|r| r.split().0);
    (indentation, run_properties)
}

fn convert_num(n: NumXml) -> NumberingInstance {
    let abstract_num_id = n
        .abstract_num_id
        .map(|v| AbstractNumId::new(v.val))
        .unwrap_or_else(|| AbstractNumId::new(0));
    let level_overrides = n
        .overrides
        .into_iter()
        .filter_map(|o| {
            o.lvl.map(|mut lvl| {
                lvl.ilvl = o.ilvl; // legacy parser used override's @ilvl, not inner lvl's
                NumberingLevelDefinition::from(lvl)
            })
        })
        .collect();
    NumberingInstance {
        abstract_num_id,
        level_overrides,
    }
}
