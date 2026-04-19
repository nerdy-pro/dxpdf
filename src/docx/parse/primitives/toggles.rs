//! OOXML `ST_OnOff` (§17.18.68) toggle element plus attribute-level booleans.
//!
//! A toggle element without `@val` is "on" by default — the element's
//! presence alone asserts the property. With `@val`, the spec lists
//! `true`/`1`/`on` and `false`/`0`/`off` as the valid values.
//!
//! Per plan §Decisions, `OnOff` is the single spec-driven exception to
//! strict enum handling: unknown `@val` values resolve to `true` rather than
//! failing deserialization, matching legacy Word's tolerant toggle behavior.
//!
//! [`AttrBool`] is the attribute-level counterpart: a plain boolean
//! attribute like `@rotWithShape="1"` that reads as a string and maps
//! `"1"`/`"true"`/`"on"` to `true`, everything else to `false`.

use serde::{Deserialize, Deserializer};

/// Toggle value wrapper. `.0` is the resolved bool.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct OnOff(pub bool);

impl<'de> Deserialize<'de> for OnOff {
    fn deserialize<D: Deserializer<'de>>(d: D) -> Result<Self, D::Error> {
        #[derive(Deserialize)]
        struct Raw {
            #[serde(rename = "@val", default)]
            val: Option<String>,
        }
        let raw = Raw::deserialize(d)?;
        let on = match raw.val.as_deref() {
            None => true,
            Some("true" | "1" | "on") => true,
            Some("false" | "0" | "off") => false,
            Some(_) => true,
        };
        Ok(OnOff(on))
    }
}

/// Attribute-level boolean. Accepts `"1"`, `"true"`, `"on"` as true;
/// anything else (including `"0"`, `"false"`, `"off"`) as false.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct AttrBool(pub bool);

impl<'de> Deserialize<'de> for AttrBool {
    fn deserialize<D: Deserializer<'de>>(d: D) -> Result<Self, D::Error> {
        let s = String::deserialize(d)?;
        Ok(Self(matches!(s.as_str(), "1" | "true" | "on")))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[derive(Deserialize)]
    struct Toggle {
        #[serde(rename = "flag", default)]
        flag: Option<OnOff>,
    }

    fn flag(xml: &str) -> Option<bool> {
        let t: Toggle = quick_xml::de::from_str(xml).unwrap();
        t.flag.map(|o| o.0)
    }

    #[test]
    fn present_without_val_is_on() {
        assert_eq!(flag(r#"<x><flag/></x>"#), Some(true));
    }

    #[test]
    fn val_true_family_is_on() {
        assert_eq!(flag(r#"<x><flag val="true"/></x>"#), Some(true));
        assert_eq!(flag(r#"<x><flag val="1"/></x>"#), Some(true));
        assert_eq!(flag(r#"<x><flag val="on"/></x>"#), Some(true));
    }

    #[test]
    fn val_false_family_is_off() {
        assert_eq!(flag(r#"<x><flag val="false"/></x>"#), Some(false));
        assert_eq!(flag(r#"<x><flag val="0"/></x>"#), Some(false));
        assert_eq!(flag(r#"<x><flag val="off"/></x>"#), Some(false));
    }

    #[test]
    fn unknown_val_defaults_to_on() {
        assert_eq!(flag(r#"<x><flag val="garbage"/></x>"#), Some(true));
    }

    #[test]
    fn absent_is_none() {
        assert_eq!(flag(r#"<x/>"#), None);
    }
}
