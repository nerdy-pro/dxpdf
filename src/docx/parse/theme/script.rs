//! ISO 15924 script-code mapping for `<a:font script="...">` (§20.1.4.1.16).

use crate::docx::model::ScriptTag;

/// Known ISO 15924 script codes referenced by OOXML theme font schemes.
///
/// The list matches the variants enumerated on `ScriptTag`. Unknown codes
/// fall through to `ScriptTag::Other`.
const SCRIPT_TAGS: &[(&str, ScriptTag)] = &[
    ("Arab", ScriptTag::Arab),
    ("Armn", ScriptTag::Armn),
    ("Beng", ScriptTag::Beng),
    ("Bopo", ScriptTag::Bopo),
    ("Bugi", ScriptTag::Bugi),
    ("Cans", ScriptTag::Cans),
    ("Cher", ScriptTag::Cher),
    ("Deva", ScriptTag::Deva),
    ("Ethi", ScriptTag::Ethi),
    ("Geor", ScriptTag::Geor),
    ("Gujr", ScriptTag::Gujr),
    ("Guru", ScriptTag::Guru),
    ("Hang", ScriptTag::Hang),
    ("Hans", ScriptTag::Hans),
    ("Hant", ScriptTag::Hant),
    ("Hebr", ScriptTag::Hebr),
    ("Java", ScriptTag::Java),
    ("Jpan", ScriptTag::Jpan),
    ("Khmr", ScriptTag::Khmr),
    ("Knda", ScriptTag::Knda),
    ("Laoo", ScriptTag::Laoo),
    ("Lisu", ScriptTag::Lisu),
    ("Mlym", ScriptTag::Mlym),
    ("Mong", ScriptTag::Mong),
    ("Mymr", ScriptTag::Mymr),
    ("Nkoo", ScriptTag::Nkoo),
    ("Olck", ScriptTag::Olck),
    ("Orya", ScriptTag::Orya),
    ("Osma", ScriptTag::Osma),
    ("Phag", ScriptTag::Phag),
    ("Sinh", ScriptTag::Sinh),
    ("Sora", ScriptTag::Sora),
    ("Syre", ScriptTag::Syre),
    ("Syrj", ScriptTag::Syrj),
    ("Syrn", ScriptTag::Syrn),
    ("Syrc", ScriptTag::Syrc),
    ("Tale", ScriptTag::Tale),
    ("Talu", ScriptTag::Talu),
    ("Taml", ScriptTag::Taml),
    ("Telu", ScriptTag::Telu),
    ("Tfng", ScriptTag::Tfng),
    ("Thaa", ScriptTag::Thaa),
    ("Thai", ScriptTag::Thai),
    ("Tibt", ScriptTag::Tibt),
    ("Uigh", ScriptTag::Uigh),
    ("Viet", ScriptTag::Viet),
    ("Yiii", ScriptTag::Yiii),
];

/// Parse an ISO 15924 script code into a `ScriptTag`.
pub fn parse_script_tag(s: &str) -> ScriptTag {
    if let Some((_, tag)) = SCRIPT_TAGS.iter().find(|(name, _)| *name == s) {
        return tag.clone();
    }
    log::warn!("theme: unrecognized script tag {s:?}");
    ScriptTag::Other(s.to_string().into_boxed_str())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn known_tag_is_mapped() {
        assert_eq!(parse_script_tag("Hans"), ScriptTag::Hans);
        assert_eq!(parse_script_tag("Arab"), ScriptTag::Arab);
    }

    #[test]
    fn unknown_tag_preserves_full_string() {
        match parse_script_tag("Xxxx-long") {
            ScriptTag::Other(s) => assert_eq!(&*s, "Xxxx-long"),
            other => panic!("expected Other, got {other:?}"),
        }
    }

    #[test]
    fn empty_tag_is_other() {
        match parse_script_tag("") {
            ScriptTag::Other(s) => assert_eq!(&*s, ""),
            other => panic!("expected Other, got {other:?}"),
        }
    }

    #[test]
    fn all_listed_tags_are_unique() {
        let mut names: Vec<&str> = SCRIPT_TAGS.iter().map(|(n, _)| *n).collect();
        names.sort_unstable();
        let before = names.len();
        names.dedup();
        assert_eq!(names.len(), before, "duplicate script tag in table");
    }
}
