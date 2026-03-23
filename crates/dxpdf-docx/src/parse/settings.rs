//! Parser for `word/settings.xml`.

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::dimension::Dimension;
use crate::error::Result;
use crate::model::DocumentSettings;
use crate::xml;

pub fn parse_settings(data: &[u8]) -> Result<DocumentSettings> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut settings = DocumentSettings::default();

    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"defaultTabStop" => {
                        if let Some(val) = xml::optional_attr_i64(e, b"val")? {
                            settings.default_tab_stop = Dimension::new(val);
                        }
                    }
                    b"evenAndOddHeaders" => {
                        let enabled = xml::optional_attr_bool(e, b"val")?.unwrap_or(true);
                        settings.even_and_odd_headers = enabled;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(settings)
}
