//! Parser for `word/settings.xml`.

use log::warn;
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::dimension::Dimension;
use crate::error::Result;
use crate::model::{DocumentSettings, RevisionSaveId};
use crate::xml;

pub fn parse_settings(data: &[u8]) -> Result<DocumentSettings> {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut settings = DocumentSettings::default();

    loop {
        match xml::next_event(&mut reader, &mut buf)? {
            Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                if local.as_slice() == b"rsids" {
                    parse_rsids(&mut reader, &mut buf, &mut settings)?;
                }
            }
            Event::Empty(ref e) => {
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

fn parse_rsids(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
    settings: &mut DocumentSettings,
) -> Result<()> {
    loop {
        match xml::next_event(reader, buf)? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let local = xml::local_name(e.name().as_ref()).to_vec();
                match local.as_slice() {
                    b"rsidRoot" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            settings.rsid_root = RevisionSaveId::from_hex(&val);
                        }
                    }
                    b"rsid" => {
                        if let Some(val) = xml::optional_attr(e, b"val")? {
                            if let Some(rsid) = RevisionSaveId::from_hex(&val) {
                                settings.rsids.push(rsid);
                            }
                        }
                    }
                    _ => {
                        warn!(
                            "rsids: unsupported element <{}>",
                            String::from_utf8_lossy(&local)
                        );
                    }
                }
            }
            Event::End(ref e) if xml::local_name(e.name().as_ref()) == b"rsids" => break,
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(())
}
