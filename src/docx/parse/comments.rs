//! Parser for `word/comments.xml` (§11.3.1, issue #154).
//!
//! The shape of `notes.rs` with the attributes notes lack: each `<w:comment>`
//! carries `@w:id`, `@w:author` and `@w:initials`, and holds ordinary block
//! content converted through the body machinery. `<w:annotationRef>` inside
//! the comment body — the reference mark Word renders there — has its own
//! schema variant and is dropped at conversion: the balloon labels itself
//! with the author instead. Modern-comments sibling parts (`commentsExtended`, `people`) are
//! not read; the core part carries everything the balloon shows.

use std::collections::HashMap;

use serde::Deserialize;

use crate::docx::error::Result;
use crate::docx::model::{Comment, CommentId};
use crate::docx::parse::body;
use crate::docx::parse::body_schema::BlockChildXml;
use crate::docx::parse::serde_xml::from_xml;

/// Parse a comments part into comment content keyed by id.
pub fn parse_comments(data: &[u8]) -> Result<HashMap<CommentId, Comment>> {
    let file: CommentsFileXml = from_xml(data)?;
    let mut out = HashMap::new();
    for c in file.comments {
        let Some(id) = c.id else {
            // An id-less comment is unreachable from any anchor.
            continue;
        };
        let mut ctx = body::ConvertCtx::new();
        let (content, _) = body::convert_container(c.content, &mut ctx);
        out.insert(
            CommentId::new(id),
            Comment {
                author: c.author.unwrap_or_default(),
                initials: c.initials.unwrap_or_default(),
                content,
            },
        );
    }
    Ok(out)
}

#[derive(Deserialize, Default)]
struct CommentsFileXml {
    #[serde(rename = "comment", default)]
    comments: Vec<CommentXml>,
}

#[derive(Deserialize, Default)]
struct CommentXml {
    #[serde(rename = "@id")]
    id: Option<i64>,
    #[serde(rename = "@author")]
    author: Option<String>,
    #[serde(rename = "@initials")]
    initials: Option<String>,
    #[serde(rename = "$value", default)]
    content: Vec<BlockChildXml>,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{Block, Inline, RunElement};

    fn text_of(blocks: &[Block]) -> String {
        let mut s = String::new();
        for b in blocks {
            if let Block::Paragraph(p) = b {
                for i in &p.content {
                    if let Inline::TextRun(tr) = i {
                        for e in &tr.content {
                            if let RunElement::Text(t) = e {
                                s.push_str(t);
                            }
                        }
                    }
                }
            }
        }
        s
    }

    #[test]
    fn a_comment_parses_with_author_and_content() {
        let xml = r#"<w:comments xmlns:w="x">
            <w:comment w:id="1" w:author="Ann" w:initials="A">
                <w:p><w:r><w:annotationRef/></w:r><w:r><w:t>note text</w:t></w:r></w:p>
            </w:comment>
        </w:comments>"#;
        let comments = parse_comments(xml.as_bytes()).expect("parses");
        let c = &comments[&CommentId::new(1)];
        assert_eq!(c.author, "Ann");
        assert_eq!(c.initials, "A");
        assert_eq!(text_of(&c.content), "note text");
    }

    #[test]
    fn an_id_less_comment_is_dropped() {
        let xml = r#"<w:comments xmlns:w="x"><w:comment w:author="A"/></w:comments>"#;
        assert!(parse_comments(xml.as_bytes()).expect("parses").is_empty());
    }
}
