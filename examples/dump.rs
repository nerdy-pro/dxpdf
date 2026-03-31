use dxpdf::model::*;

fn main() {
    env_logger::Builder::from_env(env_logger::Env::default().default_filter_or("warn"))
        .format_target(false)
        .init();
    let dir = concat!(env!("CARGO_MANIFEST_DIR"), "/test-files");
    let mut entries: Vec<_> = std::fs::read_dir(dir)
        .unwrap()
        .filter_map(|e| e.ok())
        .filter(|e| e.path().extension().map(|x| x == "docx").unwrap_or(false))
        .collect();
    entries.sort_by_key(|e| e.file_name());

    for entry in entries {
        let path = entry.path();
        let name = path.file_name().unwrap().to_string_lossy();
        let data = std::fs::read(&path).unwrap();
        match dxpdf::docx::parse(&data) {
            Ok(doc) => dump(&name, &doc),
            Err(e) => println!("=== {name} ===\n  ERROR: {e}\n"),
        }
    }
}

fn dump(name: &str, doc: &Document) {
    println!("╔══════════════════════════════════════════════════════════════");
    println!("║ {name}");
    println!("╠══════════════════════════════════════════════════════════════");

    // Settings
    println!("║ Settings:");
    println!(
        "║   default tab stop: {} twips",
        doc.settings.default_tab_stop.raw()
    );
    println!(
        "║   even/odd headers: {}",
        doc.settings.even_and_odd_headers
    );

    // Rsids
    if let Some(root) = doc.settings.rsid_root {
        println!("║   rsid root: {:08X}", root.value());
    }
    println!("║   rsid history: {} sessions", doc.settings.rsids.len());

    // Theme
    if let Some(theme) = &doc.theme {
        println!("║ Theme:");
        println!("║   major font: {}", theme.major_font.latin);
        println!("║   minor font: {}", theme.minor_font.latin);
        println!(
            "║   colors: dk1=#{:06X} lt1=#{:06X} accent1=#{:06X}",
            theme.color_scheme.dark1, theme.color_scheme.light1, theme.color_scheme.accent1
        );
    } else {
        println!("║ Theme: none");
    }

    // Section
    let s = &doc.final_section;
    println!("║ Final section:");
    if let Some(ref ps) = s.page_size {
        println!(
            "║   page: {:?}×{:?} twips ({:?})",
            ps.width, ps.height, ps.orientation
        );
    }
    if let Some(ref pm) = s.page_margins {
        println!(
            "║   margins: top={:?} right={:?} bottom={:?} left={:?} twips",
            pm.top, pm.right, pm.bottom, pm.left
        );
    }

    // Body stats
    let (paras, tables, images, hyperlinks, fields, tabs, breaks) = count_body(&doc.body);
    println!("║ Body:");
    println!("║   top-level blocks: {}", doc.body.len());
    println!("║   paragraphs (total): {paras}");
    println!("║   tables (total): {tables}");
    println!("║   images: {images}");
    println!("║   hyperlinks: {hyperlinks}");
    println!("║   fields: {fields}");
    println!("║   tabs: {tabs}");
    println!("║   line/page/col breaks: {breaks}");

    // Headers/footers
    println!(
        "║ Headers: {} | Footers: {}",
        doc.headers.len(),
        doc.footers.len()
    );
    for (rel_id, blocks) in &doc.headers {
        let (p, t, i, ..) = count_body(blocks);
        println!(
            "║   header {}: {p} paragraphs, {t} tables, {i} images",
            rel_id.as_str()
        );
    }
    for (rel_id, blocks) in &doc.footers {
        let (p, t, i, ..) = count_body(blocks);
        println!(
            "║   footer {}: {p} paragraphs, {t} tables, {i} images",
            rel_id.as_str()
        );
    }

    // Footnotes/endnotes
    if !doc.footnotes.is_empty() {
        println!("║ Footnotes: {}", doc.footnotes.len());
    }
    if !doc.endnotes.is_empty() {
        println!("║ Endnotes: {}", doc.endnotes.len());
    }

    // Media
    if !doc.media.is_empty() {
        println!("║ Media: {} entries", doc.media.len());
        for (rel_id, data) in &doc.media {
            let kind = if data.starts_with(&[0x89, 0x50, 0x4E, 0x47]) {
                "PNG"
            } else if data.starts_with(&[0xFF, 0xD8]) {
                "JPEG"
            } else if data.starts_with(b"GIF") {
                "GIF"
            } else {
                "unknown"
            };
            println!("║   {} → {} bytes ({})", rel_id.as_str(), data.len(), kind);
        }
    }

    // First few paragraphs of text
    println!("║ Text preview:");
    let mut lines = 0;
    for block in &doc.body {
        if lines >= 10 {
            println!("║   ...");
            break;
        }
        if let Block::Paragraph(p) = block {
            let text: String = p
                .content
                .iter()
                .filter_map(|i| {
                    if let Inline::TextRun(r) = i {
                        Some(
                            r.content
                                .iter()
                                .filter_map(|e| match e {
                                    RunElement::Text(t) => Some(t.as_str()),
                                    _ => None,
                                })
                                .collect::<String>(),
                        )
                    } else {
                        None
                    }
                })
                .collect();
            if !text.trim().is_empty() {
                let display = if text.len() > 100 {
                    format!("{}...", &text[..100])
                } else {
                    text
                };
                println!("║   │ {display}");
                lines += 1;
            }
        }
    }

    // Sample paragraph properties
    println!("║ Sample paragraph properties:");
    let mut shown = 0;
    for block in &doc.body {
        if shown >= 3 {
            break;
        }
        if let Block::Paragraph(p) = block {
            let has_text = p.content.iter().any(|i| {
                matches!(i, Inline::TextRun(r) if r.content.iter().any(|e| matches!(e, RunElement::Text(t) if !t.trim().is_empty())))
            });
            if !has_text {
                continue;
            }
            let pp = &p.properties;
            print!("║   ¶ align={:?}", pp.alignment);
            if let Some(ref sp) = pp.spacing {
                if let Some(before) = sp.before {
                    print!(" before={}", before.raw());
                }
                if let Some(after) = sp.after {
                    print!(" after={}", after.raw());
                }
            }
            if let Some(ref ind) = pp.indentation {
                if let Some(start) = ind.start {
                    print!(" indent.start={}", start.raw());
                }
            }
            if pp.numbering.is_some() {
                print!(" [numbered]");
            }
            if pp.keep_next == Some(true) {
                print!(" [keepNext]");
            }
            println!();

            // Run properties of first text run
            if let Some(Inline::TextRun(run)) =
                p.content.iter().find(|i| matches!(i, Inline::TextRun(_)))
            {
                let rp = &run.properties;
                if let Some(sz) = rp.font_size {
                    print!("║     run: size={}hp", sz.raw());
                } else {
                    print!("║     run:");
                }
                if rp.bold == Some(true) {
                    print!(" bold");
                }
                if rp.italic == Some(true) {
                    print!(" italic");
                }
                if let Some(ref u) = rp.underline {
                    print!(" underline={u:?}");
                }
                if let Some(ref f) = rp.fonts.ascii {
                    print!(" font=\"{f}\"");
                }
                print!(" color={:?}", rp.color);
                println!();
            }
            shown += 1;
        }
    }

    println!("╚══════════════════════════════════════════════════════════════");
    println!();
}

fn count_body(blocks: &[Block]) -> (usize, usize, usize, usize, usize, usize, usize) {
    let mut paras = 0;
    let mut tables = 0;
    let mut images = 0;
    let mut hyperlinks = 0;
    let mut fields = 0;
    let mut tabs = 0;
    let mut breaks = 0;

    for block in blocks {
        match block {
            Block::Paragraph(p) => {
                paras += 1;
                for inline in &p.content {
                    match inline {
                        Inline::Image(_) => images += 1,
                        Inline::Hyperlink(h) => {
                            hyperlinks += 1;
                            for i in &h.content {
                                if matches!(i, Inline::Image(_)) {
                                    images += 1;
                                }
                            }
                        }
                        Inline::Field(_) => fields += 1,
                        Inline::TextRun(r) => {
                            for elem in &r.content {
                                match elem {
                                    RunElement::Tab => tabs += 1,
                                    RunElement::LineBreak(_)
                                    | RunElement::PageBreak
                                    | RunElement::ColumnBreak => breaks += 1,
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            Block::Table(t) => {
                tables += 1;
                for row in &t.rows {
                    for cell in &row.cells {
                        let (p, tb, i, h, f, ta, b) = count_body(&cell.content);
                        paras += p;
                        tables += tb;
                        images += i;
                        hyperlinks += h;
                        fields += f;
                        tabs += ta;
                        breaks += b;
                    }
                }
            }
            Block::SectionBreak(_) => {}
        }
    }
    (paras, tables, images, hyperlinks, fields, tabs, breaks)
}
