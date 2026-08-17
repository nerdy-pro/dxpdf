//! §17.4.38 / §17.4.66: the structural invariants of a table's border network.
//!
//! ECMA-376 says which edges a cell paints and [MS-OI29500] §17.4.66 says which
//! of two facing cells wins a shared one, but neither says anything about the
//! square where a vertical border crosses a horizontal one. That is this
//! engine's own convention, and it has now been reported wrong three times, in
//! three different shapes, always as the same symptom: a 1–2px notch at a cell
//! corner. Each report was a square that the two edges of the cell owning it had
//! both been emptied of, while the borders that actually meet there belonged to
//! the neighbouring row and the neighbouring column.
//!
//! So this file asserts properties rather than any one shape of them, over whole
//! rendered documents, and nothing here knows which cell a rect came from —
//! that knowledge is exactly what each of the three defects had too little of.
//! Three properties, which between them say the network is a set of lines that
//! meet cleanly:
//!
//! 1. **Every junction square is ink.** A junction is where a vertical border
//!    rect and a horizontal one touch or overlap; the square they join in is the
//!    vertical's x-band crossed with the horizontal's y-band. A hole there is
//!    the reported notch.
//! 2. **No two border rects overlap.** Where two do, which colour reaches the
//!    page is decided by emission *order* rather than by any rule — so a
//!    document renders one way and the same borders in another arrangement
//!    render another. Painting a square twice is the other half of the same
//!    defect as painting it not at all.
//! 3. **No collinear pair is separated by a gap narrower than itself.** A break
//!    in a line one border-width long is never deliberate: real space between
//!    two tables is orders of magnitude wider. This is the notch that the
//!    junction audit cannot see, because it is a hole in a line with nothing
//!    crossing it.
//!
//! The fixture is the reporter's own document. `test-cases/` is untracked
//! (private customer documents), so those tests are gated on its presence and
//! are a no-op in CI; the same invariants are asserted on in-memory tables by
//! `render::layout::table::emit`'s own tests, which do run there.

use dxpdf::render::layout::draw_command::{DrawCommand, LayoutedPage};
use dxpdf::render::resolve_and_layout;

/// A rect as `(x0, x1, y0, y1)`. Only thin ones are borders; a shading rect or
/// an image would swamp the junction search with squares that no border meets.
const MAX_BORDER_THICKNESS: f32 = 3.0;
const EPS: f32 = 0.001;

type Rect = (f32, f32, f32, f32);

/// The committed corpus. Absolute, so the audits do not depend on the working
/// directory a test binary happens to be run from.
const FIXTURES: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/test-files");

fn border_rects(page: &LayoutedPage) -> Vec<Rect> {
    page.commands
        .iter()
        .filter_map(|c| match c {
            DrawCommand::Rect { rect, .. } => {
                let (w, h) = (rect.size.width.raw(), rect.size.height.raw());
                (w.min(h) <= MAX_BORDER_THICKNESS && w.min(h) > 0.0).then(|| {
                    (
                        rect.origin.x.raw(),
                        rect.origin.x.raw() + w,
                        rect.origin.y.raw(),
                        rect.origin.y.raw() + h,
                    )
                })
            }
            _ => None,
        })
        .collect()
}

/// Whether `square` is entirely painted by `rects` — by their **union**, not by
/// any one of them.
///
/// The union matters: two tables whose grids differ by a tenth of a point leave
/// junction squares straddling the seam between two abutting horizontals, which
/// a single-rect test reports as holes that are not there. Exact rather than
/// sampled — rect edges are the only discontinuities, so testing one x inside
/// each slab between them decides the whole slab.
fn covered(square: Rect, rects: &[Rect]) -> bool {
    let (sx0, sx1, sy0, sy1) = square;
    let mut xs = vec![sx0, sx1];
    for (x0, x1, ..) in rects {
        for x in [*x0, *x1] {
            if x > sx0 && x < sx1 {
                xs.push(x);
            }
        }
    }
    xs.sort_by(f32::total_cmp);

    for pair in xs.windows(2) {
        let (a, b) = (pair[0], pair[1]);
        if b - a <= EPS {
            continue;
        }
        let mid = (a + b) * 0.5;
        let mut spans: Vec<(f32, f32)> = rects
            .iter()
            .filter(|(x0, x1, ..)| *x0 <= mid && mid <= *x1)
            .map(|(_, _, y0, y1)| (*y0, *y1))
            .collect();
        spans.sort_by(|p, q| p.0.total_cmp(&q.0));
        let mut reached = sy0;
        for (y0, y1) in spans {
            if y0 > reached + EPS {
                break;
            }
            reached = reached.max(y1);
        }
        if reached < sy1 - EPS {
            return false;
        }
    }
    true
}

/// `(junctions_checked, unpainted)` for one page.
fn junctions(rects: &[Rect]) -> (usize, Vec<Rect>) {
    let (vertical, horizontal): (Vec<_>, Vec<_>) = rects
        .iter()
        .copied()
        .partition(|(x0, x1, y0, y1)| x1 - x0 < y1 - y0);

    let mut checked = 0usize;
    let mut missing: Vec<Rect> = Vec::new();
    for (vx0, vx1, vy0, vy1) in vertical {
        for (hx0, hx1, hy0, hy1) in horizontal.iter().copied() {
            if vx1 < hx0 - EPS || vx0 > hx1 + EPS || hy1 < vy0 - EPS || hy0 > vy1 + EPS {
                continue;
            }
            checked += 1;
            let square = (vx0, vx1, hy0, hy1);
            if !covered(square, rects) && !missing.contains(&square) {
                missing.push(square);
            }
        }
    }
    (checked, missing)
}

/// Whether two rects share positive area. Touching along an edge is not
/// overlapping — abutting rects are how every line in this engine is built.
fn overlaps(a: Rect, b: Rect) -> bool {
    let (ax0, ax1, ay0, ay1) = a;
    let (bx0, bx1, by0, by1) = b;
    ax1 - bx0 > EPS && bx1 - ax0 > EPS && ay1 - by0 > EPS && by1 - ay0 > EPS
}

/// The one overlap that is the document's fault rather than the renderer's:
/// **two parallel lines closer together than they are thick**.
///
/// Two boundaries a hair apart, each carrying a border wider than the gap, cross
/// whatever a renderer does with them — there is no decomposition that separates
/// them, because the geometry the author asked for is impossible. `hRule="exact"`
/// on a row shorter than its own borders is one way to write it (§17.4.80); an
/// empty `<w:tr/>` between two rows is another (§17.4.66,
/// `issue-157-empty-row-edge`).
///
/// Recognised narrowly, and the narrowness is the whole of it. The two must run
/// the same extent along **their long axis** and sit at different positions
/// across their short one — two parallel lines, stacked through their own
/// thickness.
///
/// Matching on "same extent along *some* axis" is not enough, and the difference
/// is not hypothetical: a junction overrunning a horizontal segment shares the
/// segment's **short** axis (both are that boundary's 0.5pt band) and differs
/// along the long one. That is exactly the defect this audit caught in
/// `sample-docx-files-sample1.docx`, and the loose form exempts it.
fn is_parallel_crowding(a: Rect, b: Rect) -> bool {
    let same = |p: (f32, f32), q: (f32, f32)| (p.0 - q.0).abs() <= EPS && (p.1 - q.1).abs() <= EPS;
    let (ax, ay) = ((a.0, a.1), (a.2, a.3));
    let (bx, by) = ((b.0, b.1), (b.2, b.3));
    // A square is ambiguous and `>=` calls it horizontal; both rects must agree
    // with the shared axis either way, so the ambiguity cannot let a pair in.
    let long_is_x = |r: Rect| r.1 - r.0 >= r.3 - r.2;
    let stacked = |shared_is_x: bool| long_is_x(a) == shared_is_x && long_is_x(b) == shared_is_x;
    (same(ax, bx) && !same(ay, by) && stacked(true))
        || (same(ay, by) && !same(ax, bx) && stacked(false))
}

/// `(pairs_examined, overlapping pairs, reported as their intersection)`.
///
/// The intersection is the useful report: it is the square whose colour is
/// undecided, and it is usually a junction, which is what says *why* the pair
/// exists.
fn overlapping(rects: &[Rect]) -> (usize, Vec<Rect>) {
    let mut examined = 0usize;
    let mut bad = Vec::new();
    for (i, a) in rects.iter().copied().enumerate() {
        for b in rects[i + 1..].iter().copied() {
            examined += 1;
            if overlaps(a, b) && !is_parallel_crowding(a, b) {
                let square = (a.0.max(b.0), a.1.min(b.1), a.2.max(b.2), a.3.min(b.3));
                if !bad.contains(&square) {
                    bad.push(square);
                }
            }
        }
    }
    (examined, bad)
}

/// `(collinear pairs examined, gaps in a line narrower than the line is thick)`.
///
/// Grouped by *overlapping band* rather than by exact coordinate, because two
/// segments of one grid line may differ in width — a 3pt `w:left` above a 1pt
/// one is still one line — and their bands then share only their overlap. Only
/// gaps shorter than the thicker of the two are reported: that is the size a
/// notch is, and it is far below any real space between two tables.
fn collinear_gaps(rects: &[Rect]) -> (usize, Vec<Rect>) {
    let mut examined = 0usize;
    let mut bad: Vec<Rect> = Vec::new();

    // (along-axis extractor, across-axis extractor, rebuild) for each of the two
    // orientations, so the same walk serves both.
    let vertical = |r: &Rect| r.1 - r.0 < r.3 - r.2;
    for want_vertical in [true, false] {
        let group: Vec<Rect> = rects
            .iter()
            .copied()
            .filter(|r| vertical(r) == want_vertical)
            .collect();
        for (i, a) in group.iter().copied().enumerate() {
            for b in group[i + 1..].iter().copied() {
                // `across` is the band the two must share to be collinear;
                // `along` is the direction the gap is measured in.
                let (across_a, across_b, along_a, along_b) = if want_vertical {
                    ((a.0, a.1), (b.0, b.1), (a.2, a.3), (b.2, b.3))
                } else {
                    ((a.2, a.3), (b.2, b.3), (a.0, a.1), (b.0, b.1))
                };
                let share = across_a.1.min(across_b.1) - across_a.0.max(across_b.0);
                if share <= EPS {
                    continue;
                }
                let (lo, hi) = if along_a.0 <= along_b.0 {
                    (along_a, along_b)
                } else {
                    (along_b, along_a)
                };
                let gap = hi.0 - lo.1;
                let thickness = (across_a.1 - across_a.0).max(across_b.1 - across_b.0);
                if gap <= EPS || gap >= thickness {
                    continue;
                }
                examined += 1;
                let hole = if want_vertical {
                    (
                        across_a.0.max(across_b.0),
                        across_a.1.min(across_b.1),
                        lo.1,
                        hi.0,
                    )
                } else {
                    (
                        lo.1,
                        hi.0,
                        across_a.0.max(across_b.0),
                        across_a.1.min(across_b.1),
                    )
                };
                if !covered(hole, rects) && !bad.contains(&hole) {
                    bad.push(hole);
                }
            }
        }
    }
    (examined, bad)
}

/// What one page is asked. `(things examined, violations)` — the count is what
/// makes a clean run non-vacuous.
type PageCheck = fn(&[Rect]) -> (usize, Vec<Rect>);

/// `(pages, things checked, one report line per page that has a violation)` for
/// one document.
fn audit(path: &std::path::Path, check: PageCheck) -> (usize, usize, Vec<String>) {
    let bytes = std::fs::read(path).unwrap_or_else(|e| panic!("read {}: {e}", path.display()));
    let doc =
        dxpdf::docx::parse(&bytes).unwrap_or_else(|e| panic!("parse {}: {e}", path.display()));
    let (_, pages) = resolve_and_layout(doc);

    let mut total_checked = 0usize;
    let mut failures = Vec::new();
    for (i, page) in pages.iter().enumerate() {
        let (checked, missing) = check(&border_rects(page));
        total_checked += checked;
        if !missing.is_empty() {
            failures.push(format!(
                "{} page {}: {missing:?}",
                path.file_name().unwrap_or_default().to_string_lossy(),
                i + 1
            ));
        }
    }
    (pages.len(), total_checked, failures)
}

/// Every `.docx` in a directory, sorted, or none when the directory is absent.
fn corpus(dir: &str) -> Vec<std::path::PathBuf> {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return Vec::new();
    };
    let mut v: Vec<_> = entries
        .filter_map(|e| e.ok().map(|e| e.path()))
        .filter(|p| p.extension().is_some_and(|e| e == "docx"))
        .filter(|p| !is_word_owner_file(p))
        .collect();
    v.sort();
    v
}

/// Word writes a `~$`-prefixed owner file beside any document it has open, with
/// the same `.docx` extension and no ZIP inside it. Anyone comparing a fixture
/// against Word therefore drops one into the corpus directory, and a scan that
/// picked it up would fail the audit with a parse error that has nothing to do
/// with borders.
fn is_word_owner_file(p: &std::path::Path) -> bool {
    p.file_name()
        .and_then(|n| n.to_str())
        .is_some_and(|n| n.starts_with("~$"))
}

/// Run one check over a corpus and print a line per document, so a run with
/// `--nocapture` is the corpus-wide report and not only a pass/fail.
///
/// Returns `(things checked, failure lines)`. The count is the non-vacuity
/// guard: a check that examined nothing has proved nothing, and the rect filter
/// silently ceasing to match is a realistic way for that to happen.
fn audit_corpus(dir: &str, unit: &str, check: PageCheck) -> (usize, Vec<String>) {
    let mut failures = Vec::new();
    let mut checked_total = 0usize;
    for path in corpus(dir) {
        let (pages, checked, mut bad) = audit(&path, check);
        checked_total += checked;
        println!(
            "{:>7} {unit} {:>3} pages  {:>2} bad  {}",
            checked,
            pages,
            bad.len(),
            path.file_name().unwrap_or_default().to_string_lossy()
        );
        failures.append(&mut bad);
    }
    println!(
        "{dir}: {checked_total} {unit} checked, {} bad",
        failures.len()
    );
    (checked_total, failures)
}

/// The whole committed corpus, every page: no junction is painted by nobody.
///
/// This is the check that would have caught all three reported corner defects at
/// once, and it is here rather than in a scratch script for that reason — the
/// class stays closed only while something asks the question on every render.
#[test]
fn no_committed_fixture_has_an_unpainted_border_junction() {
    let (checked, failures) = audit_corpus(FIXTURES, "junctions", junctions);
    assert!(checked > 100, "audit examined only {checked} junctions");
    assert!(
        failures.is_empty(),
        "border junctions painted by nobody:\n{}",
        failures.join("\n")
    );
}

/// The same over the untracked local corpus, which is where all three reports
/// came from. A no-op without it; the loop above still runs in CI.
#[test]
fn no_local_corpus_document_has_an_unpainted_border_junction() {
    if corpus("test-cases").is_empty() {
        eprintln!("SKIPPED: test-cases/ not present");
        return;
    }
    let (_, failures) = audit_corpus("test-cases", "junctions", junctions);
    assert!(
        failures.is_empty(),
        "border junctions painted by nobody:\n{}",
        failures.join("\n")
    );
}

/// The whole committed corpus: **no two border rects overlap**.
///
/// The other half of the junction defect. A square painted twice has no defined
/// colour — the later command wins, so the answer is emission order, which is
/// not a rule anyone can state or a reader can predict. It is also the exact
/// symptom of a model in which two different things both believe they own a
/// square, which is what a per-cell emitter is whenever the square is on a
/// shared edge.
#[test]
fn no_committed_fixture_paints_a_border_square_twice() {
    let (checked, failures) = audit_corpus(FIXTURES, "rect pairs", overlapping);
    assert!(checked > 1000, "audit examined only {checked} pairs");
    assert!(
        failures.is_empty(),
        "border rects overlap (intersections):\n{}",
        failures.join("\n")
    );
}

// There is deliberately **no local-corpus twin of the overlap audit**, and the
// reason is a real limit rather than an oversight.
//
// Every check in this file works from draw commands, which is the whole point —
// no knowledge of which cell, or which *table*, painted a rect. For the junction
// and gap audits that costs nothing. For the overlap audit it costs the one case
// the untracked corpus is full of and the committed one has none of: a **nested
// table** flush with its parent's edge. Both tables draw that edge, correctly and
// independently, and the two rects land on the same line. From the command
// stream that is indistinguishable from a junction overrunning a segment inside
// one table — the defect this audit caught in `sample-docx-files-sample1.docx` —
// because both are two rects of one colour partly covering each other, with
// neither containing the other.
//
// So the audit runs strictly on the corpus where the distinction cannot arise.
// Making it run on the other one needs the rects tagged with the table that drew
// them, which is a change to `DrawCommand` for a test's benefit, and not one to
// make without a defect asking for it.

/// The whole committed corpus: **a line is not broken by a gap narrower than
/// itself**.
///
/// The notch the junction audit cannot see. Where two segments of one grid line
/// stop short of each other with nothing crossing between them, no junction
/// exists to be checked — the hole is in the line itself. Bounding the gap by
/// the line's own thickness is what keeps real space out of it: two tables
/// separated by a paragraph are orders of magnitude further apart than a border
/// is thick.
#[test]
fn no_committed_fixture_breaks_a_border_line_by_less_than_its_width() {
    let (_, failures) = audit_corpus(FIXTURES, "collinear gaps", collinear_gaps);
    assert!(
        failures.is_empty(),
        "border lines broken by a sub-width gap:\n{}",
        failures.join("\n")
    );
}

/// The same over the untracked local corpus. A no-op without it.
#[test]
fn no_local_corpus_document_breaks_a_border_line_by_less_than_its_width() {
    if corpus("test-cases").is_empty() {
        eprintln!("SKIPPED: test-cases/ not present");
        return;
    }
    let (_, failures) = audit_corpus("test-cases", "collinear gaps", collinear_gaps);
    assert!(
        failures.is_empty(),
        "border lines broken by a sub-width gap:\n{}",
        failures.join("\n")
    );
}

/// The reporter's document: "a still-missing cell corner on page 1, at the cell
/// labelled *Location GPS:*".
///
/// Measured off the rendered page-1 content stream, the notch is the square
/// x = 265.598…266.098, y = 206.152…206.652 — the right edge of the form's
/// spacer column crossed with the band under the short gutter row above it. The
/// x is grid-derived and fixed; the y is a sum of measured row heights and so is
/// a property of the host's fonts, which is why the assertion below is the
/// property over the whole document rather than that one square: on any host
/// where the notch exists at all, it is a junction, and the audit finds it.
#[test]
fn ip05_trenches_has_no_unpainted_border_junction() {
    let path = std::path::Path::new("test-cases/IP 05 Trenches_Bad Harzburg_03-06-2026.docx");
    if !path.exists() {
        eprintln!("SKIPPED: {} not present", path.display());
        return;
    }
    let bytes = std::fs::read(path).expect("read fixture");
    let doc = dxpdf::docx::parse(&bytes).expect("parse fixture");
    let (_, pages) = resolve_and_layout(doc);
    assert!(!pages.is_empty(), "expected at least one page");

    let mut total_checked = 0usize;
    let mut failures: Vec<String> = Vec::new();
    for (i, page) in pages.iter().enumerate() {
        let (checked, missing) = junctions(&border_rects(page));
        total_checked += checked;
        if !missing.is_empty() {
            failures.push(format!("page {}: {missing:?}", i + 1));
        }
    }

    // Non-vacuity: this document's tables are drawn with `Tabellenraster`, so
    // every page carries hundreds of junctions. A run that found none would
    // mean the rect filter above stopped matching, not that the borders are
    // sound.
    assert!(
        total_checked > 100,
        "expected the audit to find junctions to check, got {total_checked}"
    );
    assert!(
        failures.is_empty(),
        "border junctions painted by nobody:\n{}",
        failures.join("\n")
    );
}
