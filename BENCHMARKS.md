# Benchmark Results

Conversion benchmarks for `dxpdf` using the test document:
**Protokoll test DIN VDE 0100-600 Erstinbetriebnahme.docx**

## Test Document

| Property | Value |
|---|---|
| Input file | Protokoll test DIN VDE 0100-600 Erstinbetriebnahme.docx |
| Input size | 41 KB |
| Output size | 572 KB |
| Pages | 3 (2 portrait A4 + 1 landscape A4) |
| Tables | 11 |
| Images | 2 PNG |
| Sections | 2 (portrait → landscape) |

## System

| Property | Value |
|---|---|
| CPU | Apple M3 Max |
| OS | macOS 26.3.1 |
| Rust | 1.93.1 |
| Binary size | 4.7 MB (release) |

## Results

| Date | Commit | Mean time | Min | Max | Peak RSS | Notes |
|---|---|---|---|---|---|---|
| 2026-03-20 | initial | 88.9 ms ± 2.2 ms | 84.9 ms | 91.9 ms | 19.3 MB | Baseline with Skia text measurement, Carlito font substitution |
| 2026-03-20 | borders | 85.0 ms ± 3.1 ms | 80.4 ms | 90.1 ms | 18.8 MB | +table borders, cell margins, font subs, units.rs, code review refactor |
| 2026-03-20 | spacing | 84.1 ms ± 0.7 ms | 82.4 ms | 85.3 ms | 18.7 MB | +row heights, adjacent tables, hyphen breaks, natural line spacing |
| 2026-03-20 | shading | 85.4 ms ± 3.5 ms | 83.0 ms | 98.9 ms | 18.8 MB | +cell shading/backgrounds, DrawCommand::Rect |
| 2026-03-20 | spaces | 89.6 ms ± 0.8 ms | 88.2 ms | 90.8 ms | 18.8 MB | +space handling, render order fix (shading→content→borders), hyphen wrapping |
| 2026-03-20 | hdr/ftr | 112.3 ms ± 2.1 ms | 109.7 ms | 119.5 ms | 19.0 MB | +headers/footers, signed EMU offsets, float align, header extent push-down |
| 2026-03-20 | review | 114.4 ms ± 2.4 ms | 111.7 ms | 120.5 ms | 19.0 MB | Code review: constants, DocDefaultsLayout factory, private newtype fields, Deref impls |
| 2026-03-20 | arc | 113.7 ms ± 1.8 ms | 112.3 ms | 119.2 ms | 19.0 MB | Arc\<Vec\<u8\>\> for image data — avoids deep cloning in layout pipeline |
| 2026-03-20 | latest | 114.0 ms ± 1.3 ms | 111.0 ms | 117.8 ms | 18.8 MB | Rc\<Vec\<u8\>\> for images, Rc\<str\> for font families — no atomic overhead |
| 2026-03-21 | pipeline | 118.4 ms ± 1.0 ms | 116.1 ms | 120.0 ms | 19.1 MB | Measure→layout→paint pipeline for all elements, shared measure_lines, after-table spacing fix, 86 unit tests |
| 2026-03-21 | latest | 113.3 ms ± 1.4 ms | 111.9 ms | 118.0 ms | 19.3 MB | +hyperlinks (PDF link annotations), superscript/subscript, pctPosVOffset, unsupported feature warnings, 98 unit tests |
| 2026-03-21 | v0.1.3 | 114.2 ms ± 0.6 ms | 113.3 ms | 115.3 ms | 19.3 MB | +char styles, paragraph borders, font-first resolution, field codes, 104 unit tests |

### Methodology

- **Timing**: `hyperfine` with 3 warmup runs and 20 measured runs, output to `/dev/null`
- **Memory**: macOS `/usr/bin/time -l` reporting maximum resident set size
- **Build**: `cargo build --release` (optimized, no debug info)
