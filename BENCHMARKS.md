# Benchmark Results

Conversion benchmarks for `docx-pdf` using the test document:
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
| 2026-03-20 | latest | 84.1 ms ± 0.7 ms | 82.4 ms | 85.3 ms | 18.7 MB | +row heights, adjacent tables, hyphen breaks, natural line spacing |

### Methodology

- **Timing**: `hyperfine` with 3 warmup runs and 20 measured runs, output to `/dev/null`
- **Memory**: macOS `/usr/bin/time -l` reporting maximum resident set size
- **Build**: `cargo build --release` (optimized, no debug info)
