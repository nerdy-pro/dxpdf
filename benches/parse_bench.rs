use criterion::{criterion_group, criterion_main, BenchmarkId, Criterion, Throughput};

const TEST_DIR: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/test-files");

const FILES: &[&str] = &[
    "sample-docx-files-sample1.docx",
    "sample-docx-files-sample2.docx",
    "sample-docx-files-sample3.docx",
    "sample-docx-files-sample4.docx",
    "sample-docx-files-sample-4.docx",
    "sample-docx-files-sample-5.docx",
    "sample-docx-files-sample-6.docx",
];

fn bench_parse(c: &mut Criterion) {
    let mut group = c.benchmark_group("parse");

    for name in FILES {
        let path = format!("{TEST_DIR}/{name}");
        let data = std::fs::read(&path).unwrap();
        let short = name
            .trim_start_matches("sample-docx-files-")
            .trim_end_matches(".docx");

        group.throughput(Throughput::Bytes(data.len() as u64));
        group.bench_with_input(BenchmarkId::new("file", short), &data, |b, data| {
            b.iter(|| dxpdf::docx::parse(data).unwrap());
        });
    }

    group.finish();
}

criterion_group!(benches, bench_parse);
criterion_main!(benches);
