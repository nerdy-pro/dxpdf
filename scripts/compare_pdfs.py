#!/usr/bin/env python3
"""Compare generated PDFs against reference (_real) PDFs pixel-by-pixel.

Usage:
    python3 scripts/compare_pdfs.py [--dpi 150] [--threshold 70] [--update-baseline]

Requires: pdftoppm (poppler), Pillow
Install: brew install poppler && pip3 install Pillow
"""

import subprocess
import os
import sys
import glob
import json
from datetime import datetime

try:
    from PIL import Image
except ImportError:
    print("Error: Pillow not installed. Run: pip3 install Pillow")
    sys.exit(1)


def rasterize_pdf(pdf_path, output_prefix, dpi=150):
    """Rasterize a PDF to PNG files, one per page."""
    result = subprocess.run(
        ["pdftoppm", "-png", "-r", str(dpi), pdf_path, output_prefix],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(f"  Error rasterizing {pdf_path}: {result.stderr}")
        return []
    parent = os.path.dirname(output_prefix)
    prefix = os.path.basename(output_prefix)
    return sorted([
        os.path.join(parent, f)
        for f in os.listdir(parent)
        if f.startswith(prefix + "-") and f.endswith(".png")
    ])


def compare_images(path_a, path_b, tolerance=30):
    """Compare two images pixel-by-pixel. Returns match percentage."""
    img_a = Image.open(path_a).convert("RGB")
    img_b = Image.open(path_b).convert("RGB")

    # Crop to common size (handles 1px rounding differences)
    w = min(img_a.width, img_b.width)
    h = min(img_a.height, img_b.height)
    img_a = img_a.crop((0, 0, w, h))
    img_b = img_b.crop((0, 0, w, h))

    pix_a = img_a.load()
    pix_b = img_b.load()
    total = w * h
    matching = 0

    for y in range(h):
        for x in range(w):
            ra, ga, ba = pix_a[x, y]
            rb, gb, bb = pix_b[x, y]
            if abs(ra - rb) + abs(ga - gb) + abs(ba - bb) <= tolerance:
                matching += 1

    return matching / total * 100 if total > 0 else 0.0


def compare_pdfs(generated_pdf, reference_pdf, dpi=150):
    """Compare two PDFs page-by-page. Returns list of per-page results."""
    tmp = "/tmp/pdf_compare"
    os.makedirs(tmp, exist_ok=True)

    gen_pages = rasterize_pdf(generated_pdf, os.path.join(tmp, "gen"), dpi)
    ref_pages = rasterize_pdf(reference_pdf, os.path.join(tmp, "ref"), dpi)

    results = []
    for i, (gp, rp) in enumerate(zip(gen_pages, ref_pages)):
        pct = compare_images(gp, rp)
        results.append({"page": i + 1, "match_pct": round(pct, 1)})

    page_count_match = len(gen_pages) == len(ref_pages)
    if not page_count_match:
        results.append({
            "page": 0,
            "match_pct": 0,
            "note": f"page count: {len(gen_pages)} vs {len(ref_pages)}",
        })

    # Cleanup
    for f in os.listdir(tmp):
        os.remove(os.path.join(tmp, f))

    avg = sum(r["match_pct"] for r in results if r["page"] > 0) / max(
        len([r for r in results if r["page"] > 0]), 1
    )

    return {
        "generated_pages": len(gen_pages),
        "reference_pages": len(ref_pages),
        "page_count_match": page_count_match,
        "average_match": round(avg, 1),
        "min_match": round(
            min((r["match_pct"] for r in results if r["page"] > 0), default=0), 1
        ),
        "pages": results,
    }


def find_test_pairs(test_dir="test-cases"):
    """Find pairs of generated and reference PDFs."""
    pairs = []
    for real_pdf in sorted(glob.glob(os.path.join(test_dir, "*_real.pdf"))):
        base = real_pdf.replace("_real.pdf", "")
        # Try common naming patterns
        candidates = [
            base + "_.pdf",
            base + ".pdf",
        ]
        generated = None
        for c in candidates:
            if os.path.exists(c):
                generated = c
                break
        if generated:
            name = os.path.basename(base)
            pairs.append((name, generated, real_pdf))
    return pairs


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Compare generated PDFs against references")
    parser.add_argument("--dpi", type=int, default=150, help="Rendering DPI (default: 150)")
    parser.add_argument("--threshold", type=float, default=70.0,
                        help="Minimum match %% to pass (default: 70)")
    parser.add_argument("--update-baseline", action="store_true",
                        help="Write results to VISUAL_COMPARISON.md")
    parser.add_argument("--json", action="store_true", help="Output JSON")
    args = parser.parse_args()

    pairs = find_test_pairs()
    if not pairs:
        print("No test pairs found (need *_real.pdf files in test-cases/)")
        sys.exit(1)

    all_results = {}
    all_pass = True

    for name, generated, reference in pairs:
        short_name = name[:40]
        if not args.json:
            print(f"\n{short_name}:")
        result = compare_pdfs(generated, reference, args.dpi)
        all_results[name] = result

        if args.json:
            continue

        pages_str = f"{result['generated_pages']}/{result['reference_pages']}"
        page_match = "ok" if result["page_count_match"] else "MISMATCH"
        print(f"  Pages: {pages_str} ({page_match})")
        print(f"  Average match: {result['average_match']}%")
        print(f"  Min match: {result['min_match']}%")

        for p in result["pages"]:
            if p["page"] == 0:
                print(f"  WARNING: {p.get('note', '')}")
            elif p["match_pct"] < args.threshold:
                print(f"  Page {p['page']}: {p['match_pct']}% BELOW THRESHOLD")

        if result["min_match"] < args.threshold:
            all_pass = False
            print(f"  FAIL: min match {result['min_match']}% < threshold {args.threshold}%")
        elif not result["page_count_match"]:
            all_pass = False
            print(f"  FAIL: page count mismatch")
        else:
            print(f"  PASS")

    if args.json:
        print(json.dumps(all_results, indent=2))
        return

    if args.update_baseline:
        write_comparison_report(all_results)

    print(f"\n{'='*50}")
    print(f"Overall: {'PASS' if all_pass else 'FAIL'}")

    if not all_pass:
        sys.exit(1)


def write_comparison_report(results):
    """Write comparison results to VISUAL_COMPARISON.md."""
    with open("VISUAL_COMPARISON.md", "w") as f:
        f.write("# Visual Comparison Results\n\n")
        f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n")
        f.write("| Document | Pages | Avg Match | Min Match | Status |\n")
        f.write("|---|---|---|---|---|\n")
        for name, result in results.items():
            pages = f"{result['generated_pages']}/{result['reference_pages']}"
            status = "PASS" if result["min_match"] >= 70 and result["page_count_match"] else "FAIL"
            f.write(f"| {name[:40]} | {pages} | {result['average_match']}% | {result['min_match']}% | {status} |\n")
        f.write(f"\nThreshold: 70% pixel match (tolerance: 30/255 per channel)\n")
    print("\nWrote VISUAL_COMPARISON.md")


if __name__ == "__main__":
    main()
