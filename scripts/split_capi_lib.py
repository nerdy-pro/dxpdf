#!/usr/bin/env python3
"""Split an oversized go/internal/capi/lib/<platform>/libdxpdf.a into several
libdxpdf_partN.a archives, each safely under GitHub's 100 MB per-file limit.

Why this exists: linux/amd64 and linux/arm64's static libraries come out of
`cargo build --release --features capi` at ~120-127 MB — over the limit, and
neither symbol-stripping (~3-7%) nor LTO (which made it *larger*, since a
`staticlib`'s consumer is a foreign linker that can't exploit the embedded
bitcode) closes anywhere near that gap. The size is ~2,150 object files
spread across Skia/HarfBuzz/ICU/fontcull with no single prunable component,
so there is nothing to cut without dropping real functionality.

An `ar` archive is just a container of `.o` members with a symbol index; a
linker resolving `-lfoo` doesn't care whether all of a library's members
live in one archive file or several, as long as it can find them. So this
splits the members across multiple smaller archives instead, each under the
GitHub limit with real margin (see CAP_BYTES). The corresponding
`go/cgo_linux_*.go` files link every part inside a single `-Wl,--start-group
... --end-group`, which is required and not optional: GNU ld resolves
archives in one left-to-right pass per archive by default, and splitting
what was one archive scatters mutually-referencing object files across the
new ones (e.g. one Skia class's definition ends up in part 1, a method that
calls back into it in part 2) — `--start-group`/`--end-group` makes ld keep
re-scanning the group until nothing new resolves, which a single unsplit
archive never needed. macOS's static libraries (darwin_amd64/darwin_arm64)
are under the limit as built and are never split — this is Linux-only both
because ld64 doesn't use the same archive-once-per-scan model and because
there's no need to route around a problem that doesn't exist there.

Verified end-to-end (built the split archives, linked and ran `go test
./... -race` against them for real) before this script was written from
that working process — see AGENTS.md's Go-bindings note.

Requires GNU `ar`/`ranlib` (binutils) — i.e. run this on Linux or in a
Linux container, matching the platform whose archive is being split. macOS
`ar`/ranlib operate on Mach-O and won't produce a usable ELF archive.

Usage:

    python3 scripts/split_capi_lib.py go/internal/capi/lib/linux_amd64/libdxpdf.a

Replaces the input `libdxpdf.a` with `libdxpdf_part1.a`, `libdxpdf_part2.a`,
... in the same directory (deletes the unsplit original). Re-run
`UPDATE_CAPI_LIB_HASH=1 cargo test --test capi_lib_freshness` and commit
alongside — the split doesn't touch the Rust source the hash covers, but
`every_supported_platform_has_a_committed_library` checks for exactly this
shape (`libdxpdf.a` or `libdxpdf_part*.a`) per platform.
"""
from __future__ import annotations

import collections
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# Target *member payload* per part, before `ar`'s own per-member headers and
# symbol index are added back in. Measured overhead on the linux/amd64
# archive was ~8% at this size — 55 MiB of members came out to a 60-63 MB
# archive — so this leaves real margin under the 100 MB limit rather than
# aiming right at it.
CAP_BYTES = 55 * 1024 * 1024


def run(*args: str) -> None:
    subprocess.run(args, check=True)


def split(archive: Path) -> list[Path]:
    with tempfile.TemporaryDirectory(prefix="split_capi_lib_") as tmp:
        extract_dir = Path(tmp)
        members = subprocess.run(
            ["ar", "t", str(archive)], capture_output=True, text=True, check=True
        ).stdout.splitlines()
        if not members:
            raise SystemExit(f"{archive}: no members found — not a valid archive?")

        # Member names collide across a Rust archive with ~2000 objects
        # (e.g. two vendored copies of the same upstream file), so members
        # are extracted by *occurrence* (ar's `-N` modifier), not by name,
        # and immediately renamed to a globally unique, order-preserving
        # filename. The index prefix is what preserves original order once
        # everything lands in one flat directory.
        seen: collections.Counter[str] = collections.Counter()
        extracted: list[tuple[Path, int]] = []
        for index, name in enumerate(members):
            seen[name] += 1
            run("ar", "x", "-N", str(seen[name]), str(archive), name, "--output", str(extract_dir))
            member_path = extract_dir / name
            unique_path = extract_dir / f"{index:05d}__{name}"
            member_path.rename(unique_path)
            extracted.append((unique_path, unique_path.stat().st_size))

        parts: list[list[Path]] = [[]]
        part_size = [0]
        for path, size in extracted:
            if part_size[-1] + size > CAP_BYTES and parts[-1]:
                parts.append([])
                part_size.append(0)
            parts[-1].append(path)
            part_size[-1] += size

        stem = archive.stem  # "libdxpdf"
        out_paths = []
        for i, group in enumerate(parts, start=1):
            out_path = archive.parent / f"{stem}_part{i}{archive.suffix}"
            # A response file, not argv, since a part can hold hundreds of
            # members — comfortably past comfortable command-line lengths.
            rsp = extract_dir / f"part{i}.rsp"
            rsp.write_text("\n".join(str(p) for p in group))
            run("ar", "rcs", str(out_path), f"@{rsp}")
            run("ranlib", str(out_path))
            out_paths.append(out_path)
        return out_paths


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit(f"usage: {sys.argv[0]} <path/to/libdxpdf.a>")
    archive = Path(sys.argv[1]).resolve()
    if not archive.is_file():
        raise SystemExit(f"{archive}: not found")
    if shutil.which("ar") is None or shutil.which("ranlib") is None:
        raise SystemExit("ar/ranlib not found — run this on Linux (or in a Linux container)")

    out_paths = split(archive)
    archive.unlink()
    print(f"replaced {archive} with:")
    for p in out_paths:
        print(f"  {p}  ({p.stat().st_size / 1024 / 1024:.1f} MiB)")


if __name__ == "__main__":
    main()
