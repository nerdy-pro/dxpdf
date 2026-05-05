#!/usr/bin/env python3
"""Verify a built dxpdf wheel has FreeType properly embedded.

With ``skia-safe[embed-freetype]`` enabled, FreeType is statically linked
into Skia. The dxpdf extension module must therefore:

  1. not list ``libfreetype.so.6`` (Linux) or any ``libfreetype`` dylib
     (macOS) as a *direct* dynamic dependency, and
  2. contain no UNDEFINED ``FT_*`` symbols in its dynamic symbol table.

A bundled libfreetype may still appear in ``dxpdf.libs/`` because
auditwheel/delocate also pulls in transitive dependencies — fontconfig
links libfreetype, so the loader will load one regardless. That's
harmless as long as the extension itself does not reference FT symbols.

Re-introducing direct FreeType linkage brings back the "undefined symbol:
FT_Palette_Data_Get" error on hosts with older system FreeType.
"""
from __future__ import annotations

import argparse
import struct
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


def parse_elf(path: Path) -> tuple[list[str], list[str]]:
    """Return (DT_NEEDED libraries, undefined FT_* symbols) from an ELF64 file."""
    data = path.read_bytes()
    if data[:4] != b"\x7fELF" or data[4] != 2:
        raise ValueError(f"{path} is not ELF64")
    e_shoff = struct.unpack_from("<Q", data, 0x28)[0]
    e_shentsize = struct.unpack_from("<H", data, 0x3A)[0]
    e_shnum = struct.unpack_from("<H", data, 0x3C)[0]
    e_shstrndx = struct.unpack_from("<H", data, 0x3E)[0]

    sections = []
    for i in range(e_shnum):
        off = e_shoff + i * e_shentsize
        sh_name = struct.unpack_from("<I", data, off)[0]
        sh_offset = struct.unpack_from("<Q", data, off + 0x18)[0]
        sh_size = struct.unpack_from("<Q", data, off + 0x20)[0]
        sections.append((sh_name, sh_offset, sh_size))

    shstr = data[sections[e_shstrndx][1] : sections[e_shstrndx][1] + sections[e_shstrndx][2]]

    def name_at(idx: int) -> str:
        return shstr[idx : shstr.find(b"\x00", idx)].decode("ascii", "replace")

    by_name = {name_at(s[0]): s for s in sections}
    dyn = by_name.get(".dynamic")
    dynstr = by_name.get(".dynstr")
    dynsym = by_name.get(".dynsym")
    if not (dyn and dynstr and dynsym):
        raise ValueError(f"{path} missing dynamic linking sections")

    dynstr_data = data[dynstr[1] : dynstr[1] + dynstr[2]]

    needed = []
    for i in range(0, dyn[2], 16):
        d_tag, d_val = struct.unpack_from("<qQ", data, dyn[1] + i)
        if d_tag == 0:
            break
        if d_tag == 1:  # DT_NEEDED
            end = dynstr_data.find(b"\x00", d_val)
            needed.append(dynstr_data[d_val:end].decode("ascii", "replace"))

    undef_ft = []
    for i in range(dynsym[2] // 24):
        off = dynsym[1] + i * 24
        st_name = struct.unpack_from("<I", data, off)[0]
        st_shndx = struct.unpack_from("<H", data, off + 0x06)[0]
        if st_shndx == 0:  # SHN_UNDEF
            end = dynstr_data.find(b"\x00", st_name)
            sym = dynstr_data[st_name:end].decode("ascii", "replace")
            if sym.startswith("FT_"):
                undef_ft.append(sym)
    return needed, undef_ft


def macos_otool_deps(dylib: Path) -> list[str]:
    out = subprocess.check_output(["otool", "-L", str(dylib)], text=True)
    deps = []
    for line in out.splitlines()[1:]:
        line = line.strip()
        if line:
            deps.append(line.split(" ", 1)[0])
    return deps


def is_elf(path: Path) -> bool:
    with open(path, "rb") as f:
        return f.read(4) == b"\x7fELF"


def fail(msg: str) -> None:
    print(f"FAIL: {msg}", file=sys.stderr)
    sys.exit(1)


def verify(wheel: Path) -> None:
    if not wheel.is_file():
        fail(f"{wheel} not found")

    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        with zipfile.ZipFile(wheel) as zf:
            zf.extractall(root)

        ext_candidates = [
            p
            for p in [*root.rglob("dxpdf*.so"), *root.rglob("dxpdf*.dylib")]
            if "dxpdf.libs" not in p.parts
        ]
        if not ext_candidates:
            fail(f"no dxpdf extension module found in {wheel.name}")
        ext = ext_candidates[0]

        print(f"verifying {wheel.name} :: {ext.relative_to(root)}")

        if is_elf(ext):
            needed, undef_ft = parse_elf(ext)
            print(f"  DT_NEEDED ({len(needed)}): {', '.join(needed)}")
            ft_needed = [n for n in needed if "freetype" in n.lower()]
            if ft_needed:
                fail(f"{ext.name} directly links {ft_needed}")
            if undef_ft:
                shown = ", ".join(undef_ft[:5])
                more = f" (+{len(undef_ft) - 5} more)" if len(undef_ft) > 5 else ""
                fail(f"{ext.name} has {len(undef_ft)} unresolved FT_* symbols: {shown}{more}")
        elif sys.platform == "darwin":
            deps = macos_otool_deps(ext)
            ft_deps = [d for d in deps if "freetype" in d.lower()]
            if ft_deps:
                fail(f"{ext.name} directly links {ft_deps}")
        else:
            fail(f"unsupported binary format for {ext}")

        bundled = sorted(p.relative_to(root) for p in root.rglob("libfreetype*"))
        if bundled:
            print(f"  note: wheel bundles libfreetype (transitive via fontconfig): {bundled[0]}")

        print(f"OK: {wheel.name} — no direct FreeType linkage")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("wheels", type=Path, nargs="+", help="path(s) to .whl files")
    args = ap.parse_args()
    for w in args.wheels:
        verify(w)


if __name__ == "__main__":
    main()
