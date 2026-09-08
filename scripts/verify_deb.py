#!/usr/bin/env python3
"""Verify a built dxpdf .deb is shaped the way Debian expects.

The sibling of ``verify_wheel.py``, and deliberately the same shape: pure
standard library, parsing the container formats (``ar`` + ``tar``) directly, so
it runs on a developer's macOS box with no ``dpkg`` installed as well as inside
the CI container.

What it is guarding, in rough order of how expensive the failure is:

  1. **Nothing under /usr/lib.** The crate is ``crate-type = ["rlib",
     "cdylib", "staticlib"]``, so a release build also emits ``libdxpdf.so``
     and ``libdxpdf.a`` — the bodies of the PyO3 extension module and the Go
     bindings' cgo target, neither with a soname, headers or an ABI promise.
     cargo-deb's *default* asset set sweeps up C-ABI libraries, so the
     explicit ``assets`` list in Cargo.toml exists to keep both out of the
     package. This proves they stayed out.

  2. **No FreeType in Depends.** ``skia-safe[embed-freetype]`` static-links
     FreeType into Skia. If that feature is ever dropped, ``dpkg-shlibdeps``
     starts emitting ``libfreetype6`` and the package silently acquires a
     dependency on whatever FreeType the host has — the same class of bug
     ``verify_wheel.py`` was written for, which showed up there as "undefined
     symbol: FT_Palette_Data_Get" on older hosts.

  3. **A font package in Recommends.** dxpdf resolves fonts through fontconfig
     at render time, so what is installed on the machine decides what the PDF
     looks like. A bare install is not font-less — ``libfontconfig1`` pulls in
     ``fonts-dejavu-core`` — so the symptom is not blank text but the wrong
     face: measured on a bookworm container, a document asking for Courier New
     got DejaVu Sans, a proportional face substituted for a monospace one.
     Liberation is metric-compatible with the three faces most documents ask
     for, and recommending it is a packaging decision, so it is checked here.

  4. **The Policy-required furniture** — a man page for the binary (§12.1), a
     copyright file (§12.5), a changelog (§12.7). These are what a Debian
     reviewer looks for first.

Usage::

    python3 scripts/verify_deb.py target/debian/*.deb

The expected version defaults to the one in ``Cargo.toml``, so CI does not have
to extract it; pass ``--expect-version`` to override.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import sys
import tarfile
import tempfile
from pathlib import Path

# `parse_elf` is shared with the wheel checker rather than reimplemented: the
# FreeType invariant is the same one, and it should not be possible to fix it in
# one place and not the other. Both scripts live in this directory, which is on
# sys.path whenever either is run as a script.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from verify_wheel import parse_elf  # noqa: E402

#: Debian architectures we publish. `all` is deliberately absent — the package
#: contains a compiled binary and is emphatically not architecture-independent.
KNOWN_ARCHITECTURES = {"amd64", "arm64"}

#: Every path the package is allowed to contain, and each one is required.
#: Written as an exact set rather than a subset check so that a stray file is
#: as loud as a missing one.
EXPECTED_PAYLOAD = {
    "./usr/bin/dxpdf",
    "./usr/share/man/man1/dxpdf.1.gz",
    "./usr/share/doc/dxpdf/copyright",
    "./usr/share/doc/dxpdf/changelog.Debian.gz",
    "./usr/share/doc/dxpdf/README.md",
    "./usr/share/lintian/overrides/dxpdf",
}

#: Fields Debian Policy §5.3 requires in a binary package's control file, plus
#: the ones that make `apt show` useful.
REQUIRED_CONTROL_FIELDS = [
    "Package",
    "Version",
    "Architecture",
    "Maintainer",
    "Description",
    "Section",
    "Priority",
    "Homepage",
]


class Failure(Exception):
    """A verification failure, reported with the .deb it came from."""


def check(condition: bool, message: str) -> None:
    if not condition:
        raise Failure(message)


# --------------------------------------------------------------------------
# `ar` — the outer container. Six lines of format, so parsing it beats taking a
# dependency on `dpkg-deb` that would confine this script to Debian hosts.
# --------------------------------------------------------------------------


def read_ar(path: Path) -> list[tuple[str, bytes]]:
    """Return ``[(member_name, member_bytes), ...]`` in archive order."""
    data = path.read_bytes()
    check(data[:8] == b"!<arch>\n", f"{path.name} is not an ar archive")

    members: list[tuple[str, bytes]] = []
    offset = 8
    while offset + 60 <= len(data):
        header = data[offset : offset + 60]
        check(header[58:60] == b"`\n", f"{path.name}: corrupt ar header at byte {offset}")
        # Names are space-padded and conventionally terminated by `/`.
        name = header[0:16].decode("ascii", "replace").strip().rstrip("/")
        size = int(header[48:58].decode("ascii", "replace").strip())
        start = offset + 60
        members.append((name, data[start : start + size]))
        # Members are padded to an even offset.
        offset = start + size + (size % 2)
    return members


def open_tar(name: str, blob: bytes) -> tarfile.TarFile:
    """Open a ``control.tar.*`` / ``data.tar.*`` member.

    ``tarfile`` handles gzip and xz from the standard library. Zstd needs
    Python 3.14's ``compression.zstd``, so an older interpreter gets told what
    to do about it rather than a stack trace.
    """
    if name.endswith(".zst") and sys.version_info < (3, 14):
        raise Failure(
            f"{name} is zstd-compressed and this Python ({sys.version_info.major}."
            f"{sys.version_info.minor}) cannot read it. Build with "
            "`cargo deb --compress-type xz`, which is what CI uses."
        )
    return tarfile.open(fileobj=io.BytesIO(blob), mode="r:*")


def parse_control(text: str) -> dict[str, str]:
    """Parse a Debian control stanza (RFC 822-ish, with folded continuations)."""
    fields: dict[str, str] = {}
    key = ""
    for line in text.splitlines():
        if line[:1] in (" ", "\t") and key:
            fields[key] += "\n" + line.strip()
        elif ":" in line:
            key, _, value = line.partition(":")
            key = key.strip()
            fields[key] = value.strip()
    return fields


def dependency_names(field: str) -> list[str]:
    """Package names from a Depends/Recommends field, alternatives flattened.

    ``libc6 (>= 2.36), fonts-liberation2 | fonts-liberation`` becomes
    ``["libc6", "fonts-liberation2", "fonts-liberation"]``.
    """
    names = []
    for clause in field.split(","):
        for alternative in clause.split("|"):
            name = alternative.strip().split(" ")[0].split("(")[0].strip()
            if name:
                names.append(name)
    return names


# --------------------------------------------------------------------------
# The checks
# --------------------------------------------------------------------------


def verify(deb: Path, expect_version: str | None) -> None:
    check(deb.is_file(), f"{deb} not found")
    print(f"verifying {deb.name}")

    members = read_ar(deb)
    names = [name for name, _ in members]
    check(
        len(names) >= 3 and names[0] == "debian-binary",
        f"expected debian-binary first, got {names}",
    )
    check(
        names[1].startswith("control.tar") and names[2].startswith("data.tar"),
        f"expected control.tar.* then data.tar.*, got {names[1:3]}",
    )
    blobs = dict(members)
    check(
        blobs["debian-binary"] == b"2.0\n",
        f"unexpected deb format version {blobs['debian-binary']!r}",
    )

    control_tar = open_tar(names[1], blobs[names[1]])
    data_tar = open_tar(names[2], blobs[names[2]])

    _verify_control(control_tar, expect_version)
    _verify_payload(data_tar)
    _verify_md5sums(control_tar, data_tar)
    print(f"OK: {deb.name}")


def _verify_control(control_tar: tarfile.TarFile, expect_version: str | None) -> None:
    entry = control_tar.extractfile("./control")
    check(entry is not None, "control.tar has no ./control")
    fields = parse_control(entry.read().decode("utf-8"))

    missing = [f for f in REQUIRED_CONTROL_FIELDS if not fields.get(f)]
    check(not missing, f"control is missing required field(s): {missing}")

    check(fields["Package"] == "dxpdf", f"Package is {fields['Package']!r}, expected 'dxpdf'")
    check(
        fields["Architecture"] in KNOWN_ARCHITECTURES,
        f"Architecture {fields['Architecture']!r} is not one of {sorted(KNOWN_ARCHITECTURES)}",
    )
    if expect_version:
        # cargo-deb appends a Debian revision, so 0.4.0 becomes 0.4.0-1.
        check(
            fields["Version"].split("-")[0] == expect_version,
            f"Version is {fields['Version']!r}, expected upstream {expect_version}",
        )
    check(
        "@" in fields["Maintainer"] and "<" in fields["Maintainer"],
        f"Maintainer {fields['Maintainer']!r} is not a 'Name <address>' pair",
    )
    # A one-line Description is legal but useless in `apt show`; the extended
    # body is what a user reads before installing.
    check(
        "\n" in fields["Description"],
        "Description has no extended body (set `extended-description` in Cargo.toml)",
    )
    print(f"  {fields['Package']} {fields['Version']} {fields['Architecture']}")

    depends = dependency_names(fields.get("Depends", ""))
    print(f"  Depends ({len(depends)}): {', '.join(depends)}")
    freetype = [d for d in depends if "freetype" in d]
    check(
        not freetype,
        f"Depends names {freetype} — `embed-freetype` should have "
        "static-linked FreeType into Skia. See this script's module docstring.",
    )
    for required in ("libc6", "libfontconfig1"):
        check(
            required in depends,
            f"Depends does not name {required}; `depends = \"$auto\"` may not have run "
            "(dpkg-shlibdeps only works on a Debian host)",
        )

    recommends = dependency_names(fields.get("Recommends", ""))
    check(
        any(d.startswith("fonts-") for d in recommends),
        f"Recommends names no font package (got {recommends}) — on a minimal "
        "Debian install the converter would emit text-free PDFs",
    )
    print(f"  Recommends: {', '.join(recommends)}")


def _verify_payload(data_tar: tarfile.TarFile) -> None:
    entries = {m.name: m for m in data_tar.getmembers() if not m.isdir()}

    check(
        set(entries) == EXPECTED_PAYLOAD,
        "payload mismatch:\n"
        f"    unexpected: {sorted(set(entries) - EXPECTED_PAYLOAD)}\n"
        f"    missing:    {sorted(EXPECTED_PAYLOAD - set(entries))}",
    )
    # Redundant with the exact-set check above while EXPECTED_PAYLOAD stays as
    # it is, but it is the assertion that must survive someone adding a file to
    # that set for an unrelated reason.
    stray = [name for name in entries if name.startswith("./usr/lib")]
    check(not stray, f"package ships {stray} under /usr/lib — the cdylib must not be packaged")

    binary = entries["./usr/bin/dxpdf"]
    check(binary.isfile(), "./usr/bin/dxpdf is not a regular file")
    check(binary.mode & 0o777 == 0o755, f"./usr/bin/dxpdf has mode {binary.mode:o}, expected 755")

    with tempfile.TemporaryDirectory() as td:
        extracted = Path(td) / "dxpdf"
        source = data_tar.extractfile(binary)
        check(source is not None, "cannot read ./usr/bin/dxpdf out of data.tar")
        extracted.write_bytes(source.read())

        if extracted.read_bytes()[:4] != b"\x7fELF":
            # Building on macOS produces a Mach-O binary inside a Debian
            # container format. Everything above still holds; the linkage
            # checks below do not apply.
            print("  note: binary is not ELF (cross-format build) — skipping linkage checks")
            return

        needed, undef_ft = parse_elf(extracted)
        print(f"  DT_NEEDED ({len(needed)}): {', '.join(needed)}")
        ft_needed = [n for n in needed if "freetype" in n.lower()]
        check(not ft_needed, f"/usr/bin/dxpdf directly links {ft_needed}")
        check(
            not undef_ft,
            f"/usr/bin/dxpdf has {len(undef_ft)} unresolved FT_* symbols: "
            f"{', '.join(undef_ft[:5])}",
        )


def _verify_md5sums(control_tar: tarfile.TarFile, data_tar: tarfile.TarFile) -> None:
    """Check `md5sums` if there is one — a *stale* one is worse than none.

    cargo-deb 3.6.2 writes no `md5sums` at all: the control archive holds only
    `./control`. That costs `dpkg --verify` and `debsums` their ability to say
    whether an installed file has been altered, and lintian notes it as
    `no-md5sums-control-file` — at severity *info*, so it is not something the
    package is failing to do so much as something it is not doing. Fixing it
    means repacking the archive after cargo-deb, which is a lot of machinery
    for an info-level tag, so this is a soft check rather than a hard one: it
    reports the absence and verifies the contents if a future cargo-deb starts
    emitting them.
    """
    names = {m.name.lstrip("./") for m in control_tar.getmembers()}
    if "md5sums" not in names:
        print("  note: no md5sums in control.tar (cargo-deb does not write one)")
        return

    entry = control_tar.extractfile("./md5sums")
    check(entry is not None, "control.tar lists md5sums but it cannot be read")
    recorded = {}
    for line in entry.read().decode("utf-8").splitlines():
        digest, _, path = line.partition("  ")
        recorded[f"./{path}"] = digest

    payload = {m.name for m in data_tar.getmembers() if m.isfile()}
    check(
        set(recorded) == payload,
        f"md5sums does not cover the payload:\n"
        f"    unlisted: {sorted(payload - set(recorded))}\n"
        f"    phantom:  {sorted(set(recorded) - payload)}",
    )
    for name, digest in recorded.items():
        member = data_tar.extractfile(name)
        check(member is not None, f"md5sums lists {name}, which is not readable")
        actual = hashlib.md5(member.read()).hexdigest()
        check(actual == digest, f"md5sums for {name} is {digest}, actual {actual}")
    print(f"  md5sums: {len(recorded)} file(s), all matching")


def manifest_version() -> str | None:
    """The crate version from Cargo.toml, or None if it cannot be read.

    ``tomllib`` is standard library from Python 3.11, which is what bookworm
    ships; on anything older this returns None and the version check is skipped
    rather than the whole run failing over a default.
    """
    try:
        import tomllib
    except ImportError:
        return None
    manifest = Path(__file__).resolve().parent.parent / "Cargo.toml"
    try:
        with open(manifest, "rb") as f:
            return tomllib.load(f)["package"]["version"]
    except (OSError, KeyError):
        return None


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("debs", type=Path, nargs="+", help="path(s) to .deb files")
    ap.add_argument(
        "--expect-version",
        default=manifest_version(),
        help="upstream version to require (default: the version in Cargo.toml)",
    )
    args = ap.parse_args()
    if not args.expect_version:
        print("note: could not read a version from Cargo.toml; not checking Version")

    failures = 0
    for deb in args.debs:
        try:
            verify(deb, args.expect_version)
        except Failure as exc:
            print(f"FAIL: {deb.name}: {exc}", file=sys.stderr)
            failures += 1
    if failures:
        sys.exit(1)


if __name__ == "__main__":
    main()
