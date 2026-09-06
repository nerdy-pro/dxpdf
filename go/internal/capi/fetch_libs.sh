#!/usr/bin/env bash
# Downloads the prebuilt static library for the current (or requested)
# GOOS/GOARCH from this repo's GitHub Release assets, so `go build`/`go
# test` in the parent module has something for cgo to link against.
#
# Go modules have no build-script hook that could do this automatically at
# `go build` time (unlike Cargo's build.rs) — this is the same tradeoff
# other cgo wrapper packages around a large native library make (e.g.
# wasmtime-go): a `go generate` step the consumer runs once, not a network
# fetch hidden inside the build.
#
# Usage:
#
#     go generate ./...
#
# or, to fetch for a different target than the host (e.g. cross-compiling):
#
#     GOOS=linux GOARCH=arm64 go generate ./...
set -euo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")"

GOOS="${GOOS:-$(go env GOOS)}"
GOARCH="${GOARCH:-$(go env GOARCH)}"

# Matches the `label` column in .github/workflows/go.yml's build matrix.
case "${GOOS}-${GOARCH}" in
  linux-amd64) LABEL=linux-x86_64 ;;
  linux-arm64) LABEL=linux-aarch64 ;;
  darwin-arm64) LABEL=macos-arm64 ;;
  darwin-amd64) LABEL=macos-x86_64 ;;
  *)
    echo "fetch_libs.sh: no prebuilt dxpdf static library for ${GOOS}/${GOARCH}" >&2
    echo "(Windows is not yet supported by the Go bindings — see AGENTS.md)" >&2
    exit 1
    ;;
esac

DEST="lib/${GOOS}_${GOARCH}"
if [ -f "${DEST}/libdxpdf.a" ]; then
  echo "fetch_libs.sh: ${DEST}/libdxpdf.a already present, skipping download"
  exit 0
fi

VERSION="$(sed -nE 's/^const releaseVersion = "(.*)"$/\1/p' ../../dxpdf.go)"
if [ -z "${VERSION}" ]; then
  echo "fetch_libs.sh: couldn't read releaseVersion from go/dxpdf.go" >&2
  exit 1
fi

# Same GitHub Release the crate itself is published from (tag `vX.Y.Z`) —
# `.github/workflows/go.yml` attaches these tarballs to it alongside the
# .deb packages `deb.yml` attaches. Not the separate `go/vX.Y.Z` *git* tag,
# which exists only so `go get .../dxpdf/go@vX.Y.Z` resolves a commit; it
# carries no release assets of its own.
URL="https://github.com/nerdy-pro/dxpdf/releases/download/v${VERSION}/libdxpdf-go-${LABEL}.tar.gz"
echo "fetch_libs.sh: downloading ${URL}"
mkdir -p "${DEST}"
curl -fsSL "${URL}" | tar xz -C "${DEST}"
echo "fetch_libs.sh: wrote ${DEST}/libdxpdf.a"
