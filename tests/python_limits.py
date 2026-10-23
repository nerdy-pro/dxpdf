"""Run with python3 tests/python_limits.py /path/to/libdxpdf.so."""

import importlib.util
import io
from pathlib import Path
import sys
import tempfile
import unittest
import zipfile


def load_extension(path):
    spec = importlib.util.spec_from_file_location("dxpdf", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


dxpdf = load_extension(sys.argv.pop(1))


def document():
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as package:
        package.writestr(
            "_rels/.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="word/document.xml"/></Relationships>',
        )
        package.writestr(
            "word/document.xml",
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>",
        )
    return output.getvalue()


class ResourceLimitsTests(unittest.TestCase):
    def test_defaults_and_read_only_properties(self):
        limits = dxpdf.ResourceLimits()
        self.assertEqual(limits.max_input_bytes, 32 * 1024 * 1024)
        self.assertEqual(limits.max_zip_entries, 4096)
        self.assertEqual(limits.max_zip_entry_bytes, 64 * 1024 * 1024)
        self.assertEqual(limits.max_zip_total_bytes, 256 * 1024 * 1024)
        self.assertEqual(limits.max_image_pixels, 25_000_000)
        with self.assertRaises(AttributeError):
            limits.max_input_bytes = 1

    def test_invalid_limits_are_explicit_errors(self):
        for field in (
            "max_input_bytes",
            "max_zip_entries",
            "max_zip_entry_bytes",
            "max_zip_total_bytes",
            "max_image_pixels",
        ):
            for value in (-1, 2**64):
                with self.subTest(field=field, value=value):
                    with self.assertRaises(OverflowError):
                        dxpdf.ResourceLimits(**{field: value})

    def test_input_budget_precedes_parsing(self):
        with self.assertRaisesRegex(RuntimeError, "input bytes"):
            dxpdf.convert(b"not a zip", limits=dxpdf.ResourceLimits(max_input_bytes=8))

    def test_zip_budgets_reach_conversion(self):
        data = document()
        for field, resource in (
            ("max_zip_entries", "ZIP entries"),
            ("max_zip_entry_bytes", "ZIP entry bytes"),
            ("max_zip_total_bytes", "ZIP total bytes"),
        ):
            with self.subTest(field=field):
                with self.assertRaisesRegex(RuntimeError, resource):
                    dxpdf.convert(data, limits=dxpdf.ResourceLimits(**{field: 0}))

    def test_defaults_exact_input_boundary_and_positional_dpi(self):
        data = document()
        self.assertTrue(dxpdf.convert(data).startswith(b"%PDF-"))
        self.assertTrue(
            dxpdf.convert(
                data,
                96,
                limits=dxpdf.ResourceLimits(max_input_bytes=len(data)),
            ).startswith(b"%PDF-")
        )

    def test_file_budget_does_not_overwrite_output(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.docx"
            target = Path(directory) / "output.pdf"
            source.write_bytes(document())
            target.write_bytes(b"keep")
            for field, value, resource in (
                ("max_input_bytes", 8, "input bytes"),
                ("max_zip_entries", 0, "ZIP entries"),
                ("max_zip_entry_bytes", 0, "ZIP entry bytes"),
                ("max_zip_total_bytes", 0, "ZIP total bytes"),
            ):
                with self.subTest(field=field):
                    with self.assertRaisesRegex(RuntimeError, resource):
                        dxpdf.convert_file(
                            str(source),
                            str(target),
                            limits=dxpdf.ResourceLimits(**{field: value}),
                        )
                    self.assertEqual(target.read_bytes(), b"keep")

    def test_file_conversion_accepts_limits_and_dpi(self):
        data = document()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.docx"
            target = Path(directory) / "output.pdf"
            source.write_bytes(data)
            dxpdf.convert_file(
                str(source),
                str(target),
                96,
                limits=dxpdf.ResourceLimits(max_input_bytes=len(data)),
            )
            self.assertTrue(target.read_bytes().startswith(b"%PDF-"))

    def test_image_budget_reaches_painter(self):
        fixture = Path(__file__).resolve().parents[1] / "test-files/sample-docx-files-sample1.docx"
        limits = dxpdf.ResourceLimits(max_image_pixels=0)
        with self.assertRaisesRegex(RuntimeError, "image pixels"):
            dxpdf.convert(fixture.read_bytes(), limits=limits)
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "output.pdf"
            target.write_bytes(b"keep")
            with self.assertRaisesRegex(RuntimeError, "image pixels"):
                dxpdf.convert_file(str(fixture), str(target), limits=limits)
            self.assertEqual(target.read_bytes(), b"keep")


if __name__ == "__main__":
    unittest.main()
