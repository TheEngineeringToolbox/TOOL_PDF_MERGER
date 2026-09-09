"""Regression tests using temporary files; Microsoft Office is not started."""

import tempfile
import unittest
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from app import find_attachment_folders, merge_pdfs, validate_main_docx


class AttachmentTests(unittest.TestCase):
    def test_attachment_and_file_order_and_ignored_files(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            for name in [
                "Bijlage A — Overzicht",
                "Bijlage 10 - Details",
                "Bijlage 2 – Plan",
            ]:
                folder = project / "Bijlage" / name
                folder.mkdir(parents=True)
                for filename in [
                    "10 Tekening.pdf",
                    "2 Tekening.pdf",
                    "~$draft.docx",
                    "desktop.ini",
                    "test.tmp",
                    "test.bak",
                ]:
                    (folder / filename).touch()
            attachments = find_attachment_folders(project)
            self.assertEqual([item[0] for item in attachments], ["2", "10", "A"])
            self.assertEqual(attachments[0][1], "Plan")
            self.assertEqual(
                [p.name for p in attachments[0][3]],
                ["2 Tekening.pdf", "10 Tekening.pdf"],
            )
            self.assertEqual(find_attachment_folders(project / "Bijlage"), attachments)

    def test_missing_attachment_root(self):
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(FileNotFoundError):
                find_attachment_folders(Path(directory))

    def test_unrecognized_attachment_folder(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "Bijlage" / "Onbekend").mkdir(parents=True)
            with self.assertRaises(RuntimeError):
                find_attachment_folders(root)

    def test_main_report_validation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with self.assertRaises(FileNotFoundError):
                validate_main_docx(root / "missing.docx")
            for filename in ["report.pdf", "~$report.docx"]:
                path = root / filename
                path.touch()
                with self.assertRaises(ValueError):
                    validate_main_docx(path)
            report = root / "Report.DOCX"
            report.touch()
            self.assertEqual(validate_main_docx(report), report.resolve())


class MergeTests(unittest.TestCase):
    def test_merge_preserves_page_order_and_dimensions(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            sources = []
            for index, widths in enumerate([[100, 200], [300]]):
                writer = PdfWriter()
                for width in widths:
                    writer.add_blank_page(width=width, height=400)
                source = root / f"{index}.pdf"
                with source.open("wb") as stream:
                    writer.write(stream)
                sources.append(source)
            output = root / "output" / "merged.pdf"
            merge_pdfs(sources, output)
            reader = PdfReader(output)
            self.assertEqual(
                [float(page.mediabox.width) for page in reader.pages], [100, 200, 300]
            )


if __name__ == "__main__":
    unittest.main()
