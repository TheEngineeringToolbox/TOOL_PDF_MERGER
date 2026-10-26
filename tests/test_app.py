"""Regression tests using temporary files; Microsoft Office is not started."""

import tempfile
import unittest
import zipfile
from pathlib import Path

from lxml import etree as ET
from pypdf import PdfReader, PdfWriter

from app import (
    NS,
    TOOL_FIELDS_NS,
    W_NS,
    add_cover_metadata,
    align_cover_bars,
    find_attachment_folders,
    merge_pdfs,
    read_report_metadata,
    validate_main_docx,
    validate_report_metadata,
)


class MetadataTests(unittest.TestCase):
    def test_cover_bars_center_on_page_without_moving_small_artwork(self):
        wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        root = ET.fromstring(
            f'<w:hdr xmlns:w="{W_NS}" xmlns:wp="{wp}"><w:p><w:r><w:drawing>'
            '<wp:anchor><wp:positionH relativeFrom="page"><wp:posOffset>137795</wp:posOffset>'
            '</wp:positionH><wp:extent cx="7200000" cy="486000"/></wp:anchor>'
            '<wp:anchor><wp:positionH relativeFrom="margin"><wp:posOffset>42</wp:posOffset>'
            '</wp:positionH><wp:extent cx="100000" cy="100000"/></wp:anchor>'
            "</w:drawing></w:r></w:p></w:hdr>"
        )
        result = ET.fromstring(align_cover_bars(ET.tostring(root), 11906))
        positions = result.findall(f".//{{{wp}}}positionH")
        self.assertEqual(positions[0].get("relativeFrom"), "page")
        self.assertEqual(positions[0].find(f"{{{wp}}}align").text, "center")
        self.assertIsNone(positions[0].find(f"{{{wp}}}posOffset"))
        self.assertEqual(positions[1].find(f"{{{wp}}}posOffset").text, "42")

    def make_report(self, path, fields, bound=True):
        with zipfile.ZipFile(path, "w") as archive:
            for index, (store_id, values) in enumerate(fields, 1):
                root = ET.Element(f"{{{TOOL_FIELDS_NS}}}mydata")
                for key, value in values.items():
                    ET.SubElement(root, f"{{{TOOL_FIELDS_NS}}}{key}").text = value
                archive.writestr(f"customXml/item{index}.xml", ET.tostring(root))
                archive.writestr(
                    f"customXml/_rels/item{index}.xml.rels",
                    '<Relationships><Relationship Type="test/customXmlProps" '
                    f'Target="itemProps{index}.xml"/></Relationships>',
                )
                archive.writestr(
                    f"customXml/itemProps{index}.xml",
                    '<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/'
                    f'officeDocument/2006/customXml" ds:itemID="{store_id}"/>',
                )
            root = ET.Element(f"{{{W_NS}}}document", nsmap={"w": W_NS})
            if bound:
                binding = ET.SubElement(root, f"{{{W_NS}}}dataBinding")
                binding.set(f"{{{W_NS}}}prefixMappings", f"xmlns:t='{TOOL_FIELDS_NS}'")
                binding.set(f"{{{W_NS}}}storeItemID", "active")
            archive.writestr("word/document.xml", ET.tostring(root))

    def test_bound_store_wins_over_old_template_values_and_keeps_zero(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.docx"
            self.make_report(
                path,
                [
                    ("old", {"Documentnummer": "STALE", "Documentversienummer": "9"}),
                    (
                        "active",
                        {
                            "DocumentNummer": "T-250801-013",
                            "DocumentVersienummer": "0",
                            "DocumentVersiedatum": "",
                        },
                    ),
                ],
            )
            self.assertEqual(
                read_report_metadata(path),
                {
                    "kenmerk": "T-250801-013",
                    "versie": "0",
                    "datum": "",
                },
            )

    def test_missing_or_ambiguous_xml_does_not_use_example_values(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.docx"
            self.make_report(path, [])
            self.assertEqual(
                read_report_metadata(path),
                dict.fromkeys(("kenmerk", "versie", "datum"), ""),
            )
            self.make_report(path, [("old", {"Documentnummer": "STALE"})])
            self.assertEqual(read_report_metadata(path)["kenmerk"], "")
            self.make_report(
                path,
                [("a", {"Documentnummer": "A"}), ("b", {"Documentnummer": "B"})],
                bound=False,
            )
            self.assertEqual(read_report_metadata(path)["kenmerk"], "")

    def test_date_validation_and_required_fields(self):
        data = {
            "kenmerk": "T-250801-013",
            "versie": "0",
            "datum": "2026-10-12T00:00:00",
        }
        self.assertEqual(validate_report_metadata(data)["datum"], "12-10-2026")
        for value in ("", "31-02-2026"):
            with self.assertRaises(ValueError):
                validate_report_metadata({**data, "datum": value})

    def test_header_preserves_artwork_and_adds_two_lines(self):
        header = f'<w:hdr xmlns:w="{W_NS}"><w:p><w:r><w:drawing/></w:r></w:p></w:hdr>'
        data = {"kenmerk": "T-250801-013 & X", "versie": "v0", "datum": "12-10-2026"}
        root = ET.fromstring(add_cover_metadata(header.encode(), data))
        self.assertEqual(
            root.xpath("//w:t/text()", namespaces=NS),
            ["T-250801-013 & X v0", "dd. 12-10-2026"],
        )
        self.assertEqual(len(root.xpath("//w:drawing", namespaces=NS)), 1)
        self.assertEqual(len(root.xpath("//w:br", namespaces=NS)), 1)


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
