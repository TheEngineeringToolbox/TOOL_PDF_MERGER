from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile
import time
import tkinter as tk
import zipfile
from copy import deepcopy
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

try:
    import win32com.client as win32
except ImportError:
    win32 = None

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader = PdfWriter = None

from lxml import etree as ET

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

COMPANY_NAME = "TOOL Engineers B.V."
APP_NAME = "Rapportage Merger Tool"
APP_VERSION = "0.0.3"
ATTACHMENT_DIR_NAME = "Bijlage"
TEMPLATE_PATH = Path(r"C:\TOOL\Templates\Rapportage_merge_tool\Bijlage voorbladen.docx")

LOGO_PATH = Path(
    r"D:\OD\TOOL Engineers BV\TOOL - Documenten\00 - TOOL\01 - PR TOOL\01 - Logo\TOOL Engineers 15-07-2024.png"
)
WD_EXPORT_FORMAT_PDF = 17
WD_FORMAT_DOCUMENT_DEFAULT = 16
WD_DO_NOT_SAVE_CHANGES = 0
WD_ALERTS_NONE = 0

SUPPORTED_OFFICE = {
    ".doc": "word",
    ".docx": "word",
    ".xls": "excel",
    ".xlsx": "excel",
    ".xlsm": "excel",
    ".ppt": "powerpoint",
    ".pptx": "powerpoint",
}

IGNORED_NAMES = {
    "desktop.ini",
    "thumbs.db",
}


# ---------------------------------------------------------------------------
# General helpers
# ---------------------------------------------------------------------------


def natural_key(path: Path):
    parts = re.split(r"(\d+)", path.name.lower())
    return [int(p) if p.isdigit() else p for p in parts]


def get_template_path():
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            "Het centrale voorblad-template is niet gevonden:\n\n"
            f"{TEMPLATE_PATH}\n\n"
            "Controleer of de centrale locatie bereikbaar is."
        )

    return TEMPLATE_PATH


def clean_attachment_title(filename: str) -> str:
    """
    Turns:
        '01 Overzicht Google Earth.pdf'
        '2 - Berekening.pdf'
    into:
        'Overzicht Google Earth'
        'Berekening'
    """
    stem = Path(filename).stem.strip()
    stem = re.sub(r"^\s*\d+\s*[-_.:)]\s*", "", stem)
    stem = re.sub(r"^\s*\d+\s+", "", stem)
    return stem.strip()


def validate_main_docx(main_docx: Path) -> Path:
    """Validate the DOCX selected by the user."""
    main_docx = Path(main_docx).resolve()

    if not main_docx.exists():
        raise FileNotFoundError(
            f"Het geselecteerde DOCX-bestand bestaat niet:\n{main_docx}"
        )

    if main_docx.suffix.lower() != ".docx":
        raise ValueError("Het gekozen hoofdrapport moet een .docx-bestand zijn.")

    if main_docx.name.startswith("~$"):
        raise ValueError("Het gekozen bestand is een tijdelijk Word-bestand.")

    return main_docx


def attachment_sort_key(identifier: str):
    s = str(identifier).strip()
    if s.isdigit():
        return (0, int(s))
    if len(s) == 1 and s.isalpha():
        return (1, s.upper())
    return (2, s.lower())


def find_attachment_folders(project_dir: Path):
    """
    Supports:
      <project>\\Bijlage\\Bijlage A - Titel
      <project>\\Bijlage\\Bijlage 1 - Titel

    Hyphen variants -, – and — are accepted.
    """
    selected = project_dir.resolve()

    # Find the Bijlage root.
    if selected.name.lower() == ATTACHMENT_DIR_NAME.lower():
        root = selected
    else:
        root = selected / ATTACHMENT_DIR_NAME
        if not root.exists():
            candidates = [
                p
                for p in selected.iterdir()
                if p.is_dir() and p.name.lower() == ATTACHMENT_DIR_NAME.lower()
            ]
            if candidates:
                root = candidates[0]

    if not root.exists() or not root.is_dir():
        raise FileNotFoundError(
            "Kan de map 'Bijlage' niet vinden.\n\n"
            "Verwachte structuur:\n"
            "<project>\\Rapportage.docx\n"
            "<project>\\Bijlage\\Bijlage A - Locatie overzicht\n\n"
            f"Gekozen map:\n{selected}"
        )

    result = []

    # Accept ASCII hyphen, en dash and em dash.
    pattern = re.compile(
        r"^\s*Bijlage\s+([A-Za-z0-9]+)\s*[-–—]\s*(.+?)\s*$",
        re.IGNORECASE,
    )

    for folder in root.iterdir():
        if not folder.is_dir():
            continue

        # Windows can contain a non-breaking space copied from documents.
        folder_name = folder.name.replace("\u00a0", " ").strip()
        match = pattern.match(folder_name)
        if not match:
            continue

        identifier = match.group(1).strip()
        title = match.group(2).strip()

        files = [
            p
            for p in folder.iterdir()
            if p.is_file()
            and not p.name.startswith("~$")
            and p.name.lower() not in IGNORED_NAMES
            and p.suffix.lower() not in {".tmp", ".bak"}
        ]
        files.sort(key=natural_key)

        result.append((identifier, title, folder, files))

    result.sort(key=lambda x: attachment_sort_key(x[0]))

    if not result:
        found_dirs = [p.name for p in root.iterdir() if p.is_dir()]
        listing = "\n".join(f" - {name}" for name in found_dirs) or " (geen submappen)"
        raise RuntimeError(
            "Geen bijlagemappen herkend in:\n"
            f"{root}\n\n"
            "Verwacht bijvoorbeeld:\n"
            "  Bijlage A - Locatie overzicht\n"
            "  Bijlage B - Tekeningen\n\n"
            "Submappen die daadwerkelijk zijn gevonden:\n"
            f"{listing}"
        )

    return result


# ---------------------------------------------------------------------------
# Word/Excel/PowerPoint conversion using Microsoft Office COM
# ---------------------------------------------------------------------------


class OfficeConverter:
    def __init__(self):
        if win32 is None:
            raise RuntimeError(
                "pywin32 ontbreekt. Installeer de dependencies met:\n"
                "pip install -r requirements.txt"
            )
        self.word = None
        self.excel = None
        self.powerpoint = None

    def start(self):
        # DispatchEx prevents us from taking over the user's existing Office app.
        self.word = win32.DispatchEx("Word.Application")
        self.word.Visible = False
        self.word.DisplayAlerts = WD_ALERTS_NONE

        self.excel = win32.DispatchEx("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False

        self.powerpoint = win32.DispatchEx("PowerPoint.Application")

    def stop(self):
        for app in (self.powerpoint, self.excel, self.word):
            if app is None:
                continue
            try:
                app.Quit()
            except Exception:
                pass

        self.powerpoint = None
        self.excel = None
        self.word = None

    def word_to_pdf(self, source: Path, target: Path):
        doc = None
        try:
            doc = self.word.Documents.Open(
                str(source.resolve()),
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )
            doc.ExportAsFixedFormat(
                str(target.resolve()),
                WD_EXPORT_FORMAT_PDF,
                False,  # OpenAfterExport
            )
        finally:
            if doc is not None:
                try:
                    doc.Close(WD_DO_NOT_SAVE_CHANGES)
                except Exception:
                    pass

    def excel_to_pdf(self, source: Path, target: Path):
        book = None
        try:
            book = self.excel.Workbooks.Open(
                str(source.resolve()),
                ReadOnly=True,
                UpdateLinks=0,
                AddToMru=False,
            )
            book.ExportAsFixedFormat(
                WD_EXPORT_FORMAT_PDF,
                str(target.resolve()),
                0,  # quality: standard
                True,  # include doc properties
                False,  # ignore print areas
            )
        finally:
            if book is not None:
                try:
                    book.Close(False)
                except Exception:
                    pass

    def powerpoint_to_pdf(self, source: Path, target: Path):
        presentation = None
        try:
            presentation = self.powerpoint.Presentations.Open(
                str(source.resolve()),
                WithWindow=False,
            )
            presentation.SaveAs(
                str(target.resolve()),
                32,  # ppSaveAsPDF
            )
        finally:
            if presentation is not None:
                try:
                    presentation.Close()
                except Exception:
                    pass

    def convert(self, source: Path, target: Path):
        suffix = source.suffix.lower()

        if suffix == ".pdf":
            shutil.copy2(source, target)
            return

        kind = SUPPORTED_OFFICE.get(suffix)
        if kind == "word":
            self.word_to_pdf(source, target)
        elif kind == "excel":
            self.excel_to_pdf(source, target)
        elif kind == "powerpoint":
            self.powerpoint_to_pdf(source, target)
        else:
            raise RuntimeError(
                f"Bestandstype wordt niet ondersteund: {source.name}\n"
                f"Ondersteund: PDF, Word, Excel en PowerPoint."
            )


# ---------------------------------------------------------------------------
# Dynamic cover page
#
# The supplied Voorbladen.docx contains the visual layout used by the user.
# We take the first page as the master page, preserve its header/footer,
# borders, fonts and styles, then replace the title and bullet list.
# ---------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W_NS, "r": R_NS}

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def set_paragraph_text_preserve_format(paragraph, text: str):
    """
    Replace the visible text while retaining the paragraph/run formatting.
    The first text run is retained; additional text runs are removed.
    """
    runs = paragraph.xpath("./w:r", namespaces=NS)

    if not runs:
        r = ET.Element(f"{{{W_NS}}}r")
        t = ET.SubElement(r, f"{{{W_NS}}}t")
        t.text = text
        paragraph.append(r)
        return

    first_run = runs[0]

    text_nodes = first_run.xpath(".//w:t", namespaces=NS)
    if text_nodes:
        text_nodes[0].text = text
        if text.startswith(" ") or text.endswith(" "):
            text_nodes[0].set(XML_SPACE, "preserve")
        for node in text_nodes[1:]:
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)
    else:
        t = ET.SubElement(first_run, f"{{{W_NS}}}t")
        t.text = text

    # Remove all later runs to prevent old text remaining.
    for run in runs[1:]:
        paragraph.remove(run)


def make_dynamic_cover(
    template: Path,
    output_docx: Path,
    attachment_number: str,
    attachment_title: str,
    item_titles: list[str],
):
    """
    Create a one-page cover from the first page of Voorbladen.docx.

    Important: the supplied template uses a page-break paragraph after the
    first cover page. We locate that break directly with an XPath query
    instead of assuming it is a direct child in a particular XML structure.
    This is more robust against Word changing the XML structure.
    """
    with zipfile.ZipFile(template, "r") as zin:
        document_xml = zin.read("word/document.xml")
        root = ET.fromstring(document_xml)

        body = root.find(f"{{{W_NS}}}body")
        if body is None:
            raise RuntimeError("Ongeldig Word-template: <w:body> ontbreekt.")

        # Find the first paragraph that contains an explicit page break.
        # The supplied Voorbladen.docx has this after the first cover page.
        page_break_paragraphs = body.xpath(
            "./w:p[.//w:br[@w:type='page']]",
            namespaces=NS,
        )

        if not page_break_paragraphs:
            # Some Word versions can serialize the break differently.
            # Fall back to locating the first w:br with type=page and then
            # walking up to its paragraph.
            breaks = body.xpath(
                ".//w:br[@w:type='page']",
                namespaces=NS,
            )
            if breaks:
                node = breaks[0]
                while node is not None and node.tag != f"{{{W_NS}}}p":
                    node = node.getparent()
                if node is not None:
                    page_break_paragraphs = [node]

        if not page_break_paragraphs:
            raise RuntimeError(
                "Het template bevat geen herkenbaar pagina-einde. "
                "Controleer Voorbladen.docx."
            )

        break_paragraph = page_break_paragraphs[0]
        children = list(body)

        try:
            break_index = children.index(break_paragraph)
        except ValueError:
            raise RuntimeError(
                "Het pagina-einde van het template kon niet worden gelokaliseerd."
            )

        # Everything before the page-break paragraph belongs to page 1.
        first_page_children = children[:break_index]
        paragraphs = [c for c in first_page_children if c.tag == f"{{{W_NS}}}p"]

        if not paragraphs:
            raise RuntimeError("Geen inhoud op de eerste templatepagina gevonden.")

        # In the supplied template:
        #   p[0] = title
        #   p[1..] = bullet examples
        title_paragraph = paragraphs[0]

        # Find a paragraph carrying Word's numbering/bullet definition.
        bullet_master = None
        for p in paragraphs[1:]:
            if p.xpath("./w:pPr/w:numPr", namespaces=NS):
                bullet_master = p
                break

        if bullet_master is None:
            # Fallback: first paragraph after the title.
            if len(paragraphs) >= 2:
                bullet_master = paragraphs[1]
            else:
                raise RuntimeError(
                    "Het template bevat geen voorbeeld voor de bijlagenlijst."
                )

        # Preserve section properties. These control page size, margins,
        # header/footer and other page-level formatting.
        sectPr = body.find(f"{{{W_NS}}}sectPr")

        # Clear body content.
        for child in list(body):
            body.remove(child)

        # 1. Title
        title = deepcopy(title_paragraph)
        set_paragraph_text_preserve_format(
            title,
            f"Bijlage {attachment_number}: {attachment_title}",
        )
        body.append(title)

        # 2. Dynamic bullet list
        for item in item_titles:
            p = deepcopy(bullet_master)
            set_paragraph_text_preserve_format(p, item)
            body.append(p)

        # # 3. Add a blank paragraph using the formatting of the last
        # #    non-title paragraph. This preserves the visual breathing room.
        # blank_source = paragraphs[-1]
        # if blank_source is not title_paragraph:
        #     blank = deepcopy(blank_source)
        #     for node in blank.xpath(".//w:t", namespaces=NS):
        #         node.text = ""
        #     body.append(blank)

        if sectPr is not None:
            body.append(sectPr)

        new_document_xml = ET.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )

        output_docx.parent.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = (
                    new_document_xml
                    if item.filename == "word/document.xml"
                    else zin.read(item.filename)
                )
                zout.writestr(item, data)


def format_eta(seconds):
    """Format seconds as HH:MM:SS or MM:SS."""
    if seconds is None:
        return "--:--"
    try:
        seconds = max(0, int(round(float(seconds))))
    except (TypeError, ValueError):
        return "--:--"

    minutes, secs = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)

    if hours:
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    return f"{minutes:02d}:{secs:02d}"


class PipelineProgress:
    """Track the complete generation pipeline and calculate a global ETA."""

    def __init__(self, callback, total_units):
        self.callback = callback
        self.total_units = max(1, int(total_units))
        self.completed_units = 0
        self.start_time = time.monotonic()
        self.last_gui_update = 0.0
        self.update_interval = 0.20
        self.last_message = ""

    def update(self, message, completed=None, force=False):
        if completed is not None:
            self.completed_units = max(0, min(self.total_units, completed))

        now = time.monotonic()
        if not force and now - self.last_gui_update < self.update_interval:
            return

        elapsed = now - self.start_time
        fraction = self.completed_units / self.total_units
        rate = (
            self.completed_units / elapsed
            if elapsed > 0 and self.completed_units
            else 0
        )
        remaining = (
            ((self.total_units - self.completed_units) / rate) if rate > 0 else None
        )
        percent = fraction * 100

        self.last_gui_update = now
        self.last_message = str(message)

        if self.callback:
            self.callback(
                str(message),
                percent,
                remaining,
                elapsed,
            )

    def force_update(self, message, completed=None):
        self.update(message, completed=completed, force=True)


# ---------------------------------------------------------------------------
# PDF merge
# ---------------------------------------------------------------------------


def merge_pdfs(
    pdf_files: list[Path],
    output_pdf: Path,
    pipeline=None,
    merge_start_units=0,
    merge_units=1,
):
    if PdfReader is None or PdfWriter is None:
        raise RuntimeError(
            "pypdf ontbreekt. Installeer de dependencies met:\n"
            "pip install -r requirements.txt"
        )

    writer = PdfWriter()

    # Count pages first so the merge stage can contribute meaningful progress.
    total_pages = 0
    page_counts = []
    for pdf in pdf_files:
        reader = PdfReader(str(pdf))
        count = len(reader.pages)
        page_counts.append((pdf, count))
        total_pages += count

    processed_pages = 0

    for pdf, count in page_counts:
        reader = PdfReader(str(pdf))
        for page in reader.pages:
            writer.add_page(page)
            processed_pages += 1

            if pipeline and total_pages:
                # Merge gets its own share of the global pipeline.
                fraction = processed_pages / total_pages
                completed = merge_start_units + fraction * merge_units
                pipeline.update(
                    f"PDF's samenvoegen: pagina {processed_pages} van {total_pages}",
                    completed=completed,
                )

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with open(output_pdf, "wb") as f:
        writer.write(f)

    if pipeline:
        pipeline.force_update(
            f"PDF's samenvoegen: pagina {total_pages} van {total_pages}",
            completed=merge_start_units + merge_units,
        )


# ---------------------------------------------------------------------------
# Main generation pipeline
# ---------------------------------------------------------------------------


def generate(project_dir: Path, main_docx: Path, progress_callback=None) -> Path:
    project_dir = project_dir.resolve()
    main_docx = validate_main_docx(main_docx)
    attachment_folders = find_attachment_folders(project_dir)

    template = get_template_path()
    if not template.exists():
        raise FileNotFoundError(f"Template ontbreekt: {template}")

    output_dir = project_dir / "PDF rapportage"
    output_dir.mkdir(exist_ok=True)

    # Work units: main report + one cover per attachment + each attachment file.
    conversion_units = 1 + sum(1 + len(files) for _, _, _, files in attachment_folders)
    # Give merge a separate unit. The page-level progress inside that unit is
    # based on actual pages.
    total_units = conversion_units + 1
    pipeline = PipelineProgress(progress_callback, total_units)

    temp_dir = Path(tempfile.mkdtemp(prefix="rapportage_generator_"))
    converter = OfficeConverter()
    converter.start()

    try:
        completed = 0

        # 1. Main report.
        report_pdf = temp_dir / "00_Rapportage.pdf"
        pipeline.force_update(
            f"Rapportage converteren: {main_docx.name}",
            completed=completed,
        )
        converter.word_to_pdf(main_docx, report_pdf)
        completed += 1
        pipeline.force_update(
            f"Rapportage gereed: {main_docx.name}",
            completed=completed,
        )

        final_parts = [report_pdf]

        # 2. Attachments.
        for number, title, folder, files in attachment_folders:
            pipeline.force_update(
                f"Bijlage {number}: {title} ({len(files)} bestand(en))",
                completed=completed,
            )

            if not files:
                raise RuntimeError(f"Bijlage {number} - {title} bevat geen bestanden.")

            item_titles = [p.stem for p in files]

            cover_docx = temp_dir / f"{str(number):0>3}_voorblad.docx"
            cover_pdf = temp_dir / f"{str(number):0>3}_voorblad.pdf"

            make_dynamic_cover(
                template=template,
                output_docx=cover_docx,
                attachment_number=number,
                attachment_title=title,
                item_titles=item_titles,
            )

            converter.word_to_pdf(cover_docx, cover_pdf)
            final_parts.append(cover_pdf)
            completed += 1
            pipeline.force_update(
                f"Bijlage {number}: voorblad gereed",
                completed=completed,
            )

            for item_index, source in enumerate(files, start=1):
                target_pdf = temp_dir / (
                    f"{str(number):0>3}_{item_index:03d}_{source.stem}.pdf"
                )

                # Keep the log compact: no per-file log message. The live
                # status still tells the user what is being processed.
                pipeline.update(
                    f"Bijlage {number}: {source.name}",
                    completed=completed,
                    force=True,
                )

                converter.convert(source, target_pdf)
                final_parts.append(target_pdf)
                completed += 1
                pipeline.force_update(
                    f"Bijlage {number}: gereed",
                    completed=completed,
                )

        # 3. Merge. The ETA continues from the original pipeline start time.
        output_name = main_docx.stem + "_compleet.pdf"
        output_pdf = output_dir / output_name

        pipeline.force_update("PDF's samenvoegen...", completed=completed)

        merge_pdfs(
            final_parts,
            output_pdf,
            pipeline=pipeline,
            merge_start_units=completed,
            merge_units=1,
        )

        pipeline.force_update(
            f"Klaar: {output_pdf}",
            completed=total_units,
        )

        return output_pdf

    finally:
        converter.stop()
        shutil.rmtree(temp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Simple Windows GUI
# ---------------------------------------------------------------------------


def build_preview_items(project_dir: Path, main_docx: Path):
    """Build a simple tree representation of the PDF that will be generated."""
    items = [("report", main_docx.name)]
    try:
        attachments = find_attachment_folders(project_dir)
    except Exception:
        return items
    for number, title, folder, files in attachments:
        items.append(("cover", f"Voorblad — Bijlage {number}: {title}"))
        for index, source in enumerate(files, 1):
            items.append(("file", f"{index:02d} — {source.name}"))
    return items


class App:
    def __init__(self, root):
        self.root = root
        # Title
        self.root.title(f"{COMPANY_NAME} {APP_NAME} v{APP_VERSION}")
        # Startup size
        self.root.geometry("1000x1000")
        self.root.minsize(620, 360)

        self.project_var = tk.StringVar()
        self.docx_var = tk.StringVar()
        self.status_var = tk.StringVar(
            value="Selecteer eerst het gewenste Word-bestand (.docx)."
        )

        self.progress_var = tk.DoubleVar(value=0.0)
        self._last_progress_gui_update = 0.0
        self._progress_update_interval = 0.2
        self._logged_progress_messages = set()

        frame = tk.Frame(root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        # Header: logo links, applicatienaam rechts daarvan
        header = tk.Frame(frame)
        header.pack(fill="x", pady=(0, 15))

        if LOGO_PATH.exists():
            self.logo_image = Image.open(LOGO_PATH)

            # Alleen de breedte bepalen; hoogte blijft evenredig.
            logo_width = 100
            ratio = logo_width / self.logo_image.width
            logo_height = int(self.logo_image.height * ratio)
            self.logo_image = self.logo_image.resize((logo_width, logo_height))

            self.logo = ImageTk.PhotoImage(self.logo_image)

            tk.Label(
                header,
                image=self.logo,
            ).pack(side="left", padx=(0, 15))

        tk.Label(
            header,
            text=f"{APP_NAME} v{APP_VERSION}",
            font=("Segoe UI", 18, "bold"),
        ).pack(side="left", anchor="center")

        tk.Label(
            frame,
            text=(
                "Microsoft Word/Office wordt gebruikt voor de conversie. "
                "De mapstructuur bepaalt automatisch de volgorde."
            ),
            font=("Segoe UI", 10),
            wraplength=650,
            justify="left",
        ).pack(anchor="w", pady=(5, 20))

        tk.Label(
            frame,
            text="1. Hoofdrapport (.docx)",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

        docx_row = tk.Frame(frame)
        docx_row.pack(fill="x", pady=(4, 10))

        tk.Entry(
            docx_row,
            textvariable=self.docx_var,
            font=("Segoe UI", 10),
        ).pack(side="left", fill="x", expand=True)

        tk.Button(
            docx_row,
            text="Kies DOCX...",
            command=self.choose_docx,
            width=14,
        ).pack(side="left", padx=(8, 0))

        tk.Label(
            frame,
            text="2. Projectmap (automatisch bepaald)",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

        project_row = tk.Frame(frame)
        project_row.pack(fill="x", pady=(4, 10))

        tk.Entry(
            project_row,
            textvariable=self.project_var,
            font=("Segoe UI", 10),
            state="readonly",
        ).pack(side="left", fill="x", expand=True)

        # Preview van de uiteindelijke PDF-opbouw
        tk.Label(
            frame,
            text="3. Preview PDF-opbouw",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(6, 4))

        preview_frame = tk.Frame(frame)
        preview_frame.pack(fill="both", expand=True, pady=(0, 10))

        self.preview_tree = ttk.Treeview(
            preview_frame,
            show="tree",
            height=12,
        )
        self.preview_tree.pack(side="left", fill="both", expand=True)

        preview_scroll = ttk.Scrollbar(
            preview_frame,
            orient="vertical",
            command=self.preview_tree.yview,
        )
        preview_scroll.pack(side="right", fill="y")
        self.preview_tree.configure(yscrollcommand=preview_scroll.set)

        self.preview_tree.insert(
            "",
            "end",
            text="Selecteer een DOCX om de PDF-opbouw te bekijken.",
        )

        self.generate_button = tk.Button(
            frame,
            text="GENEREER COMPLETE PDF",
            command=self.run,
            font=("Segoe UI", 11, "bold"),
            height=2,
        )
        self.generate_button.pack(fill="x", pady=(20, 12))

        self.status = tk.Label(
            frame,
            textvariable=self.status_var,
            anchor="w",
            justify="left",
            wraplength=650,
        )
        self.status.pack(fill="x")

        progress_frame = tk.Frame(frame)
        progress_frame.pack(fill="x", pady=(8, 0))
        self.progress = ttk.Progressbar(
            progress_frame, variable=self.progress_var, maximum=100, mode="determinate"
        )
        self.progress.pack(side="left", fill="x", expand=True)
        self.progress_percent_var = tk.StringVar(value="0%")
        tk.Label(
            progress_frame,
            textvariable=self.progress_percent_var,
            width=5,
            anchor="e",
            font=("Segoe UI", 9),
        ).pack(side="left", padx=(8, 0))
        self.eta_var = tk.StringVar(value="Geschatte resterende tijd: --")
        tk.Label(
            frame, textvariable=self.eta_var, anchor="w", font=("Segoe UI", 9)
        ).pack(fill="x", pady=(3, 0))

        self.log = tk.Text(
            frame,
            height=12,
            state="disabled",
            font=("Consolas", 9),
        )
        self.log.pack(fill="both", expand=True, pady=(12, 0))

    def choose_docx(self):
        filename = filedialog.askopenfilename(
            title="Kies het hoofdrapport",
            filetypes=[
                ("Word-documenten", "*.docx"),
                ("Alle bestanden", "*.*"),
            ],
        )
        if not filename:
            return

        # Log leegmaken bij het selecteren van een nieuw document
        self.clear_log()

        docx = Path(filename).resolve()
        self.docx_var.set(str(docx))

        project_dir = docx.parent
        candidate = project_dir

        while True:
            try:
                has_bijlage = any(
                    p.is_dir() and p.name.lower() == ATTACHMENT_DIR_NAME.lower()
                    for p in candidate.iterdir()
                )
            except OSError:
                has_bijlage = False

            if has_bijlage:
                project_dir = candidate
                break

            if candidate == candidate.parent:
                break
            candidate = candidate.parent

        self.project_var.set(str(project_dir))
        self.update_preview()
        self.status_var.set(
            "DOCX geselecteerd. Controleer de preview en klik daarna op "
            "'GENEREER COMPLETE PDF'."
        )

    def update_preview(self):
        """Refresh the tree showing the expected PDF sequence."""
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)

        docx_text = self.docx_var.get().strip()
        folder = self.project_var.get().strip()

        if not docx_text:
            self.preview_tree.insert(
                "",
                "end",
                text="Selecteer een DOCX om de PDF-opbouw te bekijken.",
            )
            return

        main_docx = Path(docx_text)
        project_dir = Path(folder) if folder else main_docx.parent

        if not main_docx.exists():
            self.preview_tree.insert(
                "",
                "end",
                text="Het geselecteerde DOCX-bestand bestaat niet.",
            )
            return

        try:
            attachments = find_attachment_folders(project_dir)
        except Exception as exc:
            self.preview_tree.insert(
                "",
                "end",
                text=f"Bijlagen niet gevonden: {exc}",
            )
            return

        root = self.preview_tree.insert(
            "",
            "end",
            text=f"Hoofdrapport — {main_docx.name}",
            open=True,
        )

        for number, title, folder, files in attachments:
            cover = self.preview_tree.insert(
                root,
                "end",
                text=f"Voorblad — Bijlage {number}: {title}",
                open=True,
            )

            for index, source in enumerate(files, 1):
                self.preview_tree.insert(
                    cover,
                    "end",
                    text=source.name,
                )

        if not attachments:
            self.preview_tree.insert(root, "end", text="Geen bijlagen gevonden.")

    # Clear log
    def clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.configure(state="disabled")

    def update_progress(self, message, percent, eta=None, elapsed=None):
        """Update live progress and write only compact, high-level log entries."""
        try:
            pct = max(0.0, min(100.0, float(percent)))
        except (TypeError, ValueError):
            pct = 0.0

        self.progress_var.set(pct)
        self.progress_percent_var.set(f"{pct:.0f}%")

        if pct >= 100.0:
            eta_text = "Gereed"
        elif eta is None:
            eta_text = "Geschatte resterende tijd: wordt berekend..."
        else:
            eta_text = f"Geschatte resterende tijd: ± {format_eta(eta)}"

        if elapsed is not None:
            self.eta_var.set(f"Verstreken tijd: {format_eta(elapsed)}  |  {eta_text}")
        else:
            self.eta_var.set(eta_text)

        self.status_var.set(str(message))

        # Compacte log: alleen hoofdrapport/bijlagen/fasen, nooit elk bronbestand.
        msg = str(message)
        log_message = None

        if msg.startswith("Rapportage gereed:"):
            log_message = "Hoofdrapport gereed."
        elif msg.startswith("Bijlage ") and ": " in msg:
            prefix, detail = msg.split(": ", 1)

            # "Bijlage A: Titel (3 bestand(en))" -> één regel.
            if detail.endswith("bestand(en))") or "bestand(en)" in detail:
                log_message = msg

            # Gereedmeldingen van individuele bijlagen niet in de log zetten.
            # "Bijlage A: voorblad gereed" en "Bijlage A: gereed" blijven
            # uitsluitend zichtbaar via de live status.
            # "Bijlage A: bestandsnaam.pdf" -> niet loggen.
        elif msg.startswith("PDF's samenvoegen"):
            log_message = "PDF's samenvoegen..."
        elif msg.startswith("Klaar:"):
            log_message = msg

        if log_message and log_message not in self._logged_progress_messages:
            self._logged_progress_messages.add(log_message)
            self.log.configure(state="normal")
            self.log.insert("end", log_message + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")

        self.root.update_idletasks()

    def write_log(self, text, percent=None, eta=None, elapsed=None):
        """Schrijf normale logregels; live voortgang loopt via update_progress()."""
        message = str(text)

        # Een eventuele progress-call wordt centraal afgehandeld.
        if percent is not None:
            self.update_progress(message, percent, eta, elapsed)
            return

        # Compacte log: geen afzonderlijke bronbestanden.
        is_file_detail = (
            message.startswith("  Verwerken:")
            or message.startswith("Verwerken:")
            or message.startswith("Gereed:")
            or message.startswith("Voorblad Bijlage ")
        )

        if not is_file_detail:
            self.log.configure(state="normal")
            self.log.insert("end", message + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")

        self.root.update_idletasks()

    def run(self):
        docx_text = self.docx_var.get().strip()
        folder = self.project_var.get().strip()

        if not docx_text:
            messagebox.showwarning(
                APP_NAME, "Kies eerst het gewenste hoofdrapport (.docx)."
            )
            return

        main_docx = Path(docx_text).resolve()
        project_dir = Path(folder).resolve() if folder else main_docx.parent

        if not main_docx.exists():
            messagebox.showerror(
                APP_NAME, "Het geselecteerde DOCX-bestand bestaat niet meer."
            )
            return

        if not project_dir.exists():
            messagebox.showerror(APP_NAME, "De projectmap bestaat niet.")
            return

        self.generate_button.config(state="disabled")

        # Start elke generatie met een schone log.
        self.clear_log()
        self._logged_progress_messages.clear()

        self.progress_var.set(0.0)
        self._last_progress_gui_update = 0.0
        self.progress_var.set(0)
        self.progress_percent_var.set("0%")
        self.eta_var.set("Geschatte resterende tijd: wordt berekend...")
        try:
            self.write_log(f"Hoofdrapport: {main_docx.name}")
            self.write_log(f"Projectmap: {project_dir}")
            self.write_log("Start...")
            generate(project_dir, main_docx, self.update_progress)
        except Exception as exc:
            self.write_log(f"FOUT: {exc}")
            messagebox.showerror(APP_NAME, f"Genereren is mislukt:\n\n{exc}")
        finally:
            self.generate_button.config(state="normal")


def main():
    if os.name != "nt":
        print("Deze versie is bedoeld voor Windows met Microsoft Office.")
        sys.exit(1)

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
