"""
Word to PDF Converter  –  GUI File Picker + Page Selection + Preview + Email
=============================================================================
1. Opens a startup dialog explaining the tool.
2. Asks whether to delete the sent email from the Sent folder.
3. Opens a file-picker window so you can select one or more .docx files.
4. For each file, shows a PAGE SELECTION screen inside the same window:
   - Enter specific pages / ranges  (e.g. "1,3,5-8,10")
   - Click "Preview" to see a thumbnail of each selected page
   - Click "Confirm & Convert" to proceed
5. Converts only the selected pages to PDF via MS Word COM (updates TOC + fields).
6. Emails each PDF to you as an attachment via SMTP (Gmail / Outlook / any).

REQUIREMENTS
------------
    pip install python-docx pypdf comtypes pillow pdf2image

    Microsoft Word must be installed (used via COM for field/TOC refresh).
    Poppler must be installed and on PATH for pdf2image to render previews:
      - Windows: https://github.com/oschwartz10612/poppler-windows/releases
        Extract and add the bin/ folder to your system PATH.

EMAIL SETUP  (edit the CONFIG block below)
------------------------------------------
Gmail:
  • SMTP_HOST = "smtp.gmail.com",  SMTP_PORT = 465
  • Use an App Password (not your regular Gmail password):
    Google Account -> Security -> 2-Step Verification -> App passwords
"""

import os
import sys
import logging
import platform
import smtplib
import tempfile
import tkinter as tk
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import imaplib
from datetime import date
import uuid

# ─────────────────────────── EMAIL CONFIG ──────────────────────────────────
SMTP_HOST     = "smtp.gmail.com"
SMTP_PORT     = 465
SMTP_USER     = " "
SMTP_PASSWORD = " "
EMAIL_TO      = " "
EMAIL_FROM    = SMTP_USER
# ────────────────────────────────────────────────────────────────────────────

OUTPUT_DIR = r"C:\Users\30316376\Downloads"
LOG_FILE   = "./converter.log"

# ────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        *(
            [logging.FileHandler(LOG_FILE, encoding="utf-8")]
            if LOG_FILE else []
        ),
    ],
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════
#  UTILITIES
# ══════════════════════════════════════════════════════════════════

def parse_page_input(raw: str, total_pages: int) -> list[int]:
    """
    Parse a page-range string like "1,3,5-8,10" into a sorted list of
    0-based page indices.  Raises ValueError on bad input.
    """
    pages = set()
    for part in raw.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            a, b = int(a), int(b)
            if a < 1 or b > total_pages or a > b:
                raise ValueError(f"Range {a}-{b} is out of bounds (1–{total_pages})")
            pages.update(range(a - 1, b))          # convert to 0-based
        else:
            n = int(part)
            if n < 1 or n > total_pages:
                raise ValueError(f"Page {n} is out of bounds (1–{total_pages})")
            pages.add(n - 1)
    return sorted(pages)


def get_total_pages_via_word(docx_path: Path) -> int:
    """Return the total page count using Word COM (most accurate)."""
    try:
        import comtypes.client
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        try:
            count = doc.ComputeStatistics(2)   # wdStatisticPages = 2
        finally:
            doc.Close(SaveChanges=False)
            word.Quit()
        return count
    except Exception:
        return 0


def render_page_preview(full_pdf_path: Path, page_index: int,
                         thumb_width: int = 380) -> "ImageTk.PhotoImage | None":
    """
    Render a single page of a PDF to a Tkinter-compatible PhotoImage.
    Uses pdf2image (poppler).  Returns None on failure.
    """
    try:
        from pdf2image import convert_from_path
        from PIL import Image, ImageTk

        images = convert_from_path(
            str(full_pdf_path),
            first_page=page_index + 1,
            last_page=page_index + 1,
            dpi=96,
        )
        if not images:
            return None

        img = images[0]
        ratio = thumb_width / img.width
        new_h = int(img.height * ratio)
        img = img.resize((thumb_width, new_h), Image.LANCZOS)
        return ImageTk.PhotoImage(img)

    except ImportError:
        return None
    except Exception as e:
        log.warning(f"Preview render failed: {e}")
        return None


# ══════════════════════════════════════════════════════════════════
#  CONVERSION HELPERS
# ══════════════════════════════════════════════════════════════════

def _convert_with_msword(docx_path: Path, output_dir: Path,
                          page_indices: list[int] | None = None) -> Path:
    """
    Convert docx -> PDF via Word COM.
    If page_indices is given, extract only those pages from the full PDF.
    """
    try:
        import comtypes.client
    except ImportError:
        raise ImportError("comtypes not found.\nRun:  pip install comtypes")

    abs_docx    = str(docx_path.resolve())
    full_pdf    = str((output_dir / (docx_path.stem + "_FULL_TMP.pdf")).resolve())
    final_pdf   = str((output_dir / (docx_path.stem + ".pdf")).resolve())

    log.info("  Opening Word via COM ...")
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False

    try:
        doc = word.Documents.Open(abs_docx, ReadOnly=False)

        log.info("  Updating all fields and TOC ...")
        doc.Fields.Update()
        for toc in doc.TablesOfContents:
            toc.Update()
        for tof in doc.TablesOfFigures:
            tof.Update()

        log.info("  Exporting full PDF ...")
        wdExportFormatPDF = 17
        doc.ExportAsFixedFormat(
            OutputFileName=full_pdf,
            ExportFormat=wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=1,
            DocStructureTags=True,
        )
        log.info("  Word COM export complete.")
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()

    full_pdf_path = Path(full_pdf)

    if page_indices is not None:
        log.info(f"  Extracting pages: {[p+1 for p in page_indices]}")
        from pypdf import PdfReader, PdfWriter
        reader = PdfReader(str(full_pdf_path))
        writer = PdfWriter()
        for idx in page_indices:
            if 0 <= idx < len(reader.pages):
                writer.add_page(reader.pages[idx])
        if reader.metadata:
            writer.add_metadata(reader.metadata)
        with open(final_pdf, "wb") as fh:
            writer.write(fh)
        full_pdf_path.unlink(missing_ok=True)
    else:
        full_pdf_path.rename(final_pdf)

    return Path(final_pdf)


def _convert_with_libreoffice(docx_path: Path, output_dir: Path,
                                page_indices: list[int] | None = None) -> Path:
    import subprocess
    cmd = [
        "libreoffice", "--headless",
        "--infilter=writer_pdf_Export",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(docx_path),
    ]
    log.info("  Running LibreOffice conversion ...")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice error:\n{result.stderr}")

    full_pdf = output_dir / (docx_path.stem + ".pdf")

    if page_indices is not None:
        from pypdf import PdfReader, PdfWriter
        tmp = output_dir / (docx_path.stem + "_FULL_TMP.pdf")
        full_pdf.rename(tmp)
        reader = PdfReader(str(tmp))
        writer = PdfWriter()
        for idx in page_indices:
            if 0 <= idx < len(reader.pages):
                writer.add_page(reader.pages[idx])
        with open(full_pdf, "wb") as fh:
            writer.write(fh)
        tmp.unlink(missing_ok=True)

    return full_pdf


def convert_docx_to_pdf(docx_path: Path,
                         page_indices: list[int] | None = None) -> Path:
    out_dir = Path(OUTPUT_DIR) if OUTPUT_DIR else docx_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    if platform.system() == "Windows":
        try:
            return _convert_with_msword(docx_path, out_dir, page_indices)
        except Exception as exc:
            log.warning(f"  MS Word COM failed ({exc}); trying LibreOffice ...")

    return _convert_with_libreoffice(docx_path, out_dir, page_indices)


def add_bookmarks_to_pdf(pdf_path: Path, docx_path: Path) -> Path:
    from docx import Document
    from pypdf import PdfReader, PdfWriter

    log.info("  Injecting bookmarks ...")
    doc    = Document(str(docx_path))
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)
    if reader.metadata:
        writer.add_metadata(reader.metadata)

    headings = [
        {"text": p.text.strip(), "level": int(p.style.name[-1])}
        for p in doc.paragraphs
        if p.style.name in ("Heading 1", "Heading 2") and p.text.strip()
    ]

    current_h1 = None
    added = 0
    for h in headings:
        needle = h["text"].lower()
        page_idx = next(
            (i for i, pg in enumerate(reader.pages)
             if needle in (pg.extract_text() or "").lower()),
            None,
        )
        if page_idx is None:
            continue
        if h["level"] == 1:
            current_h1 = writer.add_outline_item(h["text"], page_idx)
            added += 1
        else:
            writer.add_outline_item(h["text"], page_idx, parent=current_h1)
            added += 1

    with open(pdf_path, "wb") as fh:
        writer.write(fh)

    log.info(f"  Bookmarks written: {added}")
    return pdf_path


# ══════════════════════════════════════════════════════════════════
#  EMAIL
# ══════════════════════════════════════════════════════════════════

def send_pdf_by_email(pdf_path: Path, original_docx_name: str,
                       delete_sent: bool = False) -> None:
    log.info(f"  Sending email to {EMAIL_TO} ...")

    msg = EmailMessage()
    msg["Subject"] = f"Converted PDF: {pdf_path.name}"
    msg["From"]    = EMAIL_FROM
    msg["To"]      = EMAIL_TO
    unique_id = f"<{uuid.uuid4()}@word-to-pdf-converter>"
    msg["Message-ID"] = unique_id
    # Then pass it to the delete function:
    if delete_sent:
        _delete_from_sent(message_id=unique_id)
    msg.set_content(
        f"Hi,\n\n"
        f"Please find attached the converted PDF for:\n"
        f"  {original_docx_name}\n\n"
        f"Converted on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')}.\n\n"
        f"This email was sent automatically by the Word-to-PDF Converter."
    )

    with open(pdf_path, "rb") as fh:
        msg.add_attachment(
            fh.read(),
            maintype="application",
            subtype="pdf",
            filename=pdf_path.name,
        )

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as smtp:
        smtp.login(SMTP_USER, SMTP_PASSWORD)
        smtp.send_message(msg)
    log.info("  Email sent successfully.")

    if delete_sent:
        _delete_from_sent(message_id=unique_id)


def _delete_from_sent(message_id: str) -> None:
    SENT_FOLDER_CANDIDATES = [
        '"[Gmail]/Sent Mail"', "Sent", "Sent Items", "INBOX.Sent",
    ]
    IMAP_HOST = "imap.gmail.com"

    log.info("  Connecting to IMAP to delete sent email ...")
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST)
        mail.login(SMTP_USER, SMTP_PASSWORD)

        selected = False
        for folder in SENT_FOLDER_CANDIDATES:
            result, _ = mail.select(folder)
            if result == "OK":
                log.info(f"  Opened sent folder: {folder}")
                selected = True
                break

        if not selected:
            log.warning("  Could not find Sent folder — skipping delete.")
            mail.logout()
            return

        # Search by unique Message-ID instead of subject or date
        result, data = mail.search(None, f'HEADER Message-ID "{message_id}"')
        if result != "OK" or not data[0]:
            log.warning("  Sent message not found — skipping delete.")
            mail.logout()
            return

        for num in data[0].split():
            mail.store(num, "+FLAGS", "\\Deleted")
        mail.expunge()
        log.info("  Sent email permanently deleted from Sent folder.")
        mail.logout()

    except Exception as exc:
        log.warning(f"  Could not delete from Sent folder: {exc}")


# ══════════════════════════════════════════════════════════════════
#  MAIN APPLICATION  (multi-page Tkinter wizard)
# ══════════════════════════════════════════════════════════════════

class ConverterApp(tk.Tk):
    """
    A single Tk window that acts as a wizard:

      Page 0  –  Welcome / options (delete sent? all pages or specific?)
      Page 1  –  File picker confirmation (shown briefly, then auto-advances)
      Page 2  –  Per-file page-selection + preview
      Page 3  –  Conversion progress
    """

    WIN_W, WIN_H = 560, 620
    PREVIEW_W    = 400   # thumbnail width in pixels

    def __init__(self):
        super().__init__()
        self.title("Word to PDF Converter")
        self.resizable(False, False)
        self.geometry(f"{self.WIN_W}x{self.WIN_H}")
        self._centre_window()
        self.attributes("-topmost", True)

        # ── State ──
        self.delete_sent   : bool          = False
        self.select_pages  : bool          = False
        self.docx_files    : list[Path]    = []
        self.file_pages    : dict          = {}   # path -> list[int] (0-based) or None
        self._current_file_idx : int       = 0
        self._tmp_full_pdfs: list[Path]    = []   # temp PDFs used for preview
        self._preview_images              = []    # keep refs alive (GC protection)

        # ── Container for pages ──
        self.container = tk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self._show_welcome()

    # ─────────────────────────────────────────────────────────────
    #  Helpers
    # ─────────────────────────────────────────────────────────────

    def _centre_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  // 2) - (self.WIN_W // 2)
        y = (self.winfo_screenheight() // 2) - (self.WIN_H // 2)
        self.geometry(f"{self.WIN_W}x{self.WIN_H}+{x}+{y}")

    def _clear(self):
        for w in self.container.winfo_children():
            w.destroy()

    def _header(self, text: str, step: str = ""):
        frm = tk.Frame(self.container, bg="#1565C0")
        frm.pack(fill="x")
        tk.Label(
            frm, text=text,
            font=("Segoe UI", 13, "bold"),
            bg="#1565C0", fg="white",
            pady=14, padx=20, anchor="w",
        ).pack(side="left", fill="x", expand=True)
        if step:
            tk.Label(
                frm, text=step,
                font=("Segoe UI", 9),
                bg="#1565C0", fg="#90CAF9",
                padx=14,
            ).pack(side="right")

    def _btn(self, parent, text, cmd, color="#1565C0", width=22, fg="white"):
        return tk.Button(
            parent, text=text, command=cmd,
            font=("Segoe UI", 9, "bold"), width=width,
            bg=color, fg=fg,
            activebackground=color, activeforeground=fg,
            relief="flat", cursor="hand2", pady=6,
        )

    def _separator(self):
        ttk.Separator(self.container, orient="horizontal").pack(
            fill="x", padx=20, pady=6)

    # ─────────────────────────────────────────────────────────────
    #  PAGE 0  –  Welcome
    # ─────────────────────────────────────────────────────────────

    def _show_welcome(self):
        self._clear()
        self._header("Word to PDF Converter", "Step 1 of 3")

        desc = (
            "This tool will:\n"
            "  1.  Let you select one or more Word (.docx) files.\n"
            "  2.  Update all fields and Table of Contents.\n"
            "  3.  Convert to PDF (all pages or specific pages you choose).\n"
            "  4.  Show you a preview before converting.\n"
            "  5.  Email the PDF to you as an attachment.\n"
        )
        tk.Label(
            self.container, text=desc,
            font=("Segoe UI", 9), justify="left",
            wraplength=500, anchor="w",
            padx=24, pady=12,
        ).pack(fill="x")

        self._separator()

        # ── Page selection preference ──
        tk.Label(
            self.container,
            text="Which pages should be converted?",
            font=("Segoe UI", 9, "bold"), padx=24, anchor="w",
        ).pack(fill="x")

        self._page_mode = tk.StringVar(value="all")
        row = tk.Frame(self.container)
        row.pack(fill="x", padx=36, pady=4)
        tk.Radiobutton(
            row, text="All pages (default)",
            variable=self._page_mode, value="all",
            font=("Segoe UI", 9),
        ).pack(anchor="w")
        tk.Radiobutton(
            row, text="Specific pages / ranges  (e.g. 1, 3, 5-8)",
            variable=self._page_mode, value="specific",
            font=("Segoe UI", 9),
        ).pack(anchor="w")

        self._separator()

        # ── Delete sent ──
        tk.Label(
            self.container,
            text="After sending, delete the email from your Sent folder?",
            font=("Segoe UI", 9, "bold"), padx=24, anchor="w",
        ).pack(fill="x")

        self._delete_var = tk.BooleanVar(value=False)
        row2 = tk.Frame(self.container)
        row2.pack(fill="x", padx=36, pady=4)
        tk.Radiobutton(
            row2, text="No — keep in Sent folder",
            variable=self._delete_var, value=False,
            font=("Segoe UI", 9),
        ).pack(anchor="w")
        tk.Radiobutton(
            row2, text="Yes — delete after sending",
            variable=self._delete_var, value=True,
            font=("Segoe UI", 9),
        ).pack(anchor="w")

        self._separator()

        btn_row = tk.Frame(self.container)
        btn_row.pack(pady=14)
        self._btn(
            btn_row, "Select Files →", self._on_welcome_next,
            color="#1565C0", width=30,
        ).pack()

    def _on_welcome_next(self):
        self.select_pages = (self._page_mode.get() == "specific")
        self.delete_sent  = bool(self._delete_var.get())

        # Open file picker (hides window briefly)
        self.withdraw()
        paths = filedialog.askopenfilenames(
            title="Select Word document(s) to convert",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        self.deiconify()

        if not paths:
            log.info("No files selected.")
            return

        self.docx_files = [Path(p) for p in paths]
        self.file_pages  = {}

        if self.select_pages:
            self._current_file_idx = 0
            self._show_page_selector()
        else:
            for f in self.docx_files:
                self.file_pages[f] = None   # None = all pages
            self._show_progress()

    # ─────────────────────────────────────────────────────────────
    #  PAGE 2  –  Page Selector + Preview
    # ─────────────────────────────────────────────────────────────

    def _show_page_selector(self):
        self._clear()
        idx   = self._current_file_idx
        total = len(self.docx_files)
        f     = self.docx_files[idx]

        self._header(
            f"Select Pages: {f.name}",
            f"File {idx+1} of {total}",
        )

        # ── Page count ──
        self._page_count_var = tk.StringVar(value="Counting pages …")
        tk.Label(
            self.container, textvariable=self._page_count_var,
            font=("Segoe UI", 8, "italic"), fg="#555",
            padx=24, anchor="w",
        ).pack(fill="x", pady=(6, 0))

        # ── Input row ──
        input_frm = tk.Frame(self.container)
        input_frm.pack(fill="x", padx=24, pady=6)

        tk.Label(
            input_frm, text="Pages:",
            font=("Segoe UI", 9, "bold"),
        ).grid(row=0, column=0, sticky="w")

        self._page_input = tk.Entry(input_frm, font=("Segoe UI", 10), width=28)
        self._page_input.grid(row=0, column=1, padx=8, sticky="w")
        self._page_input.insert(0, "e.g. 1, 3, 5-8")
        self._page_input.bind("<FocusIn>", self._clear_placeholder)

        self._btn(
            input_frm, "Preview ▶", self._on_preview,
            color="#00695C", width=12,
        ).grid(row=0, column=2, padx=4)

        self._page_error = tk.StringVar()
        tk.Label(
            self.container, textvariable=self._page_error,
            font=("Segoe UI", 8), fg="#C62828",
            padx=24, anchor="w",
        ).pack(fill="x")

        # ── Preview area (scrollable) ──
        tk.Label(
            self.container,
            text="Page Previews",
            font=("Segoe UI", 9, "bold"),
            padx=24, anchor="w",
        ).pack(fill="x", pady=(4, 0))

        preview_outer = tk.Frame(self.container, relief="sunken", bd=1)
        preview_outer.pack(fill="both", expand=True, padx=24, pady=4)

        self._preview_canvas = tk.Canvas(
            preview_outer, bg="#f0f0f0",
            highlightthickness=0,
        )
        scrollbar = ttk.Scrollbar(
            preview_outer, orient="vertical",
            command=self._preview_canvas.yview,
        )
        self._preview_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self._preview_canvas.pack(side="left", fill="both", expand=True)

        self._preview_inner = tk.Frame(self._preview_canvas, bg="#f0f0f0")
        self._canvas_window = self._preview_canvas.create_window(
            (0, 0), window=self._preview_inner, anchor="nw"
        )
        self._preview_inner.bind(
            "<Configure>",
            lambda e: self._preview_canvas.configure(
                scrollregion=self._preview_canvas.bbox("all")
            )
        )
        self._preview_canvas.bind(
            "<Configure>",
            lambda e: self._preview_canvas.itemconfig(
                self._canvas_window, width=e.width
            )
        )

        # ── Bottom buttons ──
        self._separator()
        btn_row = tk.Frame(self.container)
        btn_row.pack(pady=6)

        if idx > 0:
            self._btn(
                btn_row, "← Back", self._on_selector_back,
                color="#546E7A", width=14,
            ).grid(row=0, column=0, padx=6)

        label = "Next File →" if idx < total - 1 else "Convert & Email ✓"
        self._confirm_btn = self._btn(
            btn_row, label, self._on_selector_confirm,
            color="#1565C0", width=20,
        )
        self._confirm_btn.grid(row=0, column=1, padx=6)

        # Count pages in background so the UI doesn't freeze
        self.after(50, self._count_pages_async)

    def _clear_placeholder(self, event):
        cur = self._page_input.get()
        if "e.g." in cur:
            self._page_input.delete(0, "end")

    def _count_pages_async(self):
        f = self.docx_files[self._current_file_idx]
        try:
            count = get_total_pages_via_word(f)
            if count:
                self._page_count_var.set(f"Document has {count} page(s). Leave blank for all pages.")
                self._total_pages = count
            else:
                self._page_count_var.set("Could not determine page count. Enter page numbers carefully.")
                self._total_pages = 9999
        except Exception:
            self._page_count_var.set("Page count unavailable.")
            self._total_pages = 9999

    def _get_full_tmp_pdf(self, docx_path: Path) -> Path | None:
        """
        Ensure a full temporary PDF exists for preview purposes.
        Returns path to the temp PDF, or None on failure.
        """
        tmp_pdf = Path(tempfile.gettempdir()) / (docx_path.stem + "_PREVIEW_TMP.pdf")
        if tmp_pdf.exists():
            return tmp_pdf

        try:
            import comtypes.client
            abs_docx = str(docx_path.resolve())
            abs_pdf  = str(tmp_pdf)
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(abs_docx, ReadOnly=True)
            try:
                doc.Fields.Update()
                doc.ExportAsFixedFormat(
                    OutputFileName=abs_pdf,
                    ExportFormat=17,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    CreateBookmarks=1,
                    DocStructureTags=True,
                )
            finally:
                doc.Close(SaveChanges=False)
                word.Quit()
            self._tmp_full_pdfs.append(tmp_pdf)
            return tmp_pdf
        except Exception as e:
            log.warning(f"Preview PDF generation failed: {e}")
            return None

    def _on_preview(self):
        self._page_error.set("")
        raw = self._page_input.get().strip()

        # If blank or placeholder, treat as "all pages"
        if not raw or "e.g." in raw:
            self._page_error.set("Enter page numbers first, then click Preview.")
            return

        try:
            indices = parse_page_input(raw, self._total_pages)
        except ValueError as e:
            self._page_error.set(str(e))
            return

        if not indices:
            self._page_error.set("No valid pages found in input.")
            return

        # Clear existing previews
        for w in self._preview_inner.winfo_children():
            w.destroy()
        self._preview_images.clear()

        f = self.docx_files[self._current_file_idx]

        # Show loading label
        loading = tk.Label(
            self._preview_inner,
            text="⏳  Generating previews (this may take a moment) …",
            font=("Segoe UI", 9, "italic"), bg="#f0f0f0", fg="#333",
            pady=20,
        )
        loading.pack()
        self.update()

        tmp_pdf = self._get_full_tmp_pdf(f)
        loading.destroy()

        if tmp_pdf is None:
            tk.Label(
                self._preview_inner,
                text="⚠  Preview unavailable (could not render PDF).\n"
                     "Conversion will still work — click Confirm to proceed.",
                font=("Segoe UI", 9), bg="#f0f0f0", fg="#B71C1C",
                pady=20, wraplength=350,
            ).pack()
            self.update()
            return

        # Render each page
        for i, page_idx in enumerate(indices):
            page_num = page_idx + 1
            frame = tk.Frame(self._preview_inner, bg="#f0f0f0", pady=8)
            frame.pack(fill="x", padx=10)

            tk.Label(
                frame,
                text=f"── Page {page_num} ──",
                font=("Segoe UI", 8, "bold"),
                bg="#f0f0f0", fg="#444",
            ).pack()

            photo = render_page_preview(tmp_pdf, page_idx, self.PREVIEW_W)
            if photo:
                self._preview_images.append(photo)   # prevent GC
                lbl = tk.Label(
                    frame, image=photo, bg="#f0f0f0",
                    relief="solid", bd=1,
                )
                lbl.pack(pady=2)
            else:
                tk.Label(
                    frame,
                    text=f"[Preview unavailable for page {page_num}]",
                    font=("Segoe UI", 8, "italic"),
                    bg="#f0f0f0", fg="#888",
                ).pack()

            if i < len(indices) - 1:
                ttk.Separator(frame, orient="horizontal").pack(
                    fill="x", padx=20, pady=4)

        self._preview_canvas.yview_moveto(0)
        self.update()

    def _on_selector_back(self):
        self._current_file_idx -= 1
        self._show_page_selector()

    def _on_selector_confirm(self):
        self._page_error.set("")
        raw = self._page_input.get().strip()
        f   = self.docx_files[self._current_file_idx]

        if not raw or "e.g." in raw:
            # Treat blank as "all pages"
            self.file_pages[f] = None
        else:
            try:
                indices = parse_page_input(raw, self._total_pages)
                self.file_pages[f] = indices if indices else None
            except ValueError as e:
                self._page_error.set(str(e))
                return

        self._current_file_idx += 1
        if self._current_file_idx < len(self.docx_files):
            self._show_page_selector()
        else:
            self._show_progress()

    # ─────────────────────────────────────────────────────────────
    #  PAGE 3  –  Progress
    # ─────────────────────────────────────────────────────────────

    def _show_progress(self):
        self._clear()
        total = len(self.docx_files)
        self._header("Converting & Sending …", f"0 of {total}")

        self._status_var = tk.StringVar(value="Starting …")
        tk.Label(
            self.container, textvariable=self._status_var,
            font=("Segoe UI", 9), wraplength=500, justify="left",
            padx=24, pady=4,
        ).pack(fill="x")

        self._progress_bar = ttk.Progressbar(
            self.container, length=500, maximum=total, mode="determinate"
        )
        self._progress_bar.pack(padx=28, pady=6)

        self._log_box = tk.Text(
            self.container, width=60, height=18,
            font=("Consolas", 8), state="disabled", bg="#f5f5f5",
        )
        self._log_box.pack(padx=24, pady=6)

        self.update()
        self.after(100, self._run_conversions)

    def _log(self, text: str):
        self._log_box.configure(state="normal")
        self._log_box.insert("end", text + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")
        self._status_var.set(text)
        self.update()

    def _run_conversions(self):
        errors = []
        total  = len(self.docx_files)

        for i, docx_path in enumerate(self.docx_files):
            page_indices = self.file_pages.get(docx_path)
            label = (
                f"all pages" if page_indices is None
                else f"pages {[p+1 for p in page_indices]}"
            )
            try:
                self._log(f"[{i+1}/{total}] Converting: {docx_path.name}  ({label})")
                pdf_path = convert_docx_to_pdf(docx_path, page_indices)

                if platform.system() != "Windows":
                    pdf_path = add_bookmarks_to_pdf(pdf_path, docx_path)

                self._log(f"  PDF saved → {pdf_path.name}")
                self._log(f"  Emailing …")
                send_pdf_by_email(pdf_path, docx_path.name, self.delete_sent)
                self._log(f"  ✓ Done: {docx_path.name}")

            except Exception as exc:
                msg = f"{docx_path.name}: {exc}"
                log.error(f"Failed – {msg}", exc_info=True)
                self._log(f"  ✗ ERROR: {msg}")
                errors.append(msg)

            self._progress_bar.step(1)
            self.update()

        # Clean up temp preview PDFs
        for tmp in self._tmp_full_pdfs:
            try:
                tmp.unlink(missing_ok=True)
            except Exception:
                pass

        self._finish(errors)

    def _finish(self, errors: list):
        if errors:
            summary = "Finished with errors:\n" + "\n".join(errors)
            messagebox.showwarning("Completed with errors", summary, parent=self)
        else:
            messagebox.showinfo(
                "Done", "All files converted and emailed successfully!", parent=self
            )
        self.destroy()


# ══════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════

def main() -> None:
    app = ConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
