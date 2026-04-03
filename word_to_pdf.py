"""
Word to PDF Converter  –  GUI File Picker + Email Delivery
===========================================================
1. Opens a file-picker window so you can select one or more .docx files.
2. Converts each file to PDF via MS Word COM (updates TOC + all fields).
3. Emails each PDF to you as an attachment via SMTP (Gmail / Outlook / any).

REQUIREMENTS
------------
    pip install python-docx pypdf comtypes

    Microsoft Word must be installed (used via COM for field/TOC refresh).

EMAIL SETUP  (edit the CONFIG block below)
------------------------------------------
Gmail:
  • SMTP_HOST = "smtp.gmail.com",  SMTP_PORT = 587
  • Use an App Password (not your regular Gmail password):
    Google Account -> Security -> 2-Step Verification -> App passwords

Outlook / Hotmail:
  • SMTP_HOST = "smtp-mail.outlook.com",  SMTP_PORT = 587

Office 365 work account:
  • SMTP_HOST = "smtp.office365.com",  SMTP_PORT = 587
"""

import os
import sys
import logging
import platform
import smtplib
import tkinter as tk
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import imaplib


# ─────────────────────────── EMAIL CONFIG ──────────────────────────────────
SMTP_HOST     = "smtp.gmail.com"           # change for Outlook / Office 365
SMTP_PORT     = 465                         # 587 = STARTTLS  |  465 = SSL
SMTP_USER     = " "             # your sending address
SMTP_PASSWORD = " "       # Gmail App Password (or normal pw)
EMAIL_TO      = " "    # where the PDF is sent
EMAIL_FROM    = SMTP_USER                   # usually same as SMTP_USER
# ────────────────────────────────────────────────────────────────────────────

# Output folder for PDFs.
# Set to a string path like r"C:\Users\You\PDFs" to save elsewhere.
# Leave as None to save the PDF in the same folder as the source .docx.
OUTPUT_DIR = r"C:\Users\30316376\Downloads"

LOG_FILE = "./converter.log"

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
#  STEP 1 – File picker GUI
# ══════════════════════════════════════════════════════════════════

def pick_docx_files() -> list:
    """Open a Windows file-picker and return selected .docx paths."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    paths = filedialog.askopenfilenames(
        title="Select Word document(s) to convert",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
    )
    root.destroy()

    if not paths:
        log.info("No files selected. Exiting.")
        sys.exit(0)

    return [Path(p) for p in paths]


# ══════════════════════════════════════════════════════════════════
#  STEP 2 – Convert DOCX -> PDF via MS Word COM
# ══════════════════════════════════════════════════════════════════

def _convert_with_msword(docx_path: Path, output_dir: Path) -> Path:
    try:
        import comtypes.client  # type: ignore
    except ImportError:
        raise ImportError("comtypes not found.\nRun:  pip install comtypes")

    abs_docx = str(docx_path.resolve())
    abs_pdf  = str((output_dir / (docx_path.stem + ".pdf")).resolve())

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

        log.info("  Exporting to PDF ...")
        wdExportFormatPDF = 17
        doc.ExportAsFixedFormat(
            OutputFileName=abs_pdf,
            ExportFormat=wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=0,       # wdExportOptimizeForPrint
            CreateBookmarks=1,   # wdExportCreateHeadingBookmarks
            DocStructureTags=True,
        )
        log.info("  Word COM export complete.")
    finally:
        doc.Close(SaveChanges=False)
        word.Quit()

    return Path(abs_pdf)


def _convert_with_libreoffice(docx_path: Path, output_dir: Path) -> Path:
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
    return output_dir / (docx_path.stem + ".pdf")


def convert_docx_to_pdf(docx_path: Path) -> Path:
    out_dir = Path(OUTPUT_DIR) if OUTPUT_DIR else docx_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    if platform.system() == "Windows":
        try:
            return _convert_with_msword(docx_path, out_dir)
        except Exception as exc:
            log.warning(f"  MS Word COM failed ({exc}); trying LibreOffice ...")

    return _convert_with_libreoffice(docx_path, out_dir)


# ══════════════════════════════════════════════════════════════════
#  STEP 3 – Add PDF bookmarks (LibreOffice / non-Windows fallback)
#           MS Word's own export already embeds heading bookmarks.
# ══════════════════════════════════════════════════════════════════

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
#  STEP 4 – Email the PDF
# ══════════════════════════════════════════════════════════════════
def send_pdf_by_email(pdf_path: Path, original_docx_name: str, delete_sent: bool = False) -> None:
    log.info(f"  Sending email to {EMAIL_TO} ...")

    msg = EmailMessage()
    msg["Subject"] = f"Converted PDF: {pdf_path.name}"
    msg["From"]    = EMAIL_FROM
    msg["To"]      = EMAIL_TO
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

    # ── Send ──
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as smtp:
        smtp.login(SMTP_USER, SMTP_PASSWORD)
        smtp.send_message(msg)
    log.info("  Email sent successfully.")

    # Only delete if user said yes
    if delete_sent:
        _delete_from_sent(subject=f"Converted PDF: {pdf_path.name}")


def _delete_from_sent(subject: str) -> None:
    """Connect via IMAP and permanently delete the email from the Sent folder."""

    # Sent folder name varies by provider
    SENT_FOLDER_CANDIDATES = [
        '"[Gmail]/Sent Mail"',   # Gmail
        "Sent",                  # Outlook / generic IMAP
        "Sent Items",            # Office 365
        "INBOX.Sent",            # some others
    ]
    IMAP_HOST = "imap.gmail.com"   # change to imap-mail.outlook.com for Outlook

    log.info("  Connecting to IMAP to delete sent email ...")
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST)
        mail.login(SMTP_USER, SMTP_PASSWORD)

        # Try each possible Sent folder name until one works
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

        # Search for the message by subject
        result, data = mail.search(None, f'SUBJECT "{subject}"')
        if result != "OK" or not data[0]:
            log.warning("  Sent message not found in Sent folder — skipping delete.")
            mail.logout()
            return

        # Mark each match as deleted and expunge (permanent delete)
        for num in data[0].split():
            mail.store(num, "+FLAGS", "\\Deleted")
        mail.expunge()
        log.info("  Sent email permanently deleted from Sent folder.")

        mail.logout()

    except Exception as exc:
        # Non-fatal — don't fail the whole job just because cleanup failed
        log.warning(f"  Could not delete from Sent folder: {exc}")


# ══════════════════════════════════════════════════════════════════
#  PROGRESS WINDOW
# ══════════════════════════════════════════════════════════════════

class ProgressWindow:
    """Simple Tkinter window that shows conversion progress."""

    def __init__(self, total: int):
        self.root = tk.Tk()
        self.root.title("Word to PDF Converter")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)

        pad = {"padx": 16, "pady": 6}

        tk.Label(
            self.root, text="Converting documents ...",
            font=("Segoe UI", 11, "bold")
        ).pack(**pad)

        self.status_var = tk.StringVar(value="Starting ...")
        tk.Label(
            self.root, textvariable=self.status_var,
            font=("Segoe UI", 9), wraplength=380, justify="left"
        ).pack(**pad)

        self.bar = ttk.Progressbar(
            self.root, length=400, maximum=total, mode="determinate"
        )
        self.bar.pack(padx=16, pady=4)

        self.log_box = tk.Text(
            self.root, width=55, height=10,
            font=("Consolas", 8), state="disabled", bg="#f5f5f5"
        )
        self.log_box.pack(padx=16, pady=6)

        self.root.update()

    def set_status(self, text: str):
        self.status_var.set(text)
        self._append_log(text)
        self.root.update()

    def step(self):
        self.bar.step(1)
        self.root.update()

    def _append_log(self, text: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def finish(self, errors: list):
        if errors:
            summary = "Finished with errors:\n" + "\n".join(errors)
            self.status_var.set("Completed with errors.")
            messagebox.showwarning("Completed with errors", summary, parent=self.root)
        else:
            self.status_var.set("All files converted and emailed successfully!")
            messagebox.showinfo("Done", "All files converted and emailed!", parent=self.root)
        self.root.destroy()


# ══════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════
def ask_delete_preference() -> bool:
    """Show a custom dialog explaining the tool and asking about sent email deletion."""
    root = tk.Tk()
    root.title("Word to PDF Converter")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    # Centre the window on screen
    root.update_idletasks()
    w, h = 480, 320
    x = (root.winfo_screenwidth()  // 2) - (w // 2)
    y = (root.winfo_screenheight() // 2) - (h // 2)
    root.geometry(f"{w}x{h}+{x}+{y}")

    result = tk.BooleanVar(value=False)

    # ── Description ──
    desc = (
        "This tool will:\n"
        "  1.  Ask you to select one or more Word (.docx) files.\n"
        "  2.  Update all fields and Table of Contents in each file.\n"
        "  3.  Convert each file to PDF with Heading bookmarks.\n"
        "  4.  Email the PDF to you as an attachment.\n\n"
        "Once you click Yes or No below, a file picker will open\n"
        "so you can select the Word document(s) to convert."
    )
    tk.Label(
        root, text=desc,
        font=("Segoe UI", 9), justify="left",
        wraplength=440, anchor="w",
        padx=20, pady=16,
    ).pack(fill="x")

    # ── Divider ──
    ttk.Separator(root, orient="horizontal").pack(fill="x", padx=20)

    # ── Question ──
    tk.Label(
        root,
        text="After sending, should the email be deleted from your Sent folder?",
        font=("Segoe UI", 9, "bold"),
        wraplength=440, justify="left",
        padx=20, pady=12,
    ).pack(fill="x")

    # ── Buttons ──
    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=10)

    def on_yes():
        result.set(True)
        root.destroy()

    def on_no():
        result.set(False)
        root.destroy()

    tk.Button(
        btn_frame, text="Yes — delete after sending",
        font=("Segoe UI", 9), width=26,
        bg="#d32f2f", fg="white", activebackground="#b71c1c", activeforeground="white",
        relief="flat", cursor="hand2",
        command=on_yes,
    ).grid(row=0, column=0, padx=8)

    tk.Button(
        btn_frame, text="No — keep in Sent folder",
        font=("Segoe UI", 9), width=26,
        bg="#1976d2", fg="white", activebackground="#0d47a1", activeforeground="white",
        relief="flat", cursor="hand2",
        command=on_no,
    ).grid(row=0, column=1, padx=8)

    root.mainloop()
    return result.get()

def main() -> None:
    # 1. Ask about sent email deletion
    delete_sent = ask_delete_preference()
    log.info(f"Delete sent email: {delete_sent}")

    # 2. Let user pick files
    docx_files = pick_docx_files()
    log.info(f"Selected {len(docx_files)} file(s).")

    # 3. Show progress window
    win = ProgressWindow(total=len(docx_files))
    errors = []

    for docx_path in docx_files:
        try:
            win.set_status(f"Converting: {docx_path.name}")
            log.info(f"Processing: {docx_path.name}")

            pdf_path = convert_docx_to_pdf(docx_path)

            if platform.system() != "Windows":
                pdf_path = add_bookmarks_to_pdf(pdf_path, docx_path)

            log.info(f"PDF saved -> {pdf_path}")

            win.set_status(f"Emailing: {pdf_path.name}")
            send_pdf_by_email(pdf_path, docx_path.name, delete_sent=delete_sent)

            win.step()
            log.info(f"Done: {docx_path.name}\n")

        except Exception as exc:
            msg = f"{docx_path.name}: {exc}"
            log.error(f"Failed - {msg}", exc_info=True)
            errors.append(msg)
            win.step()

    win.finish(errors)


if __name__ == "__main__":
    main() 
