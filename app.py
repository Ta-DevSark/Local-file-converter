from __future__ import annotations

import os
import sys
import threading
import queue
import subprocess
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from docx import Document
    from docx.shared import Inches, Pt
except ImportError:
    Document = None


APP_NAME = "Convertisseur Local"
OFFICE_EXTENSIONS = {
    ".doc", ".docx", ".odt", ".rtf", ".txt",
    ".ppt", ".pptx", ".odp",
    ".xls", ".xlsx", ".ods", ".csv"
}
PDF_EXTENSION = ".pdf"


def find_libreoffice() -> str | None:
    """Locate LibreOffice without network access."""
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(candidate)
    return None


def unique_output(path: Path) -> Path:
    """Avoid overwriting an existing file."""
    if not path.exists():
        return path
    stem, suffix, parent = path.stem, path.suffix, path.parent
    index = 1
    while True:
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def convert_office_to_pdf(source: Path, output_dir: Path, soffice: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    profile_dir = output_dir / ".lo_profile"
    profile_dir.mkdir(exist_ok=True)
    command = [
        soffice,
        f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
        "--headless", "--convert-to", "pdf", "--outdir", str(output_dir),
        str(source)
    ]
    result = subprocess.run(
        command, capture_output=True, text=True, timeout=180,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
    )
    expected = output_dir / f"{source.stem}.pdf"
    if result.returncode != 0 or not expected.exists():
        details = (result.stderr or result.stdout or "Erreur inconnue").strip()
        raise RuntimeError(f"LibreOffice n'a pas pu convertir {source.name}.\n{details}")
    final_path = unique_output(expected)
    if final_path != expected:
        expected.rename(final_path)
    return final_path


def extract_pdf_to_txt(source: Path, destination: Path) -> None:
    if fitz is None:
        raise RuntimeError("PyMuPDF est absent. Reconstruisez l'application avec les dépendances prévues.")
    with fitz.open(source) as pdf:
        chunks = []
        for page_number, page in enumerate(pdf, start=1):
            text = page.get_text("text", sort=True).strip()
            chunks.append(f"--- Page {page_number} ---\n{text}")
        content = "\n\n".join(chunks).strip()
    destination.write_text(content + ("\n" if content else ""), encoding="utf-8")


def extract_pdf_to_docx(source: Path, destination: Path) -> None:
    """
    Rebuild a text-oriented DOCX from PDF text blocks.
    Text order is approximated; original typography, tables and page geometry
    are not guaranteed to be reproduced.
    """
    if fitz is None or Document is None:
        raise RuntimeError("Dépendance PDF/DOCX absente. Reconstruisez l'application avec les dépendances prévues.")
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    normal = document.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)

    with fitz.open(source) as pdf:
        for page_index, page in enumerate(pdf):
            if page_index:
                document.add_page_break()
            blocks = page.get_text("blocks", sort=True)
            for block in blocks:
                # PyMuPDF block tuple: x0,y0,x1,y1,text,block_no,block_type,...
                if len(block) < 7 or block[6] != 0:
                    continue
                text = str(block[4]).strip()
                if not text:
                    continue
                # Preserve line breaks within each text block.
                paragraph = document.add_paragraph()
                paragraph.paragraph_format.space_after = Pt(4)
                for line_index, line in enumerate(text.splitlines()):
                    if line_index:
                        paragraph.add_run().add_break()
                    paragraph.add_run(line.rstrip())
    if not document.paragraphs:
        document.add_paragraph("Aucun texte extractible n'a été détecté dans ce PDF.")
    document.save(destination)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("760x570")
        self.minsize(680, 500)
        self.files: list[Path] = []
        self.events: queue.Queue = queue.Queue()
        self.running = False
        self._build_ui()
        self.after(120, self._poll_events)

    def _build_ui(self):
        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text=APP_NAME, font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(
            root,
            text="Conversions locales • Sans compte • Sans IA • Sans connexion Internet",
            foreground="#555"
        ).pack(anchor="w", pady=(2, 14))

        source_frame = ttk.LabelFrame(root, text="1. Fichiers à convertir", padding=10)
        source_frame.pack(fill="both", expand=True)
        toolbar = ttk.Frame(source_frame)
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Button(toolbar, text="Ajouter des fichiers…", command=self.add_files).pack(side="left")
        ttk.Button(toolbar, text="Retirer la sélection", command=self.remove_selected).pack(side="left", padx=8)
        ttk.Button(toolbar, text="Tout effacer", command=self.clear_files).pack(side="left")

        self.file_list = tk.Listbox(source_frame, height=8, selectmode=tk.EXTENDED)
        self.file_list.pack(fill="both", expand=True)

        options = ttk.LabelFrame(root, text="2. Conversion", padding=10)
        options.pack(fill="x", pady=12)

        self.mode = tk.StringVar(value="office")
        modes = ttk.Frame(options)
        modes.pack(fill="x")
        ttk.Radiobutton(
            modes, text="Documents Office / bureautique → PDF",
            variable=self.mode, value="office", command=self._update_format
        ).pack(anchor="w")
        ttk.Radiobutton(
            modes, text="PDF → document texte",
            variable=self.mode, value="pdf", command=self._update_format
        ).pack(anchor="w", pady=(4, 0))

        format_row = ttk.Frame(options)
        format_row.pack(fill="x", pady=(8, 0))
        ttk.Label(format_row, text="Format de sortie :").pack(side="left")
        self.output_format = tk.StringVar(value="PDF")
        self.format_combo = ttk.Combobox(
            format_row, textvariable=self.output_format, state="readonly",
            values=["PDF"], width=12
        )
        self.format_combo.pack(side="left", padx=8)

        output_row = ttk.Frame(options)
        output_row.pack(fill="x", pady=(8, 0))
        ttk.Label(output_row, text="Dossier de destination :").pack(side="left")
        self.output_dir = tk.StringVar(value="")
        ttk.Entry(output_row, textvariable=self.output_dir).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(output_row, text="Parcourir…", command=self.choose_output_dir).pack(side="left")

        action_row = ttk.Frame(root)
        action_row.pack(fill="x")
        self.convert_button = ttk.Button(action_row, text="Convertir", command=self.start_conversion)
        self.convert_button.pack(side="left")
        self.progress = ttk.Progressbar(action_row, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=12)

        self.status = tk.StringVar(value="En attente.")
        ttk.Label(root, textvariable=self.status).pack(anchor="w", pady=(8, 4))

        log_frame = ttk.LabelFrame(root, text="Journal", padding=8)
        log_frame.pack(fill="both", expand=True)
        self.log = tk.Text(log_frame, height=6, wrap="word", state="disabled", font=("Consolas", 9))
        self.log.pack(fill="both", expand=True)

        ttk.Label(
            root,
            text="PDF scanné : extraction du texte impossible sans OCR. La conversion PDF → DOCX reconstruit le texte, pas la mise en page originale.",
            foreground="#555", wraplength=710
        ).pack(anchor="w", pady=(10, 0))

    def _update_format(self):
        if self.mode.get() == "office":
            self.format_combo.configure(values=["PDF"])
            self.output_format.set("PDF")
        else:
            self.format_combo.configure(values=["DOCX", "TXT"])
            self.output_format.set("DOCX")

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Sélectionner les fichiers",
            filetypes=[
                ("Documents pris en charge", "*.doc *.docx *.odt *.rtf *.txt *.ppt *.pptx *.odp *.xls *.xlsx *.ods *.csv *.pdf"),
                ("Tous les fichiers", "*.*")
            ]
        )
        for item in paths:
            path = Path(item)
            if path not in self.files:
                self.files.append(path)
                self.file_list.insert("end", str(path))

    def remove_selected(self):
        indices = list(self.file_list.curselection())
        for index in reversed(indices):
            self.file_list.delete(index)
            del self.files[index]

    def clear_files(self):
        self.files.clear()
        self.file_list.delete(0, "end")

    def choose_output_dir(self):
        path = filedialog.askdirectory(title="Choisir le dossier de destination")
        if path:
            self.output_dir.set(path)

    def _write_log(self, message: str):
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start_conversion(self):
        if self.running:
            return
        if not self.files:
            messagebox.showwarning(APP_NAME, "Ajoutez au moins un fichier.")
            return
        mode = self.mode.get()
        if mode == "office":
            invalid = [p for p in self.files if p.suffix.lower() not in OFFICE_EXTENSIONS]
            if invalid:
                messagebox.showerror(APP_NAME, "Le mode Documents → PDF accepte uniquement des documents bureautiques.")
                return
            soffice = find_libreoffice()
            if not soffice:
                messagebox.showerror(
                    APP_NAME,
                    "LibreOffice est requis pour convertir les documents en PDF.\n"
                    "Installez LibreOffice, puis relancez l'application."
                )
                return
        else:
            invalid = [p for p in self.files if p.suffix.lower() != ".pdf"]
            if invalid:
                messagebox.showerror(APP_NAME, "Le mode PDF → texte accepte uniquement des fichiers PDF.")
                return
            soffice = None

        destination = Path(self.output_dir.get()) if self.output_dir.get().strip() else self.files[0].parent
        destination.mkdir(parents=True, exist_ok=True)
        self.running = True
        self.convert_button.configure(state="disabled")
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0
        self.status.set("Conversion en cours…")
        threading.Thread(
            target=self._worker,
            args=(list(self.files), mode, self.output_format.get(), destination, soffice),
            daemon=True
        ).start()

    def _worker(self, files, mode, output_format, destination, soffice):
        for index, source in enumerate(files, start=1):
            try:
                if not source.is_file():
                    raise FileNotFoundError("Fichier introuvable.")
                if mode == "office":
                    output = unique_output(destination / f"{source.stem}.pdf")
                    # Convert first to a temporary folder so that existing outputs are never overwritten.
                    temp_dir = destination / ".conversion_temp"
                    temp_dir.mkdir(exist_ok=True)
                    produced = convert_office_to_pdf(source, temp_dir, soffice)
                    produced.replace(output)
                elif output_format == "TXT":
                    output = unique_output(destination / f"{source.stem}.txt")
                    extract_pdf_to_txt(source, output)
                else:
                    output = unique_output(destination / f"{source.stem}.docx")
                    extract_pdf_to_docx(source, output)
                self.events.put(("log", f"OK  {source.name}  →  {output.name}"))
            except Exception as exc:
                self.events.put(("log", f"ERREUR  {source.name}  —  {exc}"))
            self.events.put(("progress", index))
        self.events.put(("done", None))

    def _poll_events(self):
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "log":
                    self._write_log(payload)
                elif event == "progress":
                    self.progress["value"] = payload
                    self.status.set(f"{payload}/{int(self.progress['maximum'])} fichier(s) traité(s)")
                elif event == "done":
                    self.running = False
                    self.convert_button.configure(state="normal")
                    self.status.set("Traitement terminé.")
                    self._write_log("— Fin du traitement —")
        except queue.Empty:
            pass
        self.after(120, self._poll_events)


if __name__ == "__main__":
    app = App()
    app.mainloop()
