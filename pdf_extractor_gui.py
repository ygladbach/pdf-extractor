#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF -> Excel  Extractor  v2  (Shopfloor-Edition)
=================================================

Features:
  - Multiple PDF selection
  - Auto-naming: CESU_AGUAPURI_<Fecha>_<Turno>.xlsx
  - Overwrite warning before saving
  - All PDFs -> individual .xlsx files in one output folder
"""

# ═══════════════════════════════════════════════════════════════════════
#  IMPORTS
# ═══════════════════════════════════════════════════════════════════════
import re
import os
import sys
import threading
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font


# ═══════════════════════════════════════════════════════════════════════
#  REGEX PATTERNS
# ═══════════════════════════════════════════════════════════════════════
ANCHOR_RE   = re.compile(r"Descripción:\s*Parametro:", re.IGNORECASE)
DATO_RE     = re.compile(r"Dato\s*real:\s*\*", re.IGNORECASE)
PAGE_RE     = re.compile(r"Page\s+\d+\s+of\s+\d+", re.IGNORECASE)

TAG_FOR_NUM_RE = re.compile(
    r"PF/4100(?:-?INS)?-?[A-Z0-9]+(?:-[A-Z0-9]+)*", re.IGNORECASE)
NUM_TOKEN_RE   = re.compile(r"[0-9]+(?:[\.,][0-9]+)?")

UNITS = r"(?:L/min|[Bb]ar|mV|\xb5S/cm|ppb|\xb0C|%|m3/h|litros|L)"
RANGE_RE   = re.compile(
    rf"\b\d+(?:[\.,]\d+)?\s*{UNITS}?\s*a\s*\d+(?:[\.,]\d+)?\s*{UNITS}?\b",
    re.IGNORECASE)
MAX_RE     = re.compile(
    rf"Max\.\s*\d+(?:[\.,]\d+)?\s*(?:{UNITS})?", re.IGNORECASE)
MIN_RE     = re.compile(
    rf"Min\.\s*\d+(?:[\.,]\d+)?\s*(?:{UNITS})?", re.IGNORECASE)
COMP_RE    = re.compile(
    rf"(?:\u2264|\u2265|>=|<=)\s*\d+(?:[\.,][0-9]+)?\s*(?:{UNITS})?",
    re.IGNORECASE)
ATLEAST_RE = re.compile(
    rf"Al\s+menos\s*\d+(?:[\.,][0-9]+)?\s*(?:{UNITS})?", re.IGNORECASE)


# ═══════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════
def norm(s: str) -> str:
    s = s.replace('\u00a0', ' ')
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def to_float(s: str):
    if s is None:
        return None
    s = s.replace(',', '.')
    try:
        return float(s)
    except Exception:
        return None


def pick_actual(after: str):
    cleaned = PAGE_RE.sub(' ', after)
    cleaned = TAG_FOR_NUM_RE.sub(' ', cleaned)
    cleaned = norm(cleaned)
    nums = NUM_TOKEN_RE.findall(cleaned)
    if not nums:
        return None
    if re.match(r"^Min\.", cleaned, flags=re.IGNORECASE):
        return nums[-1]
    return nums[0]


def remove_first(haystack: str, needle: str) -> str:
    idx = haystack.find(needle)
    if idx == -1:
        return haystack
    return haystack[:idx] + haystack[idx + len(needle):]


def remove_last(haystack: str, needle: str) -> str:
    idx = haystack.rfind(needle)
    if idx == -1:
        return haystack
    return haystack[:idx] + haystack[idx + len(needle):]


def normalize_tags(desc: str) -> str:
    desc = re.sub(r"PF/4100-INS-\s*F(\d{2})", r"PF/4100-INSF\1", desc)
    desc = re.sub(r"PF/4100-INS-F(\d{2})", r"PF/4100-INSF\1", desc)
    return desc


def dedupe_desc(desc: str) -> str:
    desc = norm(desc)
    words = desc.split(' ')
    if len(words) % 2 == 0 and len(words) >= 6:
        half = len(words) // 2
        if words[:half] == words[half:]:
            return norm(' '.join(words[:half]))
    return desc


def format_multitag_newline(desc: str) -> str:
    if (re.search(r"PF/4100-F\d+", desc)
            and re.search(r"PF/4100-INS-", desc)):
        desc = re.sub(
            r"\s+(PF/4100-INS-[A-Z0-9\-]+)", r"\n\1", desc, count=1)
    return desc


def strip_noise(text: str) -> str:
    text = re.split(r"\bRealiza:", text, flags=re.IGNORECASE)[0]
    text = re.split(r"\xbf", text)[0]
    text = re.split(
        r"\bContinuidad\s+del\s+negocio\b", text, flags=re.IGNORECASE)[0]
    return norm(text)


def split_desc_spec(merged: str):
    merged = merged.replace('**', ' ').replace('*', ' ')
    merged = merged.replace('turnolitros', 'turno litros')
    merged = norm(merged)
    merged = re.sub(
        r"^Valor\s+obtenido:?\s*", "", merged, flags=re.IGNORECASE)
    merged = PAGE_RE.sub('', merged)
    merged = re.sub(
        r"\bPanel\s+de\s+control\b", "", merged, flags=re.IGNORECASE)
    merged = strip_noise(merged)

    specs = []
    for m in MAX_RE.findall(merged):
        mm = norm(m)
        if mm and mm not in specs:
            specs.append(mm)
        merged = merged.replace(m, ' ')
    for m in MIN_RE.findall(merged):
        mm = norm(m)
        if mm and mm not in specs:
            specs.append(mm)
        merged = merged.replace(m, ' ')

    mr = RANGE_RE.search(merged)
    if mr:
        specs.append(norm(mr.group(0)))
        merged = merged.replace(mr.group(0), ' ')

    ma = ATLEAST_RE.search(merged)
    if ma:
        specs.append(norm(ma.group(0)))
        merged = merged.replace(ma.group(0), ' ')

    for m in COMP_RE.findall(merged):
        mm = norm(m)
        if mm and mm not in specs:
            specs.append(mm)
        merged = merged.replace(m, ' ')

    desc = norm(merged)
    desc = normalize_tags(desc)
    desc = dedupe_desc(desc)

    if not specs:
        m_unit = re.search(r"\blitros\b", desc, flags=re.IGNORECASE)
        if m_unit:
            specs.append('litros')
            desc = norm(
                re.sub(r"\blitros\b", "", desc, flags=re.IGNORECASE))

    desc = format_multitag_newline(desc)
    spec = '; '.join([s.strip(' ;') for s in specs if s]).strip(' ;')
    return desc, spec


# ═══════════════════════════════════════════════════════════════════════
#  METADATA EXTRACTION  (Fecha + Turno from PDF content)
# ═══════════════════════════════════════════════════════════════════════
TURNO_MAP = {
    '1': 'T1', '2': 'T2', '3': 'T3',
}


def extract_pdf_metadata(pdf_path: str) -> tuple:
    """
    Quick-scan the first 5 pages of a PDF to extract Fecha and Turno.
    Returns (fecha_str, turno_code) e.g. ("2026-04-16", "T3").
    Falls back to ("", "") if not found.
    """
    fecha = ""
    turno_code = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages[:5]:
                text += (page.extract_text() or "") + "\n"

            # Look for "Datos de registro" section
            datos_m = re.search(
                r"Datos de registro.*?(?=Descripci\xf3n|\Z)",
                text, re.DOTALL | re.IGNORECASE)
            if datos_m:
                section = datos_m.group(0)
                # Extract date (yyyy-MM-dd)
                fm = re.search(r"(\d{4}-\d{2}-\d{2})", section)
                if fm:
                    fecha = fm.group(1)
                # Extract turno number
                tm = re.search(r"(\d)\w*\s+turno", section, re.IGNORECASE)
                if tm:
                    turno_code = TURNO_MAP.get(tm.group(1), "")
    except Exception:
        pass
    return fecha, turno_code


def build_output_filename(pdf_path: str, out_dir: str,
                          fecha: str, turno_code: str) -> str:
    """
    Build output filename: CESU_AGUAPURI_<Fecha>_<Turno>.xlsx
    Parts CESU and AGUAPURI are extracted from the PDF filename
    (first two hyphen-separated components).
    Falls back to <original_name>_<today>.xlsx if metadata is missing.
    """
    basename = os.path.splitext(os.path.basename(pdf_path))[0]

    # Try to extract first two hyphen-separated parts
    parts = basename.split("-")
    if len(parts) >= 2:
        part1 = parts[0].strip()  # e.g. "CESU"
        part2 = parts[1].strip()  # e.g. "AGUAPURI"
    else:
        part1 = basename
        part2 = ""

    # Build name components
    components = []
    if part1:
        components.append(part1)
    if part2:
        components.append(part2)
    if fecha:
        components.append(fecha)
    else:
        components.append(datetime.date.today().isoformat())
    if turno_code:
        components.append(turno_code)

    filename = "_".join(components) + ".xlsx"
    return os.path.join(out_dir, filename)


# ═══════════════════════════════════════════════════════════════════════
#  CORE LOGIC:  PDF -> Excel
# ═══════════════════════════════════════════════════════════════════════
def process_pdf_to_excel(pdf_path: str, out_path: str,
                         progress_callback=None) -> int:
    """
    Reads the PDF, extracts records, and writes an Excel file.
    Returns the number of extracted rows.
    """
    records = []
    full_text_concat = ""

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        for pnum, page in enumerate(pdf.pages, start=1):
            if progress_callback:
                progress_callback(pnum, total_pages)

            txt = page.extract_text() or ''
            full_text_concat += txt + "\n"

            parts = re.split(r"(?=Descripci\xf3n:\s*Parametro:)", txt)
            for part in parts:
                if not ANCHOR_RE.match(part):
                    continue
                body = norm(
                    ANCHOR_RE.sub('', part, count=1).replace('\n', ' '))
                md = DATO_RE.search(body)
                if not md:
                    continue
                before = body[:md.start()].strip()
                after  = body[md.end():].strip()

                actual_str = pick_actual(after)
                if actual_str is None:
                    continue

                if re.match(r"^\s*Min\.",
                            norm(PAGE_RE.sub(' ', after)),
                            flags=re.IGNORECASE):
                    after_wo_actual = remove_last(after, actual_str)
                else:
                    after_wo_actual = remove_first(after, actual_str)

                after_wo_actual = strip_noise(after_wo_actual)
                merged = norm(before + ' ' + after_wo_actual)
                desc, spec = split_desc_spec(merged)

                records.append({
                    'Pagina': pnum,
                    'Descripci\xf3n': desc,
                    'Parametro:': spec,
                    'Valor obtenido': to_float(actual_str),
                })

    _df = pd.DataFrame(records)

    # ── Fecha/Turno from "Datos de registro" section ─────────────────
    fecha, turno = "", ""
    datos_m = re.search(
        r"Datos de registro.*?(?=Descripci\xf3n|\Z)",
        full_text_concat,
        re.DOTALL | re.IGNORECASE
    )
    if datos_m:
        section = datos_m.group(0)
        fm = re.search(r"(\d{4}-\d{2}-\d{2})", section)
        if fm:
            fecha = fm.group(1)
        tm = re.search(r"(\d+\w*\s+turno)", section, re.IGNORECASE)
        if tm:
            turno = tm.group(1)

    header_row1 = f"Datos de registro - Fecha: {fecha}; Turno: *{turno}"

    # ── Realiza global ───────────────────────────────────────────────
    realiza_row = ""
    rm = re.search(
        r"\bRealiza:\s*\*?\s*([\s\S]+?)(?=\n\n|\bDescripci\xf3n:|\bContinuidad|\xbf|\Z)",
        full_text_concat,
        re.IGNORECASE
    )
    if rm:
        tail = rm.group(1)
        tail = re.split(r"\xbf", tail)[0]
        tail = re.split(r"\bDescripci\xf3n:", tail, flags=re.IGNORECASE)[0]
        tail = re.split(
            r"\bContinuidad\s+del\s+negocio\b", tail, flags=re.IGNORECASE)[0]
        realiza_row = "Realiza: " + norm(tail)

    # ── Write Excel ──────────────────────────────────────────────────
    wb = Workbook()
    ws = wb.active
    ws.title = 'Extraccion'

    # Row 1: header
    ws['A1'] = header_row1
    ws.merge_cells('A1:D1')
    ws['A1'].font = Font(bold=True)
    ws['A1'].alignment = Alignment(
        horizontal='center', vertical='center', wrap_text=True)

    # Row 2: Realiza
    ws['A2'] = realiza_row
    ws.merge_cells('A2:D2')
    ws['A2'].alignment = Alignment(
        horizontal='center', vertical='center', wrap_text=True)

    # Row 3: column headers
    headers = ['Pagina', 'Descripci\xf3n', 'Parametro:', 'Valor obtenido']
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # Row 4+: data
    for r_idx, row in enumerate(_df.itertuples(index=False), start=4):
        ws.cell(r_idx, 1, row[0])
        dcell = ws.cell(r_idx, 2, row[1])
        dcell.alignment = Alignment(wrap_text=True)
        ws.cell(r_idx, 3, row[2])

        val = row[3]
        if val is not None:
            vcell = ws.cell(r_idx, 4, f"{val:.2f}")
            vcell.alignment = Alignment(horizontal='right')

    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 32
    ws.column_dimensions['D'].width = 18
    ws.freeze_panes = 'A4'

    wb.save(out_path)
    return len(_df)


# ═══════════════════════════════════════════════════════════════════════
#  GUI  (Tkinter)
# ═══════════════════════════════════════════════════════════════════════
class PDFExtractorApp:
    """Simple shopfloor interface – multi-file edition."""

    BG       = "#f0f0f0"
    ACCENT   = "#0078D4"
    SUCCESS  = "#107C10"
    BTN_FG   = "#ffffff"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF \u2192 Excel  Extractor")
        self.root.resizable(False, False)

        w, h = 660, 410
        sx = (self.root.winfo_screenwidth()  - w) // 2
        sy = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"{w}x{h}+{sx}+{sy}")
        self.root.configure(bg=self.BG)

        self.pdf_paths: list = []
        self.pdf_display = tk.StringVar(value="")
        self.out_dir = tk.StringVar(value="")

        self._build_ui()

    # ── Build UI ─────────────────────────────────────────────────────
    def _build_ui(self):
        pad = dict(padx=18, pady=6)

        # Title
        tk.Label(
            self.root, text="PDF \u2192 Excel  Extractor",
            font=("Segoe UI", 18, "bold"),
            bg=self.BG, fg=self.ACCENT,
        ).pack(pady=(18, 4))

        tk.Label(
            self.root,
            text="Select one or more PDF files, choose the output folder, and press Start.",
            font=("Segoe UI", 10), bg=self.BG, fg="#555555",
        ).pack(pady=(0, 12))

        # ── PDF selection (multiple) ─────────────────────────────────
        frm1 = tk.Frame(self.root, bg=self.BG)
        frm1.pack(fill="x", **pad)

        self.btn_pdf = tk.Button(
            frm1, text="\U0001f4c2  Select PDF files\u2026",
            font=("Segoe UI", 10), width=22,
            command=self._pick_pdfs,
        )
        self.btn_pdf.pack(side="left")

        self.lbl_pdf = tk.Label(
            frm1, textvariable=self.pdf_display,
            font=("Segoe UI", 9), bg=self.BG, fg="#333",
            anchor="w", wraplength=400,
        )
        self.lbl_pdf.pack(side="left", padx=(10, 0), fill="x", expand=True)

        # ── Output folder ────────────────────────────────────────────
        frm2 = tk.Frame(self.root, bg=self.BG)
        frm2.pack(fill="x", **pad)

        self.btn_out = tk.Button(
            frm2, text="\U0001f4c1  Output folder\u2026",
            font=("Segoe UI", 10), width=22,
            command=self._pick_outdir,
        )
        self.btn_out.pack(side="left")

        self.lbl_out = tk.Label(
            frm2, textvariable=self.out_dir,
            font=("Segoe UI", 9), bg=self.BG, fg="#333",
            anchor="w", wraplength=400,
        )
        self.lbl_out.pack(side="left", padx=(10, 0), fill="x", expand=True)

        # ── Progress ─────────────────────────────────────────────────
        self.progress = ttk.Progressbar(
            self.root, orient="horizontal", length=600, mode="determinate")
        self.progress.pack(pady=(16, 2))

        self.lbl_status = tk.Label(
            self.root, text="Ready.",
            font=("Segoe UI", 9), bg=self.BG, fg="#555",
        )
        self.lbl_status.pack()

        # ── Start button ─────────────────────────────────────────────
        self.btn_start = tk.Button(
            self.root, text="\u25b6   Start",
            font=("Segoe UI", 13, "bold"),
            bg=self.ACCENT, fg=self.BTN_FG,
            activebackground="#005fa3", activeforeground=self.BTN_FG,
            width=20, height=1, relief="flat",
            command=self._on_start,
        )
        self.btn_start.pack(pady=(14, 10))

    # ── File / folder dialogs ────────────────────────────────────────
    def _pick_pdfs(self):
        paths = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if paths:
            self.pdf_paths = list(paths)
            n = len(self.pdf_paths)
            if n == 1:
                self.pdf_display.set(os.path.basename(self.pdf_paths[0]))
            else:
                self.pdf_display.set(f"{n} files selected")

            # Auto-suggest output folder (same as first PDF)
            if not self.out_dir.get():
                self.out_dir.set(os.path.dirname(self.pdf_paths[0]))

    def _pick_outdir(self):
        initial = ""
        if self.pdf_paths:
            initial = os.path.dirname(self.pdf_paths[0])

        folder = filedialog.askdirectory(
            title="Select output folder",
            initialdir=initial,
        )
        if folder:
            self.out_dir.set(folder)

    # ── Validation & start ───────────────────────────────────────────
    def _on_start(self):
        if not self.pdf_paths:
            messagebox.showwarning(
                "Input missing", "Please select one or more PDF files first.")
            return

        missing = [p for p in self.pdf_paths if not os.path.isfile(p)]
        if missing:
            names = "\n".join(os.path.basename(p) for p in missing[:5])
            messagebox.showerror(
                "Files not found",
                f"The following PDF files were not found:\n\n{names}")
            return

        out = self.out_dir.get().strip()
        if not out:
            messagebox.showwarning(
                "Input missing", "Please select an output folder.")
            return
        if not os.path.isdir(out):
            messagebox.showerror(
                "Folder not found",
                f"The output folder was not found:\n{out}")
            return

        # ── Build job list with metadata & overwrite check ───────────
        self.lbl_status.config(text="Reading PDF metadata\u2026", fg="#555")
        self.root.update_idletasks()

        jobs = []  # list of (pdf_path, out_path)
        skipped = []

        for pdf_path in self.pdf_paths:
            fecha, turno_code = extract_pdf_metadata(pdf_path)
            out_path = build_output_filename(pdf_path, out, fecha, turno_code)

            if os.path.exists(out_path):
                fname = os.path.basename(out_path)
                overwrite = messagebox.askyesno(
                    "File already exists",
                    f"The file already exists:\n\n"
                    f"   {fname}\n\n"
                    f"Do you want to overwrite it?")
                if not overwrite:
                    skipped.append(fname)
                    continue

            jobs.append((pdf_path, out_path))

        if not jobs:
            if skipped:
                messagebox.showinfo(
                    "Nothing to do",
                    f"All files were skipped (not overwritten):\n\n"
                    + "\n".join(skipped))
            self.lbl_status.config(text="Ready.", fg="#555")
            return

        # ── Lock UI and start processing ─────────────────────────────
        self._set_ui_locked(True)
        self.progress["value"] = 0
        self.lbl_status.config(text="Processing\u2026", fg="#555")

        thread = threading.Thread(
            target=self._run_batch, args=(jobs, skipped), daemon=True)
        thread.start()

    def _run_batch(self, jobs: list, skipped: list):
        total_files = len(jobs)
        results = []  # (filename, row_count)
        errors = []   # (filename, error_msg)

        for file_idx, (pdf_path, out_path) in enumerate(jobs, start=1):
            fname = os.path.basename(out_path)
            try:
                def progress_cb(cur_page, total_pages,
                                _fi=file_idx, _tf=total_files):
                    self.root.after(
                        0, self._set_batch_progress,
                        _fi, _tf, cur_page, total_pages)

                row_count = process_pdf_to_excel(
                    pdf_path, out_path,
                    progress_callback=progress_cb)
                results.append((fname, row_count))

            except Exception as exc:
                errors.append((fname, str(exc)))

        self.root.after(0, self._on_batch_done, results, errors, skipped)

    # ── Progress & callbacks ─────────────────────────────────────────
    def _set_batch_progress(self, file_idx, total_files,
                            cur_page, total_pages):
        # Overall progress across all files
        file_fraction = (file_idx - 1) / total_files
        page_fraction = cur_page / total_pages / total_files
        pct = int((file_fraction + page_fraction) * 100)
        self.progress["value"] = min(pct, 100)
        self.lbl_status.config(
            text=f"File {file_idx} of {total_files} \u2013 "
                 f"page {cur_page} of {total_pages}\u2026")

    def _on_batch_done(self, results: list, errors: list, skipped: list):
        self.progress["value"] = 100
        self._set_ui_locked(False)

        # Build summary
        total_records = sum(rc for _, rc in results)
        parts = []

        if results:
            parts.append(f"Successfully created {len(results)} file(s) "
                         f"with {total_records} total records:\n")
            for fname, rc in results:
                parts.append(f"   \u2705  {fname}  ({rc} records)")

        if skipped:
            parts.append(f"\nSkipped (not overwritten): {len(skipped)}")
            for fname in skipped:
                parts.append(f"   \u23ed  {fname}")

        if errors:
            parts.append(f"\nErrors: {len(errors)}")
            for fname, err in errors:
                parts.append(f"   \u274c  {fname}: {err}")

        summary = "\n".join(parts)

        if errors and not results:
            # All failed
            self.lbl_status.config(
                text="\u274c  All files failed.", fg="#D83B01")
            messagebox.showerror("Processing error", summary)
        elif errors:
            # Partial success
            self.lbl_status.config(
                text=f"\u26a0  {len(results)} OK, {len(errors)} failed.",
                fg="#D83B01")
            messagebox.showwarning("Partially completed", summary)
        else:
            # All succeeded
            self.lbl_status.config(
                text=f"\u2705  Done \u2013 {len(results)} file(s), "
                     f"{total_records} records extracted.",
                fg=self.SUCCESS)
            messagebox.showinfo("Done!", summary)

    def _on_error(self, error_msg: str):
        self.progress["value"] = 0
        self.lbl_status.config(
            text="\u274c  An error occurred.", fg="#D83B01")
        self._set_ui_locked(False)
        messagebox.showerror(
            "Processing error",
            f"An error occurred:\n\n{error_msg}\n\n"
            f"Please verify that you selected a PDF file.\n"
            f"Please try again.")

    # ── Lock / unlock UI ─────────────────────────────────────────────
    def _set_ui_locked(self, locked: bool):
        state = "disabled" if locked else "normal"
        self.btn_pdf.config(state=state)
        self.btn_out.config(state=state)
        self.btn_start.config(state=state)


# ═══════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()
