#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
This script has an IQMS file as input (CESU-PDF) and
extractsthe parameters in an excel file.
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
#  REGEX-MUSTER
# ═══════════════════════════════════════════════════════════════════════
ANCHOR_RE   = re.compile(r"Descripción:\s*Parametro:", re.IGNORECASE)
DATO_RE     = re.compile(r"Dato\s*real:\s*\*", re.IGNORECASE)
PAGE_RE     = re.compile(r"Page\s+\d+\s+of\s+\d+", re.IGNORECASE)

TAG_FOR_NUM_RE = re.compile(
    r"PF/4100(?:-?INS)?-?[A-Z0-9]+(?:-[A-Z0-9]+)*", re.IGNORECASE)
NUM_TOKEN_RE   = re.compile(r"[0-9]+(?:[\.,][0-9]+)?")

UNITS = r"(?:L/min|[Bb]ar|mV|µS/cm|ppb|°C|%|m3/h|litros|L)"
RANGE_RE   = re.compile(
    rf"\b\d+(?:[\.,]\d+)?\s*{UNITS}?\s*a\s*\d+(?:[\.,]\d+)?\s*{UNITS}?\b",
    re.IGNORECASE)
MAX_RE     = re.compile(
    rf"Max\.\s*\d+(?:[\.,]\d+)?\s*(?:{UNITS})?", re.IGNORECASE)
MIN_RE     = re.compile(
    rf"Min\.\s*\d+(?:[\.,]\d+)?\s*(?:{UNITS})?", re.IGNORECASE)
COMP_RE    = re.compile(
    rf"(?:≤|≥|>=|<=)\s*\d+(?:[\.,][0-9]+)?\s*(?:{UNITS})?", re.IGNORECASE)
ATLEAST_RE = re.compile(
    rf"Al\s+menos\s*\d+(?:[\.,][0-9]+)?\s*(?:{UNITS})?", re.IGNORECASE)


# ═══════════════════════════════════════════════════════════════════════
#  HILFSFUNKTIONEN
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
    text = re.split(r"\bRealiza:\b", text, flags=re.IGNORECASE)[0]
    text = re.split(r"¿", text)[0]
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
#  KERN-LOGIK:  PDF einlesen  →  Excel erzeugen
# ═══════════════════════════════════════════════════════════════════════
def process_pdf_to_excel(pdf_path: str, out_path: str,
                         progress_callback=None) -> int:
    """
    Liest die PDF ein, extrahiert die Datensätze und schreibt eine
    Excel-Datei.  Gibt die Anzahl der extrahierten Zeilen zurück.
    """
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        for pnum, page in enumerate(pdf.pages, start=1):
            if progress_callback:
                progress_callback(pnum, total_pages)

            txt = page.extract_text() or ''
            parts = re.split(r"(?=Descripción:\s*Parametro:)", txt)
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
                    'Descripción': desc,
                    'Parametro:': spec,
                    'Valor obtenido': to_float(actual_str),
                })

    # ── Excel-Datei schreiben ────────────────────────────────────────
    _df = pd.DataFrame(records)
    header_text = 'Flujo a la entrada del sistema PF/4100-INSF03'

    wb = Workbook()
    ws = wb.active
    ws.title = 'Extraccion'

    ws['A1'] = header_text
    ws.merge_cells('A1:D1')
    ws['A1'].font = Font(bold=True)
    ws['A1'].alignment = Alignment(
        horizontal='center', vertical='center', wrap_text=True)

    headers = ['Pagina', 'Descripción', 'Parametro:', 'Valor obtenido']
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    for r_idx, row in enumerate(_df.itertuples(index=False), start=3):
        ws.cell(r_idx, 1, row[0])
        dcell = ws.cell(r_idx, 2, row[1])
        dcell.alignment = Alignment(wrap_text=True)
        ws.cell(r_idx, 3, row[2])
        vcell = ws.cell(r_idx, 4, row[3])
        vcell.number_format = '0.###'

    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 32
    ws.column_dimensions['D'].width = 18
    ws.freeze_panes = 'A3'

    wb.save(out_path)
    return len(_df)


# ═══════════════════════════════════════════════════════════════════════
#  GUI  (Tkinter)
# ═══════════════════════════════════════════════════════════════════════
class PDFExtractorApp:
    """Einfache Oberfläche ."""

    BG       = "#f0f0f0"
    ACCENT   = "#0078D4"
    SUCCESS  = "#107C10"
    BTN_FG   = "#ffffff"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF → Excel  Extractor")
        self.root.resizable(False, False)

        # ── Fenstergröße & Zentrierung ───────────────────────────────
        w, h = 660, 390
        sx = (self.root.winfo_screenwidth()  - w) // 2
        sy = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"{w}x{h}+{sx}+{sy}")
        self.root.configure(bg=self.BG)

        self.pdf_path = tk.StringVar(value="")
        self.out_path = tk.StringVar(value="")

        self._build_ui()

    # ── UI aufbauen ──────────────────────────────────────────────────
    def _build_ui(self):
        pad = dict(padx=18, pady=6)

        # Titel
        tk.Label(
            self.root, text="PDF → Excel  Extractor",
            font=("Segoe UI", 18, "bold"),
            bg=self.BG, fg=self.ACCENT,
        ).pack(pady=(18, 4))

        tk.Label(
            self.root,
            text="PDF-Datei auswählen, Ziel festlegen, Starten drücken.",
            font=("Segoe UI", 10), bg=self.BG, fg="#555555",
        ).pack(pady=(0, 12))

        # ── Eingabe-PDF ──────────────────────────────────────────────
        frm1 = tk.Frame(self.root, bg=self.BG)
        frm1.pack(fill="x", **pad)

        self.btn_pdf = tk.Button(
            frm1, text="📂  PDF auswählen …",
            font=("Segoe UI", 10), width=22,
            command=self._pick_pdf,
        )
        self.btn_pdf.pack(side="left")

        self.lbl_pdf = tk.Label(
            frm1, textvariable=self.pdf_path,
            font=("Segoe UI", 9), bg=self.BG, fg="#333",
            anchor="w", wraplength=400,
        )
        self.lbl_pdf.pack(side="left", padx=(10, 0), fill="x", expand=True)

        # ── Ausgabe-Excel ────────────────────────────────────────────
        frm2 = tk.Frame(self.root, bg=self.BG)
        frm2.pack(fill="x", **pad)

        self.btn_out = tk.Button(
            frm2, text="💾  Ziel-Excel festlegen …",
            font=("Segoe UI", 10), width=22,
            command=self._pick_out,
        )
        self.btn_out.pack(side="left")

        self.lbl_out = tk.Label(
            frm2, textvariable=self.out_path,
            font=("Segoe UI", 9), bg=self.BG, fg="#333",
            anchor="w", wraplength=400,
        )
        self.lbl_out.pack(side="left", padx=(10, 0), fill="x", expand=True)

        # ── Fortschrittsbalken ───────────────────────────────────────
        self.progress = ttk.Progressbar(
            self.root, orient="horizontal", length=600, mode="determinate")
        self.progress.pack(pady=(16, 2))

        self.lbl_status = tk.Label(
            self.root, text="Bereit.",
            font=("Segoe UI", 9), bg=self.BG, fg="#555",
        )
        self.lbl_status.pack()

        # ── Starten-Button ───────────────────────────────────────────
        self.btn_start = tk.Button(
            self.root, text="▶   Starten",
            font=("Segoe UI", 13, "bold"),
            bg=self.ACCENT, fg=self.BTN_FG,
            activebackground="#005fa3", activeforeground=self.BTN_FG,
            width=20, height=1, relief="flat",
            command=self._on_start,
        )
        self.btn_start.pack(pady=(14, 10))

    # ── Datei-Dialoge ────────────────────────────────────────────────
    def _pick_pdf(self):
        path = filedialog.askopenfilename(
            title="PDF-Datei auswählen",
            filetypes=[("PDF-Dateien", "*.pdf"), ("Alle Dateien", "*.*")],
        )
        if path:
            self.pdf_path.set(path)
            # automatisch einen Ausgabepfad vorschlagen
            base = os.path.splitext(os.path.basename(path))[0]
            today = datetime.date.today().isoformat()
            suggestion = os.path.join(
                os.path.dirname(path),
                f"{base}_{today}.xlsx",
            )
            if not self.out_path.get():
                self.out_path.set(suggestion)

    def _pick_out(self):
        initial_dir  = ""
        initial_file = ""
        if self.pdf_path.get():
            initial_dir = os.path.dirname(self.pdf_path.get())
            base = os.path.splitext(
                os.path.basename(self.pdf_path.get()))[0]
            today = datetime.date.today().isoformat()
            initial_file = f"{base}_{today}.xlsx"

        path = filedialog.asksaveasfilename(
            title="Excel-Datei speichern unter",
            defaultextension=".xlsx",
            initialdir=initial_dir,
            initialfile=initial_file,
            filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if path:
            self.out_path.set(path)

    # ── Validierung & Start ──────────────────────────────────────────
    def _on_start(self):
        pdf = self.pdf_path.get().strip()
        out = self.out_path.get().strip()

        if not pdf:
            messagebox.showwarning(
                "Eingabe fehlt", "Bitte zuerst eine PDF-Datei auswählen.")
            return
        if not os.path.isfile(pdf):
            messagebox.showerror(
                "Datei nicht gefunden",
                f"Die PDF-Datei wurde nicht gefunden:\n{pdf}")
            return
        if not out:
            messagebox.showwarning(
                "Eingabe fehlt",
                "Bitte einen Speicherort für die Excel-Datei festlegen.")
            return

        # UI sperren
        self._set_ui_locked(True)
        self.progress["value"] = 0
        self.lbl_status.config(text="Verarbeitung läuft …", fg="#555")

        # Verarbeitung im Hintergrund starten
        thread = threading.Thread(
            target=self._run_extraction, args=(pdf, out), daemon=True)
        thread.start()

    def _run_extraction(self, pdf_path: str, out_path: str):
        try:
            row_count = process_pdf_to_excel(
                pdf_path, out_path,
                progress_callback=self._update_progress,
            )
            self.root.after(0, self._on_success, row_count, out_path)
        except Exception as exc:
            self.root.after(0, self._on_error, str(exc))

    # ── Callbacks (werden im Hauptthread aufgerufen) ─────────────────
    def _update_progress(self, current_page: int, total_pages: int):
        pct = int(current_page / total_pages * 100)
        self.root.after(0, self._set_progress, pct, current_page, total_pages)

    def _set_progress(self, pct, cur, total):
        self.progress["value"] = pct
        self.lbl_status.config(
            text=f"Seite {cur} von {total} wird verarbeitet …")

    def _on_success(self, row_count: int, out_path: str):
        self.progress["value"] = 100
        self.lbl_status.config(
            text=f"✅  Fertig – {row_count} Datensätze extrahiert.",
            fg=self.SUCCESS)
        self._set_ui_locked(False)
        messagebox.showinfo(
            "Fertig!",
            f"Extraktion abgeschlossen.\n\n"
            f"   Datensätze:  {row_count}\n"
            f"   Datei:  {os.path.basename(out_path)}\n\n"
            f"Gespeichert unter:\n{out_path}",
        )

    def _on_error(self, error_msg: str):
        self.progress["value"] = 0
        self.lbl_status.config(text="❌  Fehler aufgetreten.", fg="#D83B01")
        self._set_ui_locked(False)
        messagebox.showerror(
            "Fehler bei der Verarbeitung",
            f"Es ist ein Fehler aufgetreten:\n\n{error_msg}\n\n"
            f"Bitte prüfen Sie, ob die richtige PDF-Datei\n"
            f"ausgewählt wurde und versuchen Sie es erneut.",
        )

    # ── UI sperren / entsperren ──────────────────────────────────────
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
