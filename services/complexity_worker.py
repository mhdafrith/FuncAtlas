"""
services/complexity_worker.py
──────────────────────────────
ComplexityAnalysisWorker  – scans source files, saves function bodies as .txt,
builds an Excel report with:
  Sheet 1 "Function_Complexity" — per-function construct counts + weighted
           complexity score + level (Low / Medium / High / Very High / Complex)
  Sheet 2 "Construct_Summary"   — total count of every construct across all
           functions
"""

import os
import re

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PySide6.QtCore import QObject, Signal

from core.utils import (
    normalize_path, normalize_name, iter_source_files,
    detect_functions_in_file, extract_function_body,
)


# ── Construct patterns (C language) ──────────────────────────────────────────
CONSTRUCTS = [
    # ── Control Flow ──────────────────────────────────────────────────────────
    ("If Statement",        r'\bif\s*\('),
    ("Else Branch",         r'\belse\b(?!\s*if)'),
    ("Else-if Chain",       r'\belse\s+if\s*\('),
    ("Switch Statement",    r'\bswitch\s*\('),
    ("Case Label",          r'\bcase\s+'),
    ("Default Label",       r'\bdefault\s*:'),
    ("Break Statement",     r'\bbreak\s*;'),
    ("For Loop",            r'\bfor\s*\('),
    ("Return Statement",    r'\breturn\b'),
    ("Continue Statement",  r'\bcontinue\s*;'),
    ("While Loop",          r'\bwhile\s*\('),
    ("Do...While",          r'\bdo\s*\{'),
    # ── Function ──────────────────────────────────────────────────────────────
    ("Function Call",       r'\b[A-Za-z_]\w*\s*\('),
    ("Function Definition", r'\b[A-Za-z_]\w*\s+[A-Za-z_]\w*\s*\([^)]*\)\s*\{'),
    ("Function Declaration",r'\b[A-Za-z_]\w*\s+[A-Za-z_]\w*\s*\([^)]*\)\s*;'),
    # ── Pointer ───────────────────────────────────────────────────────────────
    ("Address-of Operator", r'&[A-Za-z_]\w*'),
    ("Pointer Declaration", r'\b[A-Za-z_]\w*\s*\*+\s*[A-Za-z_]\w*'),
    ("Pointer Dereference", r'\*[A-Za-z_]\w*'),
    ("Arrow Operator",      r'->\s*[A-Za-z_]\w*'),
    ("Void Pointer",        r'\bvoid\s*\*'),
    ("Pointer Arithmetic",  r'\b[A-Za-z_]\w*\s*[\+\-]\s*\d+|\b[A-Za-z_]\w*\s*\+\+|\+\+\s*[A-Za-z_]\w*'),
    # ── Array ─────────────────────────────────────────────────────────────────
    ("Array Access",        r'[A-Za-z_]\w*\s*\['),
    ("Array Declaration",   r'[A-Za-z_]\w*\s+[A-Za-z_]\w*\s*\[\d*\]'),
    ("String Literal",      r'"[^"]*"'),
    ("Char Array",          r'\bchar\s+[A-Za-z_]\w*\s*\['),
    # ── Type ──────────────────────────────────────────────────────────────────
    ("Cast Operation",      r'\(\s*[A-Za-z_]\w*\s*\*?\s*\)\s*[A-Za-z_\(]'),
    ("Typedef",             r'\btypedef\b'),
    ("Enum Definition",     r'\benum\s+\w*\s*\{'),
    ("Sizeof Operator",     r'\bsizeof\s*\('),
    ("Const Declaration",   r'\bconst\b'),
    ("Signed/Unsigned",     r'\b(?:signed|unsigned)\b'),
    ("Short/Long",          r'\b(?:short|long)\b'),
    # ── Storage ───────────────────────────────────────────────────────────────
    ("Static Keyword",      r'\bstatic\b'),
    # ── Preprocessor ─────────────────────────────────────────────────────────
    ("Macro Definition",    r'#\s*define\b'),
    ("Ifdef Directive",     r'#\s*ifdef\b'),
    ("If Directive",        r'#\s*if\b(?!def)'),
    ("Elif Directive",      r'#\s*elif\b'),
    ("Else Directive",      r'#\s*else\b'),
    ("Endif Directive",     r'#\s*endif\b'),
    ("Pragma Directive",    r'#\s*pragma\b'),
    # ── Operator ──────────────────────────────────────────────────────────────
    ("Compound Assignment", r'[A-Za-z_]\w*\s*(?:\+=|-=|\*=|/=|%=|&=|\|=|\^=|<<=|>>=)'),
    ("Increment",           r'\+\+'),
    ("Decrement",           r'--'),
    ("Bitwise AND",         r'(?<![&])&(?![&])'),
    ("Bitwise OR",          r'(?<!\|)\|(?!\|)'),
    ("Bitwise XOR",         r'\^'),
    ("Bitwise NOT",         r'~'),
    ("Left Shift",          r'<<'),
    ("Right Shift",         r'>>'),
    ("Logical AND",         r'&&'),
    ("Logical OR",          r'\|\|'),
    ("Logical NOT",         r'!(?!=)'),
    # ── Safety ────────────────────────────────────────────────────────────────
    ("NULL Check",          r'\bNULL\b.*(?:==|!=)|(?:==|!=).*\bNULL\b'),
    ("NULL Assignment",     r'=\s*NULL\b'),
    # ── Memory ────────────────────────────────────────────────────────────────
    ("memset",              r'\bmemset\s*\('),
    ("memcpy",              r'\bmemcpy\s*\('),
    # ── Math ──────────────────────────────────────────────────────────────────
    ("fabs",                r'\bfabs\s*\('),
    # ── Misc ──────────────────────────────────────────────────────────────────
    ("Designated Initializer", r'\.\s*[A-Za-z_]\w*\s*='),
    ("Compound Literal",    r'\(\s*[A-Za-z_]\w+\s*\)\s*\{'),
    ("Bit Field",           r':\s*\d+\s*;'),
]


def _strip_comments(text: str) -> str:
    text = re.sub(r'/\*.*?\*/', ' ', text, flags=re.DOTALL)
    text = re.sub(r'//[^\n]*', ' ', text)
    return text


def count_constructs(body: str) -> dict:
    """Return {construct_name: count} for a function body string."""
    clean = _strip_comments(body)
    counts = {}
    for name, pattern in CONSTRUCTS:
        counts[name] = len(re.findall(pattern, clean))
    return counts


def complexity_level(score: float, bands: list) -> str:
    """Given a numeric score and band list [(label, start, end), ...], return level label."""
    for label, start, end in bands:
        if start <= score <= end:
            return label
    # above all bands
    if bands:
        return bands[-1][0]
    return "Unknown"


# ── Worker ────────────────────────────────────────────────────────────────────
class ComplexityAnalysisWorker(QObject):
    progress = Signal(int, str)   # percent, message
    log      = Signal(str)
    finished = Signal(str)        # path to generated Excel
    error    = Signal(str)

    def __init__(self, source_folder: str, output_root: str,
                 weights: list = None, bands: list = None):
        """
        source_folder : root folder to scan
        output_root   : where to create function_body/ and the Excel
        weights       : [(name, weight), ...]  — from ComplexitySettingsDialog
        bands         : [(label, start, end), ...] — from ComplexitySettingsDialog
        """
        super().__init__()
        self.source_folder = normalize_path(source_folder)
        self.output_root   = normalize_path(output_root)
        self.weights       = {n: w for n, w in (weights or [])} or \
                             {n: 1 for n, _ in CONSTRUCTS}
        self.bands         = bands or [
            ("Low",       1,   5),
            ("Medium",    6,  12),
            ("High",     13,  25),
            ("Very High", 26,  40),
            ("Complex",  41, 999),
        ]

    # ── main entry ────────────────────────────────────────────────────────────
    def run(self):
        try:
            self._run()
        except Exception as exc:
            import traceback
            self.error.emit(f"{exc}\n{traceback.format_exc()}")

    def _run(self):
        body_dir = os.path.join(self.output_root, "function_body")
        os.makedirs(body_dir, exist_ok=True)
        self.log.emit(f"📁 Output folder: {self.output_root}")

        # Scan all source files
        file_entries = list(iter_source_files(self.source_folder))
        total_files  = len(file_entries)
        if total_files == 0:
            self.error.emit("No source files found in the selected folder.")
            return

        self.log.emit(f"🔍 Found {total_files} source files — scanning …")
        all_records = []   # list of dicts for Excel

        for file_idx, (full_path, file_name) in enumerate(file_entries):
            pct = int((file_idx / total_files) * 85)
            self.progress.emit(pct, f"Scanning {file_name} …")

            functions = detect_functions_in_file(full_path)
            if not functions:
                continue

            rel_path = os.path.join(os.path.basename(self.source_folder), os.path.relpath(full_path, self.source_folder))

            for fn_name in functions:
                body = extract_function_body(full_path, fn_name)

                # Save .txt
                safe_name = re.sub(r'[\\/:*?"<>|]', '_', f"{file_name}__{fn_name}")
                txt_path  = os.path.join(body_dir, f"{safe_name}.txt")
                try:
                    with open(txt_path, "w", encoding="utf-8") as fh:
                        fh.write(f"// File     : {rel_path}\n")
                        fh.write(f"// Function : {fn_name}\n")
                        fh.write("// " + "-" * 60 + "\n\n")
                        fh.write(body)
                except Exception as e:
                    self.log.emit(f"  ⚠ Could not save {safe_name}.txt: {e}")

                counts = count_constructs(body)

                # Weighted score
                score = sum(counts.get(cn, 0) * self.weights.get(cn, 1)
                            for cn, _ in CONSTRUCTS)
                level = complexity_level(score, self.bands)

                all_records.append({
                    "file_path":   rel_path,
                    "file_name":   file_name,
                    "function":    fn_name,
                    "counts":      counts,
                    "score":       score,
                    "level":       level,
                })

        self.log.emit(f"✅ Extracted {len(all_records)} functions — building Excel …")
        self.progress.emit(90, "Writing Excel report …")

        out_path = self._write_excel(all_records)
        self.progress.emit(100, "Done")
        self.log.emit(f"📊 Report saved: {out_path}")
        self.finished.emit(out_path)

    # ── Excel writer ──────────────────────────────────────────────────────────
    def _write_excel(self, records: list) -> str:
        wb = Workbook()

        # ── Sheet 1: Function Complexity ──────────────────────────────────────
        ws1 = wb.active
        ws1.title = "Function_Complexity"

        construct_names = [c[0] for c in CONSTRUCTS]
        headers = ["File Name", "File Path", "Function Name"] + \
                  construct_names + ["Complexity Score", "Complexity Level"]

        thin    = Side(style="thin",   color="C5D8EC")
        thick   = Side(style="medium", color="1A3A5C")
        def bdr(): return Border(left=thin, right=thin, top=thin, bottom=thin)

        hdr_fill = PatternFill("solid", fgColor="1F4E78")
        hdr_font = Font(color="FFFFFF", bold=True, name="Arial", size=10)

        level_colors = {
            "Low":       "D6E4BC",
            "Medium":    "FFE699",
            "High":      "F4B183",
            "Very High": "FF7070",
            "Complex":   "CC0000",
        }
        level_font_dark = {"Complex"}

        for col_idx, hdr in enumerate(headers, start=1):
            cell = ws1.cell(row=1, column=col_idx, value=hdr)
            cell.font  = hdr_font
            cell.fill  = hdr_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = bdr()

        for row_idx, rec in enumerate(records, start=2):
            row_vals = [
                rec["file_path"],
                rec["file_name"],
                rec["function"],
            ] + [rec["counts"].get(cn, 0) for cn in construct_names] + [
                rec["score"],
                rec["level"],
            ]
            for col_idx, val in enumerate(row_vals, start=1):
                cell = ws1.cell(row=row_idx, column=col_idx, value=val)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = bdr()
                # Colour the level cell
                if col_idx == len(headers):
                    lvl = rec["level"]
                    fg  = level_colors.get(lvl, "FFFFFF")
                    cell.fill = PatternFill("solid", fgColor=fg)
                    cell.font = Font(
                        bold=True,
                        color="FFFFFF" if lvl in level_font_dark else "1A1A1A",
                        name="Arial", size=10
                    )
                # Left-align text columns
                if col_idx <= 3:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        # Auto column widths
        for col in ws1.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=10)
            ws1.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)
        ws1.row_dimensions[1].height = 30
        ws1.freeze_panes = "A2"

        # ── Sheet 2: Construct Summary ────────────────────────────────────────
        ws2 = wb.create_sheet("Construct_Summary")
        ws2.cell(row=1, column=1, value="Construct").font  = hdr_font
        ws2.cell(row=1, column=1).fill   = hdr_fill
        ws2.cell(row=1, column=1).alignment = Alignment(horizontal="center")
        ws2.cell(row=1, column=2, value="Total Count").font = hdr_font
        ws2.cell(row=1, column=2).fill   = hdr_fill
        ws2.cell(row=1, column=2).alignment = Alignment(horizontal="center")

        totals = {cn: sum(r["counts"].get(cn, 0) for r in records) for cn in construct_names}
        for row_idx, cn in enumerate(construct_names, start=2):
            ws2.cell(row=row_idx, column=1, value=cn).alignment  = Alignment(horizontal="left")
            ws2.cell(row=row_idx, column=2, value=totals[cn]).alignment = Alignment(horizontal="center")
            ws2.cell(row=row_idx, column=1).border = bdr()
            ws2.cell(row=row_idx, column=2).border = bdr()

        ws2.column_dimensions["A"].width = 28
        ws2.column_dimensions["B"].width = 16

        out_path = os.path.join(self.output_root, "FuncAtlas_Complexity_Report.xlsx")
        wb.save(out_path)
        return out_path

# ── Append Worker (adds Sheet 3 to existing FuncAtlas_Report.xlsx) ───────────
class ComplexityAppendWorker(QObject):
    """
    Scans source files, computes complexity, then appends
    'Complexity_Compatibility' as Sheet 3 to an existing FuncAtlas_Report.xlsx.
    """
    progress = Signal(int, str)
    log      = Signal(str)
    finished = Signal(str)   # path to updated Excel
    error    = Signal(str)

    def __init__(self, report_path: str, source_folder: str,
                 weights: list = None, bands: list = None,
                 handled_scenarios=None):
        super().__init__()
        self.report_path   = report_path
        self.source_folder = normalize_path(source_folder)
        self.weights       = {n: w for n, w in (weights or [])} or \
                             {n: 1 for n, _ in CONSTRUCTS}
        self.bands         = bands or [
            ("Low",       1,   5),
            ("Medium",    6,  12),
            ("High",     13,  25),
            ("Very High", 26,  40),
            ("Complex",  41, 999),
        ]
        self.handled_scenarios = set(handled_scenarios) if handled_scenarios else set()

    def run(self):
        try:
            self._run()
        except Exception as exc:
            import traceback
            self.error.emit(f"{exc}\n{traceback.format_exc()}")

    def _run(self):
        from openpyxl import load_workbook

        if not os.path.isfile(self.report_path):
            self.error.emit(f"Report file not found:\n{self.report_path}")
            return

        # ── Read Sheet 1 to collect ALL New / Reuse (Modified) rows ────────────
        # We drive Sheet 3 directly from Sheet 1 rows so the row count and
        # file/path values are identical.  Key: (fn_name_lower, file_path_lower)
        # so same-named functions in different source files stay distinct.
        wb_check = load_workbook(self.report_path, read_only=True, data_only=True)
        # sheet1_rows: list of (fn_display, file_name, file_path) for New/RM rows
        sheet1_rows = []
        # include_set: set of (fn_lower, filepath_lower) — for source lookup
        include_set = set()
        if "Function_Match_Report" in wb_check.sheetnames:
            ws_check = wb_check["Function_Match_Report"]
            for row in ws_check.iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                fn_name  = str(row[1] or "").strip()   # col B = Function Name
                fname    = str(row[0] or "").strip()   # col A = File Name
                fpath    = str(row[2] or "").strip()   # col C = Target File Path
                status   = None
                for cell_val in reversed(row):
                    if cell_val in ("Reuse", "New", "Reuse (Modified)"):
                        status = cell_val
                        break
                if status in ("New", "Reuse (Modified)") and fn_name:
                    sheet1_rows.append((fn_name, fname, fpath))
                    include_set.add((normalize_name(fn_name), fpath.lower()))
        wb_check.close()

        if not sheet1_rows:
            self.error.emit(
                "No 'New' or 'Reuse (Modified)' functions found in the report.\n"
                "All functions may be 100% Reuse — nothing to analyse."
            )
            return

        self.log.emit(f"📋 {len(sheet1_rows)} functions to analyse "
                      f"(New + Reuse Modified only) …")

        # ── Scan source files and build a lookup: (fn_lower, filepath_lower) -> body
        # We need to match each Sheet 1 row to the correct source file body.
        file_entries = list(iter_source_files(self.source_folder))
        if not file_entries:
            self.error.emit("No source files found in the target source folder.")
            return

        self.log.emit(f"🔍 Scanning {len(file_entries)} source files …")

        # body_lookup: (fn_lower, full_path_lower) -> body_text
        # Also keep a name-only fallback: fn_lower -> [(full_path, body)]
        body_lookup   = {}   # (fn_lower, full_path_lower) -> body
        body_fallback = {}   # fn_lower -> [(full_path, body)]

        for idx, (full_path, file_name) in enumerate(file_entries):
            pct = int((idx / len(file_entries)) * 75)
            self.progress.emit(pct, f"Scanning {file_name} …")
            functions = detect_functions_in_file(full_path)
            if not functions:
                continue
            for fn_name in functions:
                fn_key = normalize_name(fn_name)
                body   = extract_function_body(full_path, fn_name)
                body_lookup[(fn_key, full_path.lower())] = body
                body_fallback.setdefault(fn_key, []).append((full_path, body))

        # ── Build records list in Sheet 1 order ─────────────────────────────
        # For each Sheet 1 row, find the matching body via composite key first,
        # then fall back to name-only (picks the first matching source file).
        records = []
        for fn_display, s1_fname, s1_fpath in sheet1_rows:
            fn_key = normalize_name(fn_display)
            body   = None

            # Try composite match: function name + any source file whose path
            # ends with the same basename as the Sheet 1 file path
            s1_basename = os.path.basename(s1_fpath).lower()
            for (bk_fn, bk_path_lower), bk_body in body_lookup.items():
                if bk_fn == fn_key and os.path.basename(bk_path_lower) == s1_basename:
                    body = bk_body
                    break

            # Fallback: name-only (first occurrence in source)
            if body is None:
                candidates = body_fallback.get(fn_key, [])
                if candidates:
                    body = candidates[0][1]

            if body is None:
                body = ""

            counts = count_constructs(body)
            score  = sum(counts.get(cn, 0) * self.weights.get(cn, 1)
                         for cn, _ in CONSTRUCTS)
            level  = complexity_level(score, self.bands)
            records.append({
                "function":  fn_display,
                "file_name": s1_fname,
                "file_path": s1_fpath,
                "counts":    counts,
                "score":     score,
                "level":     level,
            })

        self.log.emit(f"✅ {len(records)} functions processed — appending Sheet 3 …")
        self.progress.emit(85, "Appending Complexity_Compatibility sheet …")

        # ── Load existing workbook and append Sheet 3 ────────────────────────
        wb = load_workbook(self.report_path)

        # Remove old Sheet 3 if it already exists
        if "Complexity_Compatibility" in wb.sheetnames:
            del wb["Complexity_Compatibility"]

        ws3 = wb.create_sheet("Complexity_Compatibility")

        thin  = Side(style="thin",   color="C5D8EC")
        thick = Side(style="medium", color="1A3A5C")
        def bdr(left=None, right=None, top=None, bottom=None):
            return Border(left=left or thin, right=right or thin,
                          top=top or thin,   bottom=bottom or thin)

        hdr_fill  = PatternFill("solid", fgColor="1F4E78")
        hdr_font  = Font(color="FFFFFF", bold=True, name="Arial", size=10)
        hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        pct_align = Alignment(horizontal="center", vertical="center")
        left_align= Alignment(horizontal="left",   vertical="center")
        sno_fill  = PatternFill("solid", fgColor="EBF3FB")
        data_font = lambda: Font(name="Arial", size=10)

        green_fill  = PatternFill("solid", fgColor="C6EFCE")
        green_font  = lambda: Font(color="276221", name="Arial", size=10, bold=True)
        yellow_fill = PatternFill("solid", fgColor="FFEB9C")
        yellow_font = lambda: Font(color="9C5700", name="Arial", size=10, bold=True)
        red_fill    = PatternFill("solid", fgColor="FFC7CE")
        red_font    = lambda: Font(color="9C0006", name="Arial", size=10, bold=True)

        level_colors = {
            "Low":       "D6E4BC",
            "Medium":    "FFE699",
            "High":      "F4B183",
            "Very High": "FF7070",
            "Complex":   "CC0000",
        }

        construct_names = [c[0] for c in CONSTRUCTS]
        headers = (
            ["File Name", "Function Name", "Target File Path"]
            + construct_names
            + ["Complexity Score", "Complexity Level"]
        )
        last_col = len(headers)

        for c_idx, h in enumerate(headers, 1):
            cell = ws3.cell(1, c_idx, h)
            cell.fill      = hdr_fill
            cell.font      = hdr_font
            cell.alignment = hdr_align
            cell.border    = bdr(
                left=thick  if c_idx == 1       else thin,
                right=thick if c_idx == last_col else thin,
                top=thick, bottom=thick
            )
        ws3.row_dimensions[1].height = 44

        for r, rec in enumerate(records, 2):
            col = 1

            # File Name (first column)
            c = ws3.cell(r, col, rec["file_name"]); col += 1
            c.fill = sno_fill; c.font = Font(name="Arial", size=10, color="2C5F8A")
            c.alignment = left_align; c.border = bdr(left=thick)

            # Function Name
            c = ws3.cell(r, col, rec["function"]); col += 1
            c.font = Font(bold=True, name="Arial", size=10)
            c.alignment = left_align; c.border = bdr()

            # Target File Path
            c = ws3.cell(r, col, rec["file_path"]); col += 1
            c.font = data_font(); c.alignment = left_align; c.border = bdr()

            # Construct counts
            for cn in construct_names:
                cnt = rec["counts"].get(cn, 0)
                c = ws3.cell(r, col, cnt); col += 1
                c.alignment = pct_align; c.border = bdr(); c.font = data_font()
                if cnt > 0:
                    c.fill = PatternFill("solid", fgColor="EBF3FB")

            # Complexity Score
            c = ws3.cell(r, col, rec["score"]); col += 1
            c.alignment = pct_align; c.border = bdr()
            c.font = Font(bold=True, name="Arial", size=10)

            # Complexity Level
            c = ws3.cell(r, col, rec["level"]); col += 1
            c.alignment = pct_align
            c.border = bdr(right=thick)
            c.fill = PatternFill("solid", fgColor=level_colors.get(rec["level"], "FFFFFF"))
            c.font = Font(bold=True,
                          color="FFFFFF" if rec["level"] == "Complex" else "1A1A1A",
                          name="Arial", size=10)

        # Column widths
        ws3.column_dimensions["A"].width = 28  # File Name
        ws3.column_dimensions["B"].width = 36  # Function Name
        ws3.column_dimensions["C"].width = 50  # Target File Path
        for ci in range(4, 4 + len(construct_names)):
            ws3.column_dimensions[get_column_letter(ci)].width = 20
        score_col = 4 + len(construct_names)
        ws3.column_dimensions[get_column_letter(score_col)].width = 18
        ws3.column_dimensions[get_column_letter(score_col + 1)].width = 18
        ws3.freeze_panes = "A2"

        # ── Sheet 4: Compatibility Score ───────────────────────────────────────────────
        if "Compatibility_Score" in wb.sheetnames:
            del wb["Compatibility_Score"]

        ws4 = wb.create_sheet("Compatibility_Score")

        compat_hdrs = [
            "File Name", "Function Name", "File Path",
            "Available Scenarios", "Handled Scenarios",
            "Unhandled Scenarios", "Compatibility %"
        ]
        for c_idx, h in enumerate(compat_hdrs, 1):
            cell = ws4.cell(1, c_idx, h)
            cell.fill = hdr_fill; cell.font = hdr_font
            cell.alignment = hdr_align
            cell.border = bdr(left=thick if c_idx == 1 else thin,
                               right=thick if c_idx == len(compat_hdrs) else thin,
                               top=thick, bottom=thick)
        ws4.row_dimensions[1].height = 32

        handled = self.handled_scenarios
        for r2, rec in enumerate(records, 2):
            # available = constructs detected in this function (count > 0)
            available = [cn for cn in construct_names if rec["counts"].get(cn, 0) > 0]
            # handled_in_fn = intersection of available and user-marked handled
            handled_in_fn = [cn for cn in available if cn in handled]
            unhandled_in_fn = [cn for cn in available if cn not in handled]
            # Score = total appearances of handled scenarios /
            #         total appearances of available scenarios * 100
            total_avail_appearances   = sum(rec["counts"].get(cn, 0) for cn in available)
            total_handled_appearances = sum(rec["counts"].get(cn, 0) for cn in handled_in_fn)
            compat_pct = round(
                (total_handled_appearances / total_avail_appearances * 100), 2
            ) if total_avail_appearances > 0 else 0.0

            row_data = [
                rec["file_name"],
                rec["function"],
                rec["file_path"],
                ", ".join(available) if available else "None",
                ", ".join(handled_in_fn) if handled_in_fn else "None",
                ", ".join(unhandled_in_fn) if unhandled_in_fn else "None",
                compat_pct,
            ]
            for c_idx2, val in enumerate(row_data, 1):
                cell2 = ws4.cell(r2, c_idx2, val)
                cell2.border = bdr(left=thick if c_idx2 == 1 else thin,
                                   right=thick if c_idx2 == len(compat_hdrs) else thin)
                cell2.alignment = left_align if c_idx2 <= 3 else pct_align
                cell2.font = data_font()
                # Colour compatibility percentage cell
                if c_idx2 == len(compat_hdrs):
                    if compat_pct >= 75:
                        cell2.fill = green_fill; cell2.font = green_font()
                    elif compat_pct >= 40:
                        cell2.fill = yellow_fill; cell2.font = yellow_font()
                    else:
                        cell2.fill = red_fill; cell2.font = red_font()
                    cell2.value = f"{compat_pct:.1f}%"

        ws4.column_dimensions["A"].width = 26
        ws4.column_dimensions["B"].width = 36
        ws4.column_dimensions["C"].width = 50
        ws4.column_dimensions["D"].width = 50
        ws4.column_dimensions["E"].width = 50
        ws4.column_dimensions["F"].width = 50
        ws4.column_dimensions["G"].width = 20
        ws4.freeze_panes = "A2"

        wb.save(self.report_path)
        wb.close()

        self.progress.emit(100, "Done")
        self.log.emit(f"📊 Sheets 3 & 4 appended: {self.report_path}")
        self.finished.emit(self.report_path)