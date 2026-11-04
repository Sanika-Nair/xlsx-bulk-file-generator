#bulk_engine.py
"""
bulk_engine.py
Core logic for bulk .xlsx generation (template scan, tracker, parallel writer).
UI-independent: used by the desktop/web front-end in BulkFileGenerator.py.
"""
import os
import re
import time
from multiprocessing import Pool, cpu_count

import openpyxl
from openpyxl.utils import column_index_from_string
import xlsxwriter

# --- CONFIGURATION ---
MAX_PREFIX_LENGTH = 13
TXN_PREFIX_MIN, TXN_PREFIX_MAX = 9, 10
DEFAULT_COUNTER_SEED = 9999999

VALID_HALIGN = {"left", "center", "right"}
VALID_VALIGN = {"top", "vcenter", "bottom"}
NAME_PATTERN = re.compile(r"^[a-zA-Z0-9_]+$")
DATE_PATTERN = re.compile(
    r"^(\d{2,4}[/-]\d{2}[/-]\d{2,4}|\d{2}[/-]\d{2,4}[/-]\d{2,4})$"
)

MODE_TXN = "STRING_SEQ"
MODE_FIXED = "FIXED"


# --- TRACKER LOGIC ---
class Tracker:
    """Persists the filename counter and TxnID counter in text files."""

    def __init__(self, tracker_dir, seed=DEFAULT_COUNTER_SEED):
        self.tracker_dir = tracker_dir
        self.filename_file = os.path.join(tracker_dir, "filename_Numbercounter.txt")
        self.txnid_file = os.path.join(tracker_dir, "TxnID_columnNumberCounter.txt")
        os.makedirs(tracker_dir, exist_ok=True)
        for path in (self.filename_file, self.txnid_file):
            if not os.path.exists(path):
                self._write(path, seed)

    @staticmethod
    def _read(path):
        with open(path, "r") as f:
            return int(f.read().strip())

    @staticmethod
    def _write(path, value):
        with open(path, "w") as f:
            f.write(str(value))

    def peek(self):
        """Current counters without changing them: (next file number, next TxnID number)."""
        return self._read(self.filename_file), self._read(self.txnid_file)

    def get_batch_setup(self, n_files):
        current_fn_num = self._read(self.filename_file)
        fn_numbers = [current_fn_num - i for i in range(n_files)]
        self._write(self.filename_file, current_fn_num - n_files)
        txn_start_num = self._read(self.txnid_file)
        return fn_numbers, txn_start_num

    def update_txn_tracker(self, next_start_val):
        self._write(self.txnid_file, next_start_val)


# --- VALIDATION ---
def validate_name(value, max_len=MAX_PREFIX_LENGTH, min_len=1):
    """Returns an error string, or None if valid."""
    if not value or len(value) < min_len or len(value) > max_len:
        if min_len > 1:
            return f"Must be {min_len}-{max_len} characters."
        return f"Must be between 1 and {max_len} characters."
    if not NAME_PATTERN.match(value):
        return "No special characters allowed (except underscore '_')."
    return None


def validate_column_letter(value):
    if not value or not value.isalpha() or not value.isascii():
        return "Column must contain letters only (e.g. A, B, AA)."
    try:
        idx = column_index_from_string(value.upper())
    except ValueError:
        return "Not a valid Excel column."
    if idx > 16384:
        return "Column is beyond Excel's limit (XFD)."
    return None


def validate_date_string(value):
    if not DATE_PATTERN.match(value or ""):
        return "Invalid date format. Use dd/mm/yyyy, mm/dd/yyyy or yyyy-mm-dd."
    return None


def validate_updates(updates):
    """Returns a list of error strings across all column rules."""
    errors = []
    seen = set()
    for i, up in enumerate(updates, start=1):
        tag = f"Column rule {i}"
        col = str(up.get("col", "")).strip()
        err = validate_column_letter(col)
        if err:
            errors.append(f"{tag}: {err}")
        elif col.upper() in seen:
            errors.append(f"{tag}: column {col.upper()} is used more than once.")
        else:
            seen.add(col.upper())
        try:
            start, end = int(up["start"]), int(up["end"])
        except (TypeError, ValueError, KeyError):
            errors.append(f"{tag}: rows must be positive whole numbers.")
        else:
            if start < 1 or end < 1:
                errors.append(f"{tag}: rows must be positive whole numbers.")
            elif start > end:
                errors.append(f"{tag}: start row is after end row.")
        logic = up.get("logic") or {}
        if logic.get("mode") == MODE_TXN:
            err = validate_name(logic.get("txn_prefix", ""), TXN_PREFIX_MAX, TXN_PREFIX_MIN)
            if err:
                errors.append(f"{tag}: TxnID prefix: {err}")
        elif logic.get("mode") == MODE_FIXED:
            err = validate_date_string(logic.get("value", ""))
            if err:
                errors.append(f"{tag}: {err}")
        else:
            errors.append(f"{tag}: choose an update logic.")
    return errors


def txn_ids_per_file(updates):
    """How many TxnIDs one generated file consumes."""
    return sum(
        int(up["end"]) - int(up["start"]) + 1
        for up in updates if up["logic"]["mode"] == MODE_TXN
    )


# --- TEMPLATE SCAN ---
def get_safe_color(color_obj):
    try:
        if color_obj and color_obj.type == "rgb":
            rgb_val = str(color_obj.rgb)
            if len(rgb_val) == 8:
                rgb_val = rgb_val[2:]
            if len(rgb_val) == 6:
                return rgb_val
    except Exception:
        pass
    return None


def preview_template(template_path, rows=8):
    """Quick read for the UI: (first rows as strings, max_row, max_col)."""
    wb = openpyxl.load_workbook(template_path, data_only=True, read_only=True)
    ws = wb.active
    data = []
    for r in ws.iter_rows(min_row=1, max_row=rows, values_only=True):
        data.append(["" if v is None else str(v) for v in r])
    dims = (ws.max_row or 0, ws.max_column or 0)
    wb.close()
    return data, dims


def _cell_style(cell):
    f_col = get_safe_color(cell.font.color) if cell.font else None
    fill_col = (
        get_safe_color(cell.fill.start_color)
        if cell.fill and cell.fill.fill_type == "solid"
        else None
    )
    h_al = cell.alignment.horizontal or None
    v_al = cell.alignment.vertical or None
    if v_al == "center":  # openpyxl says 'center', xlsxwriter wants 'vcenter'
        v_al = "vcenter"
    return (cell.font.name, cell.font.size, f_col, fill_col, h_al, v_al)


def scan_template(template_path, learning_rows=50):
    """
    Reads values and learns styles from the first rows of the active sheet.
    Returns (template_data, styles) where styles = {
        "cells": {(row0, col0): style} for the learned rows (keeps header looks),
        "cols":  {col0: style} from the last learned row (used for rows beyond),
    }
    """
    tpl_wb = openpyxl.load_workbook(template_path, data_only=True)
    tpl_ws = tpl_wb.active
    max_r, max_c = tpl_ws.max_row, tpl_ws.max_column
    cells, cols = {}, {}

    for r in range(1, min(learning_rows, max_r) + 1):
        for c in range(1, max_c + 1):
            st = _cell_style(tpl_ws.cell(row=r, column=c))
            cells[(r - 1, c - 1)] = st
            cols[c - 1] = st

    template_data = [
        [tpl_ws.cell(row=r, column=c).value for c in range(1, max_c + 1)]
        for r in range(1, max_r + 1)
    ]
    tpl_wb.close()
    return template_data, {"cells": cells, "cols": cols}


# --- WORKER FUNCTION ---
# Data shared by every file in a batch is sent to each worker process once
# (via the Pool initializer) instead of being re-sent with every single file.
_SHARED = {}


def _init_worker(shared):
    global _SHARED
    _SHARED = shared


def generate_single_file(task):
    fn_num, txn_start_val = task
    s = _SHARED
    template_data, col_styles, updates = s["template_data"], s["col_styles"], s["updates"]

    new_filename = f"{s['naming_prefix']}{fn_num}{s['extension']}"
    full_path = os.path.join(s["target_dir"], new_filename)

    workbook = xlsxwriter.Workbook(full_path)
    worksheet = workbook.add_worksheet()

    styles = col_styles
    format_cache = {}

    def fmt_for(r0, c0):
        st = styles["cells"].get((r0, c0)) or styles["cols"].get(c0)
        if st is None:
            return None
        if st not in format_cache:
            name, size, font_color, fill, h_al, v_al = st
            fmt_dict = {}
            if name:
                fmt_dict["font_name"] = name
            if size:
                fmt_dict["font_size"] = size
            if font_color:
                fmt_dict["font_color"] = "#" + font_color
            if fill:
                fmt_dict["bg_color"] = "#" + fill
            if h_al in VALID_HALIGN:
                fmt_dict["align"] = h_al
            if v_al in VALID_VALIGN:
                fmt_dict["valign"] = v_al
            format_cache[st] = workbook.add_format(fmt_dict)
        return format_cache[st]

    for r_idx, row_values in enumerate(template_data):
        for c_idx, val in enumerate(row_values):
            write_val = "" if val is None or str(val).lower() == "nan" else val
            worksheet.write(r_idx, c_idx, write_val, fmt_for(r_idx, c_idx))

    current_row_txn_id = txn_start_val
    for up in updates:
        col_idx = column_index_from_string(up["col"]) - 1
        is_seq = up["logic"]["mode"] == MODE_TXN
        for r_idx in range(up["start"], up["end"] + 1):
            row_to_write = r_idx - 1
            fmt = fmt_for(row_to_write, col_idx)
            if is_seq:
                clean_prefix = up["logic"]["txn_prefix"].rstrip("_")
                worksheet.write(row_to_write, col_idx,
                                f"{clean_prefix}_{current_row_txn_id}", fmt)
                current_row_txn_id -= 1
            else:
                worksheet.write(row_to_write, col_idx, up["logic"]["value"], fmt)

    workbook.close()
    return fn_num, current_row_txn_id, full_path


# --- BATCH RUNNER ---
def run_batch(template_path, n_files, naming_prefix, updates, output_dir,
              tracker_dir, parallel=True, progress_cb=None, stage_cb=None):
    """
    Runs the full pipeline.
    progress_cb(done, total) is called as files finish; stage_cb(text) on each step.
    Returns a summary dict.
    """
    stage = stage_cb or (lambda _t: None)
    overall_start = time.time()
    if not naming_prefix.endswith("_"):
        naming_prefix += "_"
    extension = os.path.splitext(template_path)[1]
    updates = [
        {**u, "col": str(u["col"]).strip().upper(),
         "start": int(u["start"]), "end": int(u["end"])}
        for u in updates
    ]
    os.makedirs(output_dir, exist_ok=True)

    # Step 1: scan
    stage("Scanning template…")
    t1 = time.time()
    template_data, col_styles = scan_template(template_path)
    step1 = time.time() - t1

    # Step 2: tracker
    stage("Reserving IDs from tracker…")
    t2 = time.time()
    tracker = Tracker(tracker_dir)
    fn_numbers, txn_base = tracker.get_batch_setup(n_files)
    per_file = txn_ids_per_file(updates)
    step2 = time.time() - t2

    # Step 3: generation, each file gets its own TxnID block so IDs never repeat
    stage("Generating files…")
    t3 = time.time()
    shared = {
        "template_data": template_data, "col_styles": col_styles, "updates": updates,
        "naming_prefix": naming_prefix, "extension": extension, "target_dir": output_dir,
    }
    tasks = [(fn_numbers[i], txn_base - i * per_file) for i in range(n_files)]

    results = []
    cores = 1
    used_parallel = False
    if parallel and n_files > 1:
        try:
            cores = max(1, min(cpu_count(), n_files))
            chunk = max(1, n_files // (cores * 20))
            with Pool(processes=cores, initializer=_init_worker, initargs=(shared,)) as pool:
                for res in pool.imap_unordered(generate_single_file, tasks, chunksize=chunk):
                    results.append(res)
                    if progress_cb:
                        progress_cb(len(results), n_files)
            used_parallel = True
        except Exception:
            results = []  # fall back to sequential below
    if not used_parallel:
        cores = 1
        _init_worker(shared)
        for task in tasks:
            results.append(generate_single_file(task))
            if progress_cb:
                progress_cb(len(results), n_files)

    next_txn = txn_base - n_files * per_file
    tracker.update_txn_tracker(next_txn)
    step3 = time.time() - t3

    results.sort(key=lambda r: -r[0])
    return {
        "files": [r[2] for r in results],
        "output_dir": output_dir,
        "step1": step1,
        "step2": step2,
        "step3": step3,
        "total": time.time() - overall_start,
        "cores": cores,
        "txn_range": [txn_base, next_txn + 1] if per_file else None,
        "fn_range": [fn_numbers[0], fn_numbers[-1]],
        "first_file": os.path.basename(results[0][2]) if results else None,
        "last_file": os.path.basename(results[-1][2]) if results else None,
    }
