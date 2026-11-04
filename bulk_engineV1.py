"""
bulk_engine.py
Core logic for bulk .xlsx generation (template scan, tracker, parallel writer).
UI-independent: used by the desktop/web front-end in BulkFileGenerator.py.
"""
import os
import re
import tempfile
import time
import zipfile
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
    try:
        font, fill, al = cell.font, cell.fill, cell.alignment
    except AttributeError:          # empty cells in read-only mode
        return None
    f_col = get_safe_color(font.color) if font else None
    fill_col = (
        get_safe_color(fill.start_color)
        if fill is not None and fill.fill_type == "solid"
        else None
    )
    h_al = (al.horizontal or None) if al is not None else None
    v_al = (al.vertical or None) if al is not None else None
    if v_al == "center":  # openpyxl says 'center', xlsxwriter wants 'vcenter'
        v_al = "vcenter"
    return (font.name if font else None, font.size if font else None,
            f_col, fill_col, h_al, v_al)


def scan_template(template_path, learning_rows=50):
    """
    Reads values and learns styles from the first rows of the active sheet.
    Uses openpyxl's fast read-only mode.
    Returns (template_data, styles) where styles = {
        "cells": {(row0, col0): style} for the learned rows (keeps header looks),
        "cols":  {col0: style} from the last learned row (used for rows beyond),
    }
    """
    tpl_wb = openpyxl.load_workbook(template_path, data_only=True, read_only=True)
    tpl_ws = tpl_wb.active
    cells, cols = {}, {}
    template_data = []
    for r0, row in enumerate(tpl_ws.iter_rows()):
        if r0 < learning_rows:
            for cell in row:
                c0 = getattr(cell, "column", None)
                c0 = (c0 - 1) if c0 else None
                if c0 is None:
                    continue
                st = _cell_style(cell)
                cells[(r0, c0)] = st
                cols[c0] = st
        template_data.append([cell.value for cell in row])
    tpl_wb.close()

    # Pad rows to the same width and drop trailing empty rows.
    width = max((len(r) for r in template_data), default=0)
    for row in template_data:
        if len(row) < width:
            row.extend([None] * (width - len(row)))
    while template_data and all(v is None for v in template_data[-1]):
        template_data.pop()
    # Rows that were learned but have no cell for a column: reuse that column's style.
    learned = min(learning_rows, len(template_data))
    for r0 in range(learned):
        for c0 in range(width):
            cells.setdefault((r0, c0), cols.get(c0))
    return template_data, {"cells": cells, "cols": cols}


# --- FILE BUILDING ---
# Speed-up: the template part is identical in every file; only the TxnID cells
# differ. So the workbook is built ONCE with numbered placeholders in the TxnID
# cells, and every file is then made by swapping the real TxnIDs into that
# prebuilt sheet and zipping it. No cell-by-cell rebuild per file.
_TOKEN_RE = re.compile("(\\d+)")
_SHEET_PART = "xl/worksheets/sheet1.xml"


def _write_base_workbook(full_path, template_data, row_styles, col_defaults, updates,
                         max_update_row):
    """Writes the template + fixed values; TxnID cells get placeholder tokens.
    Returns slots: [(txn_prefix, offset_k)], offset_k = order the ID is used in."""
    workbook = xlsxwriter.Workbook(full_path, {"constant_memory": True,
                                               "strings_to_urls": True})
    worksheet = workbook.add_worksheet()
    format_cache = {}

    def make_fmt(st):
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

    learned_rows = [[make_fmt(st) for st in row] for row in row_styles]
    default_fmts = [make_fmt(st) for st in col_defaults]

    # TxnIDs are used column by column, top to bottom (same order as the original
    # script); k numbers them in that order so each file can fill them in.
    overrides = {}
    slots = []
    k = 0
    for up in updates:
        col_idx = up["col_idx"]
        if up["logic"]["mode"] == MODE_TXN:
            prefix = up["logic"]["txn_prefix"].rstrip("_") + "_"
            for r0 in range(up["start"] - 1, up["end"]):
                overrides.setdefault(r0, []).append((col_idx, f"{k}"))
                slots.append(prefix)
                k += 1
        else:
            value = up["logic"]["value"]
            for r0 in range(up["start"] - 1, up["end"]):
                overrides.setdefault(r0, []).append((col_idx, value))

    total_rows = max(len(template_data), max_update_row)
    n_learned = len(learned_rows)
    write, write_number, write_blank = worksheet.write, worksheet.write_number, worksheet.write_blank

    for r0 in range(total_rows):
        row = template_data[r0] if r0 < len(template_data) else []
        ov = overrides.get(r0)
        if ov:
            row = list(row)
            for c0, v in ov:
                if c0 >= len(row):
                    row.extend([None] * (c0 + 1 - len(row)))
                row[c0] = v
        fmts = learned_rows[r0] if r0 < n_learned else default_fmts
        n_fmts = len(fmts)
        for c0, val in enumerate(row):
            fmt = fmts[c0] if c0 < n_fmts else None
            if val is None:
                if fmt is not None:
                    write_blank(r0, c0, None, fmt)
            elif type(val) is int or type(val) is float:
                write_number(r0, c0, val, fmt)
            else:
                write(r0, c0, val, fmt)

    workbook.close()
    return slots


def _prepare_base(template_data, styles, updates, work_dir):
    """Builds the base workbook once and splits its sheet around the TxnID slots."""
    width = max((len(r) for r in template_data), default=0)
    n_learned = max((r for r, _c in styles["cells"]), default=-1) + 1
    row_styles = [[styles["cells"].get((r, c)) for c in range(width)] for r in range(n_learned)]
    col_defaults = [styles["cols"].get(c) for c in range(width)]
    max_update_row = max((u["end"] for u in updates), default=0)

    base_path = os.path.join(work_dir, "_base.xlsx")
    slots = _write_base_workbook(base_path, template_data, row_styles, col_defaults,
                                 updates, max_update_row)

    entries = []
    sheet_xml = None
    with zipfile.ZipFile(base_path) as zf:
        for info in zf.infolist():
            data = zf.read(info.filename)
            if info.filename == _SHEET_PART:
                sheet_xml = data.decode("utf-8")
                entries.append((info.filename, None))
            else:
                entries.append((info.filename, data))
    os.remove(base_path)

    pieces = _TOKEN_RE.split(sheet_xml)       # text, k, text, k, ..., text
    chunks = pieces[0::2]
    order = [int(x) for x in pieces[1::2]]    # slot numbers in document order
    if len(order) != len(slots):
        raise RuntimeError("Could not place every TxnID in the sheet.")
    # Precompute (prefix, offset) in the order they appear in the sheet.
    doc_slots = [(slots[k], k) for k in order]
    return {"entries": entries, "chunks": chunks, "doc_slots": doc_slots}


# Data shared by every file in a batch is sent to each worker process once
# (via the Pool initializer) instead of being re-sent with every single file.
_SHARED = {}


def _init_worker(shared):
    global _SHARED
    _SHARED = shared


def generate_single_file(task):
    fn_num, txn_start_val = task
    s = _SHARED
    full_path = os.path.join(s["target_dir"], f"{s['naming_prefix']}{fn_num}{s['extension']}")

    chunks, doc_slots = s["chunks"], s["doc_slots"]
    if doc_slots:
        parts = [None] * (len(chunks) + len(doc_slots))
        parts[0::2] = chunks
        parts[1::2] = [f"{p}{txn_start_val - k}" for p, k in doc_slots]
        sheet = "".join(parts).encode("utf-8")
    else:
        sheet = chunks[0].encode("utf-8")

    with zipfile.ZipFile(full_path, "w", zipfile.ZIP_DEFLATED,
                         compresslevel=s["compresslevel"]) as zf:
        for name, data in s["entries"]:
            zf.writestr(name, sheet if data is None else data)

    return fn_num, txn_start_val - len(doc_slots), full_path


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
         "col_idx": column_index_from_string(str(u["col"]).strip().upper()) - 1,
         "start": int(u["start"]), "end": int(u["end"])}
        for u in updates
    ]
    os.makedirs(output_dir, exist_ok=True)

    # Step 1: scan template and prebuild the shared sheet
    stage("Scanning template…")
    t1 = time.time()
    template_data, styles = scan_template(template_path)
    template_data = [
        [None if v is None or (isinstance(v, (str, float)) and str(v).lower() == "nan") else v
         for v in row]
        for row in template_data
    ]
    stage("Preparing the base file…")
    with tempfile.TemporaryDirectory(prefix="bulkgen_") as work_dir:
        base = _prepare_base(template_data, styles, updates, work_dir)
    del template_data
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
        **base,
        "naming_prefix": naming_prefix, "extension": extension, "target_dir": output_dir,
        "compresslevel": 3,
    }
    tasks = [(fn_numbers[i], txn_base - i * per_file) for i in range(n_files)]

    results = []
    cores = 1
    used_parallel = False
    if parallel and n_files > 3:
        try:
            cores = max(1, min(cpu_count(), n_files))
            with Pool(processes=cores, initializer=_init_worker, initargs=(shared,)) as pool:
                for res in pool.imap_unordered(generate_single_file, tasks, chunksize=1):
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
