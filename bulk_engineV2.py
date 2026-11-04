"""
bulk_engine.py
Core logic for bulk .xlsx generation. UI-independent: used by BulkFileGenerator.py.

How files are made
------------------
Every generated file is an exact copy of the template: all sheets, formatting,
column widths, number formats, formulas, merged cells, frozen panes and so on.
Only the cells you choose (column + row range on the active sheet) are changed.

Where Excel keeps a cell's text in its shared text list (sharedStrings.xml), the
new value is swapped into that list entry and the cell itself is left untouched,
so the sheet stays byte-for-byte the same and files stay the size of the template.
Other target cells are rewritten in the sheet with the new text, keeping their style.

Speed: the template is prepared once per batch, with numbered placeholders for
the TxnIDs. Each file is then made by putting that file's TxnIDs into the
placeholders and writing the .xlsx package. Unchanged parts of the package are
compressed once and reused for every file.

ID rules (enforced by Tracker.reserve):
  - file numbers and TxnIDs are taken from the project's tracker files, counting down
  - every number is used once only; the counters are moved past the whole batch
    BEFORE any file is written, so a failed or cancelled batch never reuses numbers
  - a batch that would need a number below 1 is refused (so nothing negative ever appears)
"""
import os
import re
import struct
import time
import zipfile
import zlib
from multiprocessing import Pool, cpu_count
from xml.sax.saxutils import escape as xml_escape

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

# --- CONFIGURATION ---
MAX_PREFIX_LENGTH = 13
TXN_PREFIX_MIN, TXN_PREFIX_MAX = 9, 10
DEFAULT_COUNTER_SEED = 9999999
LOWEST_ALLOWED_NUMBER = 1          # IDs never go below 1, so no negative numbers anywhere (not even in the tracker files)

NAME_PATTERN = re.compile(r"^[a-zA-Z0-9_]+$")
DATE_PATTERN = re.compile(
    r"^(\d{2,4}[/-]\d{2}[/-]\d{2,4}|\d{2}[/-]\d{2,4}[/-]\d{2,4})$"
)

MODE_TXN = "STRING_SEQ"
MODE_FIXED = "FIXED"


# =====================================================================
# TRACKER
# =====================================================================
class TrackerError(Exception):
    pass


class _FileLock:
    """Cross-process lock so two copies of the app can't take the same numbers."""

    def __init__(self, path, timeout=30):
        self.path, self.timeout, self.fh = path, timeout, None

    def __enter__(self):
        self.fh = open(self.path, "a+")
        deadline = time.time() + self.timeout
        while True:
            try:
                if os.name == "nt":
                    import msvcrt
                    self.fh.seek(0)
                    msvcrt.locking(self.fh.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(self.fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                return self
            except OSError:
                if time.time() > deadline:
                    self.fh.close()
                    raise TrackerError("The tracker is busy (another batch is reserving numbers). Try again.")
                time.sleep(0.1)

    def __exit__(self, *exc):
        try:
            if os.name == "nt":
                import msvcrt
                self.fh.seek(0)
                msvcrt.locking(self.fh.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.fh.fileno(), fcntl.LOCK_UN)
        finally:
            self.fh.close()


class Tracker:
    """Persists the filename counter and TxnID counter in text files."""

    def __init__(self, tracker_dir, seed=DEFAULT_COUNTER_SEED):
        self.tracker_dir = tracker_dir
        self.filename_file = os.path.join(tracker_dir, "filename_Numbercounter.txt")
        self.txnid_file = os.path.join(tracker_dir, "TxnID_columnNumberCounter.txt")
        self.lock_file = os.path.join(tracker_dir, ".tracker.lock")
        os.makedirs(tracker_dir, exist_ok=True)
        for path in (self.filename_file, self.txnid_file):
            if not os.path.exists(path):
                self._write(path, seed)

    @staticmethod
    def _read(path):
        try:
            with open(path, "r") as f:
                text = f.read().strip()
            return int(text)
        except (ValueError, OSError):
            raise TrackerError(f"Tracker file is unreadable or not a whole number: {path}")

    @staticmethod
    def _write(path, value):
        # write to a temp file then swap it in, so a crash can't leave an empty counter
        tmp = path + ".tmp"
        with open(tmp, "w") as f:
            f.write(str(value))
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)

    def peek(self):
        """Current counters without changing them: (next file number, next TxnID number)."""
        return self._read(self.filename_file), self._read(self.txnid_file)

    def reserve(self, n_files, txn_per_file):
        """
        Reserves numbers for a whole batch and moves both counters past it
        before anything is generated. Returns (file_numbers, txn_start).
        Raises TrackerError if the batch would go below LOWEST_ALLOWED_NUMBER.
        """
        with _FileLock(self.lock_file):
            fn_next, txn_next = self.peek()
            errors = check_capacity(fn_next, txn_next, n_files, txn_per_file)
            if errors:
                raise TrackerError(" ".join(errors))
            self._write(self.filename_file, fn_next - n_files)
            self._write(self.txnid_file, txn_next - n_files * txn_per_file)
        return [fn_next - i for i in range(n_files)], txn_next


def check_capacity(fn_next, txn_next, n_files, txn_per_file):
    """The no-negative-numbers rule. Used by Tracker.reserve and by the UI before Generate."""
    errors = []
    if fn_next - n_files + 1 < LOWEST_ALLOWED_NUMBER:
        errors.append(f"Only {max(0, fn_next - LOWEST_ALLOWED_NUMBER + 1):,} file numbers are left "
                      f"(next is {fn_next}); numbers can't go below {LOWEST_ALLOWED_NUMBER}.")
    need = n_files * txn_per_file
    if need and txn_next - need + 1 < LOWEST_ALLOWED_NUMBER:
        errors.append(f"This batch needs {need:,} TxnIDs but only "
                      f"{max(0, txn_next - LOWEST_ALLOWED_NUMBER + 1):,} are left (next is {txn_next}).")
    return errors


# =====================================================================
# VALIDATION
# =====================================================================
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
            elif end > 1048576:
                errors.append(f"{tag}: end row is beyond Excel's last row (1048576).")
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


# =====================================================================
# TEMPLATE PREVIEW (for the UI)
# =====================================================================
def preview_template(template_path, rows=8):
    """Quick read for the UI: (first rows as strings, (max_row, max_col), active sheet name)."""
    wb = openpyxl.load_workbook(template_path, data_only=True, read_only=True)
    ws = wb.active
    data = []
    for r in ws.iter_rows(min_row=1, max_row=rows, values_only=True):
        data.append(["" if v is None else str(v) for v in r])
    dims = (ws.max_row or 0, ws.max_column or 0)
    name = ws.title
    wb.close()
    return data, dims, name


# =====================================================================
# TEMPLATE PACKAGE: find the active sheet and patch only the target cells
# =====================================================================
_TOKEN_OPEN, _TOKEN_CLOSE = "\ue000", "\ue001"
_TOKEN_RE = re.compile(_TOKEN_OPEN + r"(\d+)" + _TOKEN_CLOSE)

_ROW_RE = re.compile(r"<row\b(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</row>)", re.S)
_CELL_RE = re.compile(r"<c\b(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</c>)", re.S)
_R_ATTR = re.compile(r'\sr="([A-Z]+)?(\d+)"')
_ROW_R_ATTR = re.compile(r'\sr="(\d+)"')
_S_ATTR = re.compile(r'\ss="(\d+)"')
_SPANS_ATTR = re.compile(r'\sspans="[^"]*"')
_F_TAG = re.compile(r"<f\b([^>]*?)/?>")


class TemplateError(Exception):
    pass


def _col_letters_to_index0(letters):
    return column_index_from_string(letters) - 1


def _resolve_target(base_dir, target):
    if target.startswith("/"):
        return target.lstrip("/")
    parts = (base_dir + "/" + target).split("/") if base_dir else target.split("/")
    out = []
    for p in parts:
        if p in ("", "."):
            continue
        if p == "..":
            if out:
                out.pop()
        else:
            out.append(p)
    return "/".join(out)


def _rels_map(xml):
    rels = {}
    for tag in re.findall(r"<Relationship\b[^>]*?/?>", xml):
        rid = re.search(r'\sId="([^"]+)"', tag)
        tgt = re.search(r'\sTarget="([^"]+)"', tag)
        if rid and tgt:
            rels[rid.group(1)] = (tgt.group(1), tag)
    return rels


def load_template_package(template_path):
    """Reads the .xlsx package. Returns (entries, sheet_part, sheet_name)
    where entries is an ordered list of (name, bytes)."""
    try:
        with zipfile.ZipFile(template_path) as zf:
            entries = [(i.filename, zf.read(i.filename)) for i in zf.infolist() if not i.is_dir()]
    except zipfile.BadZipFile:
        raise TemplateError("The template isn't a valid .xlsx file (it may be an older .xls or password-protected).")
    names = {n: d for n, d in entries}

    root_rels = _rels_map(names.get("_rels/.rels", b"").decode("utf-8"))
    wb_part = "xl/workbook.xml"
    for _rid, (tgt, tag) in root_rels.items():
        typ = re.search(r'\sType="([^"]+)"', tag)
        if typ and typ.group(1).endswith("/officeDocument"):
            wb_part = _resolve_target("", tgt)
    if wb_part not in names:
        raise TemplateError("Couldn't find the workbook inside the template.")
    wb_xml = names[wb_part].decode("utf-8")
    wb_dir = wb_part.rsplit("/", 1)[0] if "/" in wb_part else ""
    wb_rels_part = (wb_dir + "/" if wb_dir else "") + "_rels/" + wb_part.rsplit("/", 1)[-1] + ".rels"
    wb_rels = _rels_map(names.get(wb_rels_part, b"").decode("utf-8"))

    sheets = re.findall(r"<sheet\b[^>]*?/?>", wb_xml)
    if not sheets:
        raise TemplateError("The template has no sheets.")
    m = re.search(r'<workbookView\b[^>]*\sactiveTab="(\d+)"', wb_xml)
    active = int(m.group(1)) if m else 0
    if active >= len(sheets):
        active = 0
    tag = sheets[active]
    rid = re.search(r'\s(?:\w+:)?id="([^"]+)"', tag)
    name = re.search(r'\sname="([^"]*)"', tag)
    if not rid or rid.group(1) not in wb_rels:
        raise TemplateError("Couldn't locate the active sheet inside the template.")
    sheet_part = _resolve_target(wb_dir, wb_rels[rid.group(1)][0])
    if sheet_part not in names:
        raise TemplateError("The active sheet's data is missing from the template.")
    if "chartsheet" in sheet_part:
        raise TemplateError("The active sheet is a chart, not a worksheet. Save the template with a data sheet selected.")
    return entries, sheet_part, (name.group(1) if name else "Sheet")


def _drop_calc_chain(entries):
    """Removes calcChain.xml and its references; Excel rebuilds it silently.
    Needed because replaced cells may have held formulas."""
    calc = {n for n, _ in entries if n.endswith("calcChain.xml")}
    if not calc:
        return entries
    out = []
    for n, d in entries:
        if n in calc:
            continue
        if n == "[Content_Types].xml" or n.endswith(".rels"):
            text = d.decode("utf-8")
            text = re.sub(r"<Override\b[^>]*calcChain\.xml[^>]*/>", "", text)
            text = re.sub(r"<Relationship\b[^>]*calcChain\.xml[^>]*/>", "", text)
            d = text.encode("utf-8")
        out.append((n, d))
    return out


def _cols_styles(xml):
    """Column default styles from <cols>: {col0: s}."""
    styles = {}
    for tag in re.findall(r"<col\b[^>]*?/?>", xml):
        mn = re.search(r'\smin="(\d+)"', tag)
        mx = re.search(r'\smax="(\d+)"', tag)
        st = re.search(r'\sstyle="(\d+)"', tag)
        if mn and mx and st and st.group(1) != "0":
            for c in range(int(mn.group(1)), min(int(mx.group(1)), 16384) + 1):
                styles[c - 1] = st.group(1)
    return styles


def patch_sheet_xml(xml, targets, replaced_sst=None):
    """
    targets: {row_number(1-based): {col0: text}}. Replaces / creates exactly those
    cells as text, keeping each cell's existing style. Everything else in the
    sheet is left byte-for-byte as it was.
    replaced_sst: optional dict, filled with {shared-text index: times replaced}
    for replaced cells whose old value lived in the workbook's shared text list.
    """
    if _TOKEN_OPEN in xml:
        raise TemplateError("The template contains an unsupported special character.")
    m = re.search(r"<sheetData\b[^>]*?(/>|>)", xml)
    if not m:
        raise TemplateError("The active sheet has no cell data section.")
    if m.group(1) == "/>":
        head, body, tail = xml[:m.start()] + "<sheetData>", "", "</sheetData>" + xml[m.end():]
    else:
        end = xml.index("</sheetData>", m.end())
        head, body, tail = xml[:m.end()], xml[m.end():end], xml[end:]

    col_styles = _cols_styles(head)
    replaced_shared_masters = set()

    def new_cell(col0, rnum, style, text):
        s = f' s="{style}"' if style else ""
        return f'<c r="{get_column_letter(col0 + 1)}{rnum}"{s} t="inlineStr"><is><t>{text}</t></is></c>'

    def row_default_style(attrs):
        if re.search(r'\scustomFormat="(1|true)"', attrs):
            s = _S_ATTR.search(attrs)
            if s:
                return s.group(1)
        return None

    def build_row(rnum, attrs, row_body, row_targets):
        cells = {}
        rest = []
        p = 0
        last_c = -1
        for cm in _CELL_RE.finditer(row_body or ""):
            rest.append(row_body[p:cm.start()])
            p = cm.end()
            ref = _R_ATTR.search(cm.group("attrs"))
            if ref and ref.group(1):
                col0 = _col_letters_to_index0(ref.group(1))
                cell_xml = cm.group(0)
            else:  # cell without a reference: its position follows the previous cell
                col0 = last_c + 1
                cell_xml = f'<c r="{get_column_letter(col0 + 1)}{rnum}"' + cm.group(0)[2:]
            last_c = col0
            cells[col0] = cell_xml
        rest.append((row_body or "")[p:])

        rstyle = row_default_style(attrs)
        for col0, text in row_targets.items():
            old = cells.get(col0)
            if old is not None:
                old_m = _CELL_RE.match(old)
                s = _S_ATTR.search(old_m.group("attrs"))
                style = s.group(1) if s else None
                if replaced_sst is not None and re.search(r'\st="s"', old_m.group("attrs")):
                    v = re.search(r"<v>(\d+)</v>", old_m.group("body") or "")
                    if v:
                        replaced_sst[v.group(1)] = replaced_sst.get(v.group(1), 0) + 1
                for fm in _F_TAG.finditer(old):
                    fattrs = fm.group(1) or ""
                    if 't="shared"' in fattrs and " ref=" in fattrs:
                        si = re.search(r'\ssi="(\d+)"', fattrs)
                        if si:
                            replaced_shared_masters.add(si.group(1))
            else:
                style = rstyle or col_styles.get(col0)
            cells[col0] = new_cell(col0, rnum, style, text)

        attrs = _SPANS_ATTR.sub("", attrs)
        extra = "".join(x for x in rest if x.strip())
        return f"<row{attrs}>" + "".join(cells[c] for c in sorted(cells)) + extra + "</row>"

    pending = sorted(targets)
    pi = 0
    out = [head]
    pos = 0
    last_r = 0
    max_row_seen = 0
    for rm in _ROW_RE.finditer(body):
        out.append(body[pos:rm.start()])
        pos = rm.end()
        attrs = rm.group("attrs")
        ra = _ROW_R_ATTR.search(attrs)
        rnum = int(ra.group(1)) if ra else last_r + 1
        last_r = rnum
        max_row_seen = max(max_row_seen, rnum)
        while pi < len(pending) and pending[pi] < rnum:
            out.append(build_row(pending[pi], f' r="{pending[pi]}"', "", targets[pending[pi]]))
            pi += 1
        if pi < len(pending) and pending[pi] == rnum:
            if not ra:
                attrs = f' r="{rnum}"' + attrs
            out.append(build_row(rnum, attrs, rm.group("body"), targets[rnum]))
            pi += 1
        else:
            out.append(rm.group(0))
    out.append(body[pos:])
    while pi < len(pending):
        out.append(build_row(pending[pi], f' r="{pending[pi]}"', "", targets[pending[pi]]))
        pi += 1
    out.append(tail)
    new_xml = "".join(out)

    # A replaced cell may have been the "master" of a formula shared with other cells.
    for si in replaced_shared_masters:
        if re.search(rf'<f\b[^>]*\ssi="{si}"', new_xml):
            raise TemplateError(
                "One of the columns you're updating holds a formula that other cells share. "
                "Clear that column in the template (or pick a different column) and try again."
            )

    # Extend the sheet's size marker if we wrote outside it.
    if targets:
        max_row = max(max(targets), max_row_seen)
        max_col0 = max(max(cols) for cols in targets.values())
        dm = re.search(r'<dimension\s+ref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"\s*/>', new_xml)
        if dm:
            c1, r1 = dm.group(1), int(dm.group(2))
            c2 = dm.group(3) or c1
            r2 = int(dm.group(4) or r1)
            c2i = max(_col_letters_to_index0(c2), max_col0)
            r2 = max(r2, max_row)
            new_ref = f'<dimension ref="{c1}{r1}:{get_column_letter(c2i + 1)}{r2}"/>'
            new_xml = new_xml[:dm.start()] + new_ref + new_xml[dm.end():]
    return new_xml


_SST_REF_RE = re.compile(r'\st="s"[^>]*><v>(\d+)</v>')
_SI_RE = re.compile(r"<si\b[^>]*?(?:/>|>.*?</si>)", re.S)


def _sheet_rows(xml):
    """Yields (row_number, row_match) for every row in a sheet's cell data."""
    m = re.search(r"<sheetData\b[^>]*?(/>|>)", xml)
    if not m or m.group(1) == "/>":
        return
    end = xml.index("</sheetData>", m.end())
    last_r = 0
    for rm in _ROW_RE.finditer(xml, m.end(), end):
        ra = _ROW_R_ATTR.search(rm.group("attrs"))
        last_r = int(ra.group(1)) if ra else last_r + 1
        yield last_r, rm


def _target_shared_text_refs(xml, targets):
    """{(row, col0): shared-text index} for target cells whose value lives in the shared text list."""
    out = {}
    for rnum, rm in _sheet_rows(xml):
        row_targets = targets.get(rnum)
        if not row_targets:
            continue
        last_c = -1
        for cm in _CELL_RE.finditer(rm.group("body") or ""):
            ref = _R_ATTR.search(cm.group("attrs"))
            col0 = _col_letters_to_index0(ref.group(1)) if ref and ref.group(1) else last_c + 1
            last_c = col0
            if col0 in row_targets and re.search(r'\st="s"', cm.group("attrs")):
                v = re.search(r"<v>(\d+)</v>", cm.group("body") or "")
                if v:
                    out[(rnum, col0)] = int(v.group(1))
    return out


def prepare_base(template_path, updates):
    """
    Builds everything shared by the batch.

    Each target cell is updated in the way that keeps the file closest to the template:
      - Excel stores most text in one shared list (sharedStrings.xml) and the cell just
        points at its entry. If an entry is used only by cells we're updating and they
        all get the same new value (each old TxnID, or the old date), the entry's text
        is swapped for the new value and the cells are left exactly as they are.
      - Any other target cell is rewritten in the sheet with the new text, keeping its
        style; shared-list entries left unused by that are emptied.
    TxnIDs become numbered placeholders, filled in per file.
    """
    entries, sheet_part, sheet_name = load_template_package(template_path)
    entries = _drop_calc_chain(entries)
    data_by_name = dict(entries)

    targets = {}
    k = 0  # TxnIDs are used column by column, top to bottom (same order as the original script)
    for up in updates:
        col0 = up["col_idx"]
        if up["logic"]["mode"] == MODE_TXN:
            prefix = up["logic"]["txn_prefix"].rstrip("_") + "_"
            for rnum in range(up["start"], up["end"] + 1):
                targets.setdefault(rnum, {})[col0] = f"{prefix}{_TOKEN_OPEN}{k}{_TOKEN_CLOSE}"
                k += 1
        else:
            value = xml_escape(str(up["logic"]["value"]))
            for rnum in range(up["start"], up["end"] + 1):
                targets.setdefault(rnum, {})[col0] = value

    sheet_xml = data_by_name[sheet_part].decode("utf-8")
    if _TOKEN_OPEN in sheet_xml:
        raise TemplateError("The template contains an unsupported special character.")
    sst_name = next((n for n, _ in entries if n.startswith("xl/") and n.endswith("sharedStrings.xml")), None)

    # 1) Values that can be swapped inside the shared text list
    sst_new_text = {}
    total_refs = {}
    if sst_name:
        for n, d in entries:
            if n.startswith("xl/worksheets/") and n.endswith(".xml"):
                for idx in _SST_REF_RE.findall(d.decode("utf-8")):
                    total_refs[int(idx)] = total_refs.get(int(idx), 0) + 1
        groups = {}
        for rc, idx in _target_shared_text_refs(sheet_xml, targets).items():
            groups.setdefault(idx, []).append(rc)
        for idx, cells in groups.items():
            values = {targets[r][c] for r, c in cells}
            if total_refs.get(idx, 0) == len(cells) and len(values) == 1:
                sst_new_text[idx] = values.pop()
                for r, c in cells:
                    del targets[r][c]
                    if not targets[r]:
                        del targets[r]

    # 2) Everything else is written into the sheet
    replaced_sst = {}
    if targets:
        sheet_xml = patch_sheet_xml(sheet_xml, targets, replaced_sst)
        sheet_bytes = sheet_xml.encode("utf-8")
    else:
        sheet_bytes = data_by_name[sheet_part]          # sheet untouched, byte for byte

    # 3) Update the shared text list
    sst_bytes = None
    if sst_name and (sst_new_text or replaced_sst):
        unused = {int(i) for i, c in replaced_sst.items() if total_refs.get(int(i), 0) == c}
        sst = data_by_name[sst_name].decode("utf-8")
        if _TOKEN_OPEN in sst:
            raise TemplateError("The template contains an unsupported special character.")
        counter = [-1]

        def swap(m):
            counter[0] += 1
            i = counter[0]
            if i in sst_new_text:
                return f"<si><t>{sst_new_text[i]}</t></si>"
            if i in unused:
                return "<si><t/></si>"
            return m.group(0)

        sst = _SI_RE.sub(swap, sst)
        removed_refs = sum(replaced_sst.values())
        cm = re.search(r'(<sst\b[^>]*\scount=")(\d+)(")', sst)
        if cm and removed_refs:
            sst = sst[:cm.start(2)] + str(max(0, int(cm.group(2)) - removed_refs)) + sst[cm.end(2):]
        sst_bytes = sst.encode("utf-8")

    # 4) Split parts that hold TxnID placeholders into chunks for fast per-file filling
    parts, templated, placed = [], {}, 0
    for name, data in entries:
        if name == sheet_part:
            data = sheet_bytes
        elif name == sst_name and sst_bytes is not None:
            data = sst_bytes
        if name in (sheet_part, sst_name):
            text = data.decode("utf-8")
            if _TOKEN_OPEN in text:
                pieces = _TOKEN_RE.split(text)
                templated[name] = (pieces[0::2], [int(x) for x in pieces[1::2]])
                placed += len(templated[name][1])
                data = None
        parts.append((name, data))
    if placed != k:
        raise TemplateError("Couldn't place every TxnID in the file.")
    return {"parts": parts, "templated": templated, "txn_count": k, "sheet_name": sheet_name,
            "swapped_in_list": len(sst_new_text), "rewritten_cells": sum(len(v) for v in targets.values())}


# =====================================================================
# MINIMAL ZIP WRITER (reuses pre-compressed parts for every file)
# =====================================================================
def _dos_datetime(ts=None):
    t = time.localtime(ts or time.time())
    dtime = (t.tm_hour << 11) | (t.tm_min << 5) | (t.tm_sec // 2)
    ddate = ((max(t.tm_year, 1980) - 1980) << 9) | (t.tm_mon << 5) | t.tm_mday
    return dtime, ddate


def _precompress(name, data, level=6):
    c = zlib.compressobj(level, zlib.DEFLATED, -15)
    comp = c.compress(data) + c.flush()
    return (name.encode("utf-8"), zlib.crc32(data) & 0xFFFFFFFF, len(comp), len(data), comp)


def _write_zip(path, parts, dtime, ddate):
    """parts: list of (name_bytes, crc, csize, usize, comp_bytes)."""
    central = []
    offset = 0
    with open(path, "wb") as f:
        for name, crc, csize, usize, comp in parts:
            flags = 0x800 if any(b > 127 for b in name) else 0
            header = struct.pack("<IHHHHHIIIHH", 0x04034B50, 20, flags, 8, dtime, ddate,
                                 crc, csize, usize, len(name), 0)
            f.write(header)
            f.write(name)
            f.write(comp)
            central.append(struct.pack("<IHHHHHHIIIHHHHHII", 0x02014B50, 20, 20, flags, 8,
                                       dtime, ddate, crc, csize, usize, len(name),
                                       0, 0, 0, 0, 0, offset) + name)
            offset += len(header) + len(name) + csize
            if offset >= 0xFFFFFFFF:
                raise ValueError("Generated file would be larger than 4 GB.")
        cd = b"".join(central)
        f.write(cd)
        f.write(struct.pack("<IHHHHIIH", 0x06054B50, 0, 0, len(parts), len(parts),
                            len(cd), offset, 0))


# =====================================================================
# WORKER
# =====================================================================
# Data shared by every file in a batch is sent to each worker process once.
_SHARED = {}


def _init_worker(shared):
    global _SHARED
    _SHARED = shared


def generate_single_file(task):
    fn_num, txn_start_val = task
    s = _SHARED
    full_path = os.path.join(s["target_dir"], f"{s['naming_prefix']}{fn_num}{s['extension']}")

    zip_parts = []
    for item in s["parts"]:
        if isinstance(item, str):                      # a part holding TxnID placeholders
            chunks, slots = s["templated"][item]
            pieces = [None] * (len(chunks) + len(slots))
            pieces[0::2] = chunks
            pieces[1::2] = [str(txn_start_val - k) for k in slots]
            zip_parts.append(_precompress(item, "".join(pieces).encode("utf-8"), s["level"]))
        else:
            zip_parts.append(item)
    _write_zip(full_path, zip_parts, s["dtime"], s["ddate"])
    return fn_num, txn_start_val - s["txn_count"], full_path


# =====================================================================
# BATCH RUNNER
# =====================================================================
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

    # Step 1: read the template and prepare the shared copy
    stage("Preparing template...")
    t1 = time.time()
    base = prepare_base(template_path, updates)
    level = 6   # same compression strength as Excel, so files stay close to the template size
    parts = [name if data is None else _precompress(name, data, level) for name, data in base["parts"]]
    step1 = time.time() - t1

    # Step 2: reserve numbers (counters move past the whole batch before generating)
    stage("Reserving numbers from tracker...")
    t2 = time.time()
    per_file = base["txn_count"]
    tracker = Tracker(tracker_dir)
    fn_numbers, txn_base = tracker.reserve(n_files, per_file)
    step2 = time.time() - t2

    # Step 3: generation, each file gets its own TxnID block so IDs never repeat
    stage("Generating files...")
    t3 = time.time()
    dtime, ddate = _dos_datetime()
    shared = {
        "parts": parts, "templated": base["templated"], "txn_count": per_file, "level": level,
        "naming_prefix": naming_prefix, "extension": extension, "target_dir": output_dir,
        "dtime": dtime, "ddate": ddate,
    }
    tasks = [(fn_numbers[i], txn_base - i * per_file) for i in range(n_files)]

    results = []
    cores = 1
    if parallel and n_files > 3:
        try:
            cores = max(1, min(cpu_count(), n_files))
            with Pool(processes=cores, initializer=_init_worker, initargs=(shared,)) as pool:
                for res in pool.imap_unordered(generate_single_file, tasks, chunksize=1):
                    results.append(res)
                    if progress_cb:
                        progress_cb(len(results), n_files)
        except Exception:
            cores = 1  # finish whatever is left one by one
    if len(results) < n_files:
        _init_worker(shared)
        done = {r[0] for r in results}
        for task in tasks:
            if task[0] in done:
                continue
            results.append(generate_single_file(task))
            if progress_cb:
                progress_cb(len(results), n_files)
    step3 = time.time() - t3

    results.sort(key=lambda r: -r[0])
    return {
        "files": [r[2] for r in results],
        "output_dir": output_dir,
        "sheet_name": base["sheet_name"],
        "step1": step1,
        "step2": step2,
        "step3": step3,
        "total": time.time() - overall_start,
        "cores": cores,
        "txn_range": [txn_base, txn_base - n_files * per_file + 1] if per_file else None,
        "fn_range": [fn_numbers[0], fn_numbers[-1]],
        "first_file": os.path.basename(results[0][2]) if results else None,
        "last_file": os.path.basename(results[-1][2]) if results else None,
    }
