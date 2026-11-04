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
copied exactly as the template packed them; changed parts are packed about as
tightly as the template packed them, so file sizes stay with the template's.

ID rules (enforced by Tracker.reserve):
  - file numbers and TxnIDs are taken from the project's tracker files, counting down
  - every number is used once only; the counters are moved past the whole batch
    BEFORE any file is written, so a failed or cancelled batch never reuses numbers
  - a batch that would need a number below 1 is refused (so nothing negative ever appears)
"""
import os
import re
import shutil
import itertools
import struct
import uuid
import time
import zipfile
import zlib
from multiprocessing import Pool, cpu_count
from xml.sax.saxutils import escape as xml_escape

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

# --- CONFIGURATION ---
ENGINE_VERSION = 7               # BulkFileGenerator.py checks this so mismatched files are caught
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

# Date formats offered in the app -> Excel number-format codes used when a date is saved
# as a real Excel date. "/" is escaped so Excel shows a slash on every PC, whatever the
# PC's regional date separator.
DATE_FORMAT_CODES = {
    "dd/mm/yyyy": r"dd\/mm\/yyyy",
    "mm/dd/yyyy": r"mm\/dd\/yyyy",
    "yyyy-mm-dd": "yyyy-mm-dd",
    "dd-mm-yyyy": "dd-mm-yyyy",
}


class DateCell:
    """A fixed date saved as a real Excel date (a number shown with a date format)."""
    __slots__ = ("serial", "fmt")

    def __init__(self, serial, fmt):
        self.serial, self.fmt = serial, fmt


def excel_serial(iso_date, date1904=False):
    """Excel's day number for a yyyy-mm-dd date."""
    import datetime
    d = datetime.date.fromisoformat(iso_date)
    base = datetime.date(1904, 1, 1) if date1904 else datetime.date(1899, 12, 30)
    return (d - base).days


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

    def reserve(self, n_files, txn_per_file, file_numbers=True):
        """
        Reserves numbers for a whole batch and moves the counters past it
        before anything is generated. Returns (file_numbers, txn_start).
        file_numbers=False: only TxnIDs are needed (files keep their names).
        Raises TrackerError if the batch would go below LOWEST_ALLOWED_NUMBER.
        """
        with _FileLock(self.lock_file):
            fn_next, txn_next = self.peek()
            # file numbers only matter when files get new names
            errors = check_capacity(fn_next if file_numbers else 10 ** 12, txn_next, n_files, txn_per_file)
            if errors:
                raise TrackerError(" ".join(errors))
            if file_numbers:
                self._write(self.filename_file, fn_next - n_files)
            self._write(self.txnid_file, txn_next - n_files * txn_per_file)
        return ([fn_next - i for i in range(n_files)] if file_numbers else []), txn_next


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
            as_text = logic.get("as_text")
            if as_text not in (True, False):
                errors.append(f"{tag}: choose whether the date is saved as text or as an Excel date.")
            elif as_text is False:
                if logic.get("fmt") not in DATE_FORMAT_CODES:
                    errors.append(f"{tag}: choose a date format.")
                try:
                    excel_serial(str(logic.get("iso", "")))
                except ValueError:
                    errors.append(f"{tag}: the date isn't a valid calendar date.")
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
    def shown(v):
        if v is None:
            return ""
        if hasattr(v, "hour") and hasattr(v, "date") and (v.hour, v.minute, v.second) == (0, 0, 0):
            return v.date().isoformat()          # a real Excel date: show just the date
        return str(v)

    for r in ws.iter_rows(min_row=1, max_row=rows, values_only=True):
        data.append([shown(v) for v in r])
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
_ANY_CELL = re.compile(r"<c[\s>/]")


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
        raise TemplateError("This isn't a valid .xlsx file (it may be an older .xls, damaged, or password-protected).")
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


def patch_sheet_xml(xml, targets, replaced_sst=None, stats=None, date_style=None):
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
        ref = f"{get_column_letter(col0 + 1)}{rnum}"
        if isinstance(text, DateCell):        # real Excel date: a number with a date format
            return f'<c r="{ref}" s="{date_style(style, text.fmt)}"><v>{text.serial}</v></c>'
        s = f' s="{style}"' if style else ""
        return f'<c r="{ref}"{s} t="inlineStr"><is><t>{text}</t></is></c>'

    def row_default_style(attrs):
        if re.search(r'\scustomFormat="(1|true)"', attrs):
            s = _S_ATTR.search(attrs)
            if s:
                return s.group(1)
        return None

    def replacement(col0, rnum, old, text, rstyle):
        """The new cell for a target. Keeps the old cell's style and notes what it held."""
        if old is not None:
            old_m = _CELL_RE.match(old)
            s = _S_ATTR.search(old_m.group("attrs"))
            style = s.group(1) if s else None
            if replaced_sst is not None and re.search(r'\st="s"', old_m.group("attrs")):
                v = re.search(r"<v>(\d+)</v>", old_m.group("body") or "")
                if v:
                    replaced_sst[v.group(1)] = replaced_sst.get(v.group(1), 0) + 1
            if stats is not None and "<f" in (old_m.group("body") or ""):
                stats["formula_replaced"] = True
            for fm in _F_TAG.finditer(old):
                fattrs = fm.group(1) or ""
                if 't="shared"' in fattrs and " ref=" in fattrs:
                    si = re.search(r'\ssi="(\d+)"', fattrs)
                    if si:
                        replaced_shared_masters.add(si.group(1))
        else:
            style = rstyle or col_styles.get(col0)
        return new_cell(col0, rnum, style, text)

    def build_row(rnum, attrs, row_body, row_targets):
        body = row_body or ""
        rstyle = row_default_style(attrs)

        # Fast path: every cell carries its reference and every target cell already
        # exists, so only those cells are swapped and the rest of the row is kept as is.
        if body and len(_ANY_CELL.findall(body)) == body.count('<c r="'):
            spans = []
            for col0, text in row_targets.items():
                i = body.find(f'<c r="{get_column_letter(col0 + 1)}{rnum}"')
                cm = _CELL_RE.match(body, i) if i >= 0 else None
                if cm is None:
                    spans = None
                    break
                spans.append((cm.start(), cm.end(), replacement(col0, rnum, cm.group(0), text, rstyle)))
            if spans is not None:
                spans.sort()
                out, p = [], 0
                for a, b, cell in spans:
                    out.append(body[p:a])
                    out.append(cell)
                    p = b
                out.append(body[p:])
                return f"<row{_SPANS_ATTR.sub('', attrs)}>" + "".join(out) + "</row>"

        # General path: take the row apart, set the targets, put it back in column order.
        cells = {}
        rest = []
        p = 0
        last_c = -1
        for cm in _CELL_RE.finditer(body):
            rest.append(body[p:cm.start()])
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
        rest.append(body[p:])
        for col0, text in row_targets.items():
            cells[col0] = replacement(col0, rnum, cells.get(col0), text, rstyle)

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


class _StyleEditor:
    """Adds cell styles for real Excel dates. Each new style is a copy of the cell's own
    style (font, fill, border, alignment kept) with the chosen date format, and without
    the "text" apostrophe flag. Nothing else in styles.xml changes."""

    def __init__(self, xml):
        self.xml = xml
        self.changed = False
        self.cache = {}
        m = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", xml, re.S)
        if not m:
            raise TemplateError("The template's style table is missing.")
        self.xfs = re.findall(r"<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>", m.group(1), re.S)
        self.new_xfs = []
        self.fmt_ids = {}
        for tag in re.findall(r"<numFmt\b[^>]*?/?>", xml):
            fid = re.search(r'\snumFmtId="(\d+)"', tag)
            code = re.search(r'\sformatCode="([^"]*)"', tag)
            if fid and code:
                self.fmt_ids.setdefault(code.group(1), int(fid.group(1)))
        self.new_fmts = []

    def _fmt_id(self, fmt):
        code = DATE_FORMAT_CODES[fmt]
        if code not in self.fmt_ids:
            self.fmt_ids[code] = max([163] + list(self.fmt_ids.values())) + 1
            self.new_fmts.append((self.fmt_ids[code], code))
        return self.fmt_ids[code]

    @staticmethod
    def _set_attr(tag, name, value):
        if re.search(rf'\s{name}="[^"]*"', tag):
            return re.sub(rf'(\s{name}=")[^"]*(")', rf'\g<1>{value}\g<2>', tag, count=1)
        return tag.replace("<xf", f'<xf {name}="{value}"', 1)

    def date_style(self, old_style, fmt):
        key = (old_style, fmt)
        if key in self.cache:
            return self.cache[key]
        fid = self._fmt_id(fmt)
        idx = int(old_style) if old_style and int(old_style) < len(self.xfs) else 0
        base = self.xfs[idx]
        cur = re.search(r'\snumFmtId="(\d+)"', base)
        if cur and int(cur.group(1)) == fid and 'quotePrefix="1"' not in base:
            self.cache[key] = str(idx)
            return self.cache[key]
        tag = re.sub(r'\squotePrefix="[^"]*"', "", base)
        tag = self._set_attr(tag, "numFmtId", fid)
        tag = self._set_attr(tag, "applyNumberFormat", "1")
        self.new_xfs.append(tag)
        self.cache[key] = str(len(self.xfs) + len(self.new_xfs) - 1)
        self.changed = True
        return self.cache[key]

    def result(self):
        xml = self.xml
        if self.new_fmts:
            add = "".join(f'<numFmt numFmtId="{i}" formatCode="{c}"/>' for i, c in self.new_fmts)
            m = re.search(r"<numFmts\b[^>]*?(/?)>", xml)
            if m and m.group(1) == "/":            # empty list written as <numFmts count="0"/>
                xml = xml[:m.start()] + f'<numFmts count="{len(self.new_fmts)}">{add}</numFmts>' + xml[m.end():]
            elif m:
                end = xml.index("</numFmts>", m.end())
                xml = xml[:end] + add + xml[end:]
                cm = re.search(r'(<numFmts\b[^>]*\scount=")(\d+)(")', xml)
                if cm:
                    xml = xml[:cm.start(2)] + str(int(cm.group(2)) + len(self.new_fmts)) + xml[cm.end(2):]
            else:
                ss = re.search(r"<styleSheet\b[^>]*>", xml)
                xml = xml[:ss.end()] + f'<numFmts count="{len(self.new_fmts)}">{add}</numFmts>' + xml[ss.end():]
        if self.new_xfs:
            end = xml.index("</cellXfs>")
            xml = xml[:end] + "".join(self.new_xfs) + xml[end:]
            cm = re.search(r'(<cellXfs\b[^>]*\scount=")(\d+)(")', xml)
            if cm:
                xml = xml[:cm.start(2)] + str(len(self.xfs) + len(self.new_xfs)) + xml[cm.end(2):]
        return xml


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
    """{(row, col0): shared-text index} for target cells whose value lives in the shared text list.
    Only cells in the target columns are looked at. Cells written without a reference
    (rare) aren't found here; they are then simply rewritten in the sheet, which is also correct."""
    cols = {get_column_letter(c + 1) for row in targets.values() for c in row}
    if not cols:
        return {}
    pat = re.compile(r'<c r="(' + "|".join(sorted(cols, key=len, reverse=True)) +
                     r')(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
    out = {}
    for m in pat.finditer(xml):
        rnum = int(m.group(2))
        row_targets = targets.get(rnum)
        if not row_targets:
            continue
        col0 = _col_letters_to_index0(m.group(1))
        if col0 in row_targets and re.search(r'\st="s"', m.group(3)):
            v = re.search(r"<v>(\d+)</v>", m.group(4) or "")
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
    data_by_name = dict(entries)
    originals = dict(entries)

    targets = {}
    k = 0  # TxnIDs are used column by column, top to bottom (same order as the original script)
    for up in updates:
        col0 = up["col_idx"]
        if up["logic"]["mode"] == MODE_TXN:
            prefix = up["logic"]["txn_prefix"].rstrip("_") + "_"
            for rnum in range(up["start"], up["end"] + 1):
                targets.setdefault(rnum, {})[col0] = f"{prefix}{_TOKEN_OPEN}{k}{_TOKEN_CLOSE}"
                k += 1
        elif up["logic"].get("as_text", True) is False:   # real Excel date
            wb_xml = next((d for n, d in entries if n == "xl/workbook.xml"), b"").decode("utf-8", "replace")
            date1904 = bool(re.search(r'<workbookPr\b[^>]*\sdate1904="(1|true)"', wb_xml))
            value = DateCell(excel_serial(up["logic"]["iso"], date1904), up["logic"]["fmt"])
            for rnum in range(up["start"], up["end"] + 1):
                targets.setdefault(rnum, {})[col0] = value
        else:                                              # text, exactly as shown
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
            if any(not isinstance(v, str) for v in values):
                continue                                   # dates saved as numbers go in the sheet
            if total_refs.get(idx, 0) == len(cells) and len(values) == 1:
                sst_new_text[idx] = values.pop()
                for r, c in cells:
                    del targets[r][c]
                    if not targets[r]:
                        del targets[r]

    # 2) Everything else is written into the sheet
    replaced_sst = {}
    stats = {}
    styles_name = next((n for n, _ in entries if n.startswith("xl/") and n.endswith("styles.xml")), None)
    styles = _StyleEditor(data_by_name[styles_name].decode("utf-8")) if styles_name else None
    if targets:
        sheet_xml = patch_sheet_xml(sheet_xml, targets, replaced_sst, stats,
                                    styles.date_style if styles else None)
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

    # Excel's formula calculation list must go only if a formula cell was overwritten
    # (Excel rebuilds it silently); otherwise it is kept exactly as in the template.
    if stats.get("formula_replaced"):
        entries = _drop_calc_chain(entries)

    # 4) Split parts that hold TxnID placeholders into chunks for fast per-file filling
    parts, templated, placed = [], {}, 0
    for name, data in entries:
        if name == sheet_part:
            data = sheet_bytes
        elif name == sst_name and sst_bytes is not None:
            data = sst_bytes
        elif name == styles_name and styles is not None and styles.changed:
            data = styles.result().encode("utf-8")
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
            "originals": originals,
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
    """Packs one part. level 0 = stored without compression (as some templates do)."""
    crc = zlib.crc32(data) & 0xFFFFFFFF
    if level == 0:
        return (name.encode("utf-8"), crc, len(data), len(data), data, 0)
    c = zlib.compressobj(level, zlib.DEFLATED, -15)
    comp = c.compress(data) + c.flush()
    return (name.encode("utf-8"), crc, len(comp), len(data), comp, 8)


def _read_raw_parts(path):
    """The template's parts exactly as packed: {name: (name_bytes, crc, csize, usize, raw, method)}.
    Unchanged parts are copied into every file as-is, so they keep the template's exact size."""
    out = {}
    with zipfile.ZipFile(path) as zf, open(path, "rb") as f:
        for info in zf.infolist():
            if info.is_dir() or info.compress_type not in (0, 8) or info.flag_bits & 0x1:
                continue
            f.seek(info.header_offset)
            h = f.read(30)
            if len(h) < 30 or struct.unpack("<I", h[:4])[0] != 0x04034B50:
                continue
            n_len, x_len = struct.unpack("<HH", h[26:30])
            f.seek(info.header_offset + 30 + n_len + x_len)
            raw = f.read(info.compress_size)
            out[info.filename] = (info.filename.encode("utf-8"), info.CRC & 0xFFFFFFFF,
                                  info.compress_size, info.file_size, raw, info.compress_type)
    return out


def _matching_level(new_data, raw_entry):
    """The compression strength (1-9) at which this part's NEW content packs closest to
    the size the template's version of the part had (zlib's levels are coarse and not in
    order, so they are all tried). 0 = the template stored the part unpacked.
    Large parts: the 3 most promising levels are picked on a slice, then checked on the
    whole content. Runs once per batch, not per file."""
    if raw_entry is None:
        return 6
    method, csize, usize = raw_entry[5], raw_entry[2], raw_entry[3]
    if method == 0 or not usize:
        return 0

    def packed(data, level):
        c = zlib.compressobj(level, zlib.DEFLATED, -15)
        return len(c.compress(data) + c.flush())

    if len(new_data) <= 8_000_000:
        candidates = range(1, 10)
    else:
        sample = new_data[:4_000_000]
        aim = csize / usize * len(sample)
        candidates = sorted(range(1, 10), key=lambda lv: abs(packed(sample, lv) - aim))[:3]
    return min(candidates, key=lambda lv: abs(packed(new_data, lv) - csize))


def _pad_fields(pad, count):
    """Splits `pad` bytes into standard zip 'padding' extra fields (ID 0xD935, as used by
    Android's zipalign), at most PAD_PER_ENTRY bytes each, one per part. Readers skip them."""
    if pad <= 0:
        return [b""] * count
    n = -(-pad // PAD_PER_ENTRY)
    if n > count or pad < 4:
        raise ValueError("internal: padding does not fit")
    sizes = [pad // n + (1 if i < pad % n else 0) for i in range(n)]
    fields = [struct.pack("<HH", 0xD935, sz - 4) + b"\0" * (sz - 4) for sz in sizes]
    return fields + [b""] * (count - n)


def _write_zip(path, parts, dtime, ddate, pad=0):
    """parts: list of (name_bytes, crc, csize, usize, packed_bytes, method).
    pad: bytes of padding (zip extra fields in the part headers) that make the file exactly
    the template's size. Excel and other spreadsheet programs skip them."""
    central = []
    offset = 0
    extras = _pad_fields(pad, len(parts))
    with open(path, "wb") as f:
        for (name, crc, csize, usize, comp, method), extra in zip(parts, extras):
            flags = 0x800 if any(b > 127 for b in name) else 0
            header = struct.pack("<IHHHHHIIIHH", 0x04034B50, 20, flags, method, dtime, ddate,
                                 crc, csize, usize, len(name), len(extra))
            f.write(header)
            f.write(name)
            f.write(extra)
            f.write(comp)
            central.append(struct.pack("<IHHHHHHIIIHHHHHII", 0x02014B50, 20, 20, flags, method,
                                       dtime, ddate, crc, csize, usize, len(name),
                                       0, 0, 0, 0, 0, offset) + name)
            offset += len(header) + len(name) + len(extra) + csize
            if offset >= 0xFFFFFFFF:
                raise ValueError("Generated file would be larger than 4 GB.")
        cd = b"".join(central)
        f.write(cd)
        f.write(struct.pack("<IHHHHIIH", 0x06054B50, 0, 0, len(parts), len(parts),
                            len(cd), offset, 0))


# ---------------------------------------------------------------- matching the template's size
ZIP_ENTRY_OVERHEAD = 76     # per part: local header (30) + central directory entry (46), plus the name twice
ZIP_END = 22                # end-of-archive record (no comment)
PAD_PER_ENTRY = 16000       # padding per part header; LibreOffice rejects fields over ~32 KB


def _pad_ok(gap, cap):
    return gap == 0 or 4 <= gap <= cap


def _entry_size(part):
    return ZIP_ENTRY_OVERHEAD + 2 * len(part[0]) + part[2]


def _zip_total(parts):
    """Exact size of the file _write_zip writes for these parts, before any comment."""
    return sum(_entry_size(p) for p in parts) + ZIP_END


def _pad_for(parts, target):
    """Padding that makes the file exactly `target` bytes, or 0 if that isn't possible."""
    pad = target - _zip_total(parts)
    return pad if _pad_ok(pad, PAD_PER_ENTRY * len(parts)) else 0


def _plan(fixed_bytes, choices, target, margin, cap):
    """Picks one option per changed part so the file can be padded to exactly `target` bytes.
    choices: list (one per part with a choice) of lists of (key, entry_bytes).
    Prefers the smallest padding; if no combination fits, the one closest to the target.
    Returns (keys, fits)."""
    best = None
    for combo in itertools.product(*choices):
        total = fixed_bytes + sum(c[1] for c in combo) + ZIP_END
        gap = target - total
        if (gap == 0 or max(margin, 4) <= gap) and gap <= cap - margin:
            rank = (0, gap)
        elif gap > cap - margin:
            rank = (1, gap)
        else:
            rank = (2, -gap)
        if best is None or rank < best[0]:
            best = (rank, [c[0] for c in combo])
    return best[1], best[0][0] == 0


def _level_options(name, data, raw_entry=None, include_raw=False):
    """Every way to pack one part: levels 1-9 (and the template's own packing, if allowed)."""
    opts = {lv: _precompress(name, data, lv) for lv in range(1, 10)}
    if include_raw and raw_entry is not None:
        opts["raw"] = raw_entry
    return opts


def fit_parts(entries, raw, originals, target, margin, cache=None):
    """
    entries: list of (name, data) where data is the part's final content, or None when the part
    is unchanged from the original file. Returns (parts, pad, exact).
    Unchanged parts are copied exactly as packed. Changed parts are packed at the strengths
    that let the whole file be topped up to exactly `target` bytes; if the content grew too much
    for that, the largest unchanged part is repacked more tightly to make room.
    cache (dict) remembers the choice so similar files are done with a single packing each.
    """
    changed = [(n, d) for n, d in entries if d is not None]
    fixed = [raw[n] for n, d in entries if d is None]
    cap = PAD_PER_ENTRY * len(entries)

    def assemble(packed):
        out = []
        for n, d in entries:
            out.append(packed[n] if n in packed else raw[n])
        return out

    # 1) the choice that fitted the previous similar file
    if (cache and cache.get("levels") is not None
            and set(cache["levels"]) == {n for n, _ in changed}
            and all(n in raw and n in originals for n in cache.get("repack", {}))):
        packed = {}
        for n, d in changed:
            packed[n] = _precompress(n, d, cache["levels"][n])
        for n, lv in cache.get("repack", {}).items():
            packed[n] = raw[n] if lv == "raw" else _precompress(n, originals[n], lv)
        parts = assemble(packed)
        pad = target - _zip_total(parts)
        if _pad_ok(pad, cap):
            return parts, pad, True

    # 2) search the changed parts' packing strengths
    opts = {n: _level_options(n, d) for n, d in changed}
    fixed_bytes = sum(_entry_size(p) for p in fixed)
    keys, fits = _plan(fixed_bytes, [[(k, _entry_size(v)) for k, v in opts[n].items()] for n, _ in changed],
                       target, margin, cap)
    repack = {}
    if not fits:
        # 3) make room by repacking the largest unchanged part (the sheet, usually) as well
        big = max((n for n, d in entries if d is None and n in originals), key=lambda n: raw[n][2], default=None)
        if big is not None:
            opts[big] = _level_options(big, originals[big], raw[big], include_raw=True)
            fixed_bytes -= _entry_size(raw[big])
            names = [n for n, _ in changed] + [big]
            keys2, fits2 = _plan(fixed_bytes, [[(k, _entry_size(v)) for k, v in opts[n].items()] for n in names],
                                 target, margin, cap)
            if fits2:
                keys, fits = keys2[:-1], True
                repack = {big: keys2[-1]}
    packed = {n: opts[n][k] for (n, _), k in zip(changed, keys)}
    for n, k in repack.items():
        packed[n] = opts[n][k]
    if cache is not None:
        cache["levels"] = {n: k for (n, _), k in zip(changed, keys)}
        cache["repack"] = repack
    parts = assemble(packed)
    pad = _pad_for(parts, target)
    return parts, pad, pad > 0 or _zip_total(parts) == target


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
            zip_parts.append(_precompress(item, "".join(pieces).encode("utf-8"), s["levels"][item]))
        else:
            zip_parts.append(item)
    pad = _pad_for(zip_parts, s["target_size"])
    _write_zip(full_path, zip_parts, s["dtime"], s["ddate"], pad)
    return fn_num, txn_start_val - s["txn_count"], full_path, (_zip_total(zip_parts) + pad == s["target_size"])


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
    # Unchanged parts: copied exactly as packed in the template (same bytes, same size).
    # Changed parts: packed about as tightly as the template packed them.
    raw = _read_raw_parts(template_path)
    originals = base.pop("originals")
    try:                                    # realistic TxnIDs for judging how the new content packs
        sample_start = Tracker(tracker_dir).peek()[1]
    except TrackerError:
        sample_start = DEFAULT_COUNTER_SEED
    # Choose packing so every file can be made exactly the template's size: parts with TxnIDs
    # are judged on a realistic sample (their size varies a little per file, hence the margin).
    target = os.path.getsize(template_path)
    entries, samples = [], {}
    for name, data in base["parts"]:
        if data is None:
            chunks, slots = base["templated"][name]
            pieces = [None] * (len(chunks) + len(slots))
            pieces[0::2] = chunks
            pieces[1::2] = [str(sample_start - k) for k in slots]
            samples[name] = "".join(pieces).encode("utf-8")
            entries.append((name, samples[name]))
        elif name in raw and data == originals.get(name):
            entries.append((name, None))
        else:
            entries.append((name, data))
    cache = {}
    sample_parts, _pad, _ok = fit_parts(entries, raw, originals, target, margin=4096, cache=cache)
    chosen = {p[0].decode("utf-8"): p for p in sample_parts}
    parts, levels = [], {}
    for name, data in base["parts"]:
        if data is None:
            levels[name] = cache["levels"][name]
            parts.append(name)
        else:
            parts.append(chosen[name])
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
        "parts": parts, "templated": base["templated"], "txn_count": per_file, "levels": levels,
        "target_size": target,
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
        "template_size": target,
        "exact_size": sum(1 for r in results if r[3]),
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


# =====================================================================
# UPDATE EXISTING FILES (task "Update file names / column data")
# =====================================================================
WHAT_NAMES, WHAT_DATA, WHAT_BOTH = "names", "data", "both"
SAVE_IN_PLACE, SAVE_COPIES = "in_place", "copies"
_TEMP_TAG = ".bulkgen_"


def _natural_key(name):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def _quick_sheet_info(path):
    """(last row, active sheet name) read cheaply: only the start of the sheet is unpacked."""
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        wb_xml = zf.read("xl/workbook.xml").decode("utf-8", "replace")
        rels = _rels_map(zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")) \
            if "xl/_rels/workbook.xml.rels" in names else {}
        sheets = re.findall(r"<sheet\b[^>]*?/?>", wb_xml)
        m = re.search(r'<workbookView\b[^>]*\sactiveTab="(\d+)"', wb_xml)
        active = int(m.group(1)) if m else 0
        tag = sheets[active if active < len(sheets) else 0]
        rid = re.search(r'\s(?:\w+:)?id="([^"]+)"', tag)
        sheet_name = re.search(r'\sname="([^"]*)"', tag)
        part = _resolve_target("xl", rels[rid.group(1)][0]) if rid and rid.group(1) in rels else None
        rows = None
        if part in names:
            with zf.open(part) as fh:
                head = fh.read(65536).decode("utf-8", "replace")
            dm = re.search(r'<dimension\s+ref="[A-Z]+\d+(?::[A-Z]+(\d+))?"', head)
            if dm:
                rows = int(dm.group(1)) if dm.group(1) else 1
        return rows, (sheet_name.group(1) if sheet_name else "Sheet")


def scan_folder(folder, extension=".xlsx"):
    """Lists the files to update: {folder, files:[{name, size, rows, sheet | error}]}, sorted
    the way Windows Explorer sorts names. Excel's own lock files (~$...) are skipped."""
    folder = os.path.abspath(os.path.expanduser(str(folder).strip().strip('"').strip("'")))
    if not os.path.isdir(folder):
        raise TemplateError(f"This folder doesn't exist or can't be opened: {folder}")
    files = []
    for name in os.listdir(folder):
        path = os.path.join(folder, name)
        if (not os.path.isfile(path) or not name.lower().endswith(extension.lower())
                or name.startswith("~$") or name.startswith(_TEMP_TAG)):
            continue
        info = {"name": name, "size": os.path.getsize(path)}
        try:
            info["rows"], info["sheet"] = _quick_sheet_info(path)
        except Exception:
            info["error"] = "not a readable .xlsx file"
        files.append(info)
    files.sort(key=lambda f: _natural_key(f["name"]))
    return {"folder": folder, "files": files}


def _normalise_updates(updates):
    return [
        {**u, "col": str(u["col"]).strip().upper(),
         "col_idx": column_index_from_string(str(u["col"]).strip().upper()) - 1,
         "start": int(u["start"]), "end": int(u["end"])}
        for u in updates
    ]


_FIT_CACHE = {}   # per worker process: the packing choice that made the previous file fit


def build_updated_parts(src_path, updates, txn_start):
    """The package of one existing file with the column rules applied, packed so the file
    comes out exactly the size it was. Returns (parts, pad, txn_count, sheet_name, exact)."""
    base = prepare_base(src_path, updates)
    raw = _read_raw_parts(src_path)
    originals = base.pop("originals")
    entries = []
    for name, data in base["parts"]:
        if data is None:
            chunks, slots = base["templated"][name]
            pieces = [None] * (len(chunks) + len(slots))
            pieces[0::2] = chunks
            pieces[1::2] = [str(txn_start - k) for k in slots]
            entries.append((name, "".join(pieces).encode("utf-8")))
        elif name in raw and data == originals.get(name):
            entries.append((name, None))
        else:
            entries.append((name, data))
    parts, pad, exact = fit_parts(entries, raw, originals, os.path.getsize(src_path), margin=0, cache=_FIT_CACHE)
    return parts, pad, base["txn_count"], base["sheet_name"], exact


def update_single_file(task):
    """One file. task = (index, source_path, final_path, restore_path, txn_start, delete_source)
    Writes to a temporary file first and only then swaps it in, so a failure never leaves a
    half-written file. On failure a renamed source is put back under its original name."""
    idx, src, final, restore, txn_start, delete_src = task
    s = _SHARED
    size_before = os.path.getsize(src)
    out = {"index": idx, "old": os.path.basename(restore), "new": os.path.basename(final),
           "size_before": size_before, "size_after": None, "error": None, "sheet": None}
    tmp = final + ".bulkgen.tmp"
    try:
        if s["data"]:
            parts, pad, _count, sheet, _exact = build_updated_parts(src, s["updates"], txn_start)
            out["sheet"] = sheet
            _write_zip(tmp, parts, s["dtime"], s["ddate"], pad)
            os.replace(tmp, final)
            if delete_src and os.path.abspath(src) != os.path.abspath(final):
                os.remove(src)
        elif s["copies"]:
            shutil.copy2(src, tmp)
            os.replace(tmp, final)
        else:
            os.replace(src, final)          # rename only: the file itself is untouched
        out["size_after"] = os.path.getsize(final)
    except Exception as e:
        out["error"] = _friendly_error(e)
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
            if src != restore and os.path.exists(src) and not os.path.exists(restore):
                os.replace(src, restore)    # put the original name back
        except Exception:
            pass
    return out


def _friendly_error(e):
    if isinstance(e, PermissionError):
        return "file is open in another program (close it in Excel) or the folder is read-only"
    if isinstance(e, (TemplateError, TrackerError)):
        return str(e)
    return f"{type(e).__name__}: {e}"


def run_update(folder, file_names, what, naming_prefix, updates, tracker_dir,
               save_mode, output_dir=None, parallel=True, progress_cb=None, stage_cb=None):
    """
    Updates existing files.
      what: "names" (new file names only, contents byte-for-byte unchanged),
            "data"  (apply the column rules, names unchanged), or "both".
      save_mode: "in_place" (replace the files in their folder) or "copies" (write updated
            copies to output_dir and leave the originals alone).
    New names use the naming prefix + the project's file-number counter, in the order the
    files are listed (as Windows Explorer sorts them). TxnIDs come from the project's counter.
    """
    stage = stage_cb or (lambda _t: None)
    t0 = time.time()
    if what not in (WHAT_NAMES, WHAT_DATA, WHAT_BOTH):
        raise TemplateError("Choose what to update: file names, data, or both.")
    if save_mode not in (SAVE_IN_PLACE, SAVE_COPIES):
        raise TemplateError("Choose where to save: in the same folder, or copies in another folder.")
    renaming, data = what in (WHAT_NAMES, WHAT_BOTH), what in (WHAT_DATA, WHAT_BOTH)
    copies = save_mode == SAVE_COPIES
    folder = os.path.abspath(folder)
    sources = [os.path.join(folder, n) for n in file_names]
    missing = [os.path.basename(p) for p in sources if not os.path.isfile(p)]
    if missing:
        raise TemplateError(f"{len(missing)} file(s) are no longer in the folder, e.g. {missing[0]}. Search the folder again.")
    n = len(sources)
    if not n:
        raise TemplateError("There are no files to update.")
    updates = _normalise_updates(updates) if data else []
    per_file = txn_ids_per_file(updates) if data else 0
    if renaming:
        naming_prefix = naming_prefix if naming_prefix.endswith("_") else naming_prefix + "_"
    dest_dir = os.path.abspath(output_dir) if copies else folder
    if copies:
        if os.path.normcase(dest_dir) == os.path.normcase(folder):
            raise TemplateError("For copies, choose a different folder from the one the files are in.")
        os.makedirs(dest_dir, exist_ok=True)

    tracker = Tracker(tracker_dir)

    def final_names(fn_numbers):
        if renaming:
            return [f"{naming_prefix}{fn_numbers[i]}{os.path.splitext(file_names[i])[1]}" for i in range(n)]
        return list(file_names)

    def clashes(names):
        batch = {os.path.normcase(p) for p in sources} if not copies else set()
        out = []
        for nm in names:
            p = os.path.join(dest_dir, nm)
            if os.path.exists(p) and os.path.normcase(p) not in batch:
                out.append(nm)
        return out

    # Check for name clashes before taking any numbers
    stage("Checking file names...")
    fn_peek, _ = tracker.peek()
    planned = final_names([fn_peek - i for i in range(n)])
    if len({os.path.normcase(x) for x in planned}) != n:
        raise TemplateError("Two files would end up with the same name.")
    found = clashes(planned)
    if found:
        raise TemplateError(f"{len(found)} new file name(s) already exist in the destination folder, "
                            f"e.g. {found[0]}. Choose another prefix or folder.")

    # Reserve numbers (counters move past the whole batch before anything changes)
    stage("Reserving numbers from tracker...")
    fn_numbers, txn_base = tracker.reserve(n, per_file, file_numbers=renaming)
    names = final_names(fn_numbers)
    found = clashes(names)                      # numbers moved meanwhile? check again
    if found:
        raise TemplateError(f"{len(found)} new file name(s) already exist, e.g. {found[0]}.")

    # In-place renames go through temporary names first, so a new name can never land on
    # another file of the same batch that hasn't been processed yet.
    stage("Preparing files...")
    tasks, early_errors = [], []
    token = uuid.uuid4().hex[:8]
    for i, src in enumerate(sources):
        final = os.path.join(dest_dir, names[i])
        txn_start = txn_base - i * per_file
        if renaming and not copies:
            tmp_src = os.path.join(folder, f"{_TEMP_TAG}{token}_{i}{os.path.splitext(src)[1]}")
            try:
                os.replace(src, tmp_src)
            except Exception as e:
                early_errors.append({"index": i, "old": file_names[i], "new": names[i],
                                     "size_before": os.path.getsize(src), "size_after": None,
                                     "error": _friendly_error(e), "sheet": None})
                continue
            tasks.append((i, tmp_src, final, src, txn_start, True))
        else:
            tasks.append((i, src, final, src, txn_start, False))

    stage("Updating files..." if data else "Renaming files...")
    t1 = time.time()
    dtime, ddate = _dos_datetime()
    shared = {"data": data, "copies": copies, "updates": updates, "dtime": dtime, "ddate": ddate}
    results = list(early_errors)
    if progress_cb and results:
        progress_cb(len(results), n)
    cores = 1
    if data and parallel and len(tasks) > 3:
        try:
            cores = max(1, min(cpu_count(), len(tasks)))
            with Pool(processes=cores, initializer=_init_worker, initargs=(shared,)) as pool:
                for res in pool.imap_unordered(update_single_file, tasks, chunksize=1):
                    results.append(res)
                    if progress_cb:
                        progress_cb(len(results), n)
        except Exception:
            cores = 1
    done = {r["index"] for r in results}
    if any(t[0] not in done for t in tasks):
        _init_worker(shared)
        for t in tasks:
            if t[0] in done:
                continue
            results.append(update_single_file(t))
            if progress_cb:
                progress_cb(len(results), n)
    step_work = time.time() - t1

    results.sort(key=lambda r: r["index"])
    ok = [r for r in results if not r["error"]]
    failed = [r for r in results if r["error"]]
    diffs = [r["size_after"] - r["size_before"] for r in ok]
    sheets = sorted({r["sheet"] for r in ok if r["sheet"]})
    return {
        "mode": "update", "what": what, "save_mode": save_mode,
        "files": [os.path.join(dest_dir, r["new"]) for r in ok],
        "output_dir": dest_dir, "source_dir": folder,
        "updated": len(ok), "failed": [{"name": r["old"], "error": r["error"]} for r in failed],
        "renames": [[r["old"], r["new"]] for r in ok[:5]] if renaming else [],
        "size_before_total": sum(r["size_before"] for r in ok),
        "size_after_total": sum(r["size_after"] for r in ok),
        "size_diff_min": min(diffs) if diffs else 0, "size_diff_max": max(diffs) if diffs else 0,
        "exact_size": sum(1 for d in diffs if d == 0),
        "max_diff_pct": max((abs(d) / r["size_before"] * 100 for d, r in zip(diffs, ok)), default=0),
        "sheet_name": ", ".join(sheets) if sheets else None,
        "step1": 0.0, "step2": 0.0, "step3": step_work, "total": time.time() - t0, "cores": cores,
        "txn_range": [txn_base, txn_base - n * per_file + 1] if per_file else None,
        "fn_range": [fn_numbers[0], fn_numbers[-1]] if renaming else None,
        "first_file": ok[0]["new"] if ok else None, "last_file": ok[-1]["new"] if ok else None,
    }
