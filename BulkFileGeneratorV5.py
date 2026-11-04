"""
BulkFileGenerator.py: starts the Bulk File Generator.

Double-click the .exe (or run `python BulkFileGenerator.py`): a small local
server starts and the app opens in its own window (Microsoft Edge or Chrome in
"app" mode: no address bar, its own taskbar button and icon). Closing the window
stops the app, after any batch that is still running has finished. Without Edge or
Chrome it opens in the default browser and stops a few minutes after the tab is closed.
Nothing leaves this computer.

The project list (config.json) and the page (ui.html) are built INTO the .exe by
build_exe.bat (see make_bundle.py), so people using the .exe see no .py, .json or
.html files - not even in Windows' temporary folder while the app runs.
To change the projects: edit config.json in the build folder and build again.

Folders created next to the .exe:
  Tracker/     - per-project counter files (the only files users see)
  Output/      - default place for generated files
"""
import json
import multiprocessing
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import uuid
import webbrowser
import zipfile
import base64
import logging
from logging.handlers import RotatingFileHandler
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

import bulk_engine as be

try:
    # _bundled.py only exists while build_exe.bat runs (make_bundle.py writes it, then it is packed into
    # the .exe and deleted), so code editors can't find it. That's expected: the line below tells them so.
    import _bundled  # type: ignore  # noqa: F401  (project list + page, compiled into the .exe)
except ImportError:      # running from source: read config.json and ui.html from this folder
    _bundled = None

APP_NAME = "Bulk File Generator"
REQUIRED_ENGINE_VERSION = 6


def engine_mismatch():
    """bulk_engine.py, BulkFileGenerator.py and ui.html must come from the same version."""
    found = getattr(be, "ENGINE_VERSION", None)
    if found != REQUIRED_ENGINE_VERSION:
        return (f"The app's files don't match: bulk_engine.py is version {found or 'old'}, "
                f"BulkFileGenerator.py needs version {REQUIRED_ENGINE_VERSION}. Replace bulk_engine.py, "
                f"BulkFileGenerator.py and ui.html with the same latest copies, then run build_exe.bat again.")
    return None
MAX_UPLOAD_BYTES = 100 * 1024 * 1024


# ----------------------------------------------------------------- paths
def app_dir():
    """Folder of the .exe when frozen, else folder of this script."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(name):
    """Bundled read-only files (ui.html) live in _MEIPASS when frozen."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


BASE_DIR = app_dir()
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
TRACKER_ROOT = os.path.join(BASE_DIR, "Tracker")
OUTPUT_ROOT = os.path.join(BASE_DIR, "Output")
UPLOAD_DIR = tempfile.mkdtemp(prefix="bulkgen_uploads_")

UPDATE_TASK = "Update file names / column data"
BUILT_IN_TASKS = {
    "Create new bulk files": {"extensions": [".xlsx"], "kind": "create"},
    UPDATE_TASK: {"extensions": [".xlsx"], "kind": "update"},
}
DEFAULT_CONFIG = {
    "projects": ["Project Alpha", "Project Beta", "Project Gamma"],
    "tasks": BUILT_IN_TASKS,
}


def _complete_config(cfg):
    """Fills in defaults and any built-in task the config doesn't list yet.
    Returns (cfg, whether a built-in task had to be added)."""
    cfg.setdefault("projects", DEFAULT_CONFIG["projects"])
    tasks = cfg.setdefault("tasks", {})
    added = False
    for name, spec in BUILT_IN_TASKS.items():
        if name not in tasks:
            tasks[name] = dict(spec)
            added = True
    for name, spec in tasks.items():
        spec.setdefault("kind", "update" if name == UPDATE_TASK else "create")
        spec.setdefault("extensions", [".xlsx"])
    return cfg, added


def load_config():
    """Projects and tasks for the dropdowns. In the .exe they come from inside the .exe
    (nothing to see or edit next to it); from source, config.json in this folder is used."""
    if _bundled is not None:
        return _complete_config(json.loads(json.dumps(_bundled.CONFIG)))[0]
    if not os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, indent=2)
        return json.loads(json.dumps(DEFAULT_CONFIG))
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception:
        return json.loads(json.dumps(DEFAULT_CONFIG))
    cfg, added = _complete_config(cfg)
    if added:
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
        except OSError:
            pass
    return cfg


def slug(text):
    return re.sub(r"[^A-Za-z0-9_]+", "_", str(text)).strip("_") or "default"


def tracker_dir_for(project):
    return os.path.join(TRACKER_ROOT, slug(project))


# ----------------------------------------------------------------- state
UPLOADS = {}   # token -> {"path", "name"}
SCANS = {}     # token -> {"folder", "names"}  (folders searched for the update task)
JOBS = {}      # job id -> dict
TRACKER_LOCK = threading.Lock()   # one batch at a time touches the counters


def open_folder(path):
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa: S606 (Windows only)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def run_job(job_id, payload):
    job = JOBS[job_id]

    def on_progress(done, total):
        job["done"], job["total"] = done, total

    def on_stage(text):
        job["stage"] = text

    try:
        with TRACKER_LOCK:
            job["status"] = "running"
            if payload.get("kind") == "update":
                summary = be.run_update(
                    folder=payload["folder"],
                    file_names=payload["names"],
                    what=payload["what"],
                    naming_prefix=payload["naming_prefix"],
                    updates=payload["updates"],
                    tracker_dir=payload["tracker_dir"],
                    save_mode=payload["save_mode"],
                    output_dir=payload["output_dir"],
                    parallel=payload["parallel"],
                    progress_cb=on_progress,
                    stage_cb=on_stage,
                )
            else:
                summary = be.run_batch(
                    template_path=payload["template_path"],
                    n_files=payload["n_files"],
                    naming_prefix=payload["naming_prefix"],
                    updates=payload["updates"],
                    output_dir=payload["output_dir"],
                    tracker_dir=payload["tracker_dir"],
                    parallel=payload["parallel"],
                    progress_cb=on_progress,
                    stage_cb=on_stage,
                )
        job["files"] = summary.pop("files")
        job["summary"] = summary
        job["status"] = "done"
        job["stage"] = "Done"
    except Exception as e:  # report everything back to the page
        job["status"] = "error"
        if isinstance(e, (be.TrackerError, be.TemplateError)):
            job["error"] = str(e)          # plain-language message, shown as-is
        else:
            job["error"] = f"{type(e).__name__}: {e}"
        job["trace"] = traceback.format_exc()
        LOG.error("batch %s failed: %s", job_id, job["trace"])


# ----------------------------------------------------------------- HTTP
class Handler(BaseHTTPRequestHandler):
    server_version = "BulkFileGenerator/1.0"

    def log_message(self, fmt, *args):  # keep the console quiet
        pass

    # helpers
    def _json(self, obj, status=200):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(length) or b"{}")

    def _safely(self, route):
        """Runs a request; any unexpected error is logged and reported to the page as a
        message, instead of the connection just dropping ("Failed to fetch")."""
        try:
            route()
        except (BrokenPipeError, ConnectionResetError):
            pass                                  # the page went away mid-answer; nothing to report
        except Exception as e:
            LOG.error("request %s %s failed: %s", self.command, self.path, traceback.format_exc())
            try:
                self._json({"error": f"Something went wrong in the app ({type(e).__name__}: {e}). "
                                     f"Details are in {LOG_PATH}"}, 500)
            except Exception:
                pass

    def do_GET(self):
        self._safely(self._get)

    def do_POST(self):
        self._safely(self._post)

    # routes
    def _get(self):
        url = urlparse(self.path)
        q = parse_qs(url.query)
        p = url.path

        if p in ("/", "/index.html"):
            if _bundled is not None:
                body = _bundled.UI_HTML.encode("utf-8")
            else:
                with open(resource_path("ui.html"), "rb") as f:
                    body = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if p == "/favicon.ico":
            body = app_icon()
            if not body:
                return self._json({"error": "no icon"}, 404)
            self.send_response(200)
            self.send_header("Content-Type", "image/x-icon")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "max-age=86400")
            self.end_headers()
            self.wfile.write(body)
            return

        if p == "/api/ping":
            global LAST_PING
            LAST_PING = time.time()
            return self._json({"ok": True})

        if p == "/api/config":
            cfg = load_config()
            return self._json({
                "projects": cfg["projects"],
                "tasks": cfg["tasks"],
                "limits": {
                    "prefix_max": be.MAX_PREFIX_LENGTH,
                    "txn_min": be.TXN_PREFIX_MIN,
                    "txn_max": be.TXN_PREFIX_MAX,
                    "lowest": be.LOWEST_ALLOWED_NUMBER,
                },
                "config_path": None if _bundled is not None else CONFIG_PATH,
            })

        if p == "/api/tracker":
            project = (q.get("project") or [""])[0]
            if not project:
                return self._json({"error": "project required"}, 400)
            tdir = tracker_dir_for(project)
            try:
                fn_next, txn_next = be.Tracker(tdir).peek()
            except be.TrackerError as e:
                return self._json({"error": str(e)}, 500)
            return self._json({
                "fn_next": fn_next, "txn_next": txn_next, "tracker_dir": tdir,
                "default_output": os.path.join(OUTPUT_ROOT, slug(project)),
            })

        if p.startswith("/api/job/"):
            job = JOBS.get(p.rsplit("/", 1)[-1])
            if not job:
                return self._json({"error": "unknown job"}, 404)
            public = {k: v for k, v in job.items() if k not in ("files", "trace")}
            public["file_count"] = len(job.get("files") or [])
            return self._json(public)

        if p.startswith("/api/download/"):
            job = JOBS.get(p.rsplit("/", 1)[-1])
            if not job or job.get("status") != "done":
                return self._json({"error": "nothing to download"}, 404)
            zip_path = os.path.join(UPLOAD_DIR, f"{job['id']}.zip")
            if not os.path.exists(zip_path):
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fp in job["files"]:
                        zf.write(fp, arcname=os.path.basename(fp))
            size = os.path.getsize(zip_path)
            self.send_response(200)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Length", str(size))
            self.send_header("Content-Disposition",
                             f'attachment; filename="{job["zip_name"]}"')
            self.end_headers()
            with open(zip_path, "rb") as f:
                shutil.copyfileobj(f, self.wfile)
            return

        self._json({"error": "not found"}, 404)

    def _post(self):
        p = urlparse(self.path).path

        if p == "/api/upload":
            mismatch = engine_mismatch()
            if mismatch:
                return self._json({"error": mismatch}, 500)
            length = int(self.headers.get("Content-Length") or 0)
            name = os.path.basename(self.headers.get("X-Filename", "template.xlsx"))
            if length <= 0:
                return self._json({"error": "Empty file."}, 400)
            if length > MAX_UPLOAD_BYTES:
                return self._json({"error": "File is larger than 100 MB."}, 400)
            token = uuid.uuid4().hex
            folder = os.path.join(UPLOAD_DIR, token)
            os.makedirs(folder)
            path = os.path.join(folder, name)
            with open(path, "wb") as f:
                remaining = length
                while remaining:
                    chunk = self.rfile.read(min(1 << 20, remaining))
                    if not chunk:
                        break
                    f.write(chunk)
                    remaining -= len(chunk)
            try:
                rows, dims, sheet_name = be.preview_template(path)
                be.load_template_package(path)   # make sure the exact-copy engine can use it
            except be.TemplateError as e:
                return self._json({"error": str(e)}, 400)
            except Exception as e:
                return self._json({"error": f"Could not read this file as Excel: {e}"}, 400)
            UPLOADS[token] = {"path": path, "name": name}
            return self._json({"token": token, "name": name, "rows": rows, "sheet": sheet_name,
                               "max_row": dims[0], "max_col": dims[1]})

        if p == "/api/generate":
            data = self._read_json()
            errors = []
            upload = UPLOADS.get(data.get("token"))
            if not upload:
                errors.append("Upload the template again (session expired).")
            project = data.get("project") or ""
            if project not in load_config()["projects"]:
                errors.append("Select a valid project.")
            try:
                n_files = int(data.get("n_files"))
                if n_files < 1:
                    raise ValueError
            except (TypeError, ValueError):
                errors.append("Total files must be a positive whole number.")
                n_files = 0
            prefix = str(data.get("naming_prefix") or "").strip()
            err = be.validate_name(prefix)
            if err:
                errors.append(f"Naming prefix: {err}")
            updates = data.get("updates") or []
            errors += be.validate_updates(updates)
            output_dir = str(data.get("output_dir") or "").strip()
            if not output_dir:
                errors.append("Output folder is required.")
            if not errors:
                # no negative numbers: check there is room left in the counters
                try:
                    fn_next, txn_next = be.Tracker(tracker_dir_for(project)).peek()
                    errors += be.check_capacity(fn_next, txn_next, n_files, be.txn_ids_per_file(updates))
                except be.TrackerError as e:
                    errors.append(str(e))
            if errors:
                return self._json({"errors": errors}, 400)

            job_id = uuid.uuid4().hex[:12]
            JOBS[job_id] = {
                "id": job_id, "status": "queued", "stage": "Waiting...",
                "done": 0, "total": n_files,
                "zip_name": f"{slug(project)}_{slug(prefix)}_bulk_files.zip",
            }
            payload = {
                "template_path": upload["path"], "n_files": n_files,
                "naming_prefix": prefix, "updates": updates,
                "output_dir": os.path.abspath(os.path.expanduser(output_dir)),
                "tracker_dir": tracker_dir_for(project),
                "parallel": bool(data.get("parallel", True)),
            }
            threading.Thread(target=run_job, args=(job_id, payload), daemon=True).start()
            return self._json({"job": job_id})

        if p == "/api/scan-folder":
            mismatch = engine_mismatch()
            if mismatch:
                return self._json({"error": mismatch}, 500)
            data = self._read_json()
            ext = str(data.get("extension") or ".xlsx")
            try:
                scan = be.scan_folder(str(data.get("folder") or ""), ext)
            except be.TemplateError as e:
                return self._json({"error": str(e)}, 400)
            except OSError as e:
                return self._json({"error": f"Can't read that folder: {e}"}, 400)
            files = scan["files"]
            good = [f for f in files if not f.get("error")]
            bad = [f for f in files if f.get("error")]
            preview = None
            if good:
                try:
                    rows, dims, sheet = be.preview_template(os.path.join(scan["folder"], good[0]["name"]))
                    preview = {"name": good[0]["name"], "rows": rows, "sheet": sheet,
                               "max_row": dims[0], "max_col": dims[1]}
                except Exception:
                    preview = None
            token = uuid.uuid4().hex
            SCANS[token] = {"folder": scan["folder"], "names": [f["name"] for f in good]}
            row_counts = [f["rows"] for f in good if f.get("rows")]
            return self._json({
                "token": token, "folder": scan["folder"], "count": len(good),
                "total_size": sum(f["size"] for f in good),
                "first": [f["name"] for f in good[:3]], "last": good[-1]["name"] if good else None,
                "rows_min": min(row_counts) if row_counts else None,
                "rows_max": max(row_counts) if row_counts else None,
                "sheets": sorted({f.get("sheet") for f in good if f.get("sheet")}),
                "skipped": [f["name"] for f in bad][:5], "skipped_count": len(bad),
                "preview": preview,
            })

        if p == "/api/update":
            data = self._read_json()
            errors = []
            scan = SCANS.get(data.get("token"))
            if not scan:
                errors.append("Search the folder again (the list of files has expired).")
            project = data.get("project") or ""
            if project not in load_config()["projects"]:
                errors.append("Select a valid project.")
            what = data.get("what")
            if what not in (be.WHAT_NAMES, be.WHAT_DATA, be.WHAT_BOTH):
                errors.append("Choose what to update: file names, data, or both.")
            renaming = what in (be.WHAT_NAMES, be.WHAT_BOTH)
            changing_data = what in (be.WHAT_DATA, be.WHAT_BOTH)
            prefix = str(data.get("naming_prefix") or "").strip()
            if renaming:
                err = be.validate_name(prefix)
                if err:
                    errors.append(f"Naming prefix: {err}")
            updates = (data.get("updates") or []) if changing_data else []
            if changing_data:
                if not updates:
                    errors.append("Add at least one column to update.")
                errors += be.validate_updates(updates)
            save_mode = data.get("save_mode")
            if save_mode not in (be.SAVE_IN_PLACE, be.SAVE_COPIES):
                errors.append("Choose where to save the updated files.")
            output_dir = str(data.get("output_dir") or "").strip()
            if save_mode == be.SAVE_COPIES and not output_dir:
                errors.append("Choose the folder for the updated copies.")
            n = len(scan["names"]) if scan else 0
            if scan and not n:
                errors.append("That folder has no files to update.")
            if not errors:
                try:
                    fn_next, txn_next = be.Tracker(tracker_dir_for(project)).peek()
                    per_file = be.txn_ids_per_file(updates) if changing_data else 0
                    errors += be.check_capacity(fn_next if renaming else 10 ** 12, txn_next, n, per_file)
                except be.TrackerError as e:
                    errors.append(str(e))
            if errors:
                return self._json({"errors": errors}, 400)

            job_id = uuid.uuid4().hex[:12]
            JOBS[job_id] = {
                "id": job_id, "status": "queued", "stage": "Waiting...", "done": 0, "total": n,
                "zip_name": f"{slug(project)}_updated_files.zip",
            }
            payload = {
                "kind": "update", "folder": scan["folder"], "names": list(scan["names"]),
                "what": what, "naming_prefix": prefix, "updates": updates,
                "tracker_dir": tracker_dir_for(project), "save_mode": save_mode,
                "output_dir": os.path.abspath(os.path.expanduser(output_dir)) if output_dir else None,
                "parallel": bool(data.get("parallel", True)),
            }
            threading.Thread(target=run_job, args=(job_id, payload), daemon=True).start()
            return self._json({"job": job_id})

        if p == "/api/open-folder":
            data = self._read_json()
            job = JOBS.get(data.get("job"))
            if not job or job.get("status") != "done":
                return self._json({"error": "unknown job"}, 404)
            try:
                open_folder(job["summary"]["output_dir"])
                return self._json({"ok": True})
            except Exception as e:
                return self._json({"error": str(e)}, 500)

        if p == "/api/bye":                # sent by the page as its window/tab closes
            global BYE_AT
            BYE_AT = time.time()
            return self._json({"ok": True})

        if p == "/api/shutdown":
            self._json({"ok": True})
            threading.Thread(target=self.server.shutdown, daemon=True).start()
            return

        self._json({"error": "not found"}, 404)


def free_port(preferred=8765):
    for port in [preferred] + list(range(preferred + 1, preferred + 50)):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return 0


def _setup_log():
    """A small log in the user's app-data folder (no window shows errors in the .exe)."""
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), ".local", "share")
    folder = os.path.join(base, "BulkFileGenerator")
    log = logging.getLogger("bulkgen")
    log.setLevel(logging.INFO)
    try:
        os.makedirs(folder, exist_ok=True)
        h = RotatingFileHandler(os.path.join(folder, "app.log"), maxBytes=500_000, backupCount=1, encoding="utf-8")
        h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        log.addHandler(h)
    except OSError:
        log.addHandler(logging.NullHandler())
    return log, os.path.join(folder, "app.log")


LOG, LOG_PATH = _setup_log()

LAST_PING = None          # when the page last checked in
BYE_AT = None             # when the page said it was closing
IDLE_EXIT_SECONDS = int(os.environ.get("BULKGEN_IDLE_EXIT") or 900)   # stop 15 min after the last check-in


def app_icon():
    """The app icon (.ico bytes): from inside the .exe, or app.ico next to the script."""
    if _bundled is not None and getattr(_bundled, "FAVICON_B64", None):
        return base64.b64decode(_bundled.FAVICON_B64)
    try:
        with open(resource_path("app.ico"), "rb") as f:
            return f.read()
    except OSError:
        return None


def find_app_browser():
    """Microsoft Edge (on every Windows 10/11 PC) or Google Chrome, for a window without
    an address bar. None if neither is installed."""
    candidates = []
    if sys.platform.startswith("win"):
        for env in ("PROGRAMFILES(X86)", "PROGRAMFILES", "LOCALAPPDATA"):
            base = os.environ.get(env)
            if base:
                candidates += [os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe"),
                               os.path.join(base, "Google", "Chrome", "Application", "chrome.exe")]
    else:
        for name in ("microsoft-edge", "msedge", "google-chrome", "chromium", "chromium-browser"):
            found = shutil.which(name)
            if found:
                candidates.append(found)
    return next((c for c in candidates if os.path.isfile(c)), None)


def window_profile_dir():
    """A small private browser profile just for the app window, so it is its own program on
    the taskbar and closing it can be noticed. Kept in the user's app-data folder."""
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), ".local", "share")
    path = os.path.join(base, "BulkFileGenerator", "window")
    os.makedirs(path, exist_ok=True)
    return path


def open_app_window(url):
    """Opens the app in its own window. Returns the window's process, or None."""
    exe = find_app_browser()
    if not exe:
        return None
    args = [exe, f"--app={url}", f"--user-data-dir={window_profile_dir()}",
            "--no-first-run", "--no-default-browser-check", "--disable-sync",
            "--window-size=1180,920", "--disable-features=Translate"]
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        return subprocess.Popen(args, creationflags=flags,
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except OSError:
        return None


def batch_running():
    return any(j.get("status") in ("queued", "running") for j in JOBS.values())


def wait_then_stop(server, window):
    """Stops the app once its window (or tab) is really gone, but never in the middle of a batch:
      - the page said goodbye as its window closed and didn't come back within 10 s (a reload
        comes back at once); or
      - the page hasn't checked in for IDLE_EXIT_SECONDS (e.g. the PC slept with the tab closed).
    The browser process itself is NOT watched: Edge may hand the window to another process or
    keep running in the background, so it says nothing reliable about the window."""
    while True:
        time.sleep(1)
        now = time.time()
        last_ping = LAST_PING or started_at
        reason = None
        if BYE_AT is not None and BYE_AT >= last_ping and now - BYE_AT > 10:
            reason = "window closed"
        elif now - last_ping > IDLE_EXIT_SECONDS:
            reason = f"no contact from the window for {IDLE_EXIT_SECONDS // 60} minutes"
        if reason and not batch_running():
            LOG.info("stopping: %s", reason)
            server.shutdown()
            return


started_at = time.time()


def main():
    load_config()
    mismatch = engine_mismatch()
    if mismatch:
        print("WARNING: " + mismatch)
    port = int(os.environ.get("BULKGEN_PORT") or free_port())
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"{APP_NAME} is running at {url}")
    LOG.info("started %s at %s (engine %s)", APP_NAME, url, getattr(be, "ENGINE_VERSION", "?"))
    if not os.environ.get("BULKGEN_NO_BROWSER"):
        window = open_app_window(url)
        LOG.info("opened in %s", "an app window" if window is not None else "the default browser")
        if window is None:                     # no Edge/Chrome: use the default browser
            threading.Timer(0.8, lambda: webbrowser.open(url)).start()
        threading.Thread(target=wait_then_stop, args=(server, window), daemon=True).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        LOG.info("stopped")
        server.server_close()
        shutil.rmtree(UPLOAD_DIR, ignore_errors=True)


if __name__ == "__main__":
    multiprocessing.freeze_support()   # required for multiprocessing inside a .exe
    main()
