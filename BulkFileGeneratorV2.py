"""
BulkFileGenerator.py: starts the Bulk File Generator.

Double-click the .exe (or run `python BulkFileGenerator.py`): a small local
server starts and the app opens in your web browser. Nothing leaves this computer.

Folders created next to the .exe / script:
  config.json  - projects, tasks and extensions shown in the dropdowns (edit freely)
  Tracker/     - per-project counter files
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
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

import bulk_engine as be

APP_NAME = "Bulk File Generator"
REQUIRED_ENGINE_VERSION = 4


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

DEFAULT_CONFIG = {
    "projects": ["Project Alpha", "Project Beta", "Project Gamma"],
    "tasks": {
        "Create new bulk files": {"extensions": [".xlsx"]}
    },
}


def load_config():
    if not os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, indent=2)
        return DEFAULT_CONFIG
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        cfg.setdefault("projects", DEFAULT_CONFIG["projects"])
        cfg.setdefault("tasks", DEFAULT_CONFIG["tasks"])
        return cfg
    except Exception:
        return DEFAULT_CONFIG


def slug(text):
    return re.sub(r"[^A-Za-z0-9_]+", "_", str(text)).strip("_") or "default"


def tracker_dir_for(project):
    return os.path.join(TRACKER_ROOT, slug(project))


# ----------------------------------------------------------------- state
UPLOADS = {}   # token -> {"path", "name"}
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

    # routes
    def do_GET(self):
        url = urlparse(self.path)
        q = parse_qs(url.query)
        p = url.path

        if p in ("/", "/index.html"):
            with open(resource_path("ui.html"), "rb") as f:
                body = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

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
                "config_path": CONFIG_PATH,
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

    def do_POST(self):
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


def main():
    load_config()  # creates config.json on first run
    mismatch = engine_mismatch()
    if mismatch:
        print("WARNING: " + mismatch)
    port = int(os.environ.get("BULKGEN_PORT") or free_port())
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"{APP_NAME} is running at {url}")
    print("Keep this window open while you use the app. Close it (or click Quit in the app) to stop.")
    if not os.environ.get("BULKGEN_NO_BROWSER"):
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        shutil.rmtree(UPLOAD_DIR, ignore_errors=True)


if __name__ == "__main__":
    multiprocessing.freeze_support()   # required for multiprocessing inside a .exe
    main()
