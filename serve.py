#!/usr/bin/env python3
"""
serve.py - local companion server for the Use Case Roadmap.

WHY this exists:
    The web app can persist edits straight to a CSV on disk instead of making
    you download a file each time. This script serves the app AND exposes a
    tiny read/write API for exactly one CSV file - the one you point it at.

SECURITY POSTURE (this is a local single-user tool; the defaults keep it that
way on purpose):
    * Binds to 127.0.0.1 only. It is NOT reachable from the network unless you
      explicitly pass --host 0.0.0.0, which prints a warning.
    * Writes ONLY the single file you specify on the command line. No path ever
      comes from the browser, so there is no path-traversal surface.
    * Rejects cross-origin POSTs (basic anti-CSRF for a localhost tool).
    * Caps request body size.
    * Saves are atomic (temp file + fsync + os.replace) and take a timestamped
      backup first, so a crash mid-write cannot corrupt your plan.

DEPENDENCIES: none. Python 3.7+ standard library only.

USAGE:
    python3 serve.py                          # serves ./UC_roadmap.csv
    python3 serve.py --file plan.csv          # choose the CSV to read/write
    python3 serve.py --file UC_dummy_data.csv --port 8080
    python3 serve.py --no-browser
"""
import argparse
import csv
import glob
import io
import json
import os
import sys
import tempfile
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

MAX_BODY = 8 * 1024 * 1024      # 8 MB cap on POST bodies
MAX_ROWS = 100_000
MAX_COLS = 200
KEEP_BACKUPS = 10

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_HTML = os.path.join(HERE, "capability-roadmap.html")


class Config:
    """Resolved runtime configuration (the only place a file path is set)."""
    def __init__(self, csv_path, html_path, host, port):
        self.csv_path = os.path.abspath(csv_path)
        self.html_path = os.path.abspath(html_path)
        self.host = host
        self.port = port


def read_csv(path):
    """Read the CSV into {'columns': [...], 'rows': [ {col: val}, ... ]}."""
    if not os.path.exists(path):
        return {"columns": [], "rows": []}
    # utf-8-sig transparently strips a leading BOM if one is present.
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        columns = list(reader.fieldnames or [])
        rows = [dict(r) for r in reader]
    return {"columns": columns, "rows": rows}


def _prune_backups(path):
    """Keep only the most recent KEEP_BACKUPS backups for this file."""
    backups = sorted(glob.glob(path + ".*.bak"))
    for old in backups[:-KEEP_BACKUPS]:
        try:
            os.remove(old)
        except OSError:
            pass


def write_csv(path, columns, rows):
    """
    Atomically write the CSV. Returns whether a backup was taken.
    A backup of the previous file is written first, then the new content is
    streamed to a temp file in the same directory, fsync'd, and os.replace'd
    into place (atomic on POSIX and Windows).
    """
    backup_made = False
    directory = os.path.dirname(path) or "."
    os.makedirs(directory, exist_ok=True)

    if os.path.exists(path):
        stamp = time.strftime("%Y%m%d-%H%M%S")
        backup = "{}.{}.bak".format(path, stamp)
        try:
            with open(path, "rb") as src, open(backup, "wb") as dst:
                dst.write(src.read())
            backup_made = True
            _prune_backups(path)
        except OSError:
            backup_made = False  # never block a save on backup failure

    # Build the CSV text in memory (BOM for Excel friendliness).
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=columns, extrasaction="ignore")
    writer.writeheader()
    for r in rows:
        writer.writerow({c: ("" if r.get(c) is None else r.get(c)) for c in columns})
    data = ("\ufeff" + buf.getvalue()).encode("utf-8")

    fd, tmp = tempfile.mkstemp(dir=directory, suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as tf:
            tf.write(data)
            tf.flush()
            os.fsync(tf.fileno())
        os.replace(tmp, path)   # atomic swap
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    return backup_made


def make_handler(cfg):
    class Handler(BaseHTTPRequestHandler):
        server_version = "RoadmapServer/1.0"

        # -- helpers ---------------------------------------------------------
        def _send(self, code, body=b"", ctype="application/json", extra=None):
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("X-Content-Type-Options", "nosniff")
            if ctype.startswith("application/json"):
                self.send_header("Cache-Control", "no-store")
            if extra:
                for k, v in extra.items():
                    self.send_header(k, v)
            self.end_headers()
            if body:
                self.wfile.write(body)

        def _json(self, code, obj):
            self._send(code, json.dumps(obj).encode("utf-8"))

        def _origin_ok(self):
            """Allow same-origin only. Absent Origin (curl, direct nav) is allowed."""
            origin = self.headers.get("Origin")
            if not origin:
                return True
            host = urlparse(origin).netloc
            allowed = {
                "127.0.0.1:%d" % cfg.port,
                "localhost:%d" % cfg.port,
            }
            if cfg.host not in ("127.0.0.1", "localhost"):
                allowed.add("%s:%d" % (cfg.host, cfg.port))
            return host in allowed

        def log_message(self, fmt, *args):
            sys.stderr.write("  %s - %s\n" % (self.address_string(), fmt % args))

        # -- routes ----------------------------------------------------------
        def do_GET(self):
            path = urlparse(self.path).path
            if path in ("/", "/index.html", "/capability-roadmap.html"):
                return self._serve_html()
            if path == "/api/health":
                return self._json(200, {
                    "ok": True,
                    "file": os.path.basename(cfg.csv_path),
                    "path": cfg.csv_path,
                    "exists": os.path.exists(cfg.csv_path),
                })
            if path == "/api/data":
                try:
                    return self._json(200, read_csv(cfg.csv_path))
                except Exception as e:  # noqa: BLE001 - report, don't crash
                    return self._json(500, {"ok": False, "error": str(e)})
            if path == "/favicon.ico":
                return self._send(204)
            return self._json(404, {"ok": False, "error": "not found"})

        def do_POST(self):
            path = urlparse(self.path).path
            if path != "/api/data":
                return self._json(404, {"ok": False, "error": "not found"})
            if not self._origin_ok():
                return self._json(403, {"ok": False, "error": "cross-origin request refused"})
            try:
                length = int(self.headers.get("Content-Length") or 0)
            except ValueError:
                return self._json(400, {"ok": False, "error": "bad length"})
            if length <= 0 or length > MAX_BODY:
                return self._json(413, {"ok": False, "error": "empty or oversized body"})

            raw = self.rfile.read(length)
            try:
                payload = json.loads(raw.decode("utf-8"))
            except Exception:
                return self._json(400, {"ok": False, "error": "invalid JSON"})

            columns = payload.get("columns")
            rows = payload.get("rows")
            if not isinstance(columns, list) or not isinstance(rows, list):
                return self._json(400, {"ok": False, "error": "columns and rows must be arrays"})
            if len(columns) > MAX_COLS or len(rows) > MAX_ROWS:
                return self._json(413, {"ok": False, "error": "too many columns or rows"})
            if not all(isinstance(c, str) for c in columns):
                return self._json(400, {"ok": False, "error": "column names must be strings"})
            # Coerce cells to strings; reject non-object rows.
            clean = []
            for r in rows:
                if not isinstance(r, dict):
                    return self._json(400, {"ok": False, "error": "each row must be an object"})
                clean.append({k: ("" if v is None else str(v)) for k, v in r.items()})

            try:
                backup = write_csv(cfg.csv_path, columns, clean)
            except Exception as e:  # noqa: BLE001
                return self._json(500, {"ok": False, "error": "write failed: %s" % e})

            return self._json(200, {
                "ok": True,
                "path": cfg.csv_path,
                "rows": len(clean),
                "backup": backup,
                "savedAt": time.strftime("%Y-%m-%d %H:%M:%S"),
            })

        def _serve_html(self):
            try:
                with open(cfg.html_path, "rb") as f:
                    body = f.read()
            except OSError:
                return self._send(500, b"capability-roadmap.html not found next to serve.py",
                                  ctype="text/plain")
            return self._send(200, body, ctype="text/html; charset=utf-8")

    return Handler


def main():
    ap = argparse.ArgumentParser(description="Local companion server for the Use Case Roadmap.")
    ap.add_argument("csv", nargs="?", default=None,
                    help="CSV file to read from and save to (positional). "
                         "Equivalent to --file; if both are given, this wins.")
    ap.add_argument("--file", default="UC_roadmap.csv",
                    help="CSV file to read from and save to (default: ./UC_roadmap.csv)")
    ap.add_argument("--html", default=DEFAULT_HTML,
                    help="Path to capability-roadmap.html (default: alongside serve.py)")
    ap.add_argument("--host", default="127.0.0.1", help="Bind host (default: 127.0.0.1)")
    ap.add_argument("--port", type=int, default=8000, help="Port (default: 8000)")
    ap.add_argument("--no-browser", action="store_true", help="Do not open a browser automatically")
    args = ap.parse_args()

    # Positional path takes precedence over --file so `serve.py data.csv` works.
    csv_path = args.csv if args.csv else args.file
    cfg = Config(csv_path, args.html, args.host, args.port)

    if not os.path.exists(cfg.html_path):
        sys.stderr.write("ERROR: cannot find the web app at:\n  %s\n"
                         "Keep serve.py and capability-roadmap.html in the same folder,\n"
                         "or pass --html /path/to/capability-roadmap.html\n" % cfg.html_path)
        sys.exit(1)

    if cfg.host not in ("127.0.0.1", "localhost"):
        sys.stderr.write(
            "\n  WARNING: binding to %s exposes this tool (and file writes) to your\n"
            "  network. Only do this on a trusted network. Ctrl+C to abort.\n\n" % cfg.host)

    httpd = ThreadingHTTPServer((cfg.host, cfg.port), make_handler(cfg))
    url = "http://%s:%d/" % ("localhost" if cfg.host == "127.0.0.1" else cfg.host, cfg.port)

    print("\n  Use Case Roadmap - local server")
    print("  --------------------------------")
    print("  Editing file : %s" % cfg.csv_path)
    print("  Exists       : %s" % ("yes" if os.path.exists(cfg.csv_path) else "no (will be created on first Save)"))
    print("  Open in browser: %s" % url)
    print("  Backups        : <file>.<timestamp>.bak (last %d kept)" % KEEP_BACKUPS)
    print("  Press Ctrl+C to stop.\n")

    if not args.no_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n  Stopped.")
        httpd.server_close()


if __name__ == "__main__":
    main()
