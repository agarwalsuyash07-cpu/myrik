#!/usr/bin/env python3
"""
Web frontend for the order-report scripts. Runs locally or on Vercel.

    python app.py            # serves http://127.0.0.1:8765
    python app.py --selftest # runs the multipart + args + page checks

The page itself is index.html (served statically by Vercel, by this file
locally). Uploads go to a temp dir; the chosen script runs in-process exactly
as its CLI would and its stdout comes back to the page. Any workbook the
script writes comes back inline as a download link.

Serverless notes: nothing survives between requests, so responses must be
self-contained (no download tokens), and /tmp is the only writable dir.
"""

import base64
import contextlib
import http.server
import io
import json
import mimetypes
import os
import re
import runpy
import shutil
import sys
import tempfile
import threading
import traceback
import webbrowser

HERE = os.path.dirname(os.path.abspath(__file__))
PAGE = os.path.join(HERE, "index.html")
PORT = 8765

DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")


class Bad(Exception):
    """User-facing input error."""


# --- script registry: the CLI contract for each script -----------------------

def _iso_to_ddmmyyyy(iso):
    try:
        y, m, d = iso.split("-")
        return f"{d}-{m}-{y}"
    except ValueError:
        raise Bad("Pick a date.")


def _args_citywise(main, form):
    date = _iso_to_ddmmyyyy(form.get("date", ""))
    if not DATE_RE.match(date):
        raise Bad("Pick a date.")
    args = [main, date]
    if form.get("no_rikshaws"):
        args.append("--no-rikshaws")
    return args


def _args_cancelled(main, form):
    return [main]


def _args_pending(main, form):
    return [main]


def _args_split(main, form):
    return [main]  # no output path => script prints the rows instead of writing a workbook


SCRIPTS = {
    "citywise": ("citywise_deliveries.py", "Citywise deliveries", _args_citywise),
    "cancelled": ("cancelled_report.py", "Cancelled orders", _args_cancelled),
    "pending": ("pending_orders_report.py", "Total pending orders left", _args_pending),
    "split": ("split_orders_by_time.py", "Pending orders B2B sheet", _args_split),
}


# --- multipart ---------------------------------------------------------------

def parse_multipart(content_type, body):
    """Returns (fields: dict[str,str], files: dict[str,(filename, bytes)])."""
    if "boundary=" not in content_type:
        raise Bad("Malformed upload.")
    boundary = b"--" + content_type.split("boundary=")[1].strip('"; ').encode()
    fields, files = {}, {}
    for part in body.split(boundary)[1:-1]:
        head, _, data = part.partition(b"\r\n\r\n")
        if not head:
            continue
        data = data[:-2] if data.endswith(b"\r\n") else data
        disp = head.decode("utf-8", "replace")
        name = re.search(r'name="([^"]*)"', disp)
        if not name:
            continue
        fn = re.search(r'filename="([^"]*)"', disp)
        if fn is None:
            fields[name.group(1)] = data.decode("utf-8", "replace")
        elif fn.group(1):
            files[name.group(1)] = (os.path.basename(fn.group(1)), data)
    return fields, files


# --- running -----------------------------------------------------------------

def save(workdir, upload, fallback):
    name, data = upload
    if not name.lower().endswith((".xlsx", ".xlsm")):
        raise Bad(f"{name}: expected an .xlsx export.")
    path = os.path.join(workdir, re.sub(r"[^\w.\- ]", "_", name) or fallback)
    with open(path, "wb") as f:
        f.write(data)
    return path


_LOCK = threading.Lock()  # ponytail: argv/cwd/stdout are process-global; one run at a time


def run_cli(script, argv, workdir):
    """Run a report script in-process exactly as `python script.py argv...` would."""
    out, err = io.StringIO(), io.StringIO()
    with _LOCK:
        argv0, cwd0 = sys.argv, os.getcwd()
        sys.argv = [script] + argv
        os.chdir(workdir)
        try:
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                runpy.run_path(os.path.join(HERE, script), run_name="__main__")
            code = 0
        except SystemExit as e:
            code = e.code if isinstance(e.code, int) else (0 if e.code is None else 1)
            if isinstance(e.code, str):
                err.write(e.code)
        except Exception:
            traceback.print_exc(file=err)
            code = 1
        finally:
            sys.argv = argv0
            os.chdir(cwd0)
    return code, out.getvalue(), err.getvalue()


def as_download(workdir, name):
    """Inline the generated file; no server state survives between requests."""
    with open(os.path.join(workdir, name), "rb") as f:
        blob = base64.b64encode(f.read()).decode()
    ctype = mimetypes.guess_type(name)[0] or "application/octet-stream"
    return {"name": name, "url": f"data:{ctype};base64,{blob}"}


def run_script(fields, files):
    key = fields.get("script", "")
    if key not in SCRIPTS:
        raise Bad("Unknown report.")
    if "datafile" not in files:
        raise Bad("Upload the export file first.")

    script, _, build_args = SCRIPTS[key]
    workdir = tempfile.mkdtemp(prefix="myrik_")  # /tmp on Vercel, the only writable dir
    try:
        main = save(workdir, files["datafile"], "input.xlsx")

        before = set(os.listdir(workdir))
        code, stdout, stderr = run_cli(script, build_args(main, fields), workdir)
        new = sorted(set(os.listdir(workdir)) - before)

        return {
            "ok": code == 0,
            "code": code,
            "stdout": stdout.strip(),
            "stderr": stderr.strip(),
            "files": [as_download(workdir, n) for n in new],
        }
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


# --- server ------------------------------------------------------------------

class Handler(http.server.BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    # one page, one action, and Vercel rewrites the path before the function
    # sees it - so don't route on self.path, just answer by method.
    def do_GET(self):
        with open(PAGE, "rb") as f:
            self._send(200, f.read(), "text/html; charset=utf-8")

    def do_POST(self):
        try:
            body = self.rfile.read(int(self.headers.get("Content-Length", 0)))
            fields, files = parse_multipart(self.headers.get("Content-Type", ""), body)
            result = run_script(fields, files)
        except Bad as e:
            result = {"ok": False, "stdout": "", "stderr": str(e), "files": [], "code": None}
        except Exception as e:
            result = {"ok": False, "stdout": "", "stderr": f"{type(e).__name__}: {e}", "files": [], "code": None}
        self._send(200, json.dumps(result).encode("utf-8"))

    def log_message(self, fmt, *args):
        pass


# --- checks ------------------------------------------------------------------

def selftest():
    b = "----x9"
    blob = b"PK\x03\x04\r\n--not-the-boundary\x00\xff"
    body = (
        f"--{b}\r\nContent-Disposition: form-data; name=\"script\"\r\n\r\ncitywise\r\n"
        f"--{b}\r\nContent-Disposition: form-data; name=\"datafile\"; filename=\"a b.xlsx\"\r\n"
        f"Content-Type: application/octet-stream\r\n\r\n"
    ).encode() + blob + f"\r\n--{b}--\r\n".encode()

    fields, files = parse_multipart(f"multipart/form-data; boundary={b}", body)
    assert fields == {"script": "citywise"}, fields
    assert files["datafile"] == ("a b.xlsx", blob), files["datafile"]

    fields, files = parse_multipart(
        f"multipart/form-data; boundary={b}",
        f"--{b}\r\nContent-Disposition: form-data; name=\"f\"; filename=\"\"\r\n\r\n\r\n--{b}--\r\n".encode())
    assert files == {}, "empty file input must be ignored"

    assert _args_citywise("x.xlsx", {"date": "2026-08-02"}) == ["x.xlsx", "02-08-2026"]
    assert _args_citywise("x.xlsx", {"date": "2026-08-02", "no_rikshaws": "1"})[-1] == "--no-rikshaws"
    assert _args_pending("p.xlsx", {}) == ["p.xlsx"]
    assert _args_split(os.path.join("t", "in.xlsx"), {}) == [os.path.join("t", "in.xlsx")]

    # the dropdown is hand-written HTML now, so check it against the registry
    with open(PAGE, encoding="utf-8") as f:
        page = f.read()
    options = dict(re.findall(r'<option value="([^"]+)">([^<]*)</option>', page))
    assert options == {k: label for k, (_f, label, _a) in SCRIPTS.items()}, options
    for key in re.findall(r'data-for="([^"]+)"', page):
        assert key in SCRIPTS, f'data-for="{key}" matches no script'
    for _key, (script, _label, _fn) in SCRIPTS.items():
        assert os.path.isfile(os.path.join(HERE, script)), f"missing {script}"

    for bad in ("2026-8-2", "", "not-a-date"):
        try:
            _args_citywise("x.xlsx", {"date": bad}); assert False, bad
        except Bad:
            pass

    # in-process runner behaves like the CLI: usage + exit 1, argv/cwd restored
    workdir = tempfile.mkdtemp(prefix="myrik_test_")
    try:
        argv0, cwd0 = list(sys.argv), os.getcwd()
        code, out, err = run_cli("citywise_deliveries.py", [], workdir)
        assert code == 1, code
        assert "Usage:" in out, (out, err)
        assert sys.argv == argv0 and os.getcwd() == cwd0, "runner leaked argv/cwd"

        with open(os.path.join(workdir, "x.xlsx"), "wb") as f:
            f.write(b"hello")
        assert as_download(workdir, "x.xlsx")["url"].endswith(",aGVsbG8=")
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    print("selftest ok")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
        sys.exit(0)
    url = f"http://127.0.0.1:{PORT}/"
    print(f"Serving {url}  (Ctrl+C to stop)")
    webbrowser.open(url)
    http.server.ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
