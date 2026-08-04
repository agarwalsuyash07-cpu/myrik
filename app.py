#!/usr/bin/env python3
"""
Web frontend for the order-report scripts. Runs locally or on Vercel.

    python app.py            # serves http://127.0.0.1:8765
    python app.py --selftest # runs the multipart + args checks

Stdlib only. Uploaded files go to a temp dir; the chosen script runs against
them in-process (same as its CLI) and its stdout comes back to the page. Any
workbook the script writes comes back inline as a download link.

Serverless notes: nothing is kept between requests, so responses must be
self-contained (no download tokens), and /tmp is the only writable dir.
"""

import base64
import contextlib
import html
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
    def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path.split("?")[0] in ("/", "/index.html", "/api/index"):
            return self._send(200, PAGE.encode("utf-8"), "text/html; charset=utf-8")
        self._send(404, b'{"error":"not found"}')

    def do_POST(self):
        if self.path.split("?")[0] not in ("/run", "/api/index"):
            return self._send(404, b'{"error":"not found"}')
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


PAGE = """<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Order reports</title>
<style>
:root{--bg:#f7f7f5;--panel:#fff;--line:#dcdcd6;--fg:#1c1c1a;--dim:#5f5f58;
--accent:#0f766e;--accent-fg:#fff;--err:#a3312a;--r:6px}
@media(prefers-color-scheme:dark){:root{--bg:#16161a;--panel:#1e1e23;--line:#33333c;
--fg:#e9e9e4;--dim:#9a9a92;--accent:#2dd4bf;--accent-fg:#08201d;--err:#f0837a}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);
font:15px/1.5 ui-sans-serif,system-ui,"Segoe UI",sans-serif}
main{max-width:940px;margin:0 auto;padding:32px 20px 64px}
h1{font-size:20px;letter-spacing:-.01em;margin:0 0 2px}
.sub{color:var(--dim);font-size:13px;margin:0 0 24px}
.grid{display:grid;gap:20px;grid-template-columns:1fr;align-items:start}
@media(min-width:820px){.grid{grid-template-columns:340px 1fr}}
form{background:var(--panel);border:1px solid var(--line);border-radius:var(--r);padding:18px}
label{display:block;font-size:12px;font-weight:600;letter-spacing:.02em;
text-transform:uppercase;color:var(--dim);margin:0 0 6px}
.f{margin-bottom:16px}
select,input[type=date],input[type=file]{width:100%;padding:9px 10px;font:inherit;
font-size:14px;color:var(--fg);background:var(--bg);border:1px solid var(--line);border-radius:var(--r)}
input[type=file]{padding:7px 10px}
select:focus,input:focus{outline:2px solid var(--accent);outline-offset:1px;border-color:var(--accent)}
.chk{display:flex;gap:8px;align-items:center;font-size:14px;color:var(--fg);
text-transform:none;letter-spacing:0;font-weight:400}
.chk input{accent-color:var(--accent);width:16px;height:16px;margin:0}
.hint{font-size:12px;color:var(--dim);margin:6px 0 0}
button{width:100%;padding:11px;font:inherit;font-weight:600;cursor:pointer;
background:var(--accent);color:var(--accent-fg);border:0;border-radius:var(--r)}
button:hover{filter:brightness(1.08)}
button:active{transform:translateY(1px)}
button[disabled]{opacity:.55;cursor:progress;transform:none}
pre{margin:0;white-space:pre-wrap;word-break:break-word;
font:13px/1.6 ui-monospace,"Cascadia Mono",Consolas,monospace}
.dl{display:block;font-size:14px;color:var(--accent);font-weight:600}
.dl+.dl{margin-top:8px}
.box{background:var(--panel);border:1px solid var(--line);border-radius:var(--r);overflow:hidden}
.box+.box{margin-top:16px}
.box header{display:flex;align-items:center;justify-content:space-between;gap:12px;
padding:11px 14px;border-bottom:1px solid var(--line)}
h2{font-size:12px;text-transform:uppercase;letter-spacing:.04em;color:var(--dim);margin:0}
.box .body{max-height:46vh;overflow:auto;overscroll-behavior:contain}
.box .body>pre{padding:14px}
table{border-collapse:collapse;width:100%;font-variant-numeric:tabular-nums;
font:13px/1.5 ui-monospace,"Cascadia Mono",Consolas,monospace}
th{position:sticky;top:0;z-index:1;background:var(--panel);text-align:left;white-space:nowrap;
padding:9px 14px;font-size:11px;letter-spacing:.04em;text-transform:uppercase;
color:var(--dim);border-bottom:1px solid var(--line)}
td{padding:5px 14px;white-space:nowrap}
td:not(:last-child){padding-right:28px}
tbody tr:nth-child(even){background:color-mix(in srgb,var(--fg) 4%,transparent)}
tbody tr:hover{background:color-mix(in srgb,var(--accent) 14%,transparent)}
button.copy{flex:none;width:auto;padding:5px 12px;font-size:12px;
background:transparent;color:var(--accent);border:1px solid var(--line)}
button.copy:hover{border-color:var(--accent);filter:none}
button.copy:focus-visible{outline:2px solid var(--accent);outline-offset:1px}
.box.err h2,.box.err pre{color:var(--err)}
[hidden]{display:none}
</style></head><body><main>
<h1>Order reports</h1>
<p class="sub">Upload an Orders_Export workbook, pick a report, run it.</p>

<div class="grid">
<form id="f">
  <div class="f">
    <label for="script">Report</label>
    <select id="script" name="script">__OPTIONS__</select>
  </div>

  <div class="f">
    <label for="datafile">Export file (.xlsx)</label>
    <input type="file" id="datafile" name="datafile" accept=".xlsx,.xlsm" required>
  </div>

  <div data-for="citywise">
    <div class="f">
      <label for="date">Delivered on</label>
      <input type="date" id="date" name="date">
      <p class="hint">Only orders delivered on this date are counted.</p>
    </div>
    <div class="f"><label class="chk">
      <input type="checkbox" name="no_rikshaws" value="1"> Hide rikshaw counts
    </label></div>
  </div>

  <button id="go">Run report</button>
</form>

<div id="result"></div>
</div>

<script>
const f = document.getElementById('f'), sel = document.getElementById('script'),
      go = document.getElementById('go'), out = document.getElementById('result');

function sync(){ document.querySelectorAll('[data-for]').forEach(
  el => el.hidden = el.dataset.for !== sel.value ); }
sel.addEventListener('change', sync); sync();

async function copy(text, btn) {
  try { await navigator.clipboard.writeText(text); }
  catch {
    const ta = document.createElement('textarea');
    ta.value = text; ta.style.cssText = 'position:fixed;opacity:0';
    document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); } finally { ta.remove(); }
  }
  btn.textContent = 'Copied'; setTimeout(() => btn.textContent = 'Copy', 1200);
}

function table(lines) {
  const t = document.createElement('table');
  const thead = document.createElement('thead'), hr = document.createElement('tr');
  for (const c of lines[0].split('\\t')) {
    const th = document.createElement('th'); th.textContent = c; hr.appendChild(th);
  }
  thead.appendChild(hr);
  const tb = document.createElement('tbody');
  for (const line of lines.slice(1)) {
    const tr = document.createElement('tr');
    for (const c of line.split('\\t')) {
      const td = document.createElement('td'); td.textContent = c; tr.appendChild(td);
    }
    tb.appendChild(tr);
  }
  t.append(thead, tb);
  return t;
}

// one box per output section. Tab-separated bodies render as a table and copy
// as data rows only; anything else stays plain text and copies whole.
function box(title, body, opts) {
  opts = opts || {};
  const s = document.createElement('section');
  s.className = 'box' + (opts.err ? ' err' : '');
  const h = document.createElement('header'), t = document.createElement('h2');
  t.textContent = title; h.appendChild(t);

  const wrap = document.createElement('div'); wrap.className = 'body';
  let node, copyable = opts.copyable;
  if (opts.node) {
    node = opts.node;
  } else if (body.includes('\\t')) {
    const lines = body.split('\\n');
    node = table(lines);
    if (copyable === undefined) copyable = lines.slice(1).join('\\n');
  } else {
    node = document.createElement('pre'); node.textContent = body;
    if (copyable === undefined) copyable = body;
  }
  wrap.appendChild(node);

  if (copyable) {
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'copy'; b.textContent = 'Copy';
    b.addEventListener('click', () => copy(copyable, b));
    h.appendChild(b);
  }
  s.append(h, wrap);
  return s;
}

function render(d) {
  out.innerHTML = '';
  if (!d.ok) {
    out.appendChild(box('Error', d.stderr || d.stdout || 'Failed.', {err: true, copyable: ''}));
    return;
  }
  // scripts mark section breaks with a form feed; first line of each is its heading
  const sections = (d.stdout || '').split('\\f').map(s => s.trim()).filter(Boolean);
  if (!sections.length) {
    out.appendChild(box('Output', '(no output)', {copyable: ''}));
  } else if (sections.length === 1) {
    out.appendChild(box('Output', sections[0]));
  } else {
    for (const sec of sections) {
      const nl = sec.indexOf('\\n');
      out.appendChild(nl < 0 ? box(sec, '', {copyable: ''})
                             : box(sec.slice(0, nl), sec.slice(nl + 1)));
    }
  }
  if (d.stderr) out.appendChild(box('Warnings', d.stderr, {err: true, copyable: ''}));
  if (d.files.length) {
    const list = document.createElement('div');
    for (const file of d.files) {
      const a = document.createElement('a');
      a.className = 'dl'; a.href = file.url; a.download = file.name;
      a.textContent = 'Download ' + file.name; list.appendChild(a);
    }
    out.appendChild(box('Files', '', {node: list, copyable: ''}));
  }
}

f.addEventListener('submit', async e => {
  e.preventDefault();
  go.disabled = true; go.textContent = 'Running\\u2026';
  out.innerHTML = ''; out.appendChild(box('Output', 'Running\\u2026', {copyable: ''}));
  try {
    const r = await fetch('/run', {method:'POST', body:new FormData(f)});
    render(await r.json());
  } catch (err) {
    out.innerHTML = '';
    out.appendChild(box('Error', 'Server unreachable. Is app.py still running?',
                        {err: true, copyable: ''}));
  }
  go.disabled = false; go.textContent = 'Run report';
});

out.appendChild(box('Output', 'Nothing run yet.', {copyable: ''}));
</script>
</main></body></html>
"""

# dropdown is generated from SCRIPTS so renaming a label there is the only edit needed
PAGE = PAGE.replace("__OPTIONS__", "".join(
    f'<option value="{k}">{html.escape(label)}</option>' for k, (_f, label, _a) in SCRIPTS.items()))


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

    for key, (script, label, _fn) in SCRIPTS.items():
        assert os.path.isfile(os.path.join(HERE, script)), f"missing {script}"
        assert f'<option value="{key}">{html.escape(label)}</option>' in PAGE, f"{key} not in dropdown"
    for key in re.findall(r'data-for="([^"]+)"', PAGE):
        assert key in SCRIPTS, f'data-for="{key}" matches no script'

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
