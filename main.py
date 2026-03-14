import html
import os
import textwrap
from typing import Optional

from fastapi import FastAPI
from fastapi.responses import FileResponse, HTMLResponse, PlainTextResponse
from pydantic import BaseModel, Field

import win32print  # pip install pywin32


# -----------------------
# Settings
# -----------------------
PORT = int(os.environ.get("ZPL_SERVER_PORT", "8787"))

# Set this to your exact Windows printer name.
# Example: "ZDesigner ZD411-203dpi ZPL"
DEFAULT_PRINTER = os.environ.get("ZPL_PRINTER_NAME", "")

DEFAULT_WARNING = """THCa PRODUCT

HEMP-DERIVED PRODUCT-CONTAINS LESS THAN 0.3% DELTA-9 THC

21+ ONLY KEEP OUT OF REACH OF CHILDREN

THIS PRODUCT MAY CAUSE INTOXICATION WHEN HEATED

DO NOT USE WHILE DRIVING OR OPERATING HEAVY MACHINERY

CONSULT A PHYSICIAN BEFORE USE"""

# 2x1 @ 203dpi -> 406x203 dots
LABEL_WIDTH_DOTS = 406
LABEL_HEIGHT_DOTS = 203
LABEL_Y_OFFSET = 6


# -----------------------
# ZPL builder
# -----------------------
def zpl_escape(s: str) -> str:
    if s is None:
        return ""
    return str(s).replace("^", "").replace("~", "").strip()


def format_warning_lines(s: str, max_chars: int = 38, max_lines: int = 9) -> str:
    if s is None:
        return ""
    raw = str(s).replace("\r\n", "\n").replace("\r", "\n")
    paragraphs = [p.strip() for p in raw.split("\n") if p.strip()]
    lines: list[str] = []
    for paragraph in paragraphs:
        wrapped = textwrap.wrap(
            zpl_escape(paragraph),
            width=max_chars,
            break_long_words=True,
            break_on_hyphens=True,
        )
        lines.extend(wrapped or [""])
        if len(lines) >= max_lines:
            break

    lines = lines[:max_lines]
    return r"\&".join(lines)


def build_zpl_2x1_centered(
    name: str,
    price: str,
    warning: str,
    include_warning: bool,
    darkness: int = 20,
    vertical_offset: int = 0,
) -> str:
    """
    Centered 2x1 label tuned for product + price + warning.
    """
    name = zpl_escape(name)
    price = zpl_escape(price)
    warning = format_warning_lines(warning)
    # Positive values move content up; negative values move content down.
    y_offset = LABEL_Y_OFFSET - vertical_offset

    z = []
    z += ["^XA"]
    z += [f"^PW{LABEL_WIDTH_DOTS}"]
    z += [f"^LL{LABEL_HEIGHT_DOTS}"]
    z += [f"^MD{darkness}"]

    # Name
    z += [f"^FO8,{8 + y_offset}"]
    z += [f"^FB{LABEL_WIDTH_DOTS-16},2,2,C,0"]
    z += ["^A0N,20,20"]
    z += [f"^FD{name}^FS"]

    # Price
    z += [f"^FO8,{56 + y_offset}"]
    z += [f"^FB{LABEL_WIDTH_DOTS-16},1,0,C,0"]
    z += ["^A0N,22,22"]
    z += [f"^FD{price}^FS"]

    # Warning (optional)
    if include_warning and warning:
        z += [f"^FO10,{86 + y_offset}"]
        z += [f"^GB{LABEL_WIDTH_DOTS-20},1,1^FS"]
        z += [f"^FO10,{94 + y_offset}"]
        z += [f"^FB{LABEL_WIDTH_DOTS-20},9,1,C,0"]
        z += ["^A0N,10,10"]
        z += [f"^FD{warning}^FS"]

    z += ["^XZ"]
    return "\n".join(z) + "\n"


def send_raw_zpl(printer_name: str, zpl: str) -> None:
    h = win32print.OpenPrinter(printer_name)
    try:
        job = win32print.StartDocPrinter(h, 1, ("ZPL Mobile Label", "", "RAW"))
        win32print.StartPagePrinter(h)
        win32print.WritePrinter(h, zpl.encode("ascii", errors="ignore"))
        win32print.EndPagePrinter(h)
        win32print.EndDocPrinter(h)
    finally:
        win32print.ClosePrinter(h)


def list_printers() -> list[str]:
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    return [p[2] for p in win32print.EnumPrinters(flags)]


# -----------------------
# FastAPI
# -----------------------
app = FastAPI(title="ZPL Mobile Print Server")


class PrintJob(BaseModel):
    printer: Optional[str] = Field(default=None, description="Windows printer name (optional)")
    name: str = Field(default="", description="Top text")
    price: str = Field(default="", description="Bottom text")
    warning: str = Field(default=DEFAULT_WARNING, description="Health warning block")
    include_warning: bool = Field(default=True)
    copies: int = Field(default=1, ge=1, le=200)
    darkness: int = Field(default=20, ge=0, le=30)
    vertical_offset: int = Field(default=0, ge=-60, le=60, description="Shift label content in dots: positive up, negative down")


@app.get("/printers", response_class=PlainTextResponse)
def printers():
    return "\n".join(list_printers())


@app.get("/", response_class=HTMLResponse)
def root():
    # Serve the embedded mobile page (keeps beginner setup simple: one script file)
    return MOBILE_HTML


@app.get("/favicon.ico")
def favicon():
    return FileResponse("favicon.ico")


@app.get("/apple-touch-icon.png")
def apple_touch_icon():
    return FileResponse("favicon.ico")


@app.get("/apple-touch-icon-precomposed.png")
def apple_touch_icon_precomposed():
    return FileResponse("favicon.ico")


@app.post("/zpl", response_class=PlainTextResponse)
def make_zpl(job: PrintJob):
    zpl = build_zpl_2x1_centered(
        name=job.name,
        price=job.price,
        warning=job.warning,
        include_warning=job.include_warning,
        darkness=job.darkness,
        vertical_offset=job.vertical_offset,
    )
    return zpl


@app.post("/print", response_class=PlainTextResponse)
def print_label(job: PrintJob):
    printer = (job.printer or DEFAULT_PRINTER or "").strip()
    if not printer:
        return PlainTextResponse(
            "Printer not set. Set ZPL_PRINTER_NAME env var or choose a printer in the UI.",
            status_code=400,
        )

    zpl = build_zpl_2x1_centered(
        name=job.name,
        price=job.price,
        warning=job.warning,
        include_warning=job.include_warning,
        darkness=job.darkness,
        vertical_offset=job.vertical_offset,
    )

    for _ in range(job.copies):
        send_raw_zpl(printer, zpl)

    return f"Printed {job.copies} copy/copies to: {printer}"


# -----------------------
# Mobile-friendly HTML (served at "/")
# -----------------------
MOBILE_HTML = r"""
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Zebra Mobile Label Print</title>
  <link rel="icon" href="/favicon.ico" sizes="any" />
  <link rel="apple-touch-icon" href="/apple-touch-icon.png" />
  <style>
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif; margin: 16px; }
    .card { max-width: 720px; margin: 0 auto; padding: 16px; border: 1px solid #ddd; border-radius: 14px; }
    label { font-weight: 600; display: block; margin-top: 14px; }
    input, textarea, select { width: 100%; font-size: 18px; padding: 12px; border-radius: 12px; border: 1px solid #ccc; }
    textarea { min-height: 90px; }
    .row { display: flex; gap: 12px; }
    .row > div { flex: 1; }
    .btnrow { display: flex; gap: 12px; margin-top: 16px; }
    button {
      font-size: 20px;
      padding: 14px;
      border-radius: 14px;
      border: 0;
      cursor: pointer;
    }
    .btnrow button { flex: 1; }
    #printBtn { background: #111; color: #fff; }
    #zplBtn { background: #f2f2f2; }
    #status { margin-top: 12px; white-space: pre-wrap; }
    .small { font-size: 13px; color: #666; margin-top: 6px; }
    .toggle { display: flex; align-items: center; gap: 10px; margin-top: 10px; }
    .toggle input { width: auto; transform: scale(1.3); }
    .offset-control { display: flex; align-items: center; gap: 10px; margin-top: 8px; }
    .offset-control input { width: 110px; text-align: center; font-weight: 700; }
    .offset-btn { flex: 0 0 auto; font-size: 16px; padding: 10px 14px; background: #f2f2f2; border: 1px solid #ccc; }
  </style>
</head>
<body>
  <div class="card">
    <h2 style="margin-top:0">Zebra ZD411 — Mobile Label Print (2×1)</h2>

    <label>Printer</label>
    <select id="printer"></select>
    <div class="small">If blank, go to /printers on this server to see what Windows calls your printer.</div>

    <label>Item name (top)</label>
    <input id="name" placeholder="e.g. Pre-Roll - Cherry Pie" />

    <label>Price / note</label>
    <input id="price" placeholder="e.g. $5.00" />

    <label>Health warning (optional)</label>
    <textarea id="warning" placeholder="Paste your required warning here...">__DEFAULT_WARNING__</textarea>

    <div class="toggle">
      <input type="checkbox" id="include_warning" checked />
      <label for="include_warning" style="margin:0; font-weight:600;">Include warning on label</label>
    </div>

    <div class="row">
      <div>
        <label>Copies</label>
        <input id="copies" type="number" min="1" max="200" value="1" />
      </div>
      <div>
        <label>Darkness</label>
        <input id="darkness" type="number" min="0" max="30" value="20" />
      </div>
    </div>

    <label>Vertical label offset</label>
    <div class="offset-control">
      <button id="offsetUpBtn" class="offset-btn" type="button">Up</button>
      <input id="vertical_offset" type="number" min="-60" max="60" value="0" readonly />
      <button id="offsetDownBtn" class="offset-btn" type="button">Down</button>
    </div>
    <div class="small">Use positive values to move print up, negative values to move print down. 1 step = 1 dot.</div>

    <div class="btnrow">
      <button id="zplBtn" type="button">Generate ZPL</button>
      <button id="printBtn" type="button">PRINT</button>
    </div>

    <pre id="status"></pre>
  </div>

<script>
async function loadPrinters() {
  const sel = document.getElementById('printer');
  sel.innerHTML = '';
  try {
    const res = await fetch('/printers');
    const text = await res.text();
    const lines = text.split('\n').map(s => s.trim()).filter(Boolean);
    if (lines.length === 0) {
      const opt = document.createElement('option');
      opt.value = '';
      opt.textContent = '(No printers found)';
      sel.appendChild(opt);
      return;
    }
    for (const p of lines) {
      const opt = document.createElement('option');
      opt.value = p;
      opt.textContent = p;
      sel.appendChild(opt);
    }
  } catch (e) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = '(Could not load printers)';
    sel.appendChild(opt);
  }
}

function jobPayload() {
  return {
    printer: document.getElementById('printer').value || null,
    name: document.getElementById('name').value,
    price: document.getElementById('price').value,
    warning: document.getElementById('warning').value,
    include_warning: document.getElementById('include_warning').checked,
    copies: parseInt(document.getElementById('copies').value || '1', 10),
    darkness: parseInt(document.getElementById('darkness').value || '20', 10),
    vertical_offset: parseInt(document.getElementById('vertical_offset').value || '0', 10),
  };
}

async function generateZPL() {
  const status = document.getElementById('status');
  status.textContent = 'Generating ZPL...';
  const res = await fetch('/zpl', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(jobPayload())
  });
  const zpl = await res.text();

  // Download as .zpl
  const blob = new Blob([zpl], {type: 'text/plain'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const ts = new Date().toISOString().replaceAll(':','').slice(0,15);
  a.href = url;
  a.download = `label_${ts}.zpl`;
  a.click();
  URL.revokeObjectURL(url);

  status.textContent = 'ZPL downloaded.';
}

async function printLabel() {
  const status = document.getElementById('status');
  status.textContent = 'Printing...';
  const res = await fetch('/print', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(jobPayload())
  });
  const text = await res.text();
  status.textContent = text;
}

document.getElementById('zplBtn').addEventListener('click', generateZPL);
document.getElementById('printBtn').addEventListener('click', printLabel);
const verticalOffsetInput = document.getElementById('vertical_offset');

function clampOffset(value) {
  return Math.max(-60, Math.min(60, value));
}

function nudgeOffset(delta) {
  const current = parseInt(verticalOffsetInput.value || '0', 10);
  verticalOffsetInput.value = String(clampOffset(current + delta));
}

document.getElementById('offsetUpBtn').addEventListener('click', () => nudgeOffset(1));
document.getElementById('offsetDownBtn').addEventListener('click', () => nudgeOffset(-1));

loadPrinters();
</script>
</body>
</html>
""".replace("__DEFAULT_WARNING__", html.escape(DEFAULT_WARNING))


if __name__ == "__main__":
    # Run: python zpl_print_server.py
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=PORT)
