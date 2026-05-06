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
    * { box-sizing: border-box; }
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif; margin: 16px; background: #f7f7f5; color: #1f1f1f; }
    .card { max-width: 760px; margin: 0 auto; padding: 16px; border: 1px solid #ddd; border-radius: 14px; background: #fff; }
    label { font-weight: 600; display: block; margin-top: 14px; }
    input, textarea, select { width: 100%; font-size: 18px; padding: 12px; border-radius: 12px; border: 1px solid #ccc; }
    textarea { min-height: 90px; }
    .row { display: flex; gap: 12px; }
    .row > div { flex: 1; }
    .row.three > div { flex: 1 1 33%; }
    .panel { margin-top: 16px; padding: 12px; border: 1px solid #ddd; border-radius: 12px; background: #fbfbfb; }
    .panel-header, .history-header { margin-bottom: 10px; display: flex; justify-content: space-between; align-items: center; gap: 10px; }
    .panel-title, .history-title { margin: 0; font-size: 18px; }
    .quick-toolbar { display: flex; gap: 10px; align-items: center; }
    .quick-toolbar input { flex: 1; min-width: 0; }
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
    #savePinnedBtn { background: #dceee3; color: #164123; }
    #status { margin-top: 12px; white-space: pre-wrap; }
    .small { font-size: 13px; color: #666; margin-top: 6px; }
    .toggle { display: flex; align-items: center; gap: 10px; margin-top: 10px; }
    .toggle input { width: auto; transform: scale(1.3); }
    .offset-control { display: flex; align-items: center; gap: 10px; margin-top: 8px; }
    .offset-control input { width: 110px; text-align: center; font-weight: 700; }
    .offset-btn { flex: 0 0 auto; font-size: 16px; padding: 10px 14px; background: #f2f2f2; border: 1px solid #ccc; }
    .clear-history-btn, .small-action-btn { flex: 0 0 auto; font-size: 14px; padding: 8px 12px; border-radius: 10px; border: 1px solid #ccc; background: #f2f2f2; }
    #catalogList, #historyList { display: flex; flex-direction: column; gap: 10px; }
    .history-header { margin-top: 22px; }
    .history-item, .catalog-item { border: 1px solid #ddd; border-radius: 12px; padding: 12px; background: #fafafa; }
    .history-main { font-weight: 700; font-size: 16px; }
    .history-meta { margin-top: 4px; font-size: 13px; color: #555; line-height: 1.4; }
    .history-actions, .quick-actions { display: flex; gap: 8px; margin-top: 10px; }
    .history-actions button { flex: 1; font-size: 16px; padding: 10px 12px; border-radius: 12px; border: 0; }
    .quick-actions button { flex: 1; font-size: 17px; padding: 11px 10px; border-radius: 12px; border: 0; }
    .hist-load { background: #f2f2f2; }
    .hist-reprint { background: #111; color: #fff; }
    .hist-delete { background: #ffdede; color: #8b0000; }
    .quick-print { background: #111; color: #fff; }
    .quick-load { background: #f2f2f2; }
    .empty-state { padding: 8px 0; }
    @media (max-width: 560px) {
      body { margin: 8px; }
      .card { padding: 12px; }
      .row, .btnrow, .quick-toolbar { flex-direction: column; }
      .history-actions, .quick-actions { flex-wrap: wrap; }
      .history-actions button, .quick-actions button { flex: 1 1 30%; }
    }
  </style>
</head>
<body>
  <div class="card">
    <h2 style="margin-top:0">Zebra ZD411 — Mobile Label Print (2×1)</h2>

    <div class="panel">
      <div class="panel-header">
        <h3 class="panel-title">Pinned Labels</h3>
        <button id="savePinnedBtn" class="small-action-btn" type="button">Save current</button>
      </div>
      <div class="quick-toolbar">
        <input id="catalogSearch" placeholder="Search pinned labels..." />
        <button id="clearSearchBtn" class="small-action-btn" type="button">Clear</button>
      </div>
      <div class="small">Tap a saved label to load it, or print 1 / 5 / 10 copies directly.</div>
      <div id="catalogList" style="margin-top:10px;"></div>
    </div>

    <label>Printer</label>
    <select id="printer"></select>
    <div class="small">If blank, go to /printers on this server to see what Windows calls your printer.</div>

    <label>Strain / item name (top)</label>
    <input id="name" placeholder="e.g. Cherry Pie" autocomplete="off" />

    <div class="row three">
      <div>
        <label>Size</label>
        <select id="size">
          <option value="">No size</option>
          <option value="1 gram">1 gram</option>
          <option value="2 gram">2 gram</option>
          <option value="3 gram">3 gram</option>
          <option value="1/8 oz">1/8 oz</option>
          <option value="1/4 oz">1/4 oz</option>
          <option value="1 oz">1 oz</option>
        </select>
      </div>
      <div>
        <label>Type</label>
        <select id="strain_type">
          <option value="">None</option>
          <option value="Indica">Indica</option>
          <option value="Sativa">Sativa</option>
          <option value="Hybrid">Hybrid</option>
          <option value="Indica Leaning Hybrid">Indica Leaning Hybrid</option>
          <option value="Sativa Leaning Hybrid">Sativa Leaning Hybrid</option>
        </select>
      </div>
      <div>
        <label>Price preset</label>
        <select id="price_preset">
          <option value="">Custom</option>
          <option value="$5.00">$5.00</option>
          <option value="$10.00">$10.00</option>
          <option value="$18.00">$18.00</option>
          <option value="$25.00">$25.00</option>
          <option value="$32.50">$32.50</option>
          <option value="$50.00">$50.00</option>
        </select>
      </div>
    </div>

    <label>Custom price / note</label>
    <input id="price" placeholder="e.g. $10.00 or 2 for $15" />
    <div class="small">Bottom line prints size, type, and price, for example: 3.5g Hybrid - $25.00.</div>

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

    <div class="history-header">
      <h3 class="history-title">Recent Labels</h3>
      <button id="clearHistoryBtn" class="clear-history-btn" type="button">Clear all history</button>
    </div>
    <div class="quick-toolbar">
      <input id="historySearch" placeholder="Search previous printed labels..." />
      <button id="clearHistorySearchBtn" class="small-action-btn" type="button">Clear</button>
    </div>
    <div id="historyList"></div>

    <pre id="status"></pre>
  </div>

<script>
const HISTORY_KEY = 'zebra_label_history_v1';
const CATALOG_KEY = 'zebra_label_catalog_v1';
const DEFAULTS_KEY = 'zebra_label_defaults_v1';
const STRAIN_MEMORY_KEY = 'zebra_strain_type_memory_v1';
const MAX_HISTORY = 30;
const HISTORY_SCHEMA_VERSION = 1;
const CATALOG_SCHEMA_VERSION = 1;
const STRAIN_MEMORY_SCHEMA_VERSION = 1;
const LIMITS = {
  darkness: { min: 0, max: 30 },
  vertical_offset: { min: -60, max: 60 },
  copies: { min: 1, max: 200 },
};
const QUICK_COPY_COUNTS = [1, 5, 10];
const SIZE_VALUE_MAP = {
  '1 gram': '1 gram',
  '2 gram': '2 gram',
  '3 gram': '3 gram',
  '3.5 grams': '1/8 oz',
  '7 grams': '1/4 oz',
  '1 oz': '1 oz',
  Eighth: '1/8 oz',
  Quarter: '1/4 oz',
  '1/8 oz': '1/8 oz',
  '1/4 oz': '1/4 oz',
};
const SIZE_SHORT_LABELS = {
  '1 gram': '1g',
  '2 gram': '2g',
  '3 gram': '3g',
  '3.5 grams': '1/8 oz',
  '7 grams': '1/4 oz',
  '1 oz': '1 oz',
  Eighth: '1/8 oz',
  Quarter: '1/4 oz',
  '1/8 oz': '1/8 oz',
  '1/4 oz': '1/4 oz',
};
const STRAIN_TYPES = ['Indica', 'Sativa', 'Hybrid', 'Indica Leaning Hybrid', 'Sativa Leaning Hybrid'];

function setStatus(message) {
  document.getElementById('status').textContent = message;
}

function toInt(value, fallback) {
  const parsed = parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function clampNumber(value, range, fallback) {
  const parsed = toInt(value, fallback);
  return Math.max(range.min, Math.min(range.max, parsed));
}

function cleanText(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function validStrainType(value) {
  return STRAIN_TYPES.includes(value) ? value : '';
}

function normalizeSize(value) {
  const clean = cleanText(value);
  return SIZE_VALUE_MAP[clean] || clean;
}

function shortSizeLabel(value) {
  const clean = normalizeSize(value);
  return SIZE_SHORT_LABELS[clean] || clean;
}

function composeBottomLine(size, strainType, priceText) {
  const left = [shortSizeLabel(size), validStrainType(strainType)].filter(Boolean).join(' ');
  const price = cleanText(priceText);
  if (left && price) {
    return `${left} - ${price}`;
  }
  return left || price;
}

function hasStructuredFields(source) {
  return !!(
    cleanText(source.size)
    || validStrainType(cleanText(source.strain_type))
    || cleanText(source.price_preset)
    || cleanText(source.price_input)
    || cleanText(source.price_amount)
  );
}

function readJsonArray(key, schemaVersion, resetMessage) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }
    return parsed.filter((entry) =>
      entry
      && typeof entry === 'object'
      && entry.v === schemaVersion
      && entry.job
      && typeof entry.job === 'object'
    );
  } catch (e) {
    localStorage.removeItem(key);
    setStatus(resetMessage);
    return [];
  }
}

function normalizeJob(job, copyFallback = 1) {
  const source = job || {};
  const structured = hasStructuredFields(source);
  const size = normalizeSize(source.size);
  const strainType = validStrainType(cleanText(source.strain_type));
  const pricePreset = cleanText(source.price_preset);
  const priceInput = cleanText(source.price_input ?? source.price_amount ?? pricePreset);
  const legacyPrice = String(source.price || '');
  const printablePrice = structured
    ? composeBottomLine(size, strainType, priceInput || pricePreset || legacyPrice)
    : legacyPrice;
  return {
    printer: source.printer || null,
    name: String(source.name || ''),
    price: printablePrice,
    size,
    strain_type: strainType,
    price_preset: pricePreset,
    price_input: priceInput,
    warning: String(source.warning || ''),
    include_warning: source.include_warning !== false,
    copies: clampNumber(source.copies ?? copyFallback, LIMITS.copies, copyFallback),
    darkness: clampNumber(source.darkness ?? 20, LIMITS.darkness, 20),
    vertical_offset: clampNumber(source.vertical_offset ?? 0, LIMITS.vertical_offset, 0),
  };
}

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
      restoreDefaultsToForm();
      return;
    }
    for (const p of lines) {
      const opt = document.createElement('option');
      opt.value = p;
      opt.textContent = p;
      sel.appendChild(opt);
    }
    restoreDefaultsToForm();
  } catch (e) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = '(Could not load printers)';
    sel.appendChild(opt);
    restoreDefaultsToForm();
  }
}

function jobPayload() {
  return normalizeJob({
    printer: document.getElementById('printer').value || null,
    name: document.getElementById('name').value,
    size: document.getElementById('size').value,
    strain_type: document.getElementById('strain_type').value,
    price_preset: document.getElementById('price_preset').value,
    price_input: document.getElementById('price').value,
    warning: document.getElementById('warning').value,
    include_warning: document.getElementById('include_warning').checked,
    copies: document.getElementById('copies').value,
    darkness: document.getElementById('darkness').value,
    vertical_offset: document.getElementById('vertical_offset').value,
  });
}

function readHistory() {
  return readJsonArray(HISTORY_KEY, HISTORY_SCHEMA_VERSION, 'Saved label history was corrupted and has been reset.');
}

function writeHistory(list) {
  localStorage.setItem(HISTORY_KEY, JSON.stringify(list));
  renderHistory();
}

function readCatalog() {
  return readJsonArray(CATALOG_KEY, CATALOG_SCHEMA_VERSION, 'Saved pinned labels were corrupted and have been reset.');
}

function writeCatalog(list) {
  localStorage.setItem(CATALOG_KEY, JSON.stringify(list));
  renderCatalog();
}

function saveDefaults(job) {
  const defaults = normalizeJob(job || jobPayload());
  localStorage.setItem(DEFAULTS_KEY, JSON.stringify({
    printer: defaults.printer || '',
    size: defaults.size,
    price_preset: defaults.price_preset,
    price_input: defaults.price_input,
    warning: defaults.warning,
    include_warning: defaults.include_warning,
    darkness: defaults.darkness,
    vertical_offset: defaults.vertical_offset,
    copies: defaults.copies,
  }));
}

function restoreDefaultsToForm() {
  try {
    const raw = localStorage.getItem(DEFAULTS_KEY);
    if (!raw) {
      return;
    }
    const defaults = JSON.parse(raw);
    if (!defaults || typeof defaults !== 'object') {
      return;
    }
    const current = jobPayload();
    applyJobToForm({
      ...current,
      size: defaults.size ?? current.size,
      price_preset: defaults.price_preset ?? current.price_preset,
      price_input: defaults.price_input ?? current.price_input,
      warning: defaults.warning ?? current.warning,
      include_warning: defaults.include_warning !== false,
      copies: defaults.copies ?? current.copies,
      darkness: defaults.darkness ?? current.darkness,
      vertical_offset: defaults.vertical_offset ?? current.vertical_offset,
    }, defaults.printer || '');
  } catch (e) {
    localStorage.removeItem(DEFAULTS_KEY);
    setStatus('Saved defaults were corrupted and have been reset.');
  }
}

function jobFingerprint(job) {
  return JSON.stringify({
    name: job.name ?? '',
    price: job.price ?? '',
    size: job.size ?? '',
    strain_type: job.strain_type ?? '',
    price_preset: job.price_preset ?? '',
    price_input: job.price_input ?? '',
    warning: job.warning ?? '',
    include_warning: !!job.include_warning,
    darkness: Number(job.darkness ?? 0),
    vertical_offset: Number(job.vertical_offset ?? 0),
  });
}

function formatTimestamp(value) {
  if (!value) return 'Unknown time';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return value;
  return d.toLocaleString();
}

function abbreviatedWarning(text) {
  const clean = String(text || '').replace(/\s+/g, ' ').trim();
  if (!clean) return 'No warning';
  return clean.length > 40 ? `${clean.slice(0, 40)}…` : clean;
}

function normalizeStrainKey(name) {
  return cleanText(name).toLowerCase();
}

function readStrainMemory() {
  try {
    const raw = localStorage.getItem(STRAIN_MEMORY_KEY);
    if (!raw) {
      return {};
    }
    const parsed = JSON.parse(raw);
    if (!parsed || parsed.v !== STRAIN_MEMORY_SCHEMA_VERSION || typeof parsed.map !== 'object') {
      return {};
    }
    return parsed.map || {};
  } catch (e) {
    localStorage.removeItem(STRAIN_MEMORY_KEY);
    setStatus('Saved strain type memory was corrupted and has been reset.');
    return {};
  }
}

function writeStrainMemory(map) {
  localStorage.setItem(STRAIN_MEMORY_KEY, JSON.stringify({
    v: STRAIN_MEMORY_SCHEMA_VERSION,
    map,
  }));
}

function rememberStrainType(job) {
  const normalized = normalizeJob(job);
  const key = normalizeStrainKey(normalized.name);
  if (!key || !normalized.strain_type) {
    return;
  }
  const map = readStrainMemory();
  map[key] = normalized.strain_type;
  writeStrainMemory(map);
}

function autofillStrainType() {
  const select = document.getElementById('strain_type');
  if (select.value) {
    return;
  }
  const key = normalizeStrainKey(document.getElementById('name').value);
  if (!key) {
    return;
  }
  const remembered = readStrainMemory()[key];
  if (remembered && STRAIN_TYPES.includes(remembered)) {
    select.value = remembered;
  }
}

function jobSearchText(job, entry) {
  const normalized = normalizeJob(job || {});
  return [
    normalized.name,
    normalized.price,
    normalized.size,
    normalized.strain_type,
    normalized.price_input,
    normalized.warning,
    entry && entry.printer,
    entry && formatTimestamp(entry.ts),
  ].filter(Boolean).join(' ').toLowerCase();
}

function catalogMatchesSearch(entry, query) {
  if (!query) {
    return true;
  }
  return jobSearchText(entry.job || {}, entry).includes(query.toLowerCase());
}

function historyMatchesSearch(entry, query) {
  if (!query) {
    return true;
  }
  return jobSearchText(entry.job || {}, entry).includes(query.toLowerCase());
}

function applyJobToForm(job, printer) {
  const normalized = normalizeJob(job);

  document.getElementById('name').value = normalized.name;
  document.getElementById('size').value = normalized.size;
  document.getElementById('strain_type').value = normalized.strain_type;
  document.getElementById('price_preset').value = normalized.price_preset;
  document.getElementById('price').value = normalized.price_input || normalized.price;
  document.getElementById('warning').value = normalized.warning;
  document.getElementById('include_warning').checked = normalized.include_warning;
  document.getElementById('copies').value = String(normalized.copies);
  document.getElementById('darkness').value = String(normalized.darkness);
  document.getElementById('vertical_offset').value = String(normalized.vertical_offset);
  if (printer) {
    const select = document.getElementById('printer');
    const hasOption = Array.from(select.options).some((opt) => opt.value === printer);
    if (hasOption) {
      select.value = printer;
    }
  }
  saveDefaults(jobPayload());
}

function recordHistory(job) {
  const normalized = normalizeJob(job);
  const fingerprint = jobFingerprint(normalized);
  const printer = normalized.printer || '';
  const history = readHistory();
  const deduped = history.filter((entry) => {
    if (!entry || typeof entry !== 'object') return false;
    const entryFingerprint = jobFingerprint(entry.job || {});
    return !(entryFingerprint === fingerprint && (entry.printer || '') === printer);
  });
  const record = {
    v: HISTORY_SCHEMA_VERSION,
    id: `${Date.now()}_${Math.random().toString(36).slice(2, 10)}`,
    ts: new Date().toISOString(),
    job: normalized,
    printer,
  };
  deduped.unshift(record);
  writeHistory(deduped.slice(0, MAX_HISTORY));
  rememberStrainType(normalized);
}

async function printJob(job, copies, message) {
  setStatus(message || `Printing ${copies} copy/copies...`);
  const payload = {
    ...normalizeJob(job, copies),
    printer: job.printer || document.getElementById('printer').value || null,
    copies: clampNumber(copies, LIMITS.copies, 1),
  };
  const res = await fetch('/print', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(payload)
  });
  const text = await res.text();
  setStatus(text);
  if (res.ok) {
    recordHistory(payload);
    saveDefaults(payload);
  }
  return res.ok;
}

function saveCurrentAsPinned() {
  const job = normalizeJob(jobPayload());
  if (!job.name.trim() && !job.price.trim()) {
    setStatus('Enter an item name or price before saving a pinned label.');
    return;
  }
  const fingerprint = jobFingerprint(job);
  const existing = readCatalog();
  const deduped = existing.filter((entry) => jobFingerprint(entry.job || {}) !== fingerprint);
  const record = {
    v: CATALOG_SCHEMA_VERSION,
    id: `${Date.now()}_${Math.random().toString(36).slice(2, 10)}`,
    ts: new Date().toISOString(),
    job: {
      ...job,
      printer: null,
    },
  };
  deduped.unshift(record);
  writeCatalog(deduped);
  rememberStrainType(job);
  saveDefaults(job);
  setStatus('Pinned label saved.');
}

function deletePinned(entry) {
  const updated = readCatalog().filter((row) => row && row.id !== entry.id);
  writeCatalog(updated);
  setStatus('Pinned label deleted.');
}

function renderCatalog() {
  const container = document.getElementById('catalogList');
  const query = document.getElementById('catalogSearch').value.trim();
  const catalog = readCatalog().filter((entry) => catalogMatchesSearch(entry, query));
  container.innerHTML = '';

  if (catalog.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'small empty-state';
    empty.textContent = query ? 'No pinned labels match that search.' : 'No pinned labels yet.';
    container.appendChild(empty);
    return;
  }

  for (const entry of catalog) {
    if (!entry || typeof entry !== 'object') continue;
    const job = normalizeJob(entry.job || {});

    const item = document.createElement('div');
    item.className = 'catalog-item';

    const main = document.createElement('div');
    main.className = 'history-main';
    main.textContent = `${job.name || '(Unnamed item)'} — ${job.price || '(No price)'}`;

    const meta = document.createElement('div');
    meta.className = 'history-meta';
    meta.textContent = `Bottom: ${job.price || 'No bottom line'} | Warn: ${abbreviatedWarning(job.warning)} | Saved: ${formatTimestamp(entry.ts)}`;

    const quickActions = document.createElement('div');
    quickActions.className = 'quick-actions';

    const loadBtn = document.createElement('button');
    loadBtn.type = 'button';
    loadBtn.className = 'quick-load';
    loadBtn.textContent = 'Load';
    loadBtn.addEventListener('click', () => {
      applyJobToForm(job, '');
      setStatus('Pinned label loaded.');
    });
    quickActions.appendChild(loadBtn);

    for (const count of QUICK_COPY_COUNTS) {
      const printBtn = document.createElement('button');
      printBtn.type = 'button';
      printBtn.className = 'quick-print';
      printBtn.textContent = `Print ${count}`;
      printBtn.addEventListener('click', () => printJob(job, count, `Printing ${count} copy/copies from pinned label...`));
      quickActions.appendChild(printBtn);
    }

    const deleteBtn = document.createElement('button');
    deleteBtn.type = 'button';
    deleteBtn.className = 'hist-delete';
    deleteBtn.textContent = 'Delete';
    deleteBtn.addEventListener('click', () => deletePinned(entry));
    quickActions.appendChild(deleteBtn);

    item.appendChild(main);
    item.appendChild(meta);
    item.appendChild(quickActions);
    container.appendChild(item);
  }
}

function renderHistory() {
  const container = document.getElementById('historyList');
  const query = document.getElementById('historySearch').value.trim();
  const history = readHistory().filter((entry) => historyMatchesSearch(entry, query));
  container.innerHTML = '';

  if (history.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'small empty-state';
    empty.textContent = query ? 'No previous printed labels match that search.' : 'No label history yet.';
    container.appendChild(empty);
    return;
  }

  for (const entry of history) {
    if (!entry || typeof entry !== 'object') continue;
    const job = normalizeJob(entry.job || {});

    const item = document.createElement('div');
    item.className = 'history-item';

    const main = document.createElement('div');
    main.className = 'history-main';
    main.textContent = `${job.name || '(Unnamed item)'} — ${job.price || '(No price)'}`;

    const meta = document.createElement('div');
    meta.className = 'history-meta';
    meta.textContent = `Bottom: ${job.price || 'No bottom line'} | Warn: ${abbreviatedWarning(job.warning)} | ${formatTimestamp(entry.ts)} | Printer: ${entry.printer || 'Unknown'}`;

    const actions = document.createElement('div');
    actions.className = 'history-actions';

    const loadBtn = document.createElement('button');
    loadBtn.type = 'button';
    loadBtn.className = 'hist-load';
    loadBtn.textContent = 'Load';
    loadBtn.addEventListener('click', () => {
      applyJobToForm(job, entry.printer || '');
      setStatus('Recent label loaded.');
    });

    const deleteBtn = document.createElement('button');
    deleteBtn.type = 'button';
    deleteBtn.className = 'hist-delete';
    deleteBtn.textContent = 'Delete';
    deleteBtn.addEventListener('click', () => {
      const updated = readHistory().filter((row) => row && row.id !== entry.id);
      writeHistory(updated);
    });

    actions.appendChild(loadBtn);
    for (const count of QUICK_COPY_COUNTS) {
      const reprintBtn = document.createElement('button');
      reprintBtn.type = 'button';
      reprintBtn.className = 'hist-reprint';
      reprintBtn.textContent = `Reprint ${count}`;
      reprintBtn.addEventListener('click', () => {
        const payload = {
          ...job,
          printer: entry.printer || document.getElementById('printer').value || null,
        };
        printJob(payload, count, `Reprinting ${count} copy/copies...`);
      });
      actions.appendChild(reprintBtn);
    }
    actions.appendChild(deleteBtn);

    item.appendChild(main);
    item.appendChild(meta);
    item.appendChild(actions);
    container.appendChild(item);
  }
}

function clearAllHistory() {
  localStorage.removeItem(HISTORY_KEY);
  renderHistory();
  setStatus('History cleared.');
}

async function generateZPL() {
  setStatus('Generating ZPL...');
  const job = jobPayload();
  const res = await fetch('/zpl', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(job)
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

  saveDefaults(job);
  setStatus('ZPL downloaded.');
}

async function printLabel() {
  const job = jobPayload();
  await printJob(job, job.copies, 'Printing...');
}

document.getElementById('zplBtn').addEventListener('click', generateZPL);
document.getElementById('printBtn').addEventListener('click', printLabel);
document.getElementById('savePinnedBtn').addEventListener('click', saveCurrentAsPinned);
document.getElementById('catalogSearch').addEventListener('input', renderCatalog);
document.getElementById('clearSearchBtn').addEventListener('click', () => {
  document.getElementById('catalogSearch').value = '';
  renderCatalog();
});
document.getElementById('historySearch').addEventListener('input', renderHistory);
document.getElementById('clearHistorySearchBtn').addEventListener('click', () => {
  document.getElementById('historySearch').value = '';
  renderHistory();
});
const verticalOffsetInput = document.getElementById('vertical_offset');

function syncPriceFromPreset() {
  const preset = document.getElementById('price_preset').value;
  if (preset) {
    document.getElementById('price').value = preset;
  }
  saveDefaults(jobPayload());
}

function syncPresetFromPrice() {
  const price = cleanText(document.getElementById('price').value);
  const preset = document.getElementById('price_preset');
  const hasMatchingPreset = Array.from(preset.options).some((opt) => opt.value === price);
  preset.value = hasMatchingPreset ? price : '';
}

function clampOffset(value) {
  return Math.max(-60, Math.min(60, value));
}

function nudgeOffset(delta) {
  const current = parseInt(verticalOffsetInput.value || '0', 10);
  verticalOffsetInput.value = String(clampOffset(current + delta));
  saveDefaults(jobPayload());
}

document.getElementById('offsetUpBtn').addEventListener('click', () => nudgeOffset(1));
document.getElementById('offsetDownBtn').addEventListener('click', () => nudgeOffset(-1));
document.getElementById('clearHistoryBtn').addEventListener('click', clearAllHistory);
document.getElementById('name').addEventListener('input', autofillStrainType);
document.getElementById('price_preset').addEventListener('change', syncPriceFromPreset);
document.getElementById('price').addEventListener('input', () => {
  syncPresetFromPrice();
  saveDefaults(jobPayload());
});
for (const id of ['printer', 'size', 'strain_type', 'warning', 'include_warning', 'copies', 'darkness', 'vertical_offset']) {
  document.getElementById(id).addEventListener('change', () => saveDefaults(jobPayload()));
}

loadPrinters();
renderCatalog();
renderHistory();
</script>
</body>
</html>
""".replace("__DEFAULT_WARNING__", html.escape(DEFAULT_WARNING))


if __name__ == "__main__":
    # Run: python zpl_print_server.py
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=PORT)
