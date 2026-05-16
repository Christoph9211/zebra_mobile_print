import html
import os
import socket
import textwrap
import time
from collections import deque
from datetime import datetime, timezone
from typing import Any, Optional

from fastapi import FastAPI, Query
from fastapi.responses import FileResponse, HTMLResponse, PlainTextResponse
from pydantic import BaseModel, Field

import win32print  # pip install pywin32


# -----------------------
# Settings
# -----------------------
PORT = int(os.environ.get("ZPL_SERVER_PORT", "8787"))
APP_VERSION = "0.1.0"
STARTED_AT = datetime.now(timezone.utc)

# Set this to your exact Windows printer name.
# Example: "ZDesigner ZD411-203dpi ZPL"
DEFAULT_PRINTER = os.environ.get("ZPL_PRINTER_NAME", "")
DIRECT_PRINTER_HOST = os.environ.get("ZPL_PRINTER_HOST", "").strip()
DIRECT_PRINTER_PORT = int(os.environ.get("ZPL_PRINTER_PORT", "9100"))
DIRECT_PRINTER_TIMEOUT_SECONDS = float(os.environ.get("ZPL_PRINTER_TIMEOUT_SECONDS", "2.0"))

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
PRINT_LOG_LIMIT = 50
PRINT_ATTEMPT_LOG: deque[dict[str, Any]] = deque(maxlen=PRINT_LOG_LIMIT)


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


def build_test_zpl() -> str:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return "\n".join(
        [
            "^XA",
            f"^PW{LABEL_WIDTH_DOTS}",
            f"^LL{LABEL_HEIGHT_DOTS}",
            "^MD20",
            "^FO8,20",
            f"^FB{LABEL_WIDTH_DOTS-16},1,0,C,0",
            "^A0N,24,24",
            "^FDZebra test label^FS",
            "^FO8,62",
            f"^FB{LABEL_WIDTH_DOTS-16},1,0,C,0",
            "^A0N,16,16",
            f"^FD{zpl_escape(ts)}^FS",
            "^FO24,108",
            f"^GB{LABEL_WIDTH_DOTS-48},2,2^FS",
            "^XZ",
        ]
    ) + "\n"


PRINTER_STATUS_FLAGS = [
    ("paused", getattr(win32print, "PRINTER_STATUS_PAUSED", 0x00000001)),
    ("error", getattr(win32print, "PRINTER_STATUS_ERROR", 0x00000002)),
    ("pending_deletion", getattr(win32print, "PRINTER_STATUS_PENDING_DELETION", 0x00000004)),
    ("paper_jam", getattr(win32print, "PRINTER_STATUS_PAPER_JAM", 0x00000008)),
    ("paper_out", getattr(win32print, "PRINTER_STATUS_PAPER_OUT", 0x00000010)),
    ("manual_feed", getattr(win32print, "PRINTER_STATUS_MANUAL_FEED", 0x00000020)),
    ("paper_problem", getattr(win32print, "PRINTER_STATUS_PAPER_PROBLEM", 0x00000040)),
    ("offline", getattr(win32print, "PRINTER_STATUS_OFFLINE", 0x00000080)),
    ("io_active", getattr(win32print, "PRINTER_STATUS_IO_ACTIVE", 0x00000100)),
    ("busy", getattr(win32print, "PRINTER_STATUS_BUSY", 0x00000200)),
    ("printing", getattr(win32print, "PRINTER_STATUS_PRINTING", 0x00000400)),
    ("output_bin_full", getattr(win32print, "PRINTER_STATUS_OUTPUT_BIN_FULL", 0x00000800)),
    ("not_available", getattr(win32print, "PRINTER_STATUS_NOT_AVAILABLE", 0x00001000)),
    ("waiting", getattr(win32print, "PRINTER_STATUS_WAITING", 0x00002000)),
    ("processing", getattr(win32print, "PRINTER_STATUS_PROCESSING", 0x00004000)),
    ("initializing", getattr(win32print, "PRINTER_STATUS_INITIALIZING", 0x00008000)),
    ("warming_up", getattr(win32print, "PRINTER_STATUS_WARMING_UP", 0x00010000)),
    ("toner_low", getattr(win32print, "PRINTER_STATUS_TONER_LOW", 0x00020000)),
    ("no_toner", getattr(win32print, "PRINTER_STATUS_NO_TONER", 0x00040000)),
    ("page_punt", getattr(win32print, "PRINTER_STATUS_PAGE_PUNT", 0x00080000)),
    ("user_intervention", getattr(win32print, "PRINTER_STATUS_USER_INTERVENTION", 0x00100000)),
    ("out_of_memory", getattr(win32print, "PRINTER_STATUS_OUT_OF_MEMORY", 0x00200000)),
    ("door_open", getattr(win32print, "PRINTER_STATUS_DOOR_OPEN", 0x00400000)),
    ("server_unknown", getattr(win32print, "PRINTER_STATUS_SERVER_UNKNOWN", 0x00800000)),
    ("power_save", getattr(win32print, "PRINTER_STATUS_POWER_SAVE", 0x01000000)),
]

JOB_STATUS_FLAGS = [
    ("paused", getattr(win32print, "JOB_STATUS_PAUSED", 0x00000001)),
    ("error", getattr(win32print, "JOB_STATUS_ERROR", 0x00000002)),
    ("deleting", getattr(win32print, "JOB_STATUS_DELETING", 0x00000004)),
    ("spooling", getattr(win32print, "JOB_STATUS_SPOOLING", 0x00000008)),
    ("printing", getattr(win32print, "JOB_STATUS_PRINTING", 0x00000010)),
    ("offline", getattr(win32print, "JOB_STATUS_OFFLINE", 0x00000020)),
    ("paperout", getattr(win32print, "JOB_STATUS_PAPEROUT", 0x00000040)),
    ("printed", getattr(win32print, "JOB_STATUS_PRINTED", 0x00000080)),
    ("deleted", getattr(win32print, "JOB_STATUS_DELETED", 0x00000100)),
    ("blocked_device_queue", getattr(win32print, "JOB_STATUS_BLOCKED_DEVQ", 0x00000200)),
    ("user_intervention", getattr(win32print, "JOB_STATUS_USER_INTERVENTION", 0x00000400)),
    ("restart", getattr(win32print, "JOB_STATUS_RESTART", 0x00000800)),
    ("complete", getattr(win32print, "JOB_STATUS_COMPLETE", 0x00001000)),
]


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def decode_flags(value: int, flag_defs: list[tuple[str, int]]) -> list[str]:
    return [name for name, bit in flag_defs if bit and value & bit]


def format_windows_error(exc: Exception) -> str:
    parts = [str(part) for part in getattr(exc, "args", ()) if str(part)]
    if parts:
        return " | ".join(parts)
    return f"{exc.__class__.__name__}: {exc}"


def record_print_attempt(entry: dict[str, Any]) -> None:
    PRINT_ATTEMPT_LOG.appendleft(
        {
            "timestamp": now_iso(),
            **entry,
        }
    )


def get_default_printer_name() -> str:
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


def send_raw_zpl(printer_name: str, zpl: str) -> int:
    h = None
    job_id = 0
    try:
        h = win32print.OpenPrinter(printer_name)
        job = win32print.StartDocPrinter(h, 1, ("ZPL Mobile Label", "", "RAW"))
        job_id = int(job or 0)
        win32print.StartPagePrinter(h)
        win32print.WritePrinter(h, zpl.encode("ascii", errors="ignore"))
        win32print.EndPagePrinter(h)
        win32print.EndDocPrinter(h)
        return job_id
    except Exception as exc:
        raise RuntimeError(f"Windows printer error for '{printer_name}': {format_windows_error(exc)}") from exc
    finally:
        if h is not None:
            win32print.ClosePrinter(h)


def list_printers() -> list[str]:
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    return [p[2] for p in win32print.EnumPrinters(flags)]


def selected_printer_name(name: Optional[str]) -> str:
    return (name or DEFAULT_PRINTER or "").strip()


def get_printer_details(printer_name: str) -> dict[str, Any]:
    h = None
    try:
        h = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(h, 2)
        status = int(info.get("Status") or 0)
        attrs = int(info.get("Attributes") or 0)
        return {
            "ok": True,
            "name": info.get("pPrinterName") or printer_name,
            "driver": info.get("pDriverName") or "",
            "port": info.get("pPortName") or "",
            "share_name": info.get("pShareName") or "",
            "location": info.get("pLocation") or "",
            "comment": info.get("pComment") or "",
            "jobs": int(info.get("cJobs") or 0),
            "status": status,
            "status_flags": decode_flags(status, PRINTER_STATUS_FLAGS),
            "attributes": attrs,
            "default_windows_printer": get_default_printer_name(),
            "configured_default_printer": DEFAULT_PRINTER,
            "is_configured_default": bool(DEFAULT_PRINTER and DEFAULT_PRINTER == printer_name),
            "is_windows_default": bool(get_default_printer_name() == printer_name),
        }
    except Exception as exc:
        return {
            "ok": False,
            "name": printer_name,
            "error": format_windows_error(exc),
            "default_windows_printer": get_default_printer_name(),
            "configured_default_printer": DEFAULT_PRINTER,
        }
    finally:
        if h is not None:
            win32print.ClosePrinter(h)


def list_printer_details() -> list[dict[str, Any]]:
    default_windows = get_default_printer_name()
    rows: list[dict[str, Any]] = []
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    for printer in win32print.EnumPrinters(flags):
        name = printer[2]
        details = get_printer_details(name)
        details["description"] = printer[1] if len(printer) > 1 else ""
        details["comment"] = details.get("comment") or (printer[3] if len(printer) > 3 else "")
        details["is_configured_default"] = bool(DEFAULT_PRINTER and DEFAULT_PRINTER == name)
        details["is_windows_default"] = bool(default_windows and default_windows == name)
        rows.append(details)
    return rows


def list_printer_jobs(printer_name: str) -> dict[str, Any]:
    h = None
    try:
        h = win32print.OpenPrinter(printer_name)
        jobs = []
        for job in win32print.EnumJobs(h, 0, 25, 1):
            status = int(job.get("Status") or 0)
            submitted = job.get("Submitted")
            jobs.append(
                {
                    "id": job.get("JobId"),
                    "document": job.get("pDocument") or "",
                    "user": job.get("pUserName") or "",
                    "status": status,
                    "status_flags": decode_flags(status, JOB_STATUS_FLAGS),
                    "pages_printed": job.get("PagesPrinted"),
                    "total_pages": job.get("TotalPages"),
                    "submitted": str(submitted) if submitted else "",
                }
            )
        return {"ok": True, "printer": printer_name, "jobs": jobs}
    except Exception as exc:
        return {"ok": False, "printer": printer_name, "error": format_windows_error(exc), "jobs": []}
    finally:
        if h is not None:
            win32print.ClosePrinter(h)


def probe_direct_printer() -> dict[str, Any]:
    if not DIRECT_PRINTER_HOST:
        return {
            "configured": False,
            "message": "Direct Zebra network probe is not configured. Set ZPL_PRINTER_HOST to enable it.",
        }

    started = time.monotonic()
    try:
        with socket.create_connection(
            (DIRECT_PRINTER_HOST, DIRECT_PRINTER_PORT),
            timeout=DIRECT_PRINTER_TIMEOUT_SECONDS,
        ) as sock:
            sock.settimeout(DIRECT_PRINTER_TIMEOUT_SECONDS)
            status_text = ""
            try:
                sock.sendall(b'! U1 getvar "device.status"\r\n')
                status_text = sock.recv(512).decode("ascii", errors="replace").strip()
            except Exception:
                status_text = ""
        elapsed_ms = round((time.monotonic() - started) * 1000)
        return {
            "configured": True,
            "ok": True,
            "host": DIRECT_PRINTER_HOST,
            "port": DIRECT_PRINTER_PORT,
            "elapsed_ms": elapsed_ms,
            "status": status_text,
        }
    except Exception as exc:
        elapsed_ms = round((time.monotonic() - started) * 1000)
        return {
            "configured": True,
            "ok": False,
            "host": DIRECT_PRINTER_HOST,
            "port": DIRECT_PRINTER_PORT,
            "elapsed_ms": elapsed_ms,
            "error": format_windows_error(exc),
        }


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


class TestPrintJob(BaseModel):
    printer: Optional[str] = Field(default=None, description="Windows printer name (optional)")


@app.get("/health")
def health():
    return {
        "ok": True,
        "app": "ZPL Mobile Print Server",
        "version": APP_VERSION,
        "started_at": STARTED_AT.isoformat(),
        "uptime_seconds": int((datetime.now(timezone.utc) - STARTED_AT).total_seconds()),
        "port": PORT,
        "configured_default_printer": DEFAULT_PRINTER,
        "windows_default_printer": get_default_printer_name(),
        "direct_printer_configured": bool(DIRECT_PRINTER_HOST),
    }


@app.get("/printers", response_class=PlainTextResponse)
def printers():
    return "\n".join(list_printers())


@app.get("/diagnostics/printers")
def diagnostics_printers():
    try:
        return {
            "ok": True,
            "configured_default_printer": DEFAULT_PRINTER,
            "windows_default_printer": get_default_printer_name(),
            "printers": list_printer_details(),
        }
    except Exception as exc:
        return {
            "ok": False,
            "error": format_windows_error(exc),
            "configured_default_printer": DEFAULT_PRINTER,
            "windows_default_printer": get_default_printer_name(),
            "printers": [],
        }


@app.get("/diagnostics/printer")
def diagnostics_printer(name: Optional[str] = Query(default=None)):
    printer = selected_printer_name(name)
    if not printer:
        return {
            "ok": False,
            "error": "Printer not set. Choose a printer or set ZPL_PRINTER_NAME.",
            "configured_default_printer": DEFAULT_PRINTER,
            "windows_default_printer": get_default_printer_name(),
        }
    return get_printer_details(printer)


@app.get("/diagnostics/jobs")
def diagnostics_jobs(name: Optional[str] = Query(default=None)):
    printer = selected_printer_name(name)
    if not printer:
        return {"ok": False, "error": "Printer not set. Choose a printer or set ZPL_PRINTER_NAME.", "jobs": []}
    return list_printer_jobs(printer)


@app.get("/diagnostics/logs")
def diagnostics_logs():
    return {"ok": True, "limit": PRINT_LOG_LIMIT, "logs": list(PRINT_ATTEMPT_LOG)}


@app.get("/diagnostics/network")
def diagnostics_network():
    return probe_direct_printer()


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
        record_print_attempt(
            {
                "printer": "",
                "copies": job.copies,
                "success": False,
                "job_ids": [],
                "error": "Printer not set. Set ZPL_PRINTER_NAME env var or choose a printer in the UI.",
                "source": "label",
            }
        )
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

    job_ids: list[int] = []
    try:
        for _ in range(job.copies):
            job_ids.append(send_raw_zpl(printer, zpl))
    except Exception as exc:
        error = str(exc)
        record_print_attempt(
            {
                "printer": printer,
                "copies": job.copies,
                "success": False,
                "job_ids": job_ids,
                "error": error,
                "source": "label",
            }
        )
        return PlainTextResponse(error, status_code=500)

    record_print_attempt(
        {
            "printer": printer,
            "copies": job.copies,
            "success": True,
            "job_ids": job_ids,
            "error": "",
            "source": "label",
        }
    )
    return f"Printed {job.copies} copy/copies to: {printer}"


@app.post("/diagnostics/test-print", response_class=PlainTextResponse)
def diagnostics_test_print(job: TestPrintJob):
    printer = selected_printer_name(job.printer)
    if not printer:
        record_print_attempt(
            {
                "printer": "",
                "copies": 1,
                "success": False,
                "job_ids": [],
                "error": "Printer not set. Choose a printer or set ZPL_PRINTER_NAME.",
                "source": "test_label",
            }
        )
        return PlainTextResponse("Printer not set. Choose a printer or set ZPL_PRINTER_NAME.", status_code=400)

    try:
        job_id = send_raw_zpl(printer, build_test_zpl())
    except Exception as exc:
        error = str(exc)
        record_print_attempt(
            {
                "printer": printer,
                "copies": 1,
                "success": False,
                "job_ids": [],
                "error": error,
                "source": "test_label",
            }
        )
        return PlainTextResponse(error, status_code=500)

    record_print_attempt(
        {
            "printer": printer,
            "copies": 1,
            "success": True,
            "job_ids": [job_id],
            "error": "",
            "source": "test_label",
        }
    )
    return f"Sent test label to: {printer}"


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
    .collapsible-panel { padding: 0; overflow: hidden; }
    .collapsible-trigger { width: 100%; display: flex; justify-content: space-between; align-items: center; gap: 10px; padding: 14px 16px; border-radius: 0; background: transparent; color: inherit; text-align: left; font-size: 20px; font-weight: 700; }
    .collapsible-trigger span:last-child { font-size: 18px; color: #555; }
    .collapsible-content { padding: 0 16px 16px; }
    .collapsible-content[hidden] { display: none; }
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
    #status { margin-top: 12px; white-space: pre-wrap; }
    .small { font-size: 13px; color: #666; margin-top: 6px; }
    .toggle { display: flex; align-items: center; gap: 10px; margin-top: 10px; }
    .toggle input { width: auto; transform: scale(1.3); }
    .offset-control { display: flex; align-items: center; gap: 10px; margin-top: 8px; }
    .offset-control input { width: 110px; text-align: center; font-weight: 700; }
    .offset-btn { flex: 0 0 auto; font-size: 16px; padding: 10px 14px; background: #f2f2f2; border: 1px solid #ccc; }
    .clear-history-btn, .small-action-btn { flex: 0 0 auto; font-size: 14px; padding: 8px 12px; border-radius: 10px; border: 1px solid #ccc; background: #f2f2f2; }
    #historyList { display: flex; flex-direction: column; gap: 10px; }
    .history-header { margin-top: 22px; }
    .history-item { border: 1px solid #ddd; border-radius: 12px; padding: 12px; background: #fafafa; }
    .history-main { font-weight: 700; font-size: 16px; }
    .history-meta { margin-top: 4px; font-size: 13px; color: #555; line-height: 1.4; }
    .history-actions, .quick-actions { display: flex; gap: 8px; margin-top: 10px; }
    .history-actions button { flex: 1; font-size: 16px; padding: 10px 12px; border-radius: 12px; border: 0; }
    .quick-actions button { flex: 1; font-size: 17px; padding: 11px 10px; border-radius: 12px; border: 0; }
    .hist-load { background: #f2f2f2; }
    .hist-reprint { background: #111; color: #fff; }
    .hist-delete { background: #ffdede; color: #8b0000; }
    .empty-state { padding: 8px 0; }
    .diagnostic-actions { display: flex; gap: 10px; margin-top: 10px; }
    .diagnostic-actions button { flex: 1; font-size: 16px; padding: 11px 10px; border-radius: 12px; }
    #runDiagnosticsBtn { background: #e8eef8; color: #18365f; }
    #testLabelBtn { background: #111; color: #fff; }
    .diag-summary { display: grid; gap: 8px; margin-top: 10px; }
    .diag-row { padding: 10px; border-radius: 10px; border: 1px solid #ddd; background: #fff; font-size: 14px; line-height: 1.35; }
    .diag-ok { border-color: #b9ddc2; background: #f2fbf4; color: #164123; }
    .diag-warn { border-color: #e8d69a; background: #fff9df; color: #5a4200; }
    .diag-bad { border-color: #efb0aa; background: #fff0ee; color: #7a170d; }
    #diagnosticsDetails { margin-top: 10px; white-space: pre-wrap; font-size: 12px; line-height: 1.35; max-height: 260px; overflow: auto; background: #f6f6f6; border: 1px solid #ddd; border-radius: 10px; padding: 10px; }
    @media (max-width: 560px) {
      body { margin: 8px; }
      .card { padding: 12px; }
      .row, .btnrow, .quick-toolbar, .diagnostic-actions { flex-direction: column; }
      .history-actions, .quick-actions { flex-wrap: wrap; }
      .history-actions button, .quick-actions button { flex: 1 1 30%; }
    }
  </style>
</head>
<body>
  <div class="card">
    <h2 style="margin-top:0">Zebra ZD411 — Mobile Label Print (2×1)</h2>

    <label>Printer</label>
    <select id="printer"></select>
    <div class="small">If blank, go to /printers on this server to see what Windows calls your printer.</div>

    <div class="panel collapsible-panel">
      <button id="troubleshootingToggle" class="collapsible-trigger" type="button" aria-expanded="false" aria-controls="troubleshootingContent">
        <span>Troubleshooting</span>
        <span id="troubleshootingChevron">Show</span>
      </button>
      <div id="troubleshootingContent" class="collapsible-content" hidden>
        <div class="small">Run checks from this phone to see whether the web app, Windows spooler, queue, and optional Zebra network probe look healthy.</div>
        <div class="diagnostic-actions">
          <button id="runDiagnosticsBtn" type="button">Run Diagnostics</button>
          <button id="testLabelBtn" type="button">Send Test Label</button>
        </div>
        <div id="diagnosticsSummary" class="diag-summary"></div>
        <pre id="diagnosticsDetails"></pre>
      </div>
    </div>

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
          <option value="$6.00">$6.00</option>
          <option value="$10.00">$10.00</option>
          <option value="$15.00">$15.00</option>
          <option value="$18.00">$18.00</option>
          <option value="$20.00">$20.00</option>
          <option value="$25.00">$25.00</option>
          <option value="$30.00">$30.00</option>
          <option value="$32.50">$32.50</option>
          <option value="$40.00">$40.00</option>
          <option value="$50.00">$50.00</option>
          <option value="$100.00">$100.00</option>
          <option value="$125.00">$125.00</option>
          <option value="$150.00">$150.00</option>
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
const DEFAULTS_KEY = 'zebra_label_defaults_v1';
const STRAIN_MEMORY_KEY = 'zebra_strain_type_memory_v1';
const MAX_HISTORY = 100;
const HISTORY_SCHEMA_VERSION = 1;
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

function setTroubleshootingOpen(open) {
  const toggle = document.getElementById('troubleshootingToggle');
  const content = document.getElementById('troubleshootingContent');
  const chevron = document.getElementById('troubleshootingChevron');
  toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
  content.hidden = !open;
  chevron.textContent = open ? 'Hide' : 'Show';
}

function toggleTroubleshooting() {
  const currentlyOpen = document.getElementById('troubleshootingToggle').getAttribute('aria-expanded') === 'true';
  setTroubleshootingOpen(!currentlyOpen);
}

function selectedPrinter() {
  return document.getElementById('printer').value || '';
}

async function fetchJson(url, options) {
  try {
    const res = await fetch(url, options);
    const contentType = res.headers.get('content-type') || '';
    const body = contentType.includes('application/json') ? await res.json() : { text: await res.text() };
    return { ok: res.ok, status: res.status, body };
  } catch (e) {
    return { ok: false, status: 0, body: { error: String(e && e.message ? e.message : e) } };
  }
}

function addDiagnosticRow(container, level, text) {
  const row = document.createElement('div');
  row.className = `diag-row diag-${level}`;
  row.textContent = text;
  container.appendChild(row);
}

function latestLogForPrinter(logs, printer) {
  const rows = (logs && logs.body && Array.isArray(logs.body.logs)) ? logs.body.logs : [];
  if (!printer) {
    return rows[0] || null;
  }
  return rows.find((row) => row && row.printer === printer) || rows[0] || null;
}

function diagnosticSuggestion(printerInfo, jobs, network, lastLog, printer) {
  if (!printer) return 'Choose a printer, then run diagnostics again.';
  if (!printerInfo.ok || !printerInfo.body.ok) return 'Windows cannot open this printer. Check the printer name, driver, and Windows printer list.';
  const flags = printerInfo.body.status_flags || [];
  const hardFlags = ['offline', 'paused', 'paper_out', 'paper_jam', 'door_open', 'user_intervention', 'error', 'not_available'];
  const activeFlag = hardFlags.find((flag) => flags.includes(flag));
  if (activeFlag) return `Fix the printer state reported by Windows: ${activeFlag.replaceAll('_', ' ')}.`;
  const jobRows = (jobs.body && jobs.body.jobs) || [];
  const stuckJob = jobRows.find((job) => (job.status_flags || []).some((flag) => ['error', 'offline', 'paperout', 'blocked_device_queue', 'user_intervention'].includes(flag)));
  if (stuckJob) return `Clear or restart spooler job ${stuckJob.id}; Windows reports ${stuckJob.status_flags.join(', ')}.`;
  if (network.body && network.body.configured && !network.body.ok) return 'Windows may see the queue, but the direct Zebra network probe failed. Check printer IP, Wi-Fi, and power.';
  if (lastLog && lastLog.success === false) return `Last print failed: ${lastLog.error || 'unknown error'}`;
  if (jobRows.length > 0) return 'Windows has active queued jobs. If labels are not moving, open the queue and clear stuck jobs.';
  return 'No obvious server or Windows queue problem found. Send a test label to separate label content from printer hardware/media issues.';
}

function renderDiagnostics(results, printer) {
  const summary = document.getElementById('diagnosticsSummary');
  const details = document.getElementById('diagnosticsDetails');
  summary.innerHTML = '';

  const { health, printers, printerInfo, jobs, logs, network } = results;
  const effectivePrinter = printer || (health.body && health.body.configured_default_printer) || '';
  addDiagnosticRow(summary, health.ok && health.body.ok ? 'ok' : 'bad', health.ok && health.body.ok ? 'Web app is reachable.' : `Web app health check failed: ${(health.body && health.body.error) || health.status}`);

  const printerCount = printers.body && Array.isArray(printers.body.printers) ? printers.body.printers.length : 0;
  if (!printers.ok || !printers.body.ok) {
    addDiagnosticRow(summary, 'bad', `Could not read Windows printers: ${(printers.body && printers.body.error) || printers.status}`);
  } else if (printerCount === 0) {
    addDiagnosticRow(summary, 'bad', 'Windows returned no printers.');
  } else {
    addDiagnosticRow(summary, 'ok', `Windows sees ${printerCount} printer(s).`);
  }

  if (!effectivePrinter) {
    addDiagnosticRow(summary, 'warn', 'No printer is selected.');
  } else {
    if (!printer) {
      addDiagnosticRow(summary, 'ok', `Using configured default printer: ${effectivePrinter}`);
    }
    if (!printerInfo.ok || !printerInfo.body.ok) {
      addDiagnosticRow(summary, 'bad', `Selected printer could not be opened: ${(printerInfo.body && printerInfo.body.error) || printerInfo.status}`);
    } else {
      const flags = printerInfo.body.status_flags || [];
      const level = flags.some((flag) => ['offline', 'paused', 'paper_out', 'paper_jam', 'door_open', 'user_intervention', 'error', 'not_available'].includes(flag)) ? 'bad' : 'ok';
      addDiagnosticRow(summary, level, flags.length ? `Printer state: ${flags.join(', ')}` : 'Printer state has no Windows error flags.');
    }
  }

  const jobRows = jobs.body && Array.isArray(jobs.body.jobs) ? jobs.body.jobs : [];
  if (!jobs.ok || !jobs.body.ok) {
    addDiagnosticRow(summary, 'warn', `Could not read active spooler jobs: ${(jobs.body && jobs.body.error) || jobs.status}`);
  } else if (jobRows.length) {
    addDiagnosticRow(summary, 'warn', `${jobRows.length} active job(s) are in the Windows queue.`);
  } else {
    addDiagnosticRow(summary, 'ok', 'No active jobs are stuck in the Windows queue.');
  }

  if (network.body && network.body.configured) {
    addDiagnosticRow(summary, network.body.ok ? 'ok' : 'warn', network.body.ok ? `Direct Zebra TCP probe reached ${network.body.host}:${network.body.port}.` : `Direct Zebra TCP probe failed: ${network.body.error || 'unknown error'}`);
  } else {
    addDiagnosticRow(summary, 'warn', 'Direct Zebra network probe is not configured.');
  }

  const lastLog = latestLogForPrinter(logs, effectivePrinter);
  if (lastLog) {
    addDiagnosticRow(summary, lastLog.success ? 'ok' : 'bad', lastLog.success ? `Last print attempt succeeded at ${formatTimestamp(lastLog.timestamp)}.` : `Last print attempt failed at ${formatTimestamp(lastLog.timestamp)}: ${lastLog.error || 'unknown error'}`);
  } else {
    addDiagnosticRow(summary, 'warn', 'No print attempts have been logged since the server started.');
  }

  addDiagnosticRow(summary, 'warn', diagnosticSuggestion(printerInfo, jobs, network, lastLog, effectivePrinter));

  details.textContent = JSON.stringify({
    selected_printer: printer,
    effective_printer: effectivePrinter,
    health: health.body,
    printers: printers.body,
    selected_printer_status: printerInfo.body,
    jobs: jobs.body,
    direct_network_probe: network.body,
    recent_logs: logs.body,
  }, null, 2);
}

async function runDiagnostics() {
  setTroubleshootingOpen(true);
  const printer = selectedPrinter();
  setStatus('Running diagnostics...');
  const query = printer ? `?name=${encodeURIComponent(printer)}` : '';
  const [health, printers, printerInfo, jobs, logs, network] = await Promise.all([
    fetchJson('/health'),
    fetchJson('/diagnostics/printers'),
    fetchJson(`/diagnostics/printer${query}`),
    fetchJson(`/diagnostics/jobs${query}`),
    fetchJson('/diagnostics/logs'),
    fetchJson('/diagnostics/network'),
  ]);
  renderDiagnostics({ health, printers, printerInfo, jobs, logs, network }, printer);
  setStatus('Diagnostics complete.');
}

async function sendTestLabel() {
  setTroubleshootingOpen(true);
  const printer = selectedPrinter();
  setStatus('Sending test label...');
  const res = await fetch('/diagnostics/test-print', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ printer: printer || null })
  });
  const text = await res.text();
  setStatus(text);
  await runDiagnostics();
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
document.getElementById('troubleshootingToggle').addEventListener('click', toggleTroubleshooting);
document.getElementById('runDiagnosticsBtn').addEventListener('click', runDiagnostics);
document.getElementById('testLabelBtn').addEventListener('click', sendTestLabel);
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
renderHistory();
</script>
</body>
</html>
""".replace("__DEFAULT_WARNING__", html.escape(DEFAULT_WARNING))


if __name__ == "__main__":
    # Run: python zpl_print_server.py
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=PORT)
