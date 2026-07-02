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
from pydantic import BaseModel, Field, model_validator

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

DEFAULT_WARNING = """THCA PRODUCT • HEMP-DERIVED
Contains less than 0.3% Δ9 THC.
21+ only. Keep out of reach of children.
May cause intoxication when heated.
Do not use while driving or operating heavy machinery.
Consult a physician before use."""

# 2x1 @ 203dpi -> 406x203 dots; use 609x203 for 3x1 media.
LABEL_WIDTH_DOTS = int(os.environ.get("ZPL_LABEL_WIDTH_DOTS", "406"))
LABEL_HEIGHT_DOTS = int(os.environ.get("ZPL_LABEL_HEIGHT_DOTS", "203"))
LABEL_Y_OFFSET = int(os.environ.get("ZPL_LABEL_Y_OFFSET", "6"))
LABEL_MARGIN_DOTS = 16
SECTION_GAP_DOTS = 4
HEADER_SECTION_HEIGHT_DOTS = 52
DETAILS_SECTION_HEIGHT_DOTS = 18
TITLE_FONT_MAX_DOTS = 30
TITLE_FONT_MIN_DOTS = 10
SUBTITLE_FONT_DOTS = 15
DETAILS_FONT_DOTS = 14
DETAILS_FONT_MAX_DOTS = 20
PROMO_FONT_DOTS = 14
PRICE_FONT_MAX_DOTS = 42
PRICE_FONT_MIN_DOTS = 10
OLD_PRICE_FONT_DOTS = 14
WARNING_HEADING_FONT_DOTS = 20
WARNING_FONT_MAX_DOTS = 18
WARNING_FONT_MIN_DOTS = 10
PREROLL_WARNING_HEADING_FONT_DOTS = 14
PREROLL_WARNING_TOP_DOTS = 123
PREROLL_MARKDOWN_WARNING_TOP_DOTS = 126
WARNING_MAX_CHARACTERS = 600
PRINT_LOG_LIMIT = 50
PRINT_ATTEMPT_LOG: deque[dict[str, Any]] = deque(maxlen=PRINT_LOG_LIMIT)


# -----------------------
# ZPL builder
# -----------------------
def zpl_escape(s: str) -> str:
    if s is None:
        return ""
    text = str(s).replace("•", "-").replace("Δ", "DELTA-")
    return text.replace("^", "").replace("~", "").encode("ascii", errors="ignore").decode("ascii").strip()


def fitted_font_size(text: str, maximum: int, minimum: int, width_dots: int) -> int:
    # ponytail: approximate Zebra's proportional font; use printer font metrics if long text still clips.
    return max(minimum, min(maximum, (width_dots * 3 // 2) // max(len(text), 1)))


def format_product_details(size: str, strain_type: str) -> str:
    size = zpl_escape(size)
    size = {"1 gram": "1g", "2 gram": "2g", "3 gram": "3g"}.get(size, size)
    return " - ".join(filter(None, (size, zpl_escape(strain_type).upper())))


def build_price_label_sections() -> dict[str, dict[str, int]]:
    """Allocate the full label to product and price content."""
    header_top = LABEL_MARGIN_DOTS
    details_top = header_top + HEADER_SECTION_HEIGHT_DOTS + SECTION_GAP_DOTS
    price_top = details_top + DETAILS_SECTION_HEIGHT_DOTS + SECTION_GAP_DOTS
    return {
        "headerSection": {"top": header_top, "height": HEADER_SECTION_HEIGHT_DOTS},
        "detailsSection": {"top": details_top, "height": DETAILS_SECTION_HEIGHT_DOTS},
        "priceSection": {
            "top": price_top,
            "height": max(0, LABEL_HEIGHT_DOTS - LABEL_MARGIN_DOTS - price_top),
        },
    }


def build_warning_label_sections() -> dict[str, dict[str, int]]:
    """Reserve a short heading and give the rest of the label to the warning."""
    heading_top = LABEL_MARGIN_DOTS
    warning_top = heading_top + WARNING_HEADING_FONT_DOTS + 8
    return {
        "headerSection": {"top": heading_top, "height": WARNING_HEADING_FONT_DOTS},
        "warningSection": {
            "top": warning_top,
            "height": max(0, LABEL_HEIGHT_DOTS - LABEL_MARGIN_DOTS - warning_top),
        },
    }


def draw_centered_text(text: str, y: int, font_height: int, font_width: Optional[int] = None) -> list[str]:
    if not text:
        return []
    printable_width = LABEL_WIDTH_DOTS - (LABEL_MARGIN_DOTS * 2)
    return [
        f"^FO{LABEL_MARGIN_DOTS},{y}",
        f"^FB{printable_width},1,0,C,0",
        f"^A0N,{font_height},{font_width or font_height}",
        f"^FD{zpl_escape(text)}^FS",
    ]


def draw_strikethrough_text(text: str, y: int, font_height: int) -> list[str]:
    text = zpl_escape(text)
    printable_width = LABEL_WIDTH_DOTS - (LABEL_MARGIN_DOTS * 2)
    strike_width = min(printable_width, max(70, len(text) * 8))
    strike_x = LABEL_MARGIN_DOTS + ((printable_width - strike_width) // 2)
    return [
        *draw_centered_text(text, y, font_height),
        f"^FO{strike_x},{y + (font_height // 2)}",
        f"^GB{strike_width},2,2^FS",
    ]


def fit_warning_text(text: str, width_dots: int, height_dots: int) -> tuple[str, int, int, int, int]:
    clean = " ".join(zpl_escape(text).split())
    if not clean:
        raise ValueError("Health warning is required.")

    for font_height in range(WARNING_FONT_MAX_DOTS, WARNING_FONT_MIN_DOTS - 1, -1):
        font_width = max(7, font_height * 3 // 4)
        char_width = max(5, font_width * 2 // 3)
        max_chars = max(1, width_dots // char_width)
        lines = textwrap.wrap(
            clean,
            width=max_chars,
            break_long_words=True,
            break_on_hyphens=True,
        )
        line_spacing = 2
        block_height = (len(lines) * font_height) + (max(0, len(lines) - 1) * line_spacing)
        if block_height <= height_dots:
            return r"\&".join(lines), font_height, font_width, line_spacing, block_height

    raise ValueError("Health warning is too long to fit on the label.")


def build_price_label_zpl(
    name: str,
    price: str,
    darkness: int = 20,
    vertical_offset: int = 0,
    marked_down: bool = False,
    original_price: str = "",
    subtitle: str = "",
    size: str = "",
    strain_type: str = "",
    price_input: str = "",
) -> str:
    """Render the full-size retail product and price label."""
    name = zpl_escape(name).upper()
    subtitle = zpl_escape(subtitle)
    details = format_product_details(size, strain_type)
    display_price = zpl_escape(price_input) or zpl_escape(price)
    original_price = zpl_escape(original_price)
    printable_width = LABEL_WIDTH_DOTS - (LABEL_MARGIN_DOTS * 2)
    title_font = fitted_font_size(
        name,
        maximum=TITLE_FONT_MAX_DOTS,
        minimum=TITLE_FONT_MIN_DOTS,
        width_dots=printable_width,
    )
    price_font = fitted_font_size(
        display_price,
        maximum=PRICE_FONT_MAX_DOTS,
        minimum=PRICE_FONT_MIN_DOTS,
        width_dots=printable_width,
    )
    emphasize_details = not marked_down and not subtitle
    details_font = (
        fitted_font_size(details, DETAILS_FONT_MAX_DOTS, DETAILS_FONT_DOTS, printable_width)
        if emphasize_details
        else DETAILS_FONT_DOTS
    )
    # Positive values move content up; negative values move content down.
    y_offset = LABEL_Y_OFFSET - vertical_offset
    sections = build_price_label_sections()
    header_section = sections["headerSection"]
    details_section = sections["detailsSection"]
    price_section = sections["priceSection"]

    z = ["^XA", f"^PW{LABEL_WIDTH_DOTS}", f"^LL{LABEL_HEIGHT_DOTS}", f"^MD{darkness}"]

    # Header: large title stays below the rounded-corner/non-printable area.
    z += draw_centered_text(name, header_section["top"] + y_offset, title_font)
    if subtitle:
        z += draw_centered_text(
            subtitle,
            header_section["top"] + 32 + y_offset,
            SUBTITLE_FONT_DOTS,
        )

    # Details: weight and strain get their own full-width line.
    if details:
        details_y = header_section["top"] + 42 if emphasize_details else details_section["top"]
        z += draw_centered_text(details, details_y + y_offset, details_font)

    # Price: monochrome hierarchy carries the meaning; no color is required.
    if marked_down:
        z += draw_centered_text(
            "PRICE REDUCED",
            price_section["top"] + y_offset,
            PROMO_FONT_DOTS,
        )
    z += draw_centered_text(
        display_price,
        price_section["top"] + (18 if marked_down else 12 if emphasize_details else 16) + y_offset,
        price_font,
    )
    if marked_down:
        z += draw_strikethrough_text(
            f"WAS {original_price}",
            price_section["top"] + 64 + y_offset,
            OLD_PRICE_FONT_DOTS,
        )

    z += ["^XZ"]
    return "\n".join(z) + "\n"


def build_warning_label_zpl(
    warning: str,
    darkness: int = 20,
    vertical_offset: int = 0,
) -> str:
    """Render the full-size health warning label without truncating its text."""
    printable_width = LABEL_WIDTH_DOTS - (LABEL_MARGIN_DOTS * 2)
    sections = build_warning_label_sections()
    warning_section = sections["warningSection"]
    wrapped, font_height, font_width, line_spacing, block_height = fit_warning_text(
        warning,
        printable_width,
        warning_section["height"],
    )
    y_offset = LABEL_Y_OFFSET - vertical_offset
    warning_y = warning_section["top"] + ((warning_section["height"] - block_height) // 2) + y_offset
    line_count = wrapped.count(r"\&") + 1

    z = ["^XA", f"^PW{LABEL_WIDTH_DOTS}", f"^LL{LABEL_HEIGHT_DOTS}", f"^MD{darkness}"]
    z += draw_centered_text(
        "HEALTH WARNING",
        sections["headerSection"]["top"] + y_offset,
        WARNING_HEADING_FONT_DOTS,
    )
    z += [
        f"^FO{LABEL_MARGIN_DOTS},{warning_y}",
        f"^FB{printable_width},{line_count},{line_spacing},L,0",
        f"^A0N,{font_height},{font_width}",
        f"^FD{wrapped}^FS",
        "^XZ",
    ]
    return "\n".join(z) + "\n"


def build_preroll_label_zpl(
    name: str,
    price: str,
    warning: str,
    darkness: int = 20,
    vertical_offset: int = 0,
    marked_down: bool = False,
    original_price: str = "",
    price_input: str = "",
) -> str:
    """Render product, price, and health warning on one preroll label."""
    name = zpl_escape(name).upper()
    display_price = zpl_escape(price_input) or zpl_escape(price)
    printable_width = LABEL_WIDTH_DOTS - (LABEL_MARGIN_DOTS * 2)
    title_font = fitted_font_size(name, TITLE_FONT_MAX_DOTS, TITLE_FONT_MIN_DOTS, printable_width)
    price_font = fitted_font_size(
        display_price,
        34 if marked_down else PRICE_FONT_MAX_DOTS,
        PRICE_FONT_MIN_DOTS,
        printable_width,
    )
    warning_top = PREROLL_MARKDOWN_WARNING_TOP_DOTS if marked_down else PREROLL_WARNING_TOP_DOTS
    warning_height = LABEL_HEIGHT_DOTS - LABEL_MARGIN_DOTS - warning_top
    wrapped, warning_font, warning_width, line_spacing, block_height = fit_warning_text(
        warning,
        printable_width,
        warning_height,
    )
    y_offset = LABEL_Y_OFFSET - vertical_offset
    warning_y = warning_top + ((warning_height - block_height) // 2) + y_offset

    z = ["^XA", f"^PW{LABEL_WIDTH_DOTS}", f"^LL{LABEL_HEIGHT_DOTS}", f"^MD{darkness}"]
    z += draw_centered_text(name, LABEL_MARGIN_DOTS + y_offset, title_font)
    z += draw_centered_text(display_price, 52 + y_offset, price_font)
    if marked_down:
        z += draw_strikethrough_text(f"WAS {original_price}", 90 + y_offset, 12)
    z += draw_centered_text(
        "HEALTH WARNING",
        warning_top - PREROLL_WARNING_HEADING_FONT_DOTS - 5 + y_offset,
        PREROLL_WARNING_HEADING_FONT_DOTS,
    )
    z += [
        f"^FO{LABEL_MARGIN_DOTS},{warning_y}",
        f"^FB{printable_width},{wrapped.count(r'\&') + 1},{line_spacing},L,0",
        f"^A0N,{warning_font},{warning_width}",
        f"^FD{wrapped}^FS",
        "^XZ",
    ]
    return "\n".join(z) + "\n"


def build_zpl_2x1_centered(
    name: str,
    price: str,
    warning: str,
    include_warning: bool,
    darkness: int = 20,
    vertical_offset: int = 0,
    marked_down: bool = False,
    original_price: str = "",
    subtitle: str = "",
    size: str = "",
    strain_type: str = "",
    price_input: str = "",
    single_preroll_label: bool = False,
) -> str:
    """Render a price/warning pair or one combined preroll label."""
    # ponytail: keep the legacy flag in the signature for API compatibility; pairs are now mandatory.
    del include_warning
    if single_preroll_label:
        return build_preroll_label_zpl(
            name=name,
            price=price,
            warning=warning,
            darkness=darkness,
            vertical_offset=vertical_offset,
            marked_down=marked_down,
            original_price=original_price,
            price_input=price_input,
        )
    return build_price_label_zpl(
        name=name,
        price=price,
        darkness=darkness,
        vertical_offset=vertical_offset,
        marked_down=marked_down,
        original_price=original_price,
        subtitle=subtitle,
        size=size,
        strain_type=strain_type,
        price_input=price_input,
    ) + build_warning_label_zpl(
        warning=warning,
        darkness=darkness,
        vertical_offset=vertical_offset,
    )


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


def send_raw_zpl_direct(zpl: str) -> None:
    target = f"{DIRECT_PRINTER_HOST}:{DIRECT_PRINTER_PORT}"
    try:
        with socket.create_connection(
            (DIRECT_PRINTER_HOST, DIRECT_PRINTER_PORT),
            timeout=DIRECT_PRINTER_TIMEOUT_SECONDS,
        ) as sock:
            sock.sendall(zpl.encode("ascii", errors="ignore"))
    except Exception as exc:
        raise RuntimeError(f"Direct TCP printer error for '{target}': {exc}") from exc


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


def deliver_zpl(zpl: str, copies: int, printer: str) -> dict[str, Any]:
    result: dict[str, Any] = {
        "printer": printer,
        "copies": copies,
        "success": False,
        "job_ids": [],
        "path": "direct_tcp" if DIRECT_PRINTER_HOST else "windows_queue",
        "target": f"{DIRECT_PRINTER_HOST}:{DIRECT_PRINTER_PORT}" if DIRECT_PRINTER_HOST else printer,
        "fallback_from": "",
        "fallback_error": "",
        "error": "",
    }

    if DIRECT_PRINTER_HOST:
        probe = probe_direct_printer()
        if probe.get("ok"):
            try:
                for _ in range(copies):
                    send_raw_zpl_direct(zpl)
                result["success"] = True
            except Exception as exc:
                result["error"] = (
                    f"{exc} Windows fallback was not attempted because direct transmission started; "
                    "retry manually only after checking the printer."
                )
            return result

        result["fallback_from"] = "direct_tcp"
        result["fallback_error"] = probe.get("error") or "Direct TCP preflight failed."
        if not printer:
            result["error"] = (
                f"Direct TCP preflight failed for {result['target']}: {result['fallback_error']} "
                "No Windows fallback printer is configured."
            )
            return result
        result["path"] = "windows_queue"
        result["target"] = printer

    if not printer:
        result["error"] = "Printer not set. Set ZPL_PRINTER_NAME or choose a Windows printer in the UI."
        return result

    try:
        for _ in range(copies):
            result["job_ids"].append(send_raw_zpl(printer, zpl))
        result["success"] = True
    except Exception as exc:
        result["error"] = str(exc)
    return result


# -----------------------
# FastAPI
# -----------------------
app = FastAPI(title="ZPL Mobile Print Server")


class PrintJob(BaseModel):
    printer: Optional[str] = Field(default=None, description="Windows printer name (optional)")
    name: str = Field(default="", description="Top text")
    price: str = Field(default="", description="Bottom text")
    warning: str = Field(
        default=DEFAULT_WARNING,
        max_length=WARNING_MAX_CHARACTERS,
        description="Required health warning printed on the label",
    )
    include_warning: bool = Field(default=True, description="Deprecated; warning labels are always printed")
    copies: int = Field(default=1, ge=1, le=200)
    darkness: int = Field(default=20, ge=0, le=30)
    vertical_offset: int = Field(default=0, ge=-60, le=60, description="Shift label content in dots: positive up, negative down")
    marked_down: bool = Field(default=False)
    original_price: str = Field(default="")
    subtitle: str = Field(default="")
    size: str = Field(default="")
    strain_type: str = Field(default="")
    price_input: str = Field(default="")
    single_preroll_label: bool = Field(
        default=False,
        description="Print product, price, and health warning on one preroll label",
    )

    @model_validator(mode="after")
    def validate_label(self):
        if self.marked_down and not self.original_price.strip():
            raise ValueError("Original price is required for a marked-down label.")
        warning_section = build_warning_label_sections()["warningSection"]
        if self.single_preroll_label:
            warning_top = PREROLL_MARKDOWN_WARNING_TOP_DOTS if self.marked_down else PREROLL_WARNING_TOP_DOTS
            warning_section = {
                "top": warning_top,
                "height": LABEL_HEIGHT_DOTS - LABEL_MARGIN_DOTS - warning_top,
            }
        fit_warning_text(
            self.warning,
            LABEL_WIDTH_DOTS - (LABEL_MARGIN_DOTS * 2),
            warning_section["height"],
        )
        return self


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
        "direct_printer_host": DIRECT_PRINTER_HOST,
        "direct_printer_port": DIRECT_PRINTER_PORT,
        "print_route": "direct_tcp_with_windows_fallback" if DIRECT_PRINTER_HOST else "windows_queue",
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
        marked_down=job.marked_down,
        original_price=job.original_price,
        subtitle=job.subtitle,
        size=job.size,
        strain_type=job.strain_type,
        price_input=job.price_input,
        single_preroll_label=job.single_preroll_label,
    )
    return zpl


@app.post("/print", response_class=PlainTextResponse)
def print_label(job: PrintJob):
    zpl = build_zpl_2x1_centered(
        name=job.name,
        price=job.price,
        warning=job.warning,
        include_warning=job.include_warning,
        darkness=job.darkness,
        vertical_offset=job.vertical_offset,
        marked_down=job.marked_down,
        original_price=job.original_price,
        subtitle=job.subtitle,
        size=job.size,
        strain_type=job.strain_type,
        price_input=job.price_input,
        single_preroll_label=job.single_preroll_label,
    )
    result = deliver_zpl(zpl, job.copies, selected_printer_name(job.printer))
    record_print_attempt({**result, "source": "label"})
    if not result["success"]:
        status_code = 400 if not result["target"] else 500
        return PlainTextResponse(result["error"], status_code=status_code)

    fallback = (
        f" after direct TCP preflight failed: {result['fallback_error']}"
        if result["fallback_from"]
        else ""
    )
    route = "direct TCP" if result["path"] == "direct_tcp" else "Windows queue"
    quantity = (
        f"{job.copies} preroll label(s)"
        if job.single_preroll_label
        else f"{job.copies} label pair(s) ({job.copies * 2} physical labels)"
    )
    return f"Printed {quantity} via {route} to {result['target']}{fallback}."


@app.post("/diagnostics/test-print", response_class=PlainTextResponse)
def diagnostics_test_print(job: TestPrintJob):
    result = deliver_zpl(build_test_zpl(), 1, selected_printer_name(job.printer))
    record_print_attempt({**result, "source": "test_label"})
    if not result["success"]:
        status_code = 400 if not result["target"] else 500
        return PlainTextResponse(result["error"], status_code=status_code)

    fallback = (
        f" after direct TCP preflight failed: {result['fallback_error']}"
        if result["fallback_from"]
        else ""
    )
    route = "direct TCP" if result["path"] == "direct_tcp" else "Windows queue"
    return f"Sent test label via {route} to {result['target']}{fallback}."


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
    .small-action-btn { flex: 0 0 auto; font-size: 14px; padding: 8px 12px; border-radius: 10px; border: 1px solid #ccc; background: #f2f2f2; }
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

    <label>Windows printer queue</label>
    <select id="printer"></select>
    <div id="routeHint" class="small">Checking the configured print route...</div>

    <div class="panel collapsible-panel">
      <button id="troubleshootingToggle" class="collapsible-trigger" type="button" aria-expanded="false" aria-controls="troubleshootingContent">
        <span>Troubleshooting</span>
        <span id="troubleshootingChevron">Show</span>
      </button>
      <div id="troubleshootingContent" class="collapsible-content" hidden>
        <div class="small">Check the direct TCP primary path, Windows queue fallback, and the route used by recent prints.</div>
        <div class="diagnostic-actions">
          <button id="runDiagnosticsBtn" type="button">Run Diagnostics</button>
          <button id="testLabelBtn" type="button">Send Test Label</button>
        </div>
        <div id="diagnosticsSummary" class="diag-summary"></div>
        <pre id="diagnosticsDetails"></pre>
      </div>
    </div>

    <label>Product name (short title)</label>
    <input id="name" placeholder="e.g. Peach Ringz" autocomplete="off" />

    <label>Product subtitle (optional)</label>
    <input id="subtitle" placeholder="e.g. Cold Cure Live Rosin" autocomplete="off" />

    <div class="row three">
      <div>
        <label>Size</label>
        <select id="size">
          <option value="">Custom</option>
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

    <label>Custom size</label>
    <input id="size_custom" placeholder="e.g. 3.5 grams or 10 pack" />

    <label id="priceLabel" for="price">Custom price / note</label>
    <input id="price" placeholder="e.g. $10.00 or 2 for $15" />
    <div class="small">The label prints size and type separately, with this price as the largest text.</div>

    <div class="toggle">
      <input type="checkbox" id="marked_down" aria-controls="markedDownFields" aria-expanded="false" />
      <label for="marked_down" style="margin:0; font-weight:600;">Marked down price</label>
    </div>

    <div id="markedDownFields" hidden>
      <label for="original_price">Original price</label>
      <input id="original_price" placeholder="e.g. $30.00" />
      <div class="small">Standard labels show “PRICE REDUCED”; preroll labels keep the struck-out original price.</div>
    </div>

    <div class="toggle">
      <input type="checkbox" id="single_preroll_label" />
      <label for="single_preroll_label" style="margin:0; font-weight:600;">Preroll — use one sticker</label>
    </div>
    <div class="small">Combines the product name, price, and health warning on one label. Subtitle, size, and type are omitted to keep the name and price large.</div>

    <label for="warning">Health warning</label>
    <textarea id="warning" maxlength="600" required placeholder="Paste your required warning here...">__DEFAULT_WARNING__</textarea>
    <div class="small">Prerolls use one combined sticker; other products print a separate warning sticker.</div>

    <div class="row">
      <div>
        <label for="copies">Copies</label>
        <input id="copies" type="number" min="1" max="200" value="1" />
        <div class="small">Each preroll copy uses 1 sticker; other copies use 2.</div>
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

function diagnosticSuggestion(health, printerInfo, jobs, network, lastLog, printer) {
  const directConfigured = !!(health.body && health.body.direct_printer_configured);
  if (!printer && !directConfigured) return 'Choose a Windows printer, then run diagnostics again.';
  if (directConfigured && network.body && !network.body.ok && !printer) return 'Direct TCP is unavailable and no Windows fallback queue is selected.';
  if (lastLog && lastLog.success === false) return `Last print failed via ${lastLog.path || 'unknown path'}: ${lastLog.error || 'unknown error'}`;
  if (directConfigured && network.body && !network.body.ok) return 'Direct TCP is unavailable. Printing will use the selected Windows fallback queue.';
  if (!printer) return 'Direct TCP is ready; select a Windows queue if fallback is required.';
  if (!printerInfo.ok || !printerInfo.body.ok) return 'Windows cannot open this printer. Check the printer name, driver, and Windows printer list.';
  const flags = printerInfo.body.status_flags || [];
  const hardFlags = ['offline', 'paused', 'paper_out', 'paper_jam', 'door_open', 'user_intervention', 'error', 'not_available'];
  const activeFlag = hardFlags.find((flag) => flags.includes(flag));
  if (activeFlag) return `Fix the printer state reported by Windows: ${activeFlag.replaceAll('_', ' ')}.`;
  const jobRows = (jobs.body && jobs.body.jobs) || [];
  const stuckJob = jobRows.find((job) => (job.status_flags || []).some((flag) => ['error', 'offline', 'paperout', 'blocked_device_queue', 'user_intervention'].includes(flag)));
  if (stuckJob) return `Clear or restart spooler job ${stuckJob.id}; Windows reports ${stuckJob.status_flags.join(', ')}.`;
  if (jobRows.length > 0) return 'Windows has active queued jobs. If labels are not moving, open the queue and clear stuck jobs.';
  return 'No obvious print-path problem found. Send a test label to verify printer hardware and media.';
}

function renderDiagnostics(results, printer) {
  const summary = document.getElementById('diagnosticsSummary');
  const details = document.getElementById('diagnosticsDetails');
  summary.innerHTML = '';

  const { health, printers, printerInfo, jobs, logs, network } = results;
  const effectivePrinter = printer || (health.body && health.body.configured_default_printer) || '';
  const directConfigured = !!(health.body && health.body.direct_printer_configured);
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
    addDiagnosticRow(summary, 'warn', directConfigured ? 'No Windows fallback queue is selected.' : 'No Windows printer queue is selected.');
  } else {
    if (!printer) {
      addDiagnosticRow(summary, 'ok', `${directConfigured ? 'Windows fallback' : 'Configured default'} queue: ${effectivePrinter}`);
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
    addDiagnosticRow(summary, network.body.ok ? 'ok' : 'warn', network.body.ok ? `Direct TCP primary is reachable at ${network.body.host}:${network.body.port}.` : `Direct TCP primary failed; Windows fallback will be used when available: ${network.body.error || 'unknown error'}`);
  } else {
    addDiagnosticRow(summary, 'warn', 'Direct TCP is not configured; prints use the Windows queue.');
  }

  const lastLog = latestLogForPrinter(logs, effectivePrinter);
  if (lastLog) {
    const route = lastLog.path === 'direct_tcp' ? 'direct TCP' : lastLog.path === 'windows_queue' ? 'Windows queue' : 'unknown path';
    const fallback = lastLog.fallback_from ? ` Fallback reason: ${lastLog.fallback_error || 'direct TCP preflight failed'}.` : '';
    addDiagnosticRow(summary, lastLog.success ? 'ok' : 'bad', lastLog.success ? `Last print succeeded via ${route} to ${lastLog.target || 'unknown target'} at ${formatTimestamp(lastLog.timestamp)}.${fallback}` : `Last print failed via ${route} at ${formatTimestamp(lastLog.timestamp)}: ${lastLog.error || 'unknown error'}`);
  } else {
    addDiagnosticRow(summary, 'warn', 'No print attempts have been logged since the server started.');
  }

  addDiagnosticRow(summary, 'warn', diagnosticSuggestion(health, printerInfo, jobs, network, lastLog, effectivePrinter));

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

async function runDiagnostics(finalMessage = 'Diagnostics complete.') {
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
  setStatus(finalMessage);
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
  await runDiagnostics(text);
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
  const markedDown = source.marked_down === true;
  const singlePrerollLabel = source.single_preroll_label === true;
  const printablePrice = structured
    ? composeBottomLine(size, strainType, priceInput || pricePreset || legacyPrice)
    : legacyPrice;
  return {
    printer: source.printer || null,
    name: String(source.name || ''),
    subtitle: cleanText(source.subtitle),
    price: printablePrice,
    size,
    strain_type: strainType,
    price_preset: pricePreset,
    price_input: priceInput,
    marked_down: markedDown,
    single_preroll_label: singlePrerollLabel,
    original_price: cleanText(source.original_price),
    warning: String(source.warning || ''),
    include_warning: true,
    copies: clampNumber(source.copies ?? copyFallback, LIMITS.copies, copyFallback),
    darkness: clampNumber(source.darkness ?? 20, LIMITS.darkness, 20),
    vertical_offset: clampNumber(source.vertical_offset ?? 0, LIMITS.vertical_offset, 0),
  };
}

async function loadRouteHint() {
  const health = await fetchJson('/health');
  const hint = document.getElementById('routeHint');
  if (!health.ok || !health.body.ok) {
    hint.textContent = 'Could not load print-route settings.';
  } else if (health.body.direct_printer_configured) {
    hint.textContent = `Primary: direct TCP ${health.body.direct_printer_host}:${health.body.direct_printer_port}. This Windows queue is used only if the TCP preflight fails.`;
  } else {
    hint.textContent = 'Primary: this Windows queue. Set ZPL_PRINTER_HOST to use direct TCP first.';
  }
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
    subtitle: document.getElementById('subtitle').value,
    size: document.getElementById('size').value || document.getElementById('size_custom').value,
    strain_type: document.getElementById('strain_type').value,
    price_preset: document.getElementById('price_preset').value,
    price_input: document.getElementById('price').value,
    marked_down: document.getElementById('marked_down').checked,
    single_preroll_label: document.getElementById('single_preroll_label').checked,
    original_price: document.getElementById('original_price').value,
    warning: document.getElementById('warning').value,
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
    single_preroll_label: defaults.single_preroll_label,
    warning: defaults.warning,
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
      single_preroll_label: defaults.single_preroll_label ?? current.single_preroll_label,
      warning: defaults.warning ?? current.warning,
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
    subtitle: job.subtitle ?? '',
    price: job.price ?? '',
    size: job.size ?? '',
    strain_type: job.strain_type ?? '',
    price_preset: job.price_preset ?? '',
    price_input: job.price_input ?? '',
    marked_down: job.marked_down === true,
    single_preroll_label: job.single_preroll_label === true,
    original_price: job.original_price ?? '',
    warning: job.warning ?? '',
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
    normalized.subtitle,
    normalized.price,
    normalized.size,
    normalized.strain_type,
    normalized.price_input,
    normalized.original_price,
    normalized.marked_down ? 'marked down' : '',
    normalized.single_preroll_label ? 'preroll single sticker' : '',
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
  document.getElementById('subtitle').value = normalized.subtitle;
  applySizeToForm(normalized.size);
  document.getElementById('strain_type').value = normalized.strain_type;
  document.getElementById('price_preset').value = normalized.price_preset;
  document.getElementById('price').value = normalized.price_input || normalized.price;
  document.getElementById('marked_down').checked = normalized.marked_down;
  document.getElementById('single_preroll_label').checked = normalized.single_preroll_label;
  document.getElementById('original_price').value = normalized.original_price;
  updateMarkedDownUI();
  document.getElementById('warning').value = normalized.warning;
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

function recordHistory(job, result) {
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
    result,
  };
  deduped.unshift(record);
  writeHistory(deduped.slice(0, MAX_HISTORY));
  rememberStrainType(normalized);
}

async function printJob(job, copies, message) {
  if (!validateJob(job)) {
    return false;
  }
  setStatus(message || `Printing ${jobCountText(job, copies)}...`);
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
    recordHistory(payload, text);
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
    const subtitle = job.subtitle ? ` — ${job.subtitle}` : '';
    main.textContent = `${job.name || '(Unnamed item)'}${subtitle} — ${job.price_input || job.price || '(No price)'}`;

    const meta = document.createElement('div');
    meta.className = 'history-meta';
    const markdownMeta = job.marked_down ? `Marked down from ${job.original_price} | ` : '';
    const formatMeta = job.single_preroll_label ? 'Preroll: 1 sticker | ' : '';
    meta.textContent = `${formatMeta}${markdownMeta}Price: ${job.price_input || job.price || 'No price'} | Warn: ${abbreviatedWarning(job.warning)} | ${formatTimestamp(entry.ts)} | ${entry.result || `Windows printer: ${entry.printer || 'Unknown'}`}`;

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
        printJob(payload, count, `Reprinting ${jobCountText(payload, count)}...`);
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

async function generateZPL() {
  const job = jobPayload();
  if (!validateJob(job)) {
    return;
  }
  setStatus('Generating ZPL...');
  const res = await fetch('/zpl', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(job)
  });
  const zpl = await res.text();
  if (!res.ok) {
    setStatus(zpl);
    return;
  }

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
  setStatus(job.single_preroll_label ? 'Combined preroll ZPL downloaded.' : 'Paired price + warning ZPL downloaded.');
}

function jobCountText(job, copies) {
  return job.single_preroll_label
    ? `${copies} preroll label(s)`
    : `${copies} label pair(s) (${copies * 2} labels)`;
}

async function printLabel() {
  const job = jobPayload();
  await printJob(job, job.copies, `Printing ${jobCountText(job, job.copies)}...`);
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

function updateMarkedDownUI(focusOriginalPrice = false) {
  const markedDown = document.getElementById('marked_down');
  const enabled = markedDown.checked;
  document.getElementById('markedDownFields').hidden = !enabled;
  document.getElementById('priceLabel').textContent = enabled ? 'New price / note' : 'Custom price / note';
  markedDown.setAttribute('aria-expanded', String(enabled));
  if (enabled && focusOriginalPrice) {
    document.getElementById('original_price').focus();
  }
}

function validateJob(job) {
  const normalized = normalizeJob(job);
  if (normalized.marked_down && !normalized.original_price) {
    setStatus('Enter the original price for this marked-down label.');
    document.getElementById('original_price').focus();
    return false;
  }
  if (!normalized.warning.trim()) {
    setStatus('Enter the health warning for the label.');
    document.getElementById('warning').focus();
    return false;
  }
  return true;
}

function applySizeToForm(value) {
  const size = normalizeSize(value);
  const preset = document.getElementById('size');
  const hasMatchingPreset = !!size && Array.from(preset.options).some((opt) => opt.value === size);
  preset.value = hasMatchingPreset ? size : '';
  document.getElementById('size_custom').value = hasMatchingPreset ? '' : size;
}

function syncSizeFromPreset() {
  if (document.getElementById('size').value) {
    document.getElementById('size_custom').value = '';
  }
  saveDefaults(jobPayload());
}

function syncPresetFromSize() {
  const size = normalizeSize(document.getElementById('size_custom').value);
  const preset = document.getElementById('size');
  const hasMatchingPreset = !!size && Array.from(preset.options).some((opt) => opt.value === size);
  preset.value = hasMatchingPreset ? size : '';
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
document.getElementById('name').addEventListener('input', autofillStrainType);
document.getElementById('size').addEventListener('change', syncSizeFromPreset);
document.getElementById('size_custom').addEventListener('input', () => {
  syncPresetFromSize();
  saveDefaults(jobPayload());
});
document.getElementById('price_preset').addEventListener('change', syncPriceFromPreset);
document.getElementById('price').addEventListener('input', () => {
  syncPresetFromPrice();
  saveDefaults(jobPayload());
});
document.getElementById('marked_down').addEventListener('change', (event) => {
  updateMarkedDownUI(event.target.checked);
});
document.getElementById('single_preroll_label').addEventListener('change', () => saveDefaults(jobPayload()));
for (const id of ['printer', 'strain_type', 'warning', 'copies', 'darkness', 'vertical_offset']) {
  document.getElementById(id).addEventListener('change', () => saveDefaults(jobPayload()));
}

updateMarkedDownUI();
loadRouteHint();
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
