# Zebra Mobile Label Print

This app runs a small local print server for Zebra ZPL labels.

## Reset the Print Server

If the label page stops responding or printing is acting stuck, use the simple reset shortcut on the office PC:

1. Double-click `Reset Label Print Server.bat`.
2. Wait for the message `Server restarted successfully.`
3. Refresh the label page in the browser.

The reset does not erase saved browser history, defaults, strain/type memory, or label settings. It only restarts the local Python print server.

## Local and Remote URLs

On the office PC, open:

```text
http://127.0.0.1:8787
```

From a trusted Tailscale computer, use the office PC Tailscale address:

```text
http://100.78.175.1:8787
```

Remote users can use the web app diagnostics over Tailscale, but the reset shortcut is intentionally local to the office PC.

## Label Formats

Each price-label job prints two adjacent labels:

1. A full-size product and price label.
2. A full-size health-warning label.

The **Copies** field controls how many sets print, so 3 standard copies use 6 physical labels. Generated ZPL follows the same price-then-warning order. The warning is required; marked-down labels keep the reduced price and struck-through original price on the first label.

Enable **Preroll — use one sticker** to fit the product name, price, and complete health warning on a single label. Preroll copies use one physical sticker each and prioritize large name and price text.

## Troubleshooting Notes

- Website drafts are appended to `data/catalog-draft.jsonl`; set `ZPL_CATALOG_DRAFT_PATH` to override that location.
- Set `ZPL_PRINTER_HOST` to the Zebra printer IP or hostname to make direct TCP the primary print path. Port `9100` is used by default; override it with `ZPL_PRINTER_PORT`.
- Labels default to 2×1 inches at 203 dpi (`406×203` dots). Set `ZPL_LABEL_WIDTH_DOTS=609` for 3×1-inch media; `ZPL_LABEL_HEIGHT_DOTS` and `ZPL_LABEL_Y_OFFSET` are also configurable.
- When direct TCP is configured, the selected `ZPL_PRINTER_NAME` or web-page Windows queue is the fallback. Fallback happens only when the direct TCP preflight fails before label data is sent.
- If direct transmission starts and then fails, the server does not retry through Windows because the printer may already have received the label. Check the printer before retrying manually to avoid duplicate labels.
- If the reset window says another program is using port `8787`, ask someone technical to check that process.
- If the reset starts the server but the health check fails, check `logs/label-print-server.err.log`.
- If the web app is reachable but labels do not print, open `Troubleshooting`; it shows direct TCP status, Windows fallback status, and the route used by the last print.
