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

## Troubleshooting Notes

- Set `ZPL_PRINTER_HOST` to the Zebra printer IP or hostname to make direct TCP the primary print path. Port `9100` is used by default; override it with `ZPL_PRINTER_PORT`.
- When direct TCP is configured, the selected `ZPL_PRINTER_NAME` or web-page Windows queue is the fallback. Fallback happens only when the direct TCP preflight fails before label data is sent.
- If direct transmission starts and then fails, the server does not retry through Windows because the printer may already have received the label. Check the printer before retrying manually to avoid duplicate labels.
- If the reset window says another program is using port `8787`, ask someone technical to check that process.
- If the reset starts the server but the health check fails, check `logs/label-print-server.err.log`.
- If the web app is reachable but labels do not print, open `Troubleshooting`; it shows direct TCP status, Windows fallback status, and the route used by the last print.
