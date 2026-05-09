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

- If the reset window says another program is using port `8787`, ask someone technical to check that process.
- If the reset starts the server but the health check fails, check `logs/label-print-server.err.log`.
- If the web app is reachable but labels do not print, open the collapsed `Troubleshooting` panel in the web page and run diagnostics.
