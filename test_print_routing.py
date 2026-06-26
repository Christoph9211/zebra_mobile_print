import unittest
from unittest.mock import call, patch

import main


class DeliverZplTests(unittest.TestCase):
    def test_direct_tcp_is_primary(self):
        with (
            patch.object(main, "DIRECT_PRINTER_HOST", "zebra.local"),
            patch.object(main, "probe_direct_printer", return_value={"ok": True}),
            patch.object(main, "send_raw_zpl_direct") as direct,
            patch.object(main, "send_raw_zpl") as windows,
        ):
            result = main.deliver_zpl("^XA^XZ", 2, "Zebra Queue")

        self.assertTrue(result["success"])
        self.assertEqual(result["path"], "direct_tcp")
        self.assertEqual(direct.call_args_list, [call("^XA^XZ"), call("^XA^XZ")])
        windows.assert_not_called()

    def test_failed_preflight_falls_back_to_windows(self):
        with (
            patch.object(main, "DIRECT_PRINTER_HOST", "zebra.local"),
            patch.object(main, "probe_direct_printer", return_value={"ok": False, "error": "timed out"}),
            patch.object(main, "send_raw_zpl", return_value=42) as windows,
        ):
            result = main.deliver_zpl("^XA^XZ", 1, "Zebra Queue")

        self.assertTrue(result["success"])
        self.assertEqual(result["path"], "windows_queue")
        self.assertEqual(result["fallback_error"], "timed out")
        windows.assert_called_once_with("Zebra Queue", "^XA^XZ")

    def test_failed_preflight_without_windows_fallback(self):
        with (
            patch.object(main, "DIRECT_PRINTER_HOST", "zebra.local"),
            patch.object(main, "probe_direct_printer", return_value={"ok": False, "error": "refused"}),
            patch.object(main, "send_raw_zpl") as windows,
        ):
            result = main.deliver_zpl("^XA^XZ", 1, "")

        self.assertFalse(result["success"])
        self.assertEqual(result["path"], "direct_tcp")
        self.assertIn("No Windows fallback printer", result["error"])
        windows.assert_not_called()

    def test_direct_send_failure_never_falls_back(self):
        with (
            patch.object(main, "DIRECT_PRINTER_HOST", "zebra.local"),
            patch.object(main, "probe_direct_printer", return_value={"ok": True}),
            patch.object(main, "send_raw_zpl_direct", side_effect=RuntimeError("connection lost")),
            patch.object(main, "send_raw_zpl") as windows,
        ):
            result = main.deliver_zpl("^XA^XZ", 1, "Zebra Queue")

        self.assertFalse(result["success"])
        self.assertEqual(result["path"], "direct_tcp")
        self.assertIn("fallback was not attempted", result["error"])
        windows.assert_not_called()

    def test_windows_only_mode(self):
        with (
            patch.object(main, "DIRECT_PRINTER_HOST", ""),
            patch.object(main, "send_raw_zpl", return_value=7) as windows,
        ):
            result = main.deliver_zpl("^XA^XZ", 1, "Zebra Queue")

        self.assertTrue(result["success"])
        self.assertEqual(result["path"], "windows_queue")
        self.assertEqual(result["job_ids"], [7])
        windows.assert_called_once_with("Zebra Queue", "^XA^XZ")


if __name__ == "__main__":
    unittest.main()
