import unittest
from pathlib import Path
from unittest.mock import call, patch

import main


SAMPLE_WARNING = """THCA PRODUCT • HEMP-DERIVED
Contains less than 0.3% Δ9 THC.
21+ only. Keep out of reach of children.
May cause intoxication when heated.
Do not use while driving or operating heavy machinery.
Consult a physician before use."""


class MarkedDownPriceTests(unittest.TestCase):
    def test_normal_label_uses_retail_layout_without_sale_elements(self):
        zpl = main.build_zpl_2x1_centered(
            "Cherry Pie",
            "$25.00",
            "",
            False,
            size="1 gram",
            strain_type="Hybrid",
            price_input="$25.00",
        )

        self.assertNotIn("PRICE REDUCED", zpl)
        self.assertNotIn("^FDWAS ", zpl)
        self.assertNotIn("^GB", zpl)
        self.assertIn("^FDCHERRY PIE^FS", zpl)
        self.assertIn("^FD1g - HYBRID^FS", zpl)
        self.assertIn("^A0N,29,29\n^FD$25.00^FS", zpl)

    def test_peach_ringz_label_matches_golden_output(self):
        zpl = main.build_zpl_2x1_centered(
            "Peach Ringz",
            "1g Sativa - $34.99",
            SAMPLE_WARNING,
            True,
            marked_down=True,
            original_price="$40.00",
            subtitle="Cold Cure Live Rosin",
            size="1 gram",
            strain_type="Sativa",
            price_input="$34.99",
        )

        golden = Path(__file__).with_name("testdata") / "peach_ringz_2x1.zpl"
        self.assertEqual(zpl, golden.read_text(encoding="ascii"))

    def test_long_title_and_price_use_smaller_fonts(self):
        zpl = main.build_zpl_2x1_centered(
            "X" * 40,
            "Y" * 50,
            "",
            False,
            price_input="Y" * 50,
        )

        self.assertIn("^A0N,14,14\n^FD" + ("X" * 40), zpl)
        self.assertIn("^A0N,11,11\n^FD" + ("Y" * 50), zpl)

    def test_layout_sections_stay_inside_safe_margins(self):
        sections = main.build_label_sections()

        self.assertEqual(sections["headerSection"]["top"], main.LABEL_MARGIN_DOTS)
        self.assertGreater(sections["warningSection"]["height"], 0)
        self.assertLessEqual(
            sections["warningSection"]["top"] + sections["warningSection"]["height"],
            main.LABEL_HEIGHT_DOTS - main.LABEL_MARGIN_DOTS,
        )

    def test_three_inch_width_uses_full_printable_area(self):
        with patch.object(main, "LABEL_WIDTH_DOTS", 609):
            zpl = main.build_zpl_2x1_centered(
                "Peach Ringz",
                "$34.99",
                SAMPLE_WARNING,
                True,
                subtitle="Cold Cure Live Rosin",
                size="1 gram",
                strain_type="Sativa",
                price_input="$34.99",
            )

        self.assertIn("^PW609", zpl)
        self.assertIn("^FB577,1,0,C,0", zpl)
        self.assertIn("^FB577,4,0,L,0", zpl)
        self.assertNotIn("•", zpl)
        self.assertNotIn("Δ", zpl)

    def test_marked_down_job_requires_original_price(self):
        with self.assertRaisesRegex(ValueError, "Original price is required"):
            main.PrintJob(marked_down=True, original_price=" ")


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
