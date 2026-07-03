import re
import json
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import call, patch

import main


SAMPLE_WARNING = """THCA PRODUCT • HEMP-DERIVED
Contains less than 0.3% Δ9 THC.
21+ only. Keep out of reach of children.
May cause intoxication when heated.
Do not use while driving or operating heavy machinery.
Consult a physician before use."""


def split_labels(zpl: str) -> list[str]:
    return [chunk + "^XZ\n" for chunk in zpl.split("^XZ\n") if chunk]


def warning_payload(zpl: str) -> str:
    fields = re.findall(r"\^FD(.*?)\^FS", zpl)
    return " ".join(fields[fields.index("HEALTH WARNING") + 1:])


class MarkedDownPriceTests(unittest.TestCase):
    def test_normal_job_returns_price_then_warning_labels(self):
        zpl = main.build_zpl_2x1_centered(
            "Cherry Pie",
            "$25.00",
            SAMPLE_WARNING,
            True,
            size="1 gram",
            strain_type="Hybrid",
            price_input="$25.00",
        )
        price_label, warning_label = split_labels(zpl)

        self.assertEqual(len(split_labels(zpl)), 2)
        self.assertNotIn("PRICE REDUCED", price_label)
        self.assertNotIn("^FDWAS ", price_label)
        self.assertNotIn("^GB", price_label)
        self.assertIn("^FDCHERRY PIE^FS", price_label)
        self.assertIn("^FO16,64\n^FB374,1,0,C,0\n^A0N,20,20\n^FD1g - HYBRID^FS", price_label)
        self.assertIn("^FO16,112\n^FB374,1,0,C,0\n^A0N,42,42\n^FD$25.00^FS", price_label)
        self.assertNotIn("HEALTH WARNING", price_label)
        self.assertIn("^FDHEALTH WARNING^FS", warning_label)
        self.assertNotIn("CHERRY PIE", warning_label)
        self.assertNotIn("$25.00", warning_label)

    def test_details_stay_compact_when_subtitle_or_markdown_needs_the_space(self):
        subtitled = main.build_price_label_zpl(
            "Cherry Pie",
            "$25.00",
            subtitle="Live Rosin",
            size="1 gram",
            strain_type="Hybrid",
        )
        marked_down = main.build_price_label_zpl(
            "Cherry Pie",
            "$20.00",
            marked_down=True,
            original_price="$25.00",
            size="1 gram",
            strain_type="Hybrid",
        )

        for zpl in (subtitled, marked_down):
            self.assertIn("^FO16,78\n^FB374,1,0,C,0\n^A0N,14,14\n^FD1g - HYBRID^FS", zpl)

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
            SAMPLE_WARNING,
            True,
            price_input="Y" * 50,
        )
        price_label = split_labels(zpl)[0]

        self.assertIn("^A0N,14,14\n^FD" + ("X" * 40), price_label)
        self.assertIn("^A0N,11,11\n^FD" + ("Y" * 50), price_label)

    def test_layout_sections_stay_inside_safe_margins(self):
        price_sections = main.build_price_label_sections()
        warning_sections = main.build_warning_label_sections()

        self.assertEqual(price_sections["headerSection"]["top"], main.LABEL_MARGIN_DOTS)
        self.assertLessEqual(
            price_sections["priceSection"]["top"] + price_sections["priceSection"]["height"],
            main.LABEL_HEIGHT_DOTS - main.LABEL_MARGIN_DOTS,
        )
        self.assertGreater(warning_sections["warningSection"]["height"], 0)
        self.assertLessEqual(
            warning_sections["warningSection"]["top"] + warning_sections["warningSection"]["height"],
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
        price_label, warning_label = split_labels(zpl)

        self.assertIn("^PW609", zpl)
        self.assertIn("^FB577,1,0,C,0", price_label)
        self.assertRegex(warning_label, r"\^FO16,\d+\n\^A0N")
        self.assertNotIn("•", zpl)
        self.assertNotIn("Δ", zpl)

    def test_marked_down_job_requires_original_price(self):
        with self.assertRaisesRegex(ValueError, "Original price is required"):
            main.PrintJob(marked_down=True, original_price=" ")

    def test_job_requires_nonblank_warning(self):
        with self.assertRaisesRegex(ValueError, "Health warning is required"):
            main.PrintJob(warning=" ")

    def test_job_rejects_warning_that_cannot_fit_after_sanitizing(self):
        with self.assertRaisesRegex(ValueError, "too long to fit"):
            main.PrintJob(warning="Δ" * main.WARNING_MAX_CHARACTERS)

    def test_legacy_include_warning_false_still_builds_pair(self):
        job = main.PrintJob(warning=SAMPLE_WARNING, include_warning=False)
        zpl = main.build_zpl_2x1_centered(
            job.name,
            job.price,
            job.warning,
            job.include_warning,
        )

        self.assertEqual(len(split_labels(zpl)), 2)
        self.assertIn("HEALTH WARNING", split_labels(zpl)[1])

    def test_warning_is_wrapped_without_losing_text(self):
        warning_label = split_labels(
            main.build_zpl_2x1_centered("", "", SAMPLE_WARNING, True)
        )[1]
        payload = warning_payload(warning_label)
        line_positions = [
            int(y)
            for y in re.findall(
                r"\^FO16,(\d+)\n\^A0N,\d+,\d+\n\^FD",
                warning_label,
            )
        ]

        self.assertEqual(
            payload.split(),
            main.zpl_escape(SAMPLE_WARNING).split(),
        )
        self.assertGreater(len(line_positions), 1)
        self.assertEqual(line_positions, sorted(set(line_positions)))
        self.assertNotIn(r"\&", warning_label)
        self.assertNotIn("•", payload)
        self.assertNotIn("Δ", payload)

    def test_preroll_combines_readable_name_price_and_full_warning(self):
        zpl = main.build_zpl_2x1_centered(
            "Peach Ringz",
            "$10.00",
            SAMPLE_WARNING,
            True,
            price_input="$10.00",
            single_preroll_label=True,
        )
        payload = warning_payload(zpl)

        self.assertEqual(len(split_labels(zpl)), 1)
        self.assertIn("^A0N,30,30\n^FDPEACH RINGZ^FS", zpl)
        self.assertIn("^A0N,42,42\n^FD$10.00^FS", zpl)
        self.assertIn("^FDHEALTH WARNING^FS", zpl)
        self.assertEqual(
            payload.split(),
            main.zpl_escape(SAMPLE_WARNING).split(),
        )


class DeliverZplTests(unittest.TestCase):
    def test_direct_tcp_is_primary(self):
        pair = "^XA^FDPRICE^FS^XZ\n^XA^FDWARNING^FS^XZ\n"
        with (
            patch.object(main, "DIRECT_PRINTER_HOST", "zebra.local"),
            patch.object(main, "probe_direct_printer", return_value={"ok": True}),
            patch.object(main, "send_raw_zpl_direct") as direct,
            patch.object(main, "send_raw_zpl") as windows,
        ):
            result = main.deliver_zpl(pair, 2, "Zebra Queue")

        self.assertTrue(result["success"])
        self.assertEqual(result["path"], "direct_tcp")
        self.assertEqual(direct.call_args_list, [call(pair), call(pair)])
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

    def test_print_response_reports_pairs_and_physical_labels(self):
        result = {
            "success": True,
            "target": "Zebra Queue",
            "path": "windows_queue",
            "fallback_from": None,
            "fallback_error": "",
        }
        with (
            patch.object(main, "deliver_zpl", return_value=result) as deliver,
            patch.object(main, "record_print_attempt"),
        ):
            response = main.print_label(main.PrintJob(copies=3))

        self.assertIn("3 label pair(s) (6 physical labels)", response)
        sent_zpl = deliver.call_args.args[0]
        self.assertEqual(len(split_labels(sent_zpl)), 2)
        deliver.assert_called_once_with(sent_zpl, 3, main.selected_printer_name(None))

    def test_preroll_print_response_reports_one_label_per_copy(self):
        result = {
            "success": True,
            "target": "Zebra Queue",
            "path": "windows_queue",
            "fallback_from": None,
            "fallback_error": "",
        }
        with (
            patch.object(main, "deliver_zpl", return_value=result) as deliver,
            patch.object(main, "record_print_attempt"),
        ):
            response = main.print_label(
                main.PrintJob(copies=3, single_preroll_label=True)
            )

        self.assertIn("3 preroll label(s)", response)
        self.assertEqual(len(split_labels(deliver.call_args.args[0])), 1)


class CatalogDraftTests(unittest.TestCase):
    def test_invalid_flagged_draft_does_not_print(self):
        with self.assertRaisesRegex(ValueError, "require name"):
            main.PrintJob(website_draft=True)

    def test_only_successful_flagged_print_is_recorded(self):
        result = {
            "success": True, "target": "Zebra Queue", "path": "windows_queue",
            "fallback_from": None, "fallback_error": "",
        }
        with TemporaryDirectory() as directory:
            draft = Path(directory) / "catalog.jsonl"
            with (
                patch.object(main, "CATALOG_DRAFT_PATH", draft),
                patch.object(main, "deliver_zpl", return_value=result),
                patch.object(main, "record_print_attempt"),
            ):
                main.print_label(main.PrintJob(
                    name="Blue Dream", category="Flower", size="3 gram",
                    price_input="$25", website_draft=True,
                ))
                main.print_label(main.PrintJob())
            self.assertEqual(len(draft.read_text(encoding="utf-8").splitlines()), 1)

    def test_failed_print_is_not_recorded(self):
        result = {
            "success": False, "target": "Zebra Queue", "path": "windows_queue",
            "fallback_from": None, "fallback_error": "", "error": "offline",
        }
        with TemporaryDirectory() as directory:
            draft = Path(directory) / "catalog.jsonl"
            with (
                patch.object(main, "CATALOG_DRAFT_PATH", draft),
                patch.object(main, "deliver_zpl", return_value=result),
                patch.object(main, "record_print_attempt"),
            ):
                response = main.print_label(main.PrintJob(
                    name="Blue Dream", category="Flower", size="1 gram",
                    price_input="$10", website_draft=True,
                ))
            self.assertEqual(response.status_code, 500)
            self.assertFalse(draft.exists())

    def test_storage_failure_warns_not_to_reprint(self):
        result = {
            "success": True, "target": "Zebra Queue", "path": "windows_queue",
            "fallback_from": None, "fallback_error": "",
        }
        with (
            patch.object(main, "deliver_zpl", return_value=result),
            patch.object(main, "append_catalog_draft", side_effect=OSError("disk full")),
            patch.object(main, "record_print_attempt"),
        ):
            response = main.print_label(main.PrintJob(
                name="Blue Dream", category="Flower", size="1 gram",
                price_input="$10", website_draft=True,
            ))
        self.assertIn("Label printed, draft not saved; do not reprint", response)

    def test_variants_group_and_newest_duplicate_wins(self):
        with TemporaryDirectory() as directory:
            draft = Path(directory) / "catalog.jsonl"
            rows = [
                {"name": "Blue Dream", "category": "Flower", "catalog_group": "", "size": "1 gram", "price": 10},
                {"name": "blue dream", "category": "flower", "catalog_group": "", "size": "1 gram", "price": 12},
                {"name": "Blue Dream", "category": "Flower", "catalog_group": "", "size": "3 grams", "price": 25},
                {"name": "Super Lemon", "category": "Vapes & Carts", "catalog_group": "Live Resin Carts", "size": "1 gram", "price": 30},
            ]
            draft.write_text("\n".join(json.dumps(row) for row in rows) + "\nnot json\n", encoding="utf-8")

            flower, vapes = main.catalog_draft_products(draft)

            self.assertEqual(flower["size_options"], ["1 gram", "3 grams"])
            self.assertEqual(flower["prices"], {"1 gram": 12, "3 grams": 25})
            self.assertEqual(vapes["name"], "Live Resin Carts")
            self.assertEqual(vapes["size_options"], ["Super Lemon"])

    def test_grouped_category_requires_parent(self):
        with self.assertRaisesRegex(ValueError, "parent product"):
            main.PrintJob(
                name="Super Lemon", category="Other", size="1 gram",
                price_input="$20", website_draft=True,
            )

    def test_grouped_category_rejects_unknown_parent(self):
        with self.assertRaisesRegex(ValueError, "valid parent product"):
            main.PrintJob(
                name="Super Lemon", category="Vapes & Carts", size="1 gram",
                price_input="$20", catalog_group="New Group", website_draft=True,
            )

    def test_grouped_categories_accept_every_available_parent(self):
        for category, groups in main.CATALOG_GROUPS.items():
            for group in groups:
                with self.subTest(category=category, group=group):
                    job = main.PrintJob(
                        name="Variant", category=category, size="1 gram",
                        price_input="$20", catalog_group=group, website_draft=True,
                    )
                    self.assertEqual(job.catalog_group, group)

    def test_parent_group_is_a_category_dependent_dropdown(self):
        self.assertIn('<select id="catalog_group" disabled>', main.MOBILE_HTML)
        self.assertIn("select.disabled = groups.length === 0;", main.MOBILE_HTML)
        self.assertIn("select.value = groups.includes(selected) ? selected : '';", main.MOBILE_HTML)

    def test_backfill_imports_once_without_printing(self):
        with TemporaryDirectory() as directory:
            draft = Path(directory) / "catalog.jsonl"
            request = main.CatalogBackfillRequest(candidates=[{
                "source_id": "http://127.0.0.1:8787:history-1",
                "name": "Blue Dream",
                "category": "Flower",
                "size": "3 gram",
                "price_input": "$25",
                "printed_at": "2026-01-01T00:00:00Z",
            }])
            with patch.object(main, "CATALOG_DRAFT_PATH", draft):
                first = main.catalog_draft_backfill(request)
                second = main.catalog_draft_backfill(request)

            self.assertEqual(first, {"imported": 1, "skipped": 0})
            self.assertEqual(second, {"imported": 0, "skipped": 1})
            row = json.loads(draft.read_text(encoding="utf-8").splitlines()[0])
            self.assertEqual(row["size"], "3 grams")
            self.assertEqual(row["source_id"], "http://127.0.0.1:8787:history-1")

    def test_backfill_rejects_invalid_candidate(self):
        with self.assertRaisesRegex(ValueError, "category is invalid"):
            main.CatalogBackfillCandidate(
                source_id="history-1", name="Blue Dream", category="Unknown",
                size="1 gram", price_input="$10",
            )
        with self.assertRaisesRegex(ValueError, "simple number"):
            main.CatalogBackfillCandidate(
                source_id="history-2", name="Blue Dream", category="Flower",
                size="1 gram", price_input="2 for $15",
            )
        with self.assertRaisesRegex(ValueError, "size is required"):
            main.CatalogBackfillCandidate(
                source_id="history-3", name="Blue Dream", category="Flower",
                size="", price_input="$10",
            )

    def test_grouped_backfill_does_not_require_unused_size(self):
        candidate = main.CatalogBackfillCandidate(
            source_id="history-1", name="Super Lemon", category="Vapes & Carts",
            catalog_group="One Gram Carts", size="", price_input="$20",
        )
        self.assertEqual(candidate.draft_row()["size"], "")

    def test_backfill_ui_reads_history_and_requires_confirmation(self):
        self.assertIn("Review Label Backlog", main.MOBILE_HTML)
        self.assertIn("readHistory().filter", main.MOBILE_HTML)
        self.assertIn("catalogMatches(job, products)", main.MOBILE_HTML)
        self.assertIn("'/catalog-draft/backfill'", main.MOBILE_HTML)


if __name__ == "__main__":
    unittest.main()
