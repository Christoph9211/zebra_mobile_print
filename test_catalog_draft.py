import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import main


def successful_delivery():
    return {
        "success": True,
        "path": "windows_queue",
        "target": "Zebra",
        "fallback_from": "",
        "fallback_error": "",
    }


class CatalogDraftTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.path = Path(self.temp_dir.name) / "catalog-draft.jsonl"

    def job(self, **changes):
        values = {
            "name": "Cherry Pie",
            "price": "3g Hybrid - $25",
            "category": "Concentrates",
            "size": "3 gram",
            "price_input": "$25",
            "website_draft": True,
        }
        values.update(changes)
        return main.PrintJob(**values)

    def test_unflagged_and_failed_prints_are_not_recorded(self):
        with (
            patch.object(main, "CATALOG_DRAFT_PATH", self.path),
            patch.object(main, "deliver_zpl", return_value=successful_delivery()),
        ):
            main.print_label(self.job(website_draft=False))
        self.assertFalse(self.path.exists())

        failed = {**successful_delivery(), "success": False, "error": "offline"}
        with (
            patch.object(main, "CATALOG_DRAFT_PATH", self.path),
            patch.object(main, "deliver_zpl", return_value=failed),
        ):
            main.print_label(self.job())
        self.assertFalse(self.path.exists())

    def test_invalid_flagged_draft_still_prints(self):
        with patch.object(main, "deliver_zpl", return_value=successful_delivery()) as deliver:
            response = main.print_label(self.job(price_input="2 for $15"))
        self.assertIn("Label printed, draft not saved", response)
        deliver.assert_called_once()

        with patch.object(main, "deliver_zpl", return_value=successful_delivery()) as deliver:
            response = main.print_label(self.job(category="Other", catalog_group=""))
        self.assertIn("Label printed, draft not saved", response)
        deliver.assert_called_once()

    def test_storage_failure_reports_that_the_label_printed(self):
        with (
            patch.object(main, "deliver_zpl", return_value=successful_delivery()),
            patch.object(main, "append_catalog_candidate", side_effect=OSError("disk full")),
        ):
            response = main.print_label(self.job())
        self.assertIn("Label printed, draft not saved", response)
        self.assertIn("Do not reprint", response)

    def test_variants_aggregate_and_newest_duplicate_wins(self):
        rows = [
            {"name": "Cherry Pie", "category": "Concentrates", "size": "1 gram", "price": 10},
            {"name": "cherry pie", "category": "concentrates", "size": "3 gram", "price": 20},
            {"name": "CHERRY PIE", "category": "Concentrates", "size": "3 grams", "price": 25},
        ]
        for row in rows:
            main.append_catalog_candidate(row, self.path)

        self.assertEqual(
            main.build_catalog_draft(self.path),
            [{
                "name": "CHERRY PIE",
                "category": "Concentrates",
                "size_options": ["1 gram", "3 grams"],
                "prices": {"1 gram": 10, "3 grams": 25},
            }],
        )

    def test_grouped_categories_use_parent_and_printed_name(self):
        main.append_catalog_candidate(
            {
                "name": "Maui Wowie",
                "category": "Vapes & Carts",
                "catalog_group": "One Gram Carts",
                "size": "1 gram",
                "price": 18,
            },
            self.path,
        )
        self.assertEqual(
            main.build_catalog_draft(self.path),
            [{
                "name": "One Gram Carts",
                "category": "Vapes & Carts",
                "size_options": ["Maui Wowie"],
                "prices": {"Maui Wowie": 18},
            }],
        )

    def test_malformed_lines_do_not_break_export(self):
        self.path.write_text(
            "not json\n"
            + json.dumps({"name": "GMO", "category": "Flower", "size": "1/8 oz", "price": 20})
            + "\n{}\n",
            encoding="utf-8",
        )
        self.assertEqual(main.build_catalog_draft(self.path)[0]["name"], "GMO")


if __name__ == "__main__":
    unittest.main()
