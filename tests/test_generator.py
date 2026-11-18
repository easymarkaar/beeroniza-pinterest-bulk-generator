import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from generate_pins import (
    best_contrast,
    build_modifications,
    load_data,
    safe_filename,
    validate_template,
)


class SimpleTableTests(unittest.TestCase):
    def test_excel_and_csv_equal(self):
        self.assertEqual(load_data(ROOT / "pins_example.xlsx"), load_data(ROOT / "pins_example.csv"))

    def test_color_contrast(self):
        self.assertEqual(best_contrast((255, 255, 255), "#202020", "#FFFFFF"), "#202020")
        self.assertEqual(best_contrast((0, 0, 0), "#202020", "#FFFFFF"), "#FFFFFF")

    def test_filename(self):
        self.assertEqual(safe_filename("folder/first.png", "jpg"), "folder_first.jpg")

    def test_template_names(self):
        tpl = {"placeholders": [
            {"key": "background-1", "type": "image"},
            {"key": "hook", "type": "text"},
            {"key": "cta", "type": "text"},
        ]}
        cfg = {"placeholders": {"background": "background-1", "hook": "hook", "cta": "cta"}}
        validate_template(tpl, cfg)

    def test_modifications_use_exact_dynamic_names_and_colors(self):
        import requests
        row = load_data(ROOT / "pins_example.xlsx")[0]
        cfg = {"placeholders": {"background": "background-1", "hook": "hook", "cta": "cta"}}
        mods = build_modifications(row, cfg, requests.Session(), "http://localhost:8502", "not-a-secret", ROOT, {}, dry_run=True)
        self.assertEqual([x["name"] for x in mods], ["background-1", "hook", "cta"])
        self.assertEqual(mods[1]["color"], "#687720")
        self.assertEqual(mods[2]["color"], "#3D2C28")


if __name__ == "__main__":
    unittest.main()
