from __future__ import annotations

import tempfile
import unittest
import json
from copy import deepcopy
from pathlib import Path

from docx_fixer.constants import PREFACE_OUTLINE_INDENTS, TEMPLATE_OUTLINE_INDENTS
from docx_fixer.indent_settings import (
    apply_indent_settings,
    current_indent_settings,
    load_saved_indent_settings,
    save_indent_settings,
    spec_to_cm_values,
)


class IndentSettingsTests(unittest.TestCase):
    def setUp(self) -> None:
        self.original_body = deepcopy(TEMPLATE_OUTLINE_INDENTS)
        self.original_preface = deepcopy(PREFACE_OUTLINE_INDENTS)

    def tearDown(self) -> None:
        TEMPLATE_OUTLINE_INDENTS.clear()
        TEMPLATE_OUTLINE_INDENTS.update(self.original_body)
        PREFACE_OUTLINE_INDENTS.clear()
        PREFACE_OUTLINE_INDENTS.update(self.original_preface)

    def test_apply_settings_uses_number_start_and_hanging_to_compute_heading_left(self):
        settings = current_indent_settings()
        settings["body"][3]["number_start_cm"] = 2.25
        settings["body"][3]["hanging_cm"] = 1.50
        settings["body"][3]["body_left_cm"] = 9.25

        apply_indent_settings(settings)

        spec = TEMPLATE_OUTLINE_INDENTS[3]
        number_start_cm, hanging_cm, body_left_cm = spec_to_cm_values(spec)

        self.assertAlmostEqual(number_start_cm, 2.25, places=2)
        self.assertAlmostEqual(hanging_cm, 1.50, places=2)
        self.assertAlmostEqual(body_left_cm, 9.25, places=2)
        self.assertAlmostEqual(number_start_cm + hanging_cm, 3.75, places=2)

    def test_current_settings_expose_new_body_indent_defaults_for_gui(self):
        settings = current_indent_settings()
        expected = [
            (-0.04, 1.15, 0),
            (0.70, 1.12, 1.83),
            (1.47, 1.48, 2.96),
            (3.20, 0.74, 3.94),
            (3.68, 1.23, 4.91),
            (4.67, 0.74, 5.41),
            (5.16, 1.24, 6.41),
            (6.65, 0.74, 7.11),
            (7.72, 1.24, 8.96),
        ]

        for row, (number_start_cm, hanging_cm, body_left_cm) in zip(settings["body"], expected):
            with self.subTest(level=row["level"]):
                self.assertAlmostEqual(float(row["number_start_cm"]), number_start_cm, places=2)
                self.assertAlmostEqual(float(row["hanging_cm"]), hanging_cm, places=2)
                self.assertAlmostEqual(float(row["body_left_cm"]), body_left_cm, places=2)

    def test_body_left_is_independent_from_heading_left(self):
        settings = current_indent_settings()
        level_8 = settings["body"][7]

        heading_left = float(level_8["number_start_cm"]) + float(level_8["hanging_cm"])
        self.assertAlmostEqual(heading_left, 7.39, places=2)
        self.assertAlmostEqual(float(level_8["body_left_cm"]), 7.11, places=2)
        self.assertNotAlmostEqual(float(level_8["body_left_cm"]), heading_left, places=2)

    def test_save_and_load_settings_round_trips_preface_values(self):
        settings = current_indent_settings()
        settings["preface"][2]["number_start_cm"] = 3.01
        settings["preface"][2]["hanging_cm"] = 1.01
        settings["preface"][2]["body_left_cm"] = 4.88

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            save_indent_settings(settings, path)
            saved = json.loads(path.read_text(encoding="utf-8"))

            PREFACE_OUTLINE_INDENTS.clear()
            self.assertTrue(load_saved_indent_settings(path))

        self.assertIn("number_start_cm", saved["preface"][2])
        self.assertIn("hanging_cm", saved["preface"][2])
        self.assertIn("body_left_cm", saved["preface"][2])
        self.assertNotIn("left_cm", saved["preface"][2])

        number_start_cm, hanging_cm, body_left_cm = spec_to_cm_values(PREFACE_OUTLINE_INDENTS[2])
        self.assertAlmostEqual(number_start_cm, 3.01, places=2)
        self.assertAlmostEqual(hanging_cm, 1.01, places=2)
        self.assertAlmostEqual(body_left_cm, 4.88, places=2)

    def test_old_saved_settings_are_converted_to_new_format(self):
        settings = current_indent_settings()
        old_settings = {"body": [], "preface": []}
        for section in ("body", "preface"):
            for row in settings[section]:
                old_settings[section].append({
                    "level": row["level"],
                    "label": row["label"],
                    "left_cm": 4.50,
                    "number_start_cm": 2.25,
                })

        normalized = apply_indent_settings(old_settings)

        first = normalized["body"][0]
        self.assertAlmostEqual(float(first["number_start_cm"]), 2.25, places=2)
        self.assertAlmostEqual(float(first["hanging_cm"]), 2.25, places=2)
        self.assertAlmostEqual(float(first["body_left_cm"]), 4.50, places=2)

    def test_invalid_hanging_is_rejected(self):
        settings = current_indent_settings()
        settings["body"][0]["hanging_cm"] = 0

        with self.assertRaises(ValueError):
            apply_indent_settings(settings)


if __name__ == "__main__":
    unittest.main()
