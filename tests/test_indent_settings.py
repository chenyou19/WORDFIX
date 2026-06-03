from __future__ import annotations

import tempfile
import unittest
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

    def test_apply_settings_uses_left_and_number_start_to_compute_hanging(self):
        settings = current_indent_settings()
        settings["body"][3]["left_cm"] = 4.50
        settings["body"][3]["number_start_cm"] = 2.25

        apply_indent_settings(settings)

        left_cm, number_start_cm = spec_to_cm_values(TEMPLATE_OUTLINE_INDENTS[3])
        hanging_cm = left_cm - number_start_cm

        self.assertAlmostEqual(left_cm, 4.50, places=2)
        self.assertAlmostEqual(number_start_cm, 2.25, places=2)
        self.assertAlmostEqual(hanging_cm, 2.25, places=2)

    def test_current_settings_expose_updated_body_indent_defaults_for_gui(self):
        settings = current_indent_settings()
        expected = [
            (1.11, -0.04),
            (1.8, 0.69),
            (2.32, 1.32),
            (3.79, 3.05),
            (4.76, 3.53),
            (5.27, 4.52),
            (6.26, 5.02),
            (6.96, 6.2),
            (8.96, 7.72),
        ]

        for row, (left_cm, number_start_cm) in zip(settings["body"], expected):
            with self.subTest(level=row["level"]):
                self.assertAlmostEqual(float(row["left_cm"]), left_cm, places=2)
                self.assertAlmostEqual(float(row["number_start_cm"]), number_start_cm, places=2)

    def test_save_and_load_settings_round_trips_preface_values(self):
        settings = current_indent_settings()
        settings["preface"][2]["left_cm"] = 3.01
        settings["preface"][2]["number_start_cm"] = 2.00

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            save_indent_settings(settings, path)

            PREFACE_OUTLINE_INDENTS.clear()
            self.assertTrue(load_saved_indent_settings(path))

        left_cm, number_start_cm = spec_to_cm_values(PREFACE_OUTLINE_INDENTS[2])
        self.assertAlmostEqual(left_cm, 3.01, places=2)
        self.assertAlmostEqual(number_start_cm, 2.00, places=2)

    def test_invalid_number_start_is_rejected(self):
        settings = current_indent_settings()
        settings["body"][0]["left_cm"] = 1.00
        settings["body"][0]["number_start_cm"] = 1.00

        with self.assertRaises(ValueError):
            apply_indent_settings(settings)


if __name__ == "__main__":
    unittest.main()
