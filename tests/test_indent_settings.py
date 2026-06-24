from __future__ import annotations

import tempfile
import unittest
import json
from copy import deepcopy
from pathlib import Path

from docx_fixer.constants import (
    DEFAULT_HEADING_TEXT_START_OFFSET_CM,
    PREFACE_OUTLINE_INDENTS,
    TEMPLATE_OUTLINE_INDENTS,
    make_outline_indent_spec,
)
from docx_fixer.indent_settings import (
    apply_indent_settings,
    built_in_indent_settings,
    current_indent_settings,
    load_saved_indent_settings,
    save_indent_settings,
    spec_to_cm_values,
    twips_to_cm,
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

    def test_canonical_factory_uses_word_number_text_tab_and_body_positions(self):
        spec = make_outline_indent_spec(
            number_start_cm=3.49,
            text_indent_cm=3.99,
            tab_stop_cm=4.84,
            body_left_cm=3.99,
        )

        self.assertAlmostEqual(twips_to_cm(spec["number_start"]), 3.49, places=2)
        self.assertAlmostEqual(twips_to_cm(spec["left"]), 3.99, places=2)
        self.assertAlmostEqual(twips_to_cm(spec["hanging"]), 0.50, places=2)
        self.assertAlmostEqual(twips_to_cm(spec["heading_text_start"]), 4.84, places=2)
        self.assertAlmostEqual(twips_to_cm(spec["body_left"]), 3.99, places=2)

    def test_canonical_factory_rejects_text_indent_at_or_before_number_start(self):
        for text_indent_cm in (3.49, 3.48):
            with self.subTest(text_indent_cm=text_indent_cm):
                with self.assertRaisesRegex(ValueError, "文字縮排必須大於標號起點"):
                    make_outline_indent_spec(
                        number_start_cm=3.49,
                        text_indent_cm=text_indent_cm,
                        tab_stop_cm=4.84,
                        body_left_cm=3.99,
                    )

    def test_apply_settings_uses_number_start_and_hanging_to_compute_heading_left(self):
        settings = current_indent_settings()
        settings["body"][3]["number_start_cm"] = 2.25
        settings["body"][3]["hanging_cm"] = 1.50
        settings["body"][3]["heading_text_start_cm"] = 8.75
        settings["body"][3]["body_left_cm"] = 9.25

        apply_indent_settings(settings)

        spec = TEMPLATE_OUTLINE_INDENTS[3]
        number_start_cm, hanging_cm, heading_text_start_cm, body_left_cm = spec_to_cm_values(spec)

        self.assertAlmostEqual(number_start_cm, 2.25, places=2)
        self.assertAlmostEqual(hanging_cm, 1.50, places=2)
        self.assertAlmostEqual(heading_text_start_cm, 8.75, places=2)
        self.assertAlmostEqual(body_left_cm, 9.25, places=2)
        self.assertAlmostEqual(number_start_cm + hanging_cm, 3.75, places=2)

    def test_current_settings_expose_new_body_indent_defaults_for_gui(self):
        settings = current_indent_settings()
        expected = [
            (-0.04, 1.27, 1.23),
            (0.73, 1.13, 1.86),
            (1.51, 1.48, 2.99),
            (3.49, 0.50, 3.99),
            (3.74, 1.23, 4.96),
            (5.45, 0.50, 5.95),
            (4.70, 1.23, 5.94),
            (5.94, 0.49, 6.85),
            (7.72, 1.24, 8.96),
        ]

        for row, (number_start_cm, hanging_cm, body_left_cm) in zip(settings["body"], expected):
            with self.subTest(level=row["level"]):
                self.assertAlmostEqual(float(row["number_start_cm"]), number_start_cm, places=2)
                self.assertAlmostEqual(float(row["hanging_cm"]), hanging_cm, places=2)
                self.assertAlmostEqual(
                    float(row["heading_text_start_cm"]),
                    body_left_cm + DEFAULT_HEADING_TEXT_START_OFFSET_CM,
                    places=2,
                )
                self.assertAlmostEqual(float(row["body_left_cm"]), body_left_cm, places=2)

    def test_built_in_body_and_preface_heading_text_start_defaults_from_body_left(self):
        settings = built_in_indent_settings()

        for section in ("body", "preface"):
            for row in settings[section]:
                with self.subTest(section=section, level=row["level"]):
                    self.assertAlmostEqual(
                        float(row["heading_text_start_cm"]),
                        float(row["body_left_cm"]) + DEFAULT_HEADING_TEXT_START_OFFSET_CM,
                        places=2,
                    )

    def test_body_left_is_independent_from_heading_left(self):
        settings = current_indent_settings()
        level_8 = settings["body"][7]

        heading_left = float(level_8["number_start_cm"]) + float(level_8["hanging_cm"])
        self.assertAlmostEqual(heading_left, 6.43, places=2)
        self.assertAlmostEqual(float(level_8["body_left_cm"]), 6.85, places=2)
        self.assertNotAlmostEqual(float(level_8["body_left_cm"]), heading_left, places=2)

    def test_heading_left_is_computed_from_number_start_and_hanging(self):
        expected = {
            3: 3.49 + 0.50,
            5: 5.45 + 0.50,
            7: 5.94 + 0.49,
        }

        for level, heading_left_cm in expected.items():
            with self.subTest(level=level):
                spec = TEMPLATE_OUTLINE_INDENTS[level]
                number_start_cm, hanging_cm, _heading_text_start_cm, _body_left_cm = spec_to_cm_values(spec)
                self.assertAlmostEqual(number_start_cm + hanging_cm, heading_left_cm, delta=0.01)

    def test_built_in_settings_ignore_loaded_or_applied_overrides(self):
        settings = current_indent_settings()
        settings["body"][3]["number_start_cm"] = 9.00
        settings["body"][3]["hanging_cm"] = 1.00
        settings["body"][3]["heading_text_start_cm"] = 9.50
        settings["body"][3]["body_left_cm"] = 9.00
        apply_indent_settings(settings)

        builtin_level_four = built_in_indent_settings()["body"][3]

        self.assertAlmostEqual(float(builtin_level_four["number_start_cm"]), 3.49, places=2)
        self.assertAlmostEqual(float(builtin_level_four["hanging_cm"]), 0.50, places=2)
        self.assertAlmostEqual(
            float(builtin_level_four["heading_text_start_cm"]),
            float(builtin_level_four["body_left_cm"]) + DEFAULT_HEADING_TEXT_START_OFFSET_CM,
            places=2,
        )
        self.assertAlmostEqual(float(builtin_level_four["body_left_cm"]), 3.99, places=2)

    def test_save_and_load_settings_round_trips_preface_values(self):
        settings = current_indent_settings()
        settings["preface"][2]["number_start_cm"] = 3.01
        settings["preface"][2]["hanging_cm"] = 1.01
        settings["preface"][2]["heading_text_start_cm"] = 7.77
        settings["preface"][2]["body_left_cm"] = 4.88

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            save_indent_settings(settings, path)
            saved = json.loads(path.read_text(encoding="utf-8"))

            PREFACE_OUTLINE_INDENTS.clear()
            self.assertTrue(load_saved_indent_settings(path))

        saved_indent = saved["indent_settings"]
        self.assertIn("number_start_cm", saved_indent["preface"][2])
        self.assertIn("hanging_cm", saved_indent["preface"][2])
        self.assertIn("heading_text_start_cm", saved_indent["preface"][2])
        self.assertIn("body_left_cm", saved_indent["preface"][2])
        self.assertNotIn("left_cm", saved_indent["preface"][2])

        number_start_cm, hanging_cm, heading_text_start_cm, body_left_cm = spec_to_cm_values(PREFACE_OUTLINE_INDENTS[2])
        self.assertAlmostEqual(number_start_cm, 3.01, places=2)
        self.assertAlmostEqual(hanging_cm, 1.01, places=2)
        self.assertAlmostEqual(heading_text_start_cm, 7.77, places=2)
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
        self.assertAlmostEqual(
            float(first["heading_text_start_cm"]),
            4.50 + DEFAULT_HEADING_TEXT_START_OFFSET_CM,
            places=2,
        )
        self.assertAlmostEqual(float(first["body_left_cm"]), 4.50, places=2)

    def test_saved_settings_without_heading_text_start_are_upgraded(self):
        settings = current_indent_settings()
        old_settings = {"body": [], "preface": []}
        for section in ("body", "preface"):
            for row in settings[section]:
                old_settings[section].append({
                    "level": row["level"],
                    "label": row["label"],
                    "number_start_cm": row["number_start_cm"],
                    "hanging_cm": row["hanging_cm"],
                    "body_left_cm": row["body_left_cm"],
                })

        normalized = apply_indent_settings(old_settings)

        for section in ("body", "preface"):
            for row in normalized[section]:
                with self.subTest(section=section, level=row["level"]):
                    self.assertAlmostEqual(
                        float(row["heading_text_start_cm"]),
                        float(row["body_left_cm"]) + DEFAULT_HEADING_TEXT_START_OFFSET_CM,
                        places=2,
                    )

    def test_heading_text_start_round_trips_without_formula_reset(self):
        settings = current_indent_settings()
        settings["body"][3]["body_left_cm"] = 4.00
        settings["body"][3]["heading_text_start_cm"] = 8.25

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            save_indent_settings(settings, path)
            self.assertTrue(load_saved_indent_settings(path))
            save_indent_settings(current_indent_settings(), path)
            saved = json.loads(path.read_text(encoding="utf-8"))

        row = saved["indent_settings"]["body"][3]
        self.assertAlmostEqual(float(row["body_left_cm"]), 4.00, places=2)
        self.assertAlmostEqual(float(row["heading_text_start_cm"]), 8.25, places=2)
        self.assertNotAlmostEqual(
            float(row["heading_text_start_cm"]),
            float(row["body_left_cm"]) + DEFAULT_HEADING_TEXT_START_OFFSET_CM,
            places=2,
        )

    def test_save_settings_preserves_gui_defaults_in_shared_file(self):
        settings = current_indent_settings()

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            path.write_text(
                json.dumps({"gui_defaults": {"fix_table": False}}, ensure_ascii=False),
                encoding="utf-8",
            )
            save_indent_settings(settings, path)
            saved = json.loads(path.read_text(encoding="utf-8"))

        self.assertIn("indent_settings", saved)
        self.assertIn("gui_defaults", saved)
        self.assertFalse(saved["gui_defaults"]["fix_table"])

    def test_invalid_hanging_is_rejected(self):
        settings = current_indent_settings()
        settings["body"][0]["hanging_cm"] = 0

        with self.assertRaises(ValueError):
            apply_indent_settings(settings)

    def test_invalid_heading_text_start_is_rejected(self):
        settings = current_indent_settings()
        settings["body"][0]["heading_text_start_cm"] = "nan"

        with self.assertRaises(ValueError):
            apply_indent_settings(settings)


if __name__ == "__main__":
    unittest.main()
