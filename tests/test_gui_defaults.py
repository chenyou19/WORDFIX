from __future__ import annotations

import tempfile
import unittest
import json
from pathlib import Path

from docx_fixer.gui_defaults import (
    GUI_DEFAULTS_KEY,
    built_in_gui_defaults,
    load_saved_gui_defaults,
    normalize_gui_defaults,
    save_gui_defaults,
)


class GuiDefaultsTests(unittest.TestCase):
    def test_missing_file_uses_built_in_defaults(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"

            self.assertEqual(load_saved_gui_defaults(path), built_in_gui_defaults())

    def test_save_and_load_round_trip(self):
        settings = built_in_gui_defaults()
        settings["fix_table"] = False
        settings["skip_chapter_three_table_layout"] = False
        settings["skip_chapter_three_table_color"] = True
        settings["skip_chapter_three_indents"] = True
        settings["skip_nested_tables"] = False
        settings["skip_log_output"] = False

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            saved_path = save_gui_defaults(settings, path)
            loaded = load_saved_gui_defaults(path)
            raw = json.loads(path.read_text(encoding="utf-8"))

        self.assertEqual(saved_path, path)
        self.assertIn(GUI_DEFAULTS_KEY, raw)
        self.assertFalse(loaded["fix_table"])
        self.assertFalse(loaded["skip_chapter_three_table_layout"])
        self.assertTrue(loaded["skip_chapter_three_table_color"])
        self.assertTrue(loaded["skip_chapter_three_indents"])
        self.assertFalse(loaded["skip_nested_tables"])
        self.assertFalse(loaded["skip_log_output"])

    def test_save_preserves_existing_indent_settings_in_shared_file(self):
        existing = {
            "indent_settings": {
                "body": [{"level": 0, "label": "x", "number_start_cm": 1, "hanging_cm": 1, "body_left_cm": 2}],
                "preface": [{"level": 0, "label": "x", "number_start_cm": 1, "hanging_cm": 1, "body_left_cm": 2}],
            }
        }

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            path.write_text(json.dumps(existing), encoding="utf-8")
            save_gui_defaults({"fix_table": False}, path)
            raw = json.loads(path.read_text(encoding="utf-8"))

        self.assertIn("indent_settings", raw)
        self.assertIn(GUI_DEFAULTS_KEY, raw)
        self.assertFalse(raw[GUI_DEFAULTS_KEY]["fix_table"])

    def test_missing_fields_are_filled_from_built_in_defaults(self):
        normalized = normalize_gui_defaults({"fix_color": False})

        self.assertFalse(normalized["fix_color"])
        self.assertEqual(normalized["fix_table"], built_in_gui_defaults()["fix_table"])
        self.assertEqual(
            normalized["skip_chapter_three_indents"],
            built_in_gui_defaults()["skip_chapter_three_indents"],
        )
        self.assertTrue(normalized["skip_nested_tables"])
        self.assertTrue(normalized["skip_log_output"])
        self.assertFalse(normalized["skip_special_color_tables"])
        self.assertFalse(normalized["clear_special_colors_after_skip"])

    def test_old_settings_without_special_color_fields_get_built_in_defaults(self):
        normalized = normalize_gui_defaults(
            {
                "fix_table": True,
                "fix_color": True,
            }
        )

        self.assertIn("skip_special_color_tables", normalized)
        self.assertIn("clear_special_colors_after_skip", normalized)
        self.assertFalse(normalized["skip_special_color_tables"])
        self.assertFalse(normalized["clear_special_colors_after_skip"])

    def test_special_color_checkboxes_round_trip(self):
        settings = built_in_gui_defaults()
        settings["skip_special_color_tables"] = True
        settings["clear_special_colors_after_skip"] = True

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            save_gui_defaults(settings, path)
            loaded = load_saved_gui_defaults(path)

        self.assertTrue(loaded["skip_special_color_tables"])
        self.assertTrue(loaded["clear_special_colors_after_skip"])

    def test_section_three_and_note_defaults_are_off(self):
        defaults = built_in_gui_defaults()
        self.assertFalse(defaults["skip_chapter_three_adjustments"])
        self.assertFalse(defaults["move_table_notes_below"])

    def test_old_settings_without_section_three_and_note_fields_get_built_in_defaults(self):
        normalized = normalize_gui_defaults({"fix_table": True})

        self.assertIn("skip_chapter_three_adjustments", normalized)
        self.assertIn("move_table_notes_below", normalized)
        self.assertFalse(normalized["skip_chapter_three_adjustments"])
        self.assertFalse(normalized["move_table_notes_below"])

    def test_section_three_and_note_checkboxes_round_trip(self):
        settings = built_in_gui_defaults()
        settings["skip_chapter_three_adjustments"] = True
        settings["move_table_notes_below"] = True

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            save_gui_defaults(settings, path)
            loaded = load_saved_gui_defaults(path)

        self.assertTrue(loaded["skip_chapter_three_adjustments"])
        self.assertTrue(loaded["move_table_notes_below"])

    def test_built_in_defaults_skip_log_output(self):
        self.assertTrue(built_in_gui_defaults()["skip_log_output"])

    def test_bool_like_values_are_converted(self):
        normalized = normalize_gui_defaults({
            "fix_table": "false",
            "fix_color": "1",
            "fix_paragraph": 0,
        })

        self.assertFalse(normalized["fix_table"])
        self.assertTrue(normalized["fix_color"])
        self.assertFalse(normalized["fix_paragraph"])

    def test_invalid_bool_value_raises_clear_error(self):
        with self.assertRaisesRegex(ValueError, "fix_table"):
            normalize_gui_defaults({"fix_table": "maybe"})


if __name__ == "__main__":
    unittest.main()
