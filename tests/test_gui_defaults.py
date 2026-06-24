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

    def test_table_note_move_defaults_are_forced_false(self):
        defaults = built_in_gui_defaults()
        # The removed "參、不要調整" option is no longer a GUI default.
        self.assertNotIn("skip_chapter_three_adjustments", defaults)
        # The hidden table-note-move options default to False.
        self.assertFalse(defaults["move_table_notes_below"])
        self.assertFalse(defaults["skip_chapter_three_table_notes"])

    def test_old_settings_with_removed_section_three_field_do_not_error(self):
        # A saved file that still has the removed key must load without error
        # and the key is simply ignored.
        normalized = normalize_gui_defaults(
            {"fix_table": True, "skip_chapter_three_adjustments": True}
        )

        self.assertNotIn("skip_chapter_three_adjustments", normalized)
        self.assertIn("move_table_notes_below", normalized)
        self.assertIn("skip_chapter_three_table_notes", normalized)
        self.assertFalse(normalized["move_table_notes_below"])
        self.assertFalse(normalized["skip_chapter_three_table_notes"])

    def test_old_settings_with_note_move_true_are_forced_false_on_load_and_save(self):
        # Even when an old settings file stored True, loading and re-saving must
        # force the hidden table-note-move options back to False.
        old_settings = {
            "move_table_notes_below": True,
            "skip_chapter_three_table_notes": True,
        }

        normalized = normalize_gui_defaults(old_settings)
        self.assertFalse(normalized["move_table_notes_below"])
        self.assertFalse(normalized["skip_chapter_three_table_notes"])

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            # Write a raw file that still carries the True values.
            path.write_text(
                json.dumps({GUI_DEFAULTS_KEY: {**built_in_gui_defaults(), **old_settings}}),
                encoding="utf-8",
            )
            loaded = load_saved_gui_defaults(path)
            # Re-saving must persist False, never True.
            save_gui_defaults(old_settings, path)
            raw = json.loads(path.read_text(encoding="utf-8"))

        self.assertFalse(loaded["move_table_notes_below"])
        self.assertFalse(loaded["skip_chapter_three_table_notes"])
        self.assertFalse(raw[GUI_DEFAULTS_KEY]["move_table_notes_below"])
        self.assertFalse(raw[GUI_DEFAULTS_KEY]["skip_chapter_three_table_notes"])

    def test_table_footer_source_format_default_is_true(self):
        self.assertTrue(built_in_gui_defaults()["enable_table_footer_source_format"])

    def test_table_footer_source_format_round_trip(self):
        settings = built_in_gui_defaults()
        settings["enable_table_footer_source_format"] = False

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            save_gui_defaults(settings, path)
            loaded = load_saved_gui_defaults(path)

        self.assertFalse(loaded["enable_table_footer_source_format"])

    def test_old_settings_without_footer_source_field_get_default(self):
        normalized = normalize_gui_defaults({"fix_table": True})

        self.assertIn("enable_table_footer_source_format", normalized)
        self.assertTrue(normalized["enable_table_footer_source_format"])

    def test_skip_chapter_three_numbering_suffix_cleanup_default_is_true(self):
        self.assertTrue(
            built_in_gui_defaults()["skip_chapter_three_numbering_suffix_cleanup"]
        )

    def test_skip_chapter_three_numbering_suffix_cleanup_round_trip(self):
        settings = built_in_gui_defaults()
        settings["skip_chapter_three_numbering_suffix_cleanup"] = False

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            save_gui_defaults(settings, path)
            loaded = load_saved_gui_defaults(path)

        # It is a real saved option (not forced), so False round-trips.
        self.assertFalse(loaded["skip_chapter_three_numbering_suffix_cleanup"])

    def test_old_settings_without_numbering_suffix_field_get_default_true(self):
        normalized = normalize_gui_defaults({"fix_table": True})
        self.assertIn("skip_chapter_three_numbering_suffix_cleanup", normalized)
        self.assertTrue(normalized["skip_chapter_three_numbering_suffix_cleanup"])

    def test_write_note_debug_log_default_is_false(self):
        self.assertFalse(built_in_gui_defaults()["write_note_debug_log"])

    def test_old_settings_with_write_note_debug_log_true_are_forced_false(self):
        normalized = normalize_gui_defaults({"write_note_debug_log": True})
        self.assertFalse(normalized["write_note_debug_log"])

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            path.write_text(
                json.dumps(
                    {GUI_DEFAULTS_KEY: {**built_in_gui_defaults(), "write_note_debug_log": True}}
                ),
                encoding="utf-8",
            )
            loaded = load_saved_gui_defaults(path)
            save_gui_defaults({"write_note_debug_log": True}, path)
            raw = json.loads(path.read_text(encoding="utf-8"))

        self.assertFalse(loaded["write_note_debug_log"])
        self.assertFalse(raw[GUI_DEFAULTS_KEY]["write_note_debug_log"])

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
