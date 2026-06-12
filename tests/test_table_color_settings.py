from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from docx_fixer.table_color_settings import (
    TABLE_COLOR_SETTINGS_KEY,
    built_in_table_color_settings,
    format_color_list_text,
    load_saved_table_color_settings,
    normalize_hex_color,
    normalize_table_color_settings,
    parse_color_list_text,
    save_table_color_settings,
)


class TableColorSettingsTests(unittest.TestCase):
    def test_built_in_defaults(self):
        settings = built_in_table_color_settings()

        self.assertEqual(settings["keep_colors"], ["D9D9D9", "F2F2F2"])
        self.assertEqual(settings["gray_colors"], ["BFBFBF", "C0C0C0", "A6A6A6", "808080"])
        self.assertEqual(settings["gray_target"], "D9D9D9")
        self.assertEqual(settings["special_color_skip_colors"], [])

    def test_normalize_hex_color_accepts_hash_prefix_and_lowercase(self):
        self.assertEqual(normalize_hex_color("#ddebf7"), "DDEBF7")
        self.assertEqual(normalize_hex_color("DDEBF7"), "DDEBF7")
        self.assertEqual(normalize_hex_color("  c0c0c0  "), "C0C0C0")

    def test_normalize_hex_color_rejects_invalid_values(self):
        for value in ("ZZZZZZ", "12345", "", "auto", None):
            with self.assertRaises(ValueError):
                normalize_hex_color(value)

    def test_parse_color_list_text_supports_newline_and_comma(self):
        self.assertEqual(
            parse_color_list_text("DDEBF7\n#ffffff, fce4d6\nDDEBF7"),
            ["DDEBF7", "FFFFFF", "FCE4D6"],
        )
        self.assertEqual(parse_color_list_text(""), [])
        self.assertEqual(parse_color_list_text("  \n , "), [])

    def test_parse_color_list_text_rejects_invalid_token(self):
        with self.assertRaises(ValueError):
            parse_color_list_text("DDEBF7\nnot-a-color")

    def test_format_color_list_text_round_trip(self):
        colors = ["DDEBF7", "FFFFFF"]
        self.assertEqual(parse_color_list_text(format_color_list_text(colors)), colors)

    def test_normalize_fills_missing_fields_from_built_in(self):
        normalized = normalize_table_color_settings({"keep_colors": ["ffffff"]})

        self.assertEqual(normalized["keep_colors"], ["FFFFFF"])
        self.assertEqual(
            normalized["gray_colors"],
            built_in_table_color_settings()["gray_colors"],
        )
        self.assertEqual(normalized["gray_target"], "D9D9D9")
        self.assertEqual(normalized["special_color_skip_colors"], [])

    def test_normalize_none_returns_built_in(self):
        self.assertEqual(
            normalize_table_color_settings(None),
            built_in_table_color_settings(),
        )

    def test_missing_file_returns_built_in(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            self.assertEqual(
                load_saved_table_color_settings(path),
                built_in_table_color_settings(),
            )

    def test_save_and_load_round_trip(self):
        settings = {
            "keep_colors": ["#ddebf7", "FFFFFF"],
            "gray_colors": ["bfbfbf"],
            "gray_target": "#cccccc",
            "special_color_skip_colors": ["FFFF00", "ff0000"],
        }

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            saved_path = save_table_color_settings(settings, path)
            loaded = load_saved_table_color_settings(path)
            raw = json.loads(path.read_text(encoding="utf-8"))

        self.assertEqual(saved_path, path)
        self.assertIn(TABLE_COLOR_SETTINGS_KEY, raw)
        self.assertEqual(loaded["keep_colors"], ["DDEBF7", "FFFFFF"])
        self.assertEqual(loaded["gray_colors"], ["BFBFBF"])
        self.assertEqual(loaded["gray_target"], "CCCCCC")
        self.assertEqual(loaded["special_color_skip_colors"], ["FFFF00", "FF0000"])

    def test_save_preserves_other_sections_in_shared_file(self):
        existing = {
            "indent_settings": {"body": [], "preface": []},
            "gui_defaults": {"fix_table": False},
        }

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            path.write_text(json.dumps(existing), encoding="utf-8")
            save_table_color_settings({"keep_colors": ["FFFFFF"]}, path)
            raw = json.loads(path.read_text(encoding="utf-8"))

        self.assertIn("indent_settings", raw)
        self.assertIn("gui_defaults", raw)
        self.assertIn(TABLE_COLOR_SETTINGS_KEY, raw)
        self.assertEqual(raw[TABLE_COLOR_SETTINGS_KEY]["keep_colors"], ["FFFFFF"])

    def test_old_settings_file_without_color_settings_uses_built_in(self):
        existing = {"gui_defaults": {"fix_table": True}}

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "indent_defaults.json"
            path.write_text(json.dumps(existing), encoding="utf-8")
            loaded = load_saved_table_color_settings(path)

        self.assertEqual(loaded, built_in_table_color_settings())


if __name__ == "__main__":
    unittest.main()
