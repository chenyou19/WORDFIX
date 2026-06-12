from __future__ import annotations

import unittest

from pathlib import Path
from unittest.mock import Mock, patch

from docx_fixer.gui_app import (
    DEFAULT_SKIP_CHAPTER_THREE_INDENTS,
    DEFAULT_SKIP_CHAPTER_THREE_TABLE_COLOR,
    DEFAULT_SKIP_CHAPTER_THREE_TABLE_LAYOUT,
    DEFAULT_WINDOW_GEOMETRY,
    MIN_WINDOW_SIZE,
    DocxFixerApp,
    write_logs_if_enabled,
)
from docx_fixer.models import ProcessSummary


class FakeProgressBar:
    def __init__(self):
        self.configures = []
        self.started = []
        self.stop_count = 0

    def configure(self, **kwargs):
        self.configures.append(kwargs)

    def start(self, interval):
        self.started.append(interval)

    def stop(self):
        self.stop_count += 1


class FakeProgressVar:
    def __init__(self):
        self.value = None

    def set(self, value):
        self.value = value


class GuiAppTests(unittest.TestCase):
    def test_default_and_min_window_sizes_keep_action_buttons_visible(self):
        width, height = [int(part) for part in DEFAULT_WINDOW_GEOMETRY.split("x")]
        min_width, min_height = MIN_WINDOW_SIZE

        self.assertGreaterEqual(width, 900)
        self.assertGreaterEqual(height, 720)
        self.assertGreaterEqual(min_width, 820)
        self.assertGreaterEqual(min_height, 640)
        self.assertLessEqual(min_width, width)
        self.assertLessEqual(min_height, height)

    def test_progress_animation_helpers_switch_between_indeterminate_and_determinate(self):
        app = DocxFixerApp.__new__(DocxFixerApp)
        app.progress_bar = FakeProgressBar()
        app.progress_var = FakeProgressVar()

        app.start_progress_animation()

        self.assertEqual(app.progress_bar.configures[-1], {"mode": "indeterminate"})
        self.assertEqual(app.progress_bar.started, [10])

        app.stop_progress_animation(100)

        self.assertEqual(app.progress_bar.stop_count, 1)
        self.assertEqual(app.progress_bar.configures[-1], {"mode": "determinate"})
        self.assertEqual(app.progress_var.value, 100)

    def test_chapter_three_gui_options_default_to_checked_and_old_label_removed(self):
        self.assertTrue(DEFAULT_SKIP_CHAPTER_THREE_TABLE_LAYOUT)
        self.assertTrue(DEFAULT_SKIP_CHAPTER_THREE_TABLE_COLOR)
        self.assertFalse(DEFAULT_SKIP_CHAPTER_THREE_INDENTS)

        chapter_three = "".join(chr(code) for code in [0x53C3, 0x3001, 0x50F9, 0x683C, 0x5F62, 0x6210, 0x4E4B, 0x4E3B, 0x8981, 0x56E0, 0x7D20, 0x5206, 0x6790])
        gui_source = Path("docx_fixer/gui_app.py").read_text(encoding="utf-8")
        self.assertIn(f"{chapter_three}\uff1a\u8868\u683c\u7248\u9762\u4e0d\u8abf\u6574", gui_source)
        self.assertIn(f"{chapter_three}\uff1a\u8868\u683c\u984f\u8272\u4e0d\u8abf\u6574", gui_source)
        self.assertIn(f"{chapter_three}\uff1a\u7e2e\u6392\u4e0d\u8abf\u6574", gui_source)
        self.assertNotIn(f"{chapter_three}\uff1a\u8868\u683c\u4e0d\u8abf\u6574", gui_source)

    def test_default_save_buttons_are_wired_in_source(self):
        gui_source = Path("docx_fixer/gui_app.py").read_text(encoding="utf-8")

        self.assertIn("保存成預設樣式", gui_source)
        self.assertIn("command=self.save_indent_defaults", gui_source)
        self.assertIn("保存目前勾選為預設方案", gui_source)
        self.assertIn("command=self.save_gui_defaults", gui_source)
        self.assertIn("還原內建勾選預設", gui_source)
        self.assertIn("command=self.restore_builtin_gui_defaults", gui_source)
        self.assertIn("skip_log_output_var", gui_source)
        self.assertIn("不要輸出 log 檔", gui_source)
        self.assertIn("skip_nested_tables_var", gui_source)
        self.assertIn("表格中有表格不調整", gui_source)

    def test_write_logs_if_enabled_skips_all_external_log_writers(self):
        summary = ProcessSummary()
        warning_callback = Mock()

        with (
            patch("docx_fixer.gui_app.write_process_log") as write_process,
            patch("docx_fixer.gui_app.write_table_log_file") as write_table,
            patch("docx_fixer.gui_app.write_heading_suffix_log_file") as write_heading,
        ):
            result = write_logs_if_enabled(
                Path("output.docx"),
                summary,
                skip_log_output=True,
                warning_callback=warning_callback,
            )

        self.assertEqual(result, (None, None, None))
        write_process.assert_not_called()
        write_table.assert_not_called()
        write_heading.assert_not_called()
        warning_callback.assert_not_called()

    def test_write_logs_if_enabled_writes_all_existing_external_logs(self):
        summary = ProcessSummary()
        warning_callback = Mock()

        with (
            patch("docx_fixer.gui_app.write_process_log", return_value=Path("process_log.txt")) as write_process,
            patch("docx_fixer.gui_app.write_table_log_file", return_value=Path("table_log.txt")) as write_table,
            patch("docx_fixer.gui_app.write_heading_suffix_log_file", return_value=Path("heading_log.txt")) as write_heading,
        ):
            result = write_logs_if_enabled(
                Path("output.docx"),
                summary,
                skip_log_output=False,
                warning_callback=warning_callback,
            )

        self.assertEqual(result, (Path("process_log.txt"), Path("table_log.txt"), Path("heading_log.txt")))
        write_process.assert_called_once_with(Path("output.docx"), summary)
        write_table.assert_called_once_with(Path("output.docx"), summary)
        write_heading.assert_called_once_with(Path("output.docx"), summary)
        warning_callback.assert_not_called()


if __name__ == "__main__":
    unittest.main()
