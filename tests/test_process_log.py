from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from docx_fixer.models import ProcessSummary
from docx_fixer.process_log import (
    format_indent_settings_log_lines,
    format_numbering_indent_log_lines,
    format_table_log_lines,
    get_table_log_path,
    write_process_log,
    write_table_log_file,
)


class ProcessLogTests(unittest.TestCase):
    def test_numbering_indent_lines_include_text_start_number_start_and_size(self):
        summary = ProcessSummary()
        summary.numbering_measurements["body:1:壹、"] = {
            "section": "body",
            "level": 1,
            "indent_level": 1,
            "prefix": "壹、",
            "text_start_cm": 1.54,
            "number_start_cm": 0.69,
            "number_size_cm": 0.62,
            "font_name": "Microsoft JhengHei",
            "font_size_pt": 12.0,
            "count": 2,
        }

        lines = format_numbering_indent_log_lines(summary)

        self.assertTrue(any("1.54" in line for line in lines))
        self.assertTrue(any("0.69" in line for line in lines))
        self.assertTrue(any("0.62" in line for line in lines))
        self.assertTrue(any("Microsoft JhengHei" in line for line in lines))
        self.assertTrue(any("12" in line for line in lines))

    def test_process_log_writes_numbering_indent_section(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_docx = Path(tmp) / "output.docx"
            log_path = write_process_log(output_docx, ProcessSummary())
            content = log_path.read_text(encoding="utf-8")

        self.assertIn("output.docx", content)

    def test_indent_settings_snapshot_includes_level_two_body_left(self):
        lines = format_indent_settings_log_lines()

        self.assertIn("Indent settings snapshot:", lines)
        self.assertTrue(any("level=2;" in line and "body_left_cm=2.96" in line for line in lines))

    def test_process_log_writes_word_com_skip_section(self):
        summary = ProcessSummary()
        summary.word_com_body_indent_logs.append("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=no_records")

        with tempfile.TemporaryDirectory() as tmp:
            output_docx = Path(tmp) / "output.docx"
            log_path = write_process_log(output_docx, summary)
            content = log_path.read_text(encoding="utf-8")

        self.assertIn("Word COM body indent fix:", content)
        self.assertIn("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=no_records", content)

    def test_table_log_path_uses_output_docx_stem(self):
        output_docx = Path("D:/tmp/sample_fixed.docx")
        self.assertEqual(get_table_log_path(output_docx).name, "sample_fixed_table_log.txt")

    def test_table_log_file_writes_structured_table_records(self):
        summary = ProcessSummary()
        summary.table_log_records.append(
            {
                "part_name": "word/document.xml",
                "table_index": 1,
                "global_table_index": 1,
                "table_name": "壹、估價條件",
                "first_level_heading": "壹、",
                "cell_count": 24,
                "column_count": 4,
                "table_type": "skipped_first_table",
                "action": "skipped",
                "reason": "first table in word/document.xml",
                "special_layout_used": False,
                "layout_fixed": False,
                "color_fixed": False,
                "changed_to_gray": 0,
                "cleared_colors": 0,
            }
        )

        lines = format_table_log_lines(summary)
        self.assertIn("===== Table 1 =====", lines)
        self.assertIn("table_name: 壹、估價條件", lines)
        self.assertIn("first_level_heading: 壹、", lines)
        self.assertIn("table_type: skipped_first_table", lines)
        self.assertIn("special_layout_used: false", lines)

        with tempfile.TemporaryDirectory() as tmp:
            output_docx = Path(tmp) / "sample_fixed.docx"
            table_log_path = write_table_log_file(output_docx, summary)
            content = table_log_path.read_text(encoding="utf-8")

        self.assertEqual(table_log_path.name, "sample_fixed_table_log.txt")
        self.assertIn("===== Table 1 =====", content)
        self.assertIn("first_level_heading: 壹、", content)
        self.assertIn("reason: first table in word/document.xml", content)


if __name__ == "__main__":
    unittest.main()
