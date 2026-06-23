from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from docx_fixer.constants import TEMPLATE_OUTLINE_INDENTS
from docx_fixer.models import ProcessSummary
from docx_fixer.process_log import (
    format_indent_settings_log_lines,
    format_heading_suffix_log_lines,
    format_numbering_indent_log_lines,
    format_table_log_lines,
    get_heading_suffix_log_path,
    get_table_log_path,
    write_heading_suffix_log_file,
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
        summary = ProcessSummary(skipped_nested_tables=2)
        with tempfile.TemporaryDirectory() as tmp:
            output_docx = Path(tmp) / "output.docx"
            log_path = write_process_log(output_docx, summary)
            content = log_path.read_text(encoding="utf-8")

        self.assertIn("output.docx", content)
        self.assertIn("因表格中有表格而跳過的表格數：2", content)

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

    def test_heading_suffix_log_path_uses_output_docx_stem(self):
        output_docx = Path("D:/tmp/sample_fixed.docx")
        self.assertEqual(
            get_heading_suffix_log_path(output_docx).name,
            "sample_fixed_heading_suffix_log.txt",
        )

    def test_heading_suffix_log_file_writes_before_after_records(self):
        summary = ProcessSummary()
        summary.heading_suffix_before_records.append(
            {
                "part_name": "word/document.xml",
                "paragraph_index": 12,
                "source": "manual_text",
                "outline_level": 1,
                "heading_text": "一、 標題",
                "number_token": "一、",
                "suffix": "space",
                "space_count": 1,
                "tab_count": 0,
                "raw_separator_repr": "' '",
            }
        )
        summary.heading_suffix_after_records.append(
            {
                "part_name": "word/document.xml",
                "paragraph_index": 12,
                "source": "manual_text",
                "outline_level": 1,
                "heading_text": "一、標題",
                "number_token": "一、",
                "suffix": "nothing",
                "space_count": 0,
                "tab_count": 0,
                "raw_separator_repr": "''",
            }
        )

        lines = format_heading_suffix_log_lines(summary)
        self.assertIn("HEADING_SUFFIX_LOG", lines)
        self.assertIn("changed_to_nothing: 1", lines)
        self.assertIn("suffix_before: space", lines)
        self.assertIn("suffix_after: nothing", lines)
        self.assertIn("change_type: space_to_nothing", lines)

        with tempfile.TemporaryDirectory() as tmp:
            output_docx = Path(tmp) / "sample_fixed.docx"
            log_path = write_heading_suffix_log_file(output_docx, summary)
            content = log_path.read_text(encoding="utf-8")

        self.assertEqual(log_path.name, "sample_fixed_heading_suffix_log.txt")
        self.assertIn("===== SUMMARY BEFORE_FIX =====", content)
        self.assertIn("heading_text_after: 一、標題", content)

    def test_heading_suffix_log_writes_auto_raw_and_effective_suffix(self):
        summary = ProcessSummary()
        summary.heading_suffix_before_records.append(
            {
                "part_name": "word/document.xml",
                "paragraph_index": 27,
                "source": "auto_numbering_xml",
                "outline_level": 4,
                "heading_text": "自動標題",
                "number_token": "%1.",
                "suffix": "missing",
                "raw_suffix": "missing",
                "effective_suffix": "tab",
                "numId": "18",
                "ilvl": 0,
                "numFmt": "decimal",
                "lvlText": "%1.",
                "has_tab_stop": True,
                "tab_pos_twips": "2279",
                "tab_pos_cm": 4.02,
                "left_twips": "2279",
                "hanging_twips": "420",
                "number_start_twips": "1859",
                "left_cm": 4.02,
                "hanging_cm": 0.74,
                "number_start_cm": 3.28,
            }
        )
        summary.heading_suffix_after_records.append(
            {
                "part_name": "word/document.xml",
                "paragraph_index": 27,
                "source": "auto_numbering_xml",
                "outline_level": 4,
                "heading_text": "自動標題",
                "number_token": "%1.",
                "suffix": "nothing",
                "raw_suffix": "nothing",
                "effective_suffix": "nothing",
                "numId": "18",
                "ilvl": 0,
                "numFmt": "decimal",
                "lvlText": "%1.",
                "has_tab_stop": False,
                "tab_pos_twips": None,
                "tab_pos_cm": None,
                "left_twips": "2279",
                "hanging_twips": "420",
                "number_start_twips": "1859",
                "left_cm": 4.02,
                "hanging_cm": 0.74,
                "number_start_cm": 3.28,
            }
        )

        lines = format_heading_suffix_log_lines(summary)

        self.assertIn("suffix_missing: 0", lines[lines.index("===== SUMMARY AFTER_FIX ====="):])
        self.assertIn("suffix_tab: 0", lines[lines.index("===== SUMMARY AFTER_FIX ====="):])
        self.assertIn("suffix_space: 0", lines[lines.index("===== SUMMARY AFTER_FIX ====="):])
        self.assertIn("tab_stop_remaining: 0", lines[lines.index("===== SUMMARY AFTER_FIX ====="):])
        self.assertIn("after_raw_suffix_missing_count: 0", lines)
        self.assertIn("after_effective_suffix_tab_count: 0", lines)
        self.assertIn("after_suffix_space_count: 0", lines)
        self.assertIn("after_tab_stop_remaining_count: 0", lines)
        self.assertIn("after_lvlText_trailing_space_count: 0", lines)
        self.assertIn("raw_suffix_before: missing", lines)
        self.assertIn("raw_suffix_after: nothing", lines)
        self.assertIn("effective_suffix_before: tab", lines)
        self.assertIn("effective_suffix_after: nothing", lines)
        self.assertIn("has_tab_stop_after: false", lines)
        self.assertIn("change_type: missing_effective_tab_to_nothing", lines)

    def test_heading_suffix_log_warns_when_after_fix_auto_suffix_is_dirty(self):
        # outline level 4 is a nothing-suffix level, but this record still has a
        # Tab suffix and a tab stop -> a genuine rule + geometry violation.
        summary = ProcessSummary()
        summary.heading_suffix_after_records.append(
            {
                "part_name": "word/document.xml",
                "paragraph_index": 27,
                "source": "auto_numbering_xml",
                "outline_level": 4,
                "heading_text": "自動標題",
                "number_token": "%5.",
                "suffix": "tab",
                "raw_suffix": "tab",
                "effective_suffix": "tab",
                "numId": "18",
                "ilvl": 0,
                "numFmt": "decimal",
                "lvlText": "%5.",
                "lvlText_has_trailing_space": False,
                "has_tab_stop": True,
            }
        )

        lines = format_heading_suffix_log_lines(summary)

        self.assertIn("WARNING: AFTER_FIX numbering suffix/tab rule violations detected.", lines)
        self.assertIn("WARNING after_suffix_rule_violation_count=1", lines)
        self.assertIn("WARNING after_tab_geometry_violation_count=1", lines)

    def test_heading_suffix_log_does_not_warn_for_legitimate_tab_levels(self):
        # A correct level-3 Tab (suffix=tab, has_tab_stop at spec left, child
        # order ok) must NOT raise a warning.
        summary = ProcessSummary()
        summary.heading_suffix_after_records.append(
            {
                "part_name": "word/document.xml",
                "paragraph_index": 7,
                "source": "auto_numbering_xml",
                "outline_level": 3,
                "heading_text": "自動標題",
                "number_token": "%1.",
                "suffix": "tab",
                "raw_suffix": "tab",
                "effective_suffix": "tab",
                "numId": "18",
                "ilvl": 0,
                "numFmt": "decimal",
                "lvlText": "%1.",
                "lvlText_has_trailing_space": False,
                "has_tab_stop": True,
                "tab_pos_twips": TEMPLATE_OUTLINE_INDENTS[3]["left"],
                "level_child_order_ok": True,
                "ppr_child_order_ok": True,
            }
        )

        lines = format_heading_suffix_log_lines(summary)

        self.assertNotIn("WARNING: AFTER_FIX numbering suffix/tab rule violations detected.", lines)
        self.assertIn("after_expected_tab_count: 1", lines)
        self.assertIn("after_suffix_rule_violation_count: 0", lines)
        self.assertIn("after_tab_geometry_violation_count: 0", lines)

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
        self.assertIn("word_com_autofit_applied: false", lines)
        self.assertIn("word_com_autofit_sequence: none", lines)

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
