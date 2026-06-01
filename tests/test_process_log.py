from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from docx_fixer.models import ProcessSummary
from docx_fixer.process_log import format_numbering_indent_log_lines, write_process_log


class ProcessLogTests(unittest.TestCase):
    def test_numbering_indent_lines_include_text_start_number_start_and_size(self):
        summary = ProcessSummary()
        summary.numbering_measurements["壹、序言後:1:一、"] = {
            "section": "壹、序言後",
            "level": 1,
            "indent_level": 1,
            "prefix": "一、",
            "text_start_cm": 1.54,
            "number_start_cm": 0.69,
            "number_size_cm": 0.62,
            "font_name": "Microsoft JhengHei",
            "font_size_pt": 12.0,
            "count": 2,
        }

        lines = format_numbering_indent_log_lines(summary)

        self.assertIn("實際編號格式量測紀錄：", lines)
        self.assertIn("壹、序言後：", lines)
        self.assertTrue(any("一、" in line and "文字起點" in line for line in lines))
        self.assertTrue(any("編號起點" in line for line in lines))
        self.assertTrue(any("編號大小" in line for line in lines))
        self.assertTrue(any("量測次數 2" in line for line in lines))

    def test_process_log_writes_numbering_indent_section(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_docx = Path(tmp) / "output.docx"
            log_path = write_process_log(output_docx, ProcessSummary())

            content = log_path.read_text(encoding="utf-8")

        self.assertIn("實際編號格式量測紀錄：", content)
        self.assertIn("沒有量測到可見的手動編號格式。", content)


if __name__ == "__main__":
    unittest.main()
