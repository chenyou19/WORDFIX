from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from docx_fixer.constants import NS, W_NS
from docx_fixer.docx_processor import fix_docx_fast
from docx_fixer.models import ProcessOptions
from docx_fixer.models import ProcessSummary
from docx_fixer.table_cross_page import apply_cross_page_stats
from docx_fixer.table_format import (
    apply_table_format,
    process_table,
    table_cell_count,
    table_column_count,
)
from docx_fixer.xml_utils import qn


def make_table(row_columns: list[int]):
    tbl = etree.Element(qn("tbl"))
    for columns in row_columns:
        tr = etree.SubElement(tbl, qn("tr"))
        for _ in range(columns):
            tc = etree.SubElement(tr, qn("tc"))
            p = etree.SubElement(tc, qn("p"))
            r = etree.SubElement(p, qn("r"))
            t = etree.SubElement(r, qn("t"))
            t.text = "x"
    return tbl


def table_setting(tbl, name: str) -> str | None:
    element = tbl.find(f"w:tblPr/w:{name}", NS)
    if element is None:
        return None
    return element.get(qn("val")) or element.get(qn("type"))


class TableFormatTests(unittest.TestCase):
    def test_table_metrics_handle_inconsistent_rows_and_grid_span(self):
        tbl = make_table([2, 1])
        first_tc = tbl.find("w:tr/w:tc", NS)
        tc_pr = etree.SubElement(first_tc, qn("tcPr"))
        grid_span = etree.SubElement(tc_pr, qn("gridSpan"))
        grid_span.set(qn("val"), "3")

        self.assertEqual(table_cell_count(tbl), 3)
        self.assertEqual(table_column_count(tbl), 4)

    def test_special_layout_sets_autofit_and_right_alignment(self):
        tbl = make_table([2, 2, 2])
        process_table(
            tbl,
            ProcessOptions(True, False, False, False),
            special_layout=True,
        )

        self.assertEqual(table_setting(tbl, "jc"), "right")
        self.assertEqual(table_setting(tbl, "tblW"), "auto")
        self.assertEqual(table_setting(tbl, "tblLayout"), "autofit")

    def test_table_format_does_not_change_font_size(self):
        tbl = make_table([2, 2, 2])
        apply_table_format(tbl)

        self.assertFalse(tbl.xpath(".//w:sz", namespaces=NS))
        self.assertFalse(tbl.xpath(".//w:szCs", namespaces=NS))

    def test_cross_page_stats_are_applied_to_summary(self):
        summary = ProcessSummary()

        apply_cross_page_stats(
            summary,
            {
                "cross_page_tables": 2,
                "cross_page_resolved_tables": 1,
                "cross_page_still_split_tables": 1,
                "adjusted_cell_padding_tables": 2,
                "adjusted_table_spacing_tables": 2,
                "auto_height_tables": 1,
                "moved_next_page_resolved_tables": 1,
                "cannot_avoid_cross_page_tables": 1,
                "failed_cross_page_tables": 3,
            },
        )

        self.assertEqual(summary.cross_page_tables, 2)
        self.assertEqual(summary.cross_page_resolved_tables, 1)
        self.assertEqual(summary.cross_page_still_split_tables, 1)
        self.assertEqual(summary.adjusted_cell_padding_tables, 2)
        self.assertEqual(summary.adjusted_table_spacing_tables, 2)
        self.assertEqual(summary.auto_height_tables, 1)
        self.assertEqual(summary.moved_next_page_resolved_tables, 1)
        self.assertEqual(summary.cannot_avoid_cross_page_tables, 1)
        self.assertEqual(summary.failed_cross_page_tables, 3)

    def test_cross_page_global_error_marks_candidate_tables_failed(self):
        summary = ProcessSummary(tables=5, skipped_first_page_tables=1, skipped_small_tables=2)

        apply_cross_page_stats(summary, {"global_error": True})

        self.assertEqual(summary.failed_cross_page_tables, 2)

    def test_processor_skips_first_page_and_small_tables(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_table([3, 3]))
        body.append(make_table([2, 2, 2]))
        body.append(make_table([2, 2]))
        body.append(make_table([4, 4]))

        with tempfile.TemporaryDirectory() as tmp:
            input_docx = Path(tmp) / "input.docx"
            output_docx = Path(tmp) / "output.docx"
            with ZipFile(input_docx, "w", ZIP_DEFLATED) as zout:
                zout.writestr(
                    "word/document.xml",
                    etree.tostring(
                        document,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    ),
                )

            options = ProcessOptions(True, False, False, False)
            with patch("docx_fixer.docx_processor.get_word_table_start_pages", return_value=[1, 2, 2, 2]):
                summary = fix_docx_fast(input_docx, output_docx, options)

            self.assertEqual(summary.tables, 4)
            self.assertEqual(summary.skipped_first_page_tables, 1)
            self.assertEqual(summary.skipped_small_tables, 1)
            self.assertEqual(summary.special_autofit_right_tables, 1)
            self.assertEqual(summary.normal_processed_tables, 1)

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)

            self.assertIsNone(tables[0].find("w:tblPr", NS))
            self.assertEqual(table_setting(tables[1], "jc"), "right")
            self.assertIsNone(tables[2].find("w:tblPr", NS))
            self.assertEqual(table_setting(tables[3], "jc"), "center")


if __name__ == "__main__":
    unittest.main()
