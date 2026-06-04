from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from docx_fixer.constants import NS, W_NS
from docx_fixer.docx_processor import fix_docx_fast
from docx_fixer.models import ProcessOptions, ProcessSummary
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


def make_paragraph(text: str):
    p = etree.Element(qn("p"))
    r = etree.SubElement(p, qn("r"))
    t = etree.SubElement(r, qn("t"))
    t.text = text
    return p


def table_pr_element(tbl, name: str):
    return tbl.find(f"w:tblPr/w:{name}", NS)


def table_setting(tbl, name: str, attr: str = "val") -> str | None:
    element = table_pr_element(tbl, name)
    if element is None:
        return None
    return element.get(qn(attr))


class TableFormatTests(unittest.TestCase):
    def test_table_metrics_handle_inconsistent_rows_and_grid_span(self):
        tbl = make_table([2, 1])
        first_tc = tbl.find("w:tr/w:tc", NS)
        tc_pr = etree.SubElement(first_tc, qn("tcPr"))
        grid_span = etree.SubElement(tc_pr, qn("gridSpan"))
        grid_span.set(qn("val"), "3")

        self.assertEqual(table_cell_count(tbl), 3)
        self.assertEqual(table_column_count(tbl), 4)

    def test_special_layout_uses_fixed_geometry_when_available(self):
        tbl = make_table([2, 2, 2])
        process_table(
            tbl,
            ProcessOptions(True, False, False, False),
            special_layout=True,
            special_table_geometry=(720, 4000),
        )

        self.assertEqual(table_setting(tbl, "jc"), "left")
        self.assertEqual(table_setting(tbl, "tblW", "type"), "dxa")
        self.assertEqual(table_setting(tbl, "tblW", "w"), "4000")
        self.assertEqual(table_setting(tbl, "tblInd", "type"), "dxa")
        self.assertEqual(table_setting(tbl, "tblInd", "w"), "720")
        self.assertEqual(table_setting(tbl, "tblLayout", "type"), "fixed")

    def test_apply_table_format_keeps_existing_borders_and_sets_run_size_to_11pt(self):
        tbl = make_table([2, 2, 2])
        tbl_pr = etree.SubElement(tbl, qn("tblPr"))
        tbl_borders = etree.SubElement(tbl_pr, qn("tblBorders"))
        top = etree.SubElement(tbl_borders, qn("top"))
        top.set(qn("val"), "single")
        top.set(qn("color"), "FF0000")
        tbl_grid = etree.SubElement(tbl, qn("tblGrid"))
        grid_col = etree.SubElement(tbl_grid, qn("gridCol"))
        grid_col.set(qn("w"), "1200")
        first_tc_pr = tbl.find("./w:tr/w:tc/w:tcPr", NS)
        if first_tc_pr is None:
            first_tc = tbl.find("./w:tr/w:tc", NS)
            first_tc_pr = etree.SubElement(first_tc, qn("tcPr"))
        tc_w = etree.SubElement(first_tc_pr, qn("tcW"))
        tc_w.set(qn("type"), "dxa")
        tc_w.set(qn("w"), "2400")

        apply_table_format(tbl)

        self.assertEqual(table_setting(tbl, "jc"), "center")
        self.assertEqual(table_setting(tbl, "tblW", "type"), "pct")
        self.assertEqual(table_setting(tbl, "tblW", "w"), "5000")
        self.assertEqual(table_setting(tbl, "tblLayout", "type"), "autofit")
        self.assertEqual(top.get(qn("val")), "single")
        self.assertEqual(top.get(qn("color")), "FF0000")
        self.assertIsNone(tbl.find("./w:tblPr/w:tblBorders/w:left", NS))
        self.assertIsNone(tbl.find("./w:tblPr/w:tblBorders/w:right", NS))
        self.assertIsNone(tbl.find("./w:tblPr/w:tblBorders/w:bottom", NS))
        self.assertIsNone(tbl.find("./w:tblGrid", NS))
        self.assertIsNone(tbl.find(".//w:tcPr/w:tcW", NS))

        for run in tbl.xpath(".//w:r", namespaces=NS):
            r_pr = run.find("w:rPr", NS)
            self.assertIsNotNone(r_pr)
            self.assertEqual(r_pr.find("w:sz", NS).get(qn("val")), "22")
            self.assertEqual(r_pr.find("w:szCs", NS).get(qn("val")), "22")

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

    def test_processor_skips_only_first_document_table_and_still_skips_small_tables(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("壹、估價條件"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("一、比較標的資料"))
        body.append(make_table([4, 4]))
        body.append(make_paragraph("二、小表格"))
        body.append(make_table([2, 2]))
        body.append(make_paragraph("三、一般表格"))
        body.append(make_table([5, 5]))

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

            options = ProcessOptions(True, False, False, False, normalize_with_word_com=False)
            summary = fix_docx_fast(input_docx, output_docx, options)

            self.assertEqual(summary.tables, 4)
            self.assertEqual(summary.skipped_first_page_tables, 1)
            self.assertEqual(summary.skipped_small_tables, 1)
            self.assertEqual(summary.special_autofit_right_tables, 1)
            self.assertEqual(summary.normal_processed_tables, 1)
            self.assertEqual(len(summary.table_log_records), 4)
            self.assertEqual(summary.table_log_records[0]["table_type"], "skipped_first_table")
            self.assertEqual(summary.table_log_records[0]["table_name"], "壹、估價條件")
            self.assertEqual(summary.table_log_records[1]["table_type"], "special_table")
            self.assertEqual(summary.table_log_records[1]["table_name"], "一、比較標的資料")
            self.assertEqual(summary.table_log_records[2]["table_type"], "skipped_small_table")
            self.assertEqual(summary.table_log_records[2]["table_name"], "二、小表格")
            self.assertEqual(summary.table_log_records[3]["table_type"], "normal_table")
            self.assertEqual(summary.table_log_records[3]["table_name"], "三、一般表格")
            self.assertEqual(
                [record["global_table_index"] for record in summary.table_log_records],
                [1, 2, 3, 4],
            )

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)

            self.assertIsNone(tables[0].find("w:tblPr", NS))
            self.assertEqual(table_setting(tables[1], "jc"), "right")
            self.assertEqual(table_setting(tables[1], "tblLayout", "type"), "autofit")
            self.assertIsNone(tables[2].find("w:tblPr", NS))
            self.assertEqual(table_setting(tables[3], "jc"), "center")
            self.assertEqual(table_setting(tables[3], "tblLayout", "type"), "autofit")
            self.assertEqual(table_setting(tables[3], "tblW", "type"), "pct")
            self.assertEqual(table_setting(tables[3], "tblW", "w"), "5000")

    def test_header_first_table_is_not_skipped_by_document_first_table_rule(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("文件第一張表格"))
        body.append(make_table([5, 5]))

        header = etree.Element(qn("hdr"), nsmap={"w": W_NS})
        header.append(make_paragraph("頁首表格"))
        header.append(make_table([5, 5]))

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
                zout.writestr(
                    "word/header1.xml",
                    etree.tostring(
                        header,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    ),
                )

            options = ProcessOptions(True, False, False, False, normalize_with_word_com=False)
            summary = fix_docx_fast(input_docx, output_docx, options)

            self.assertEqual(summary.tables, 2)
            self.assertEqual(summary.skipped_first_page_tables, 1)
            self.assertEqual(summary.normal_processed_tables, 1)
            self.assertEqual(summary.table_log_records[0]["table_type"], "skipped_first_table")
            self.assertEqual(summary.table_log_records[1]["part_name"], "word/header1.xml")
            self.assertEqual(summary.table_log_records[1]["table_type"], "normal_table")
            self.assertEqual(summary.table_log_records[1]["table_name"], "頁首表格")

            with ZipFile(output_docx) as zin:
                header_root = etree.fromstring(zin.read("word/header1.xml"))
            header_tables = header_root.xpath(".//w:tbl", namespaces=NS)
            self.assertEqual(table_setting(header_tables[0], "jc"), "center")

    def test_color_only_tables_are_logged_as_color_only(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("第二張表格"))
        body.append(make_table([5, 5]))

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

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(False, True, False, False, normalize_with_word_com=False),
            )

            self.assertEqual(len(summary.table_log_records), 2)
            self.assertEqual(summary.table_log_records[0]["table_type"], "skipped_first_table")
            self.assertEqual(summary.table_log_records[1]["table_type"], "color_only_table")
            self.assertEqual(summary.table_log_records[1]["action"], "apply_color_only")
            self.assertEqual(
                summary.table_log_records[1]["reason"],
                "fix_table_layout disabled but fix_color enabled",
            )

    def test_table_name_uses_nearest_non_empty_paragraph_before_table(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("   "))
        body.append(make_table([5, 5]))
        long_title = "表格名稱 " + ("很長的說明" * 30)
        body.append(make_paragraph(long_title))
        body.append(make_table([5, 5]))

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

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(False, True, False, False, normalize_with_word_com=False),
            )

            self.assertEqual(summary.table_log_records[1]["table_name"], "第一張表格")
            self.assertTrue(str(summary.table_log_records[2]["table_name"]).startswith("表格名稱"))
            self.assertTrue(str(summary.table_log_records[2]["table_name"]).endswith("..."))


if __name__ == "__main__":
    unittest.main()
