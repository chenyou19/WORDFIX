from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from docx_fixer.constants import NS, W_NS
from docx_fixer.docx_processor import fix_docx_fast
from docx_fixer.models import ProcessOptions, ProcessSummary
from docx_fixer.table_cross_page import apply_cross_page_stats
from docx_fixer.table_format import (
    apply_special_table_format,
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


def make_shaded_table(row_columns: list[int], fill: str = "808080"):
    tbl = make_table(row_columns)
    first_tc = tbl.find("w:tr/w:tc", NS)
    tc_pr = first_tc.find("w:tcPr", NS)
    if tc_pr is None:
        tc_pr = etree.SubElement(first_tc, qn("tcPr"))
    shd = etree.SubElement(tc_pr, qn("shd"))
    shd.set(qn("val"), "clear")
    shd.set(qn("fill"), fill)
    shd.set(qn("color"), "auto")
    return tbl


def first_table_fill(tbl) -> str | None:
    shd = tbl.find(".//w:tcPr/w:shd", NS)
    if shd is None:
        return None
    return shd.get(qn("fill"))


def set_cell_shading(tbl, row_index: int, cell_index: int, fill: str):
    tr = tbl.findall("w:tr", NS)[row_index]
    tc = tr.findall("w:tc", NS)[cell_index]
    tc_pr = tc.find("w:tcPr", NS)
    if tc_pr is None:
        tc_pr = etree.Element(qn("tcPr"))
        tc.insert(0, tc_pr)
    shd = etree.SubElement(tc_pr, qn("shd"))
    shd.set(qn("val"), "clear")
    shd.set(qn("fill"), fill)
    shd.set(qn("color"), "auto")
    return shd


def cell_fill(tbl, row_index: int, cell_index: int) -> str | None:
    tr = tbl.findall("w:tr", NS)[row_index]
    tc = tr.findall("w:tc", NS)[cell_index]
    shd = tc.find("w:tcPr/w:shd", NS)
    if shd is None:
        return None
    return shd.get(qn("fill"))


def table_layout_signature(tbl) -> list[bytes]:
    layout_nodes = tbl.xpath(
        ".//w:tblPr | .//w:tblGrid | .//w:trPr | "
        ".//w:tcPr/w:tcW | .//w:tcPr/w:vAlign | .//w:pPr | .//w:rPr",
        namespaces=NS,
    )
    return [etree.tostring(node) for node in layout_nodes]


def make_nested_table_document():
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    body.append(make_paragraph("first table"))
    body.append(make_table([5, 5]))
    body.append(make_paragraph("nested table"))

    outer = make_table([3, 3])
    set_cell_shading(outer, 0, 0, "BFBFBF")
    inner = make_shaded_table([3, 3], fill="FF0000")
    first_tc = outer.find("w:tr/w:tc", NS)
    first_tc.append(inner)
    body.append(outer)
    return document


def make_paragraph(
    text: str,
    *,
    style: str | None = None,
    num_id: str | None = None,
    ilvl: int | None = None,
):
    p = etree.Element(qn("p"))
    if style is not None or num_id is not None:
        p_pr = etree.SubElement(p, qn("pPr"))
        if style is not None:
            p_style = etree.SubElement(p_pr, qn("pStyle"))
            p_style.set(qn("val"), style)
        if num_id is not None:
            num_pr = etree.SubElement(p_pr, qn("numPr"))
            if ilvl is not None:
                ilvl_el = etree.SubElement(num_pr, qn("ilvl"))
                ilvl_el.set(qn("val"), str(ilvl))
            num_id_el = etree.SubElement(num_pr, qn("numId"))
            num_id_el.set(qn("val"), num_id)
    r = etree.SubElement(p, qn("r"))
    t = etree.SubElement(r, qn("t"))
    t.text = text
    return p


def make_legal_traditional_numbering_xml(*num_ids: str) -> bytes:
    numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    abstract_num = etree.SubElement(numbering, qn("abstractNum"))
    abstract_num.set(qn("abstractNumId"), "10")
    lvl = etree.SubElement(abstract_num, qn("lvl"))
    lvl.set(qn("ilvl"), "0")
    start = etree.SubElement(lvl, qn("start"))
    start.set(qn("val"), "1")
    num_fmt = etree.SubElement(lvl, qn("numFmt"))
    num_fmt.set(qn("val"), "ideographLegalTraditional")
    lvl_text = etree.SubElement(lvl, qn("lvlText"))
    lvl_text.set(qn("val"), "%1、")

    for num_id in num_ids:
        num = etree.SubElement(numbering, qn("num"))
        num.set(qn("numId"), num_id)
        abstract_num_id = etree.SubElement(num, qn("abstractNumId"))
        abstract_num_id.set(qn("val"), "10")

    return etree.tostring(
        numbering,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def make_style_numbering_xml(style_id: str, *, num_id: str, ilvl: int = 0) -> bytes:
    styles = etree.Element(qn("styles"), nsmap={"w": W_NS})
    style = etree.SubElement(styles, qn("style"))
    style.set(qn("type"), "paragraph")
    style.set(qn("styleId"), style_id)
    p_pr = etree.SubElement(style, qn("pPr"))
    num_pr = etree.SubElement(p_pr, qn("numPr"))
    ilvl_el = etree.SubElement(num_pr, qn("ilvl"))
    ilvl_el.set(qn("val"), str(ilvl))
    num_id_el = etree.SubElement(num_pr, qn("numId"))
    num_id_el.set(qn("val"), num_id)
    return etree.tostring(
        styles,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def make_sect_pr(page_width: int, left_margin: int, right_margin: int):
    sect_pr = etree.Element(qn("sectPr"))
    pg_sz = etree.SubElement(sect_pr, qn("pgSz"))
    pg_sz.set(qn("w"), str(page_width))
    pg_sz.set(qn("h"), "16838")
    pg_mar = etree.SubElement(sect_pr, qn("pgMar"))
    pg_mar.set(qn("top"), "1440")
    pg_mar.set(qn("bottom"), "1440")
    pg_mar.set(qn("left"), str(left_margin))
    pg_mar.set(qn("right"), str(right_margin))
    return sect_pr


def make_section_break_paragraph(page_width: int, left_margin: int, right_margin: int):
    p = etree.Element(qn("p"))
    p_pr = etree.SubElement(p, qn("pPr"))
    p_pr.append(make_sect_pr(page_width, left_margin, right_margin))
    return p


def make_indented_paragraph(text: str, left_twips: str):
    p = make_paragraph(text)
    p_pr = etree.Element(qn("pPr"))
    p.insert(0, p_pr)
    ind = etree.SubElement(p_pr, qn("ind"))
    ind.set(qn("left"), left_twips)
    return p


def add_wide_fixed_widths(tbl, grid_width: str = "8000"):
    tbl_grid = etree.SubElement(tbl, qn("tblGrid"))
    for _ in range(table_column_count(tbl)):
        grid_col = etree.SubElement(tbl_grid, qn("gridCol"))
        grid_col.set(qn("w"), grid_width)
    for tc in tbl.xpath(".//w:tc", namespaces=NS):
        tc_pr = tc.find("w:tcPr", NS)
        if tc_pr is None:
            tc_pr = etree.Element(qn("tcPr"))
            tc.insert(0, tc_pr)
        tc_w = etree.SubElement(tc_pr, qn("tcW"))
        tc_w.set(qn("type"), "dxa")
        tc_w.set(qn("w"), grid_width)
    return tbl


def grid_column_widths(tbl) -> list[int]:
    return [
        int(grid_col.get(qn("w")))
        for grid_col in tbl.findall("w:tblGrid/w:gridCol", NS)
    ]


def cell_widths(tbl) -> list[int]:
    return [
        int(tc_w.get(qn("w")))
        for tc_w in tbl.xpath(".//w:tc/w:tcPr/w:tcW", namespaces=NS)
    ]


def set_fixed_widths_in_docx(docx_path, table_indices: list[int]) -> None:
    """Simulate a partial Word COM run that saved fixed table widths."""
    with ZipFile(docx_path) as zin:
        items = zin.infolist()
        data = {item.filename: zin.read(item.filename) for item in items}

    root = etree.fromstring(data["word/document.xml"])
    tables = root.xpath(".//w:tbl", namespaces=NS)
    for table_index in table_indices:
        tbl = tables[table_index - 1]
        tbl_pr = tbl.find("w:tblPr", NS)
        if tbl_pr is None:
            tbl_pr = etree.Element(qn("tblPr"))
            tbl.insert(0, tbl_pr)
        tbl_w = tbl_pr.find("w:tblW", NS)
        if tbl_w is None:
            tbl_w = etree.SubElement(tbl_pr, qn("tblW"))
        tbl_w.set(qn("type"), "dxa")
        tbl_w.set(qn("w"), "20000")
        tbl_grid = etree.SubElement(tbl, qn("tblGrid"))
        for _ in range(table_column_count(tbl)):
            grid_col = etree.SubElement(tbl_grid, qn("gridCol"))
            grid_col.set(qn("w"), "4000")
        first_tc = tbl.find("w:tr/w:tc", NS)
        tc_pr = first_tc.find("w:tcPr", NS)
        if tc_pr is None:
            tc_pr = etree.Element(qn("tcPr"))
            first_tc.insert(0, tc_pr)
        tc_w = etree.SubElement(tc_pr, qn("tcW"))
        tc_w.set(qn("type"), "dxa")
        tc_w.set(qn("w"), "4000")

    data["word/document.xml"] = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    with ZipFile(docx_path, "w", ZIP_DEFLATED) as zout:
        for item in items:
            zout.writestr(item, data[item.filename])


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

    def test_apply_special_table_format_clears_old_widths_and_rebuilds_grid(self):
        tbl = add_wide_fixed_widths(make_table([3, 3]), grid_width="8000")

        apply_special_table_format(tbl, left_indent_twips=720, width_twips=7586)

        widths = grid_column_widths(tbl)
        self.assertEqual(len(widths), 3)
        self.assertNotIn(8000, widths)
        self.assertEqual(sum(widths), 7586)
        self.assertLessEqual(sum(widths), 7586)

        tbl_pr = tbl.find("w:tblPr", NS)
        tbl_grid = tbl.find("w:tblGrid", NS)
        self.assertEqual(tbl.index(tbl_grid), tbl.index(tbl_pr) + 1)

        per_cell_widths = cell_widths(tbl)
        self.assertEqual(len(per_cell_widths), 6)
        self.assertNotIn(8000, per_cell_widths)
        for tc_w in tbl.xpath(".//w:tc/w:tcPr/w:tcW", namespaces=NS):
            self.assertEqual(tc_w.get(qn("type")), "dxa")
        for row_index, tr in enumerate(tbl.findall("w:tr", NS)):
            row_widths = [
                int(tc.find("w:tcPr/w:tcW", NS).get(qn("w")))
                for tc in tr.findall("w:tc", NS)
            ]
            self.assertEqual(row_widths, widths, f"row {row_index} cell widths mismatch")

    def test_apply_special_table_format_grid_span_uses_column_width_sum(self):
        tbl = make_table([2, 1])
        spanning_tc = tbl.findall("w:tr", NS)[1].find("w:tc", NS)
        tc_pr = etree.SubElement(spanning_tc, qn("tcPr"))
        grid_span = etree.SubElement(tc_pr, qn("gridSpan"))
        grid_span.set(qn("val"), "2")

        apply_special_table_format(tbl, left_indent_twips=0, width_twips=4001)

        widths = grid_column_widths(tbl)
        self.assertEqual(widths, [2001, 2000])
        self.assertEqual(sum(widths), 4001)

        first_row_widths = [
            int(tc.find("w:tcPr/w:tcW", NS).get(qn("w")))
            for tc in tbl.findall("w:tr", NS)[0].findall("w:tc", NS)
        ]
        self.assertEqual(first_row_widths, [2001, 2000])
        spanning_width = int(spanning_tc.find("w:tcPr/w:tcW", NS).get(qn("w")))
        self.assertEqual(spanning_width, 4001)

    def test_special_table_right_edge_matches_page_text_width(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        body.append(make_indented_paragraph("表格上方說明", "720"))
        body.append(add_wide_fixed_widths(make_table([3, 3, 3]), grid_width="8000"))
        body.append(make_sect_pr(11906, 1800, 1800))

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

            text_width_twips = 11906 - 1800 - 1800
            record = summary.table_log_records[1]
            self.assertEqual(record["table_type"], "special_table")
            self.assertEqual(record["special_left_indent_twips"], 720)
            self.assertEqual(record["special_width_twips"], text_width_twips - 720)
            self.assertEqual(record["special_text_width_twips"], text_width_twips)
            self.assertEqual(record["special_right_edge_twips"], text_width_twips)
            self.assertEqual(record["special_overflow_twips"], 0)

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

            self.assertEqual(table_setting(tbl, "jc"), "left")
            self.assertEqual(table_setting(tbl, "tblLayout", "type"), "fixed")
            tbl_ind = int(table_setting(tbl, "tblInd", "w"))
            tbl_w = int(table_setting(tbl, "tblW", "w"))
            self.assertEqual(tbl_ind, 720)
            self.assertLessEqual(tbl_ind + tbl_w, text_width_twips)
            self.assertEqual(tbl_ind + tbl_w, text_width_twips)

            widths = grid_column_widths(tbl)
            self.assertEqual(len(widths), 3)
            self.assertNotIn(8000, widths)
            self.assertLessEqual(sum(widths), tbl_w)
            self.assertNotIn(8000, cell_widths(tbl))

    def test_special_table_uses_current_section_properties_not_previous_section(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        # Section 1 ends here with a much wider page than section 2.
        body.append(make_section_break_paragraph(20000, 1000, 1000))
        body.append(make_indented_paragraph("表格上方說明", "720"))
        body.append(make_table([3, 3, 3]))
        body.append(make_sect_pr(11906, 1800, 1800))

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

            section_two_text_width = 11906 - 1800 - 1800
            record = summary.table_log_records[1]
            self.assertEqual(record["table_type"], "special_table")
            self.assertEqual(record["special_text_width_twips"], section_two_text_width)
            self.assertEqual(record["special_width_twips"], section_two_text_width - 720)
            self.assertEqual(record["special_overflow_twips"], 0)

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

            tbl_ind = int(table_setting(tbl, "tblInd", "w"))
            tbl_w = int(table_setting(tbl, "tblW", "w"))
            self.assertEqual(tbl_ind, 720)
            self.assertEqual(tbl_w, section_two_text_width - 720)
            self.assertLessEqual(tbl_ind + tbl_w, section_two_text_width)

    def test_special_table_falls_back_when_left_indent_reaches_page_right_edge(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        body.append(make_indented_paragraph("表格上方說明", "9000"))
        body.append(make_table([3, 3, 3]))
        body.append(make_sect_pr(11906, 1800, 1800))

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

            record = summary.table_log_records[1]
            self.assertEqual(record["table_type"], "special_table")
            self.assertIsNone(record["special_left_indent_twips"])
            self.assertIsNone(record["special_width_twips"])
            self.assertIsNone(record["special_text_width_twips"])
            self.assertIsNone(record["special_right_edge_twips"])
            self.assertIsNone(record["special_overflow_twips"])

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

            self.assertEqual(table_setting(tbl, "jc"), "right")
            self.assertEqual(table_setting(tbl, "tblW", "type"), "auto")
            self.assertEqual(table_setting(tbl, "tblLayout", "type"), "autofit")
            self.assertIsNone(table_pr_element(tbl, "tblInd"))

    def test_special_table_clamps_negative_left_indent_to_zero(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        body.append(make_indented_paragraph("表格上方說明", "-200"))
        body.append(make_table([3, 3, 3]))
        body.append(make_sect_pr(11906, 1800, 1800))

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

            text_width_twips = 11906 - 1800 - 1800
            record = summary.table_log_records[1]
            self.assertEqual(record["special_left_indent_twips"], 0)
            self.assertEqual(record["special_width_twips"], text_width_twips)
            self.assertEqual(record["special_overflow_twips"], 0)

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

            self.assertEqual(table_setting(tbl, "tblInd", "w"), "0")
            self.assertEqual(table_setting(tbl, "tblW", "w"), str(text_width_twips))

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

    def test_normal_tables_are_queued_for_word_com_autofit_only(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("first table"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("special table"))
        body.append(make_table([4, 4]))
        body.append(make_paragraph("small table"))
        body.append(make_table([2, 2]))
        body.append(make_paragraph("normal table"))
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

            captured: dict[str, object] = {}

            def fake_table_autofit(docx_path, records, stop=None):
                captured["docx_path"] = docx_path
                captured["records"] = list(records)
                return (
                    ["WORD_COM_TABLE_AUTOFIT_APPLIED global_table_index=4 sequence=content_then_window"],
                    {4},
                    set(),
                )

            with patch(
                "docx_fixer.docx_processor.apply_table_autofit_with_word_com",
                side_effect=fake_table_autofit,
            ) as autofit:
                summary = fix_docx_fast(
                    input_docx,
                    output_docx,
                    ProcessOptions(
                        fix_table_layout=True,
                        fix_color=False,
                        fix_paragraph=False,
                        normalize_with_word_com=True,
                    ),
                )

        autofit.assert_called_once()
        self.assertEqual(captured["docx_path"], output_docx)
        self.assertEqual(
            [(record["global_table_index"], record["table_name"]) for record in captured["records"]],
            [(4, "normal table")],
        )
        self.assertEqual(summary.word_com_table_autofit_records, captured["records"])
        self.assertEqual(
            [record["table_type"] for record in summary.table_log_records],
            ["skipped_first_table", "special_table", "skipped_small_table", "normal_table"],
        )
        self.assertFalse(summary.table_log_records[0]["word_com_autofit_applied"])
        self.assertFalse(summary.table_log_records[1]["word_com_autofit_applied"])
        self.assertFalse(summary.table_log_records[2]["word_com_autofit_applied"])
        self.assertTrue(summary.table_log_records[3]["word_com_autofit_applied"])
        self.assertEqual(summary.table_log_records[3]["word_com_autofit_sequence"], "content_then_window")
        self.assertEqual(summary.table_log_records[3]["word_com_autofit_status"], "word_com")
        self.assertEqual(summary.word_com_table_autofit_applied_count, 1)
        self.assertEqual(summary.word_com_table_autofit_fallback_count, 0)
        self.assertEqual(summary.word_com_table_autofit_failed_count, 0)

    def test_word_com_autofit_failure_triggers_xml_fallback(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("一般表格"))
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

            def fake_table_autofit(docx_path, records, stop=None):
                # Word COM saved fixed widths into the document before failing.
                set_fixed_widths_in_docx(docx_path, [int(r["table_index"]) for r in records])
                failed = {int(r["global_table_index"]) for r in records}
                return (
                    ["WORD_COM_TABLE_AUTOFIT_EXCEPTION type=Test message=word_not_installed"],
                    set(),
                    failed,
                )

            with patch(
                "docx_fixer.docx_processor.apply_table_autofit_with_word_com",
                side_effect=fake_table_autofit,
            ):
                summary = fix_docx_fast(
                    input_docx,
                    output_docx,
                    ProcessOptions(True, False, False, False, normalize_with_word_com=True),
                )

            record = summary.table_log_records[1]
            self.assertEqual(record["table_type"], "normal_table")
            self.assertFalse(record["word_com_autofit_applied"])
            self.assertTrue(record["word_com_autofit_fallback_applied"])
            self.assertEqual(record["word_com_autofit_status"], "xml_fallback")
            self.assertEqual(summary.word_com_table_autofit_applied_count, 0)
            self.assertEqual(summary.word_com_table_autofit_fallback_count, 1)
            self.assertEqual(summary.word_com_table_autofit_failed_count, 0)

            logs = "\n".join(summary.word_com_table_autofit_logs)
            self.assertIn("WORD_COM_TABLE_AUTOFIT_FALLBACK_STARTED failed_records_count=1", logs)
            self.assertIn("WORD_COM_TABLE_AUTOFIT_FALLBACK_APPLIED global_table_index=2", logs)
            self.assertIn("WORD_COM_TABLE_AUTOFIT_FALLBACK_DONE applied=1", logs)

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tbl = root.xpath(".//w:tbl", namespaces=NS)[1]
            self.assertEqual(table_setting(tbl, "tblW", "type"), "pct")
            self.assertEqual(table_setting(tbl, "tblW", "w"), "5000")
            self.assertEqual(table_setting(tbl, "tblLayout", "type"), "autofit")
            self.assertIsNone(tbl.find("w:tblGrid", NS))
            self.assertIsNone(tbl.find(".//w:tcPr/w:tcW", NS))

    def test_partial_word_com_failure_fallback_only_failed_tables(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("一般表格A"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("一般表格B"))
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

            def fake_table_autofit(docx_path, records, stop=None):
                # Word COM rewrote both tables with fixed widths, but only
                # confirmed AutoFit for global index 2.
                set_fixed_widths_in_docx(docx_path, [int(r["table_index"]) for r in records])
                return (
                    ["WORD_COM_TABLE_AUTOFIT_APPLIED global_table_index=2 sequence=content_then_window"],
                    {2},
                    {3},
                )

            with patch(
                "docx_fixer.docx_processor.apply_table_autofit_with_word_com",
                side_effect=fake_table_autofit,
            ):
                summary = fix_docx_fast(
                    input_docx,
                    output_docx,
                    ProcessOptions(True, False, False, False, normalize_with_word_com=True),
                )

            applied_record = summary.table_log_records[1]
            failed_record = summary.table_log_records[2]
            self.assertEqual(applied_record["word_com_autofit_status"], "word_com")
            self.assertTrue(applied_record["word_com_autofit_applied"])
            self.assertFalse(applied_record["word_com_autofit_fallback_applied"])
            self.assertEqual(failed_record["word_com_autofit_status"], "xml_fallback")
            self.assertFalse(failed_record["word_com_autofit_applied"])
            self.assertTrue(failed_record["word_com_autofit_fallback_applied"])
            self.assertEqual(summary.word_com_table_autofit_applied_count, 1)
            self.assertEqual(summary.word_com_table_autofit_fallback_count, 1)
            self.assertEqual(summary.word_com_table_autofit_failed_count, 0)

            logs = "\n".join(summary.word_com_table_autofit_logs)
            self.assertIn("WORD_COM_TABLE_AUTOFIT_FALLBACK_STARTED failed_records_count=1", logs)
            self.assertIn("WORD_COM_TABLE_AUTOFIT_FALLBACK_APPLIED global_table_index=3", logs)
            self.assertNotIn("WORD_COM_TABLE_AUTOFIT_FALLBACK_APPLIED global_table_index=2", logs)

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)

            # The Word COM "success" table keeps the widths Word saved.
            self.assertEqual(table_setting(tables[1], "tblW", "type"), "dxa")
            self.assertEqual(table_setting(tables[1], "tblW", "w"), "20000")
            self.assertIsNotNone(tables[1].find("w:tblGrid", NS))

            # The failed table is repaired back to the safe window format.
            self.assertEqual(table_setting(tables[2], "tblW", "type"), "pct")
            self.assertEqual(table_setting(tables[2], "tblW", "w"), "5000")
            self.assertEqual(table_setting(tables[2], "tblLayout", "type"), "autofit")
            self.assertIsNone(tables[2].find("w:tblGrid", NS))
            self.assertIsNone(tables[2].find(".//w:tcPr/w:tcW", NS))

    def test_word_com_runner_exception_triggers_fallback(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("一般表格"))
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

            with patch(
                "docx_fixer.docx_processor.apply_table_autofit_with_word_com",
                side_effect=RuntimeError("word_com_unavailable"),
            ):
                summary = fix_docx_fast(
                    input_docx,
                    output_docx,
                    ProcessOptions(True, False, False, False, normalize_with_word_com=True),
                )

            record = summary.table_log_records[1]
            self.assertEqual(record["word_com_autofit_status"], "xml_fallback")
            self.assertTrue(record["word_com_autofit_fallback_applied"])
            self.assertEqual(summary.word_com_table_autofit_fallback_count, 1)
            self.assertEqual(summary.word_com_table_autofit_failed_count, 0)

            logs = "\n".join(summary.word_com_table_autofit_logs)
            self.assertIn("reason=runner_failed:RuntimeError:word_com_unavailable", logs)
            self.assertIn("WORD_COM_TABLE_AUTOFIT_FALLBACK_STARTED", logs)
            self.assertIn("WORD_COM_TABLE_AUTOFIT_FALLBACK_DONE applied=1", logs)

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tbl = root.xpath(".//w:tbl", namespaces=NS)[1]
            self.assertEqual(table_setting(tbl, "tblW", "type"), "pct")
            self.assertEqual(table_setting(tbl, "tblW", "w"), "5000")
            self.assertEqual(table_setting(tbl, "tblLayout", "type"), "autofit")
            self.assertIsNone(tbl.find("w:tblGrid", NS))
            self.assertIsNone(tbl.find(".//w:tcPr/w:tcW", NS))

    def test_fallback_does_not_touch_special_table_or_skipped_table(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("特殊表格"))
        body.append(make_table([4, 4]))
        body.append(make_paragraph("小表格"))
        body.append(make_table([2, 2]))
        body.append(make_paragraph("一般表格"))
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

            def fake_table_autofit(docx_path, records, stop=None):
                failed = {int(r["global_table_index"]) for r in records}
                return (["WORD_COM_TABLE_AUTOFIT_FAILED reason=powershell_nonzero_exit"], set(), failed)

            with patch(
                "docx_fixer.docx_processor.apply_table_autofit_with_word_com",
                side_effect=fake_table_autofit,
            ):
                summary = fix_docx_fast(
                    input_docx,
                    output_docx,
                    ProcessOptions(True, False, False, False, normalize_with_word_com=True),
                )

            self.assertEqual(
                [record["table_type"] for record in summary.table_log_records],
                ["skipped_first_table", "special_table", "skipped_small_table", "normal_table"],
            )
            self.assertEqual(summary.table_log_records[0]["word_com_autofit_status"], "not_needed")
            self.assertEqual(summary.table_log_records[1]["word_com_autofit_status"], "not_needed")
            self.assertEqual(summary.table_log_records[2]["word_com_autofit_status"], "not_needed")
            self.assertEqual(summary.table_log_records[3]["word_com_autofit_status"], "xml_fallback")

            logs = "\n".join(summary.word_com_table_autofit_logs)
            self.assertIn("WORD_COM_TABLE_AUTOFIT_FALLBACK_STARTED failed_records_count=1", logs)
            self.assertIn("WORD_COM_TABLE_AUTOFIT_FALLBACK_APPLIED global_table_index=4", logs)

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)

            # First table stays untouched by both passes.
            self.assertIsNone(tables[0].find("w:tblPr", NS))
            # Special table keeps the special right-aligned fallback layout.
            self.assertEqual(table_setting(tables[1], "jc"), "right")
            self.assertEqual(table_setting(tables[1], "tblW", "type"), "auto")
            # Small table stays untouched.
            self.assertIsNone(tables[2].find("w:tblPr", NS))
            # Only the normal table is repaired by the fallback.
            self.assertEqual(table_setting(tables[3], "jc"), "center")
            self.assertEqual(table_setting(tables[3], "tblW", "type"), "pct")
            self.assertEqual(table_setting(tables[3], "tblW", "w"), "5000")

    def test_fallback_respects_chapter_three_table_layout_skip(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("壹、序言"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("參、價格形成之主要因素分析"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("肆、第四章"))
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

            captured: dict[str, object] = {}

            def fake_table_autofit(docx_path, records, stop=None):
                captured["records"] = list(records)
                failed = {int(r["global_table_index"]) for r in records}
                return (["WORD_COM_TABLE_AUTOFIT_EXCEPTION type=Test message=boom"], set(), failed)

            with patch(
                "docx_fixer.docx_processor.apply_table_autofit_with_word_com",
                side_effect=fake_table_autofit,
            ):
                summary = fix_docx_fast(
                    input_docx,
                    output_docx,
                    ProcessOptions(
                        fix_table_layout=True,
                        fix_color=False,
                        fix_paragraph=False,
                        normalize_with_word_com=True,
                        skip_chapter_three_table_layout=True,
                    ),
                )

            self.assertEqual(
                [record["global_table_index"] for record in captured["records"]],
                [3],
            )
            chapter_three_record = summary.table_log_records[1]
            self.assertEqual(chapter_three_record["table_type"], "skipped_chapter_three_table")
            self.assertEqual(chapter_three_record["word_com_autofit_status"], "not_needed")
            self.assertFalse(chapter_three_record["word_com_autofit_fallback_applied"])
            self.assertEqual(summary.table_log_records[2]["word_com_autofit_status"], "xml_fallback")

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)

            # The chapter three protected table is untouched by the fallback.
            self.assertIsNone(tables[1].find("w:tblPr", NS))
            self.assertIsNone(tables[1].find("w:tblGrid", NS))
            # The normal table outside the protected region is repaired.
            self.assertEqual(table_setting(tables[2], "tblW", "type"), "pct")
            self.assertEqual(table_setting(tables[2], "tblW", "w"), "5000")

    def test_processor_applies_nested_table_color_only_when_protection_is_enabled(self):
        document = make_nested_table_document()
        before_tables = document.xpath(".//w:tbl", namespaces=NS)
        before_outer_layout = table_layout_signature(before_tables[1])
        before_inner_layout = table_layout_signature(before_tables[2])

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

            options = ProcessOptions(
                fix_table_layout=True,
                fix_color=True,
                fix_paragraph=False,
                normalize_with_word_com=False,
                skip_nested_tables=True,
            )
            summary = fix_docx_fast(input_docx, output_docx, options)

            self.assertEqual(summary.tables, 3)
            self.assertEqual(summary.skipped_first_page_tables, 1)
            self.assertEqual(summary.skipped_nested_tables, 0)
            self.assertEqual(summary.nested_table_color_only_tables, 2)
            self.assertEqual(summary.changed_to_gray, 1)
            self.assertEqual(summary.cleared_colors, 1)
            nested_records = [
                record
                for record in summary.table_log_records
                if record["table_type"] == "nested_table_color_only"
            ]
            self.assertEqual(len(nested_records), 2)
            self.assertTrue(
                all(record["action"] == "apply_nested_table_color_only" for record in nested_records)
            )
            self.assertTrue(all(not record["layout_fixed"] for record in nested_records))
            self.assertTrue(all(record["color_fixed"] for record in nested_records))
            self.assertEqual([record["changed_to_gray"] for record in nested_records], [1, 0])
            self.assertEqual([record["cleared_colors"] for record in nested_records], [0, 1])
            self.assertTrue(all(record["shading_debug"] for record in nested_records))

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)
            outer = tables[1]
            inner = tables[2]

            self.assertEqual(table_layout_signature(outer), before_outer_layout)
            self.assertEqual(table_layout_signature(inner), before_inner_layout)
            for tbl in (outer, inner):
                self.assertIsNone(table_pr_element(tbl, "jc"))
                self.assertIsNone(table_pr_element(tbl, "tblW"))
                self.assertIsNone(table_pr_element(tbl, "tblLayout"))

            for run in inner.xpath(".//w:r", namespaces=NS):
                r_pr = run.find("w:rPr", NS)
                if r_pr is not None:
                    self.assertIsNone(r_pr.find("w:sz", NS))
                    self.assertIsNone(r_pr.find("w:szCs", NS))
            self.assertEqual(cell_fill(outer, 0, 0), "D9D9D9")
            self.assertEqual(cell_fill(inner, 0, 0), "auto")

    def test_processor_fully_skips_nested_tables_when_fix_color_is_disabled(self):
        document = make_nested_table_document()
        before_tables = document.xpath(".//w:tbl", namespaces=NS)
        before_outer_layout = table_layout_signature(before_tables[1])
        before_inner_layout = table_layout_signature(before_tables[2])

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
                ProcessOptions(
                    fix_table_layout=True,
                    fix_color=False,
                    fix_paragraph=False,
                    normalize_with_word_com=False,
                    skip_nested_tables=True,
                    skip_special_color_tables=True,
                    special_color_skip_colors=("FF0000",),
                    clear_special_colors_after_skip=True,
                ),
            )

            self.assertEqual(summary.skipped_nested_tables, 2)
            self.assertEqual(summary.nested_table_color_only_tables, 0)
            self.assertEqual(summary.special_color_skipped_tables, 0)
            self.assertEqual(summary.changed_to_gray, 0)
            self.assertEqual(summary.cleared_colors, 0)
            self.assertEqual(
                [record["table_type"] for record in summary.table_log_records],
                ["skipped_first_table", "skipped_nested_table", "skipped_nested_table"],
            )

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)
            outer = tables[1]
            inner = tables[2]
            self.assertEqual(table_layout_signature(outer), before_outer_layout)
            self.assertEqual(table_layout_signature(inner), before_inner_layout)
            self.assertEqual(cell_fill(outer, 0, 0), "BFBFBF")
            self.assertEqual(cell_fill(inner, 0, 0), "FF0000")

    def test_nested_tables_are_not_queued_for_word_com_autofit_when_protected(self):
        document = make_nested_table_document()

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

            with patch("docx_fixer.docx_processor.apply_table_autofit_with_word_com") as autofit:
                summary = fix_docx_fast(
                    input_docx,
                    output_docx,
                    ProcessOptions(
                        fix_table_layout=True,
                        fix_color=True,
                        fix_paragraph=False,
                        normalize_with_word_com=True,
                        skip_nested_tables=True,
                    ),
                )

        autofit.assert_not_called()
        self.assertEqual(summary.word_com_table_autofit_records, [])
        self.assertEqual(
            [record["table_type"] for record in summary.table_log_records],
            ["skipped_first_table", "nested_table_color_only", "nested_table_color_only"],
        )
        self.assertTrue(
            all(not record["word_com_autofit_applied"] for record in summary.table_log_records)
        )

    def test_processor_allows_nested_tables_when_protection_is_disabled(self):
        document = make_nested_table_document()

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

            options = ProcessOptions(
                fix_table_layout=True,
                fix_color=True,
                fix_paragraph=False,
                normalize_with_word_com=False,
                skip_nested_tables=False,
            )
            summary = fix_docx_fast(input_docx, output_docx, options)

            self.assertEqual(summary.tables, 3)
            self.assertEqual(summary.skipped_first_page_tables, 1)
            self.assertEqual(summary.skipped_nested_tables, 0)
            self.assertFalse(
                any(record["table_type"] == "skipped_nested_table" for record in summary.table_log_records)
            )

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)
            outer = tables[1]
            inner = tables[2]

            self.assertEqual(table_setting(outer, "jc"), "right")
            self.assertEqual(table_setting(inner, "jc"), "right")
            self.assertEqual(table_setting(inner, "tblLayout", "type"), "autofit")
            self.assertEqual(first_table_fill(inner), "auto")
            for run in inner.xpath(".//w:r", namespaces=NS):
                r_pr = run.find("w:rPr", NS)
                self.assertIsNotNone(r_pr)
                self.assertEqual(r_pr.find("w:sz", NS).get(qn("val")), "22")

    def test_processor_skips_all_table_processing_under_chapter_three(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("\u58f9\u3001\u5e8f\u8a00"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("\u53c3\u3001\u50f9\u683c\u5f62\u6210\u4e4b\u4e3b\u8981\u56e0\u7d20\u5206\u6790"))
        body.append(make_table([3, 3]))
        body.append(make_paragraph("\u8086\u3001\u7b2c\u56db\u7ae0"))
        body.append(make_table([3, 3]))

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

            options = ProcessOptions(
                fix_table_layout=True,
                fix_color=False,
                fix_paragraph=False,
                normalize_with_word_com=False,
                skip_chapter_three_tables=True,
            )
            summary = fix_docx_fast(input_docx, output_docx, options)

            self.assertEqual(summary.tables, 3)
            self.assertEqual(summary.skipped_first_page_tables, 1)
            self.assertEqual(summary.normal_processed_tables, 0)
            self.assertEqual(summary.special_autofit_right_tables, 1)
            self.assertEqual(summary.table_log_records[1]["table_type"], "skipped_chapter_three_table")
            self.assertEqual(summary.table_log_records[1]["action"], "skipped")
            self.assertEqual(summary.table_log_records[1]["special_layout_used"], False)
            self.assertEqual(
                summary.table_log_records[1]["reason"],
                "chapter three protected table; layout and color skipped",
            )
            self.assertTrue(summary.table_log_records[1]["chapter_three_table_layout_skipped"])
            self.assertTrue(summary.table_log_records[1]["chapter_three_table_color_skipped"])
            self.assertEqual(summary.table_log_records[2]["table_type"], "special_table")
            self.assertEqual(summary.table_log_records[2]["special_layout_used"], True)

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)

            self.assertIsNone(tables[1].find("w:tblPr", NS))
            self.assertIsNone(table_setting(tables[1], "tblInd", "w"))
            self.assertEqual(table_setting(tables[2], "jc"), "right")

    def test_processor_allows_color_only_under_chapter_three_when_only_layout_is_skipped(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("\u58f9\u3001\u5e8f\u8a00"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("\u53c3\u3001\u50f9\u683c\u5f62\u6210\u4e4b\u4e3b\u8981\u56e0\u7d20\u5206\u6790"))
        body.append(make_shaded_table([3, 3], fill="808080"))

        with tempfile.TemporaryDirectory() as tmp:
            input_docx = Path(tmp) / "input.docx"
            output_docx = Path(tmp) / "output.docx"
            with ZipFile(input_docx, "w", ZIP_DEFLATED) as zout:
                zout.writestr("word/document.xml", etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True))

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=True,
                    fix_color=True,
                    fix_paragraph=False,
                    normalize_with_word_com=False,
                    skip_chapter_three_table_layout=True,
                    skip_chapter_three_table_color=False,
                ),
            )

            record = summary.table_log_records[1]
            self.assertEqual(record["table_type"], "color_only_table")
            self.assertEqual(record["action"], "apply_color_only")
            self.assertEqual(record["reason"], "chapter three protected table; layout skipped; color allowed")
            self.assertFalse(record["layout_fixed"])
            self.assertTrue(record["color_fixed"])
            self.assertTrue(record["chapter_three_table_layout_skipped"])
            self.assertFalse(record["chapter_three_table_color_skipped"])

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)
            self.assertIsNone(table_setting(tables[1], "jc"))
            self.assertEqual(first_table_fill(tables[1]), "D9D9D9")

    def test_processor_allows_layout_only_under_chapter_three_when_only_color_is_skipped(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("\u58f9\u3001\u5e8f\u8a00"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("\u53c3\u3001\u50f9\u683c\u5f62\u6210\u4e4b\u4e3b\u8981\u56e0\u7d20\u5206\u6790"))
        body.append(make_shaded_table([3, 3], fill="808080"))

        with tempfile.TemporaryDirectory() as tmp:
            input_docx = Path(tmp) / "input.docx"
            output_docx = Path(tmp) / "output.docx"
            with ZipFile(input_docx, "w", ZIP_DEFLATED) as zout:
                zout.writestr("word/document.xml", etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True))

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=True,
                    fix_color=True,
                    fix_paragraph=False,
                    normalize_with_word_com=False,
                    skip_chapter_three_table_layout=False,
                    skip_chapter_three_table_color=True,
                ),
            )

            record = summary.table_log_records[1]
            self.assertEqual(record["table_type"], "special_table")
            self.assertEqual(record["action"], "apply_special_table_format")
            self.assertEqual(record["reason"], "chapter three protected table; layout allowed; color skipped")
            self.assertTrue(record["layout_fixed"])
            self.assertFalse(record["color_fixed"])
            self.assertFalse(record["chapter_three_table_layout_skipped"])
            self.assertTrue(record["chapter_three_table_color_skipped"])

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)
            self.assertEqual(table_setting(tables[1], "jc"), "right")
            self.assertEqual(first_table_fill(tables[1]), "808080")

    def test_chapter_three_color_skip_blocks_nested_table_color_only(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("\u58f9\u3001\u5e8f\u8a00"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("\u53c3\u3001\u50f9\u683c\u5f62\u6210\u4e4b\u4e3b\u8981\u56e0\u7d20\u5206\u6790"))
        outer = make_table([3, 3])
        set_cell_shading(outer, 0, 0, "BFBFBF")
        inner = make_shaded_table([3, 3], fill="FF0000")
        outer.find("w:tr/w:tc", NS).append(inner)
        body.append(outer)

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
                ProcessOptions(
                    fix_table_layout=True,
                    fix_color=True,
                    fix_paragraph=False,
                    normalize_with_word_com=False,
                    skip_nested_tables=True,
                    skip_chapter_three_table_color=True,
                ),
            )

            self.assertEqual(summary.skipped_nested_tables, 2)
            self.assertEqual(summary.nested_table_color_only_tables, 0)
            self.assertEqual(summary.changed_to_gray, 0)
            self.assertEqual(summary.cleared_colors, 0)
            self.assertEqual(
                [record["table_type"] for record in summary.table_log_records],
                ["skipped_first_table", "skipped_nested_table", "skipped_nested_table"],
            )
            for record in summary.table_log_records[1:]:
                self.assertTrue(record["chapter_three_table_color_skipped"])
                self.assertFalse(record["color_fixed"])

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)
            self.assertEqual(cell_fill(tables[1], 0, 0), "BFBFBF")
            self.assertEqual(cell_fill(tables[2], 0, 0), "FF0000")

    def test_processor_does_not_skip_other_chapter_three_titles(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("\u58f9\u3001\u5e8f\u8a00"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("\u53c3\u3001\u7b2c\u4e09\u7ae0"))
        body.append(make_table([3, 3]))
        body.append(make_paragraph("\u8086\u3001\u7b2c\u56db\u7ae0"))
        body.append(make_table([3, 3]))

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

            options = ProcessOptions(
                fix_table_layout=True,
                fix_color=False,
                fix_paragraph=False,
                normalize_with_word_com=False,
                skip_chapter_three_tables=True,
            )
            summary = fix_docx_fast(input_docx, output_docx, options)

            self.assertEqual(summary.tables, 3)
            self.assertEqual(summary.skipped_first_page_tables, 1)
            self.assertEqual(summary.normal_processed_tables, 0)
            self.assertEqual(summary.special_autofit_right_tables, 2)
            self.assertEqual(summary.table_log_records[1]["table_type"], "special_table")
            self.assertEqual(summary.table_log_records[1]["special_layout_used"], True)

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)
            self.assertEqual(table_setting(tables[1], "jc"), "right")

    def test_processor_skips_special_layout_for_auto_numbered_target_chapter_three(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("\u5e8f\u8a00", num_id="1", ilvl=0))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("\u7b2c\u4e8c\u7ae0", num_id="1", ilvl=0))
        body.append(make_paragraph("\u50f9\u683c\u5f62\u6210\u4e4b\u4e3b\u8981\u56e0\u7d20\u5206\u6790", num_id="1", ilvl=0))
        body.append(make_table([3, 3]))

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
                zout.writestr("word/numbering.xml", make_legal_traditional_numbering_xml("1"))

            options = ProcessOptions(
                fix_table_layout=True,
                fix_color=False,
                fix_paragraph=False,
                normalize_with_word_com=False,
                skip_chapter_three_tables=True,
            )
            summary = fix_docx_fast(input_docx, output_docx, options)

            self.assertEqual(summary.tables, 2)
            self.assertEqual(summary.skipped_first_page_tables, 1)
            self.assertEqual(summary.normal_processed_tables, 0)
            self.assertEqual(summary.special_autofit_right_tables, 0)
            self.assertEqual(summary.table_log_records[1]["table_type"], "skipped_chapter_three_table")
            self.assertEqual(summary.table_log_records[1]["action"], "skipped")
            self.assertEqual(summary.table_log_records[1]["special_layout_used"], False)
            self.assertEqual(
                summary.table_log_records[1]["reason"],
                "chapter three protected table; layout and color skipped",
            )

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)
            self.assertIsNone(tables[1].find("w:tblPr", NS))
            self.assertIsNone(table_setting(tables[1], "tblInd", "w"))

    def test_processor_skips_special_layout_for_style_numbered_target_chapter_three(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("\u5e8f\u8a00", style="ChapterHeading"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("\u7b2c\u4e8c\u7ae0", style="ChapterHeading"))
        body.append(make_paragraph("\u50f9\u683c\u5f62\u6210\u4e4b\u4e3b\u8981\u56e0\u7d20\u5206\u6790", style="ChapterHeading"))
        body.append(make_table([3, 3]))

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
                zout.writestr("word/numbering.xml", make_legal_traditional_numbering_xml("2"))
                zout.writestr(
                    "word/styles.xml",
                    make_style_numbering_xml("ChapterHeading", num_id="2"),
                )

            options = ProcessOptions(
                fix_table_layout=True,
                fix_color=False,
                fix_paragraph=False,
                normalize_with_word_com=False,
                skip_chapter_three_tables=True,
            )
            summary = fix_docx_fast(input_docx, output_docx, options)

            self.assertEqual(summary.tables, 2)
            self.assertEqual(summary.skipped_first_page_tables, 1)
            self.assertEqual(summary.normal_processed_tables, 0)
            self.assertEqual(summary.special_autofit_right_tables, 0)
            self.assertEqual(summary.table_log_records[1]["table_type"], "skipped_chapter_three_table")
            self.assertEqual(summary.table_log_records[1]["action"], "skipped")
            self.assertEqual(summary.table_log_records[1]["special_layout_used"], False)
            self.assertEqual(
                summary.table_log_records[1]["reason"],
                "chapter three protected table; layout and color skipped",
            )

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)
            self.assertIsNone(tables[1].find("w:tblPr", NS))
            self.assertIsNone(table_setting(tables[1], "tblInd", "w"))

    def test_processor_keeps_special_layout_for_auto_numbered_chapter_two(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("\u5e8f\u8a00", num_id="1", ilvl=0))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("\u7b2c\u4e8c\u7ae0", num_id="1", ilvl=0))
        body.append(make_table([3, 3]))

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
                zout.writestr("word/numbering.xml", make_legal_traditional_numbering_xml("1"))

            options = ProcessOptions(
                fix_table_layout=True,
                fix_color=False,
                fix_paragraph=False,
                normalize_with_word_com=False,
                skip_chapter_three_tables=True,
            )
            summary = fix_docx_fast(input_docx, output_docx, options)

            self.assertEqual(summary.tables, 2)
            self.assertEqual(summary.skipped_first_page_tables, 1)
            self.assertEqual(summary.special_autofit_right_tables, 1)
            self.assertEqual(summary.normal_processed_tables, 0)
            self.assertEqual(summary.table_log_records[1]["table_type"], "special_table")
            self.assertEqual(summary.table_log_records[1]["special_layout_used"], True)
            self.assertEqual(summary.table_log_records[1]["reason"], "column_count <= 4")

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)
            self.assertEqual(table_setting(tables[1], "jc"), "right")

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

    def test_keep_color_list_preserves_fill_in_process_table(self):
        # Built-in keep defaults are D9D9D9 and F2F2F2 only.
        tbl = make_shaded_table([3, 3], fill="D9D9D9")
        process_table(tbl, ProcessOptions(False, True, False, False))
        self.assertEqual(first_table_fill(tbl), "D9D9D9")

        tbl = make_shaded_table([3, 3], fill="F2F2F2")
        process_table(tbl, ProcessOptions(False, True, False, False))
        self.assertEqual(first_table_fill(tbl), "F2F2F2")

        # DDEBF7 is no longer a built-in keep color, so it is cleared by default.
        tbl = make_shaded_table([3, 3], fill="DDEBF7")
        process_table(tbl, ProcessOptions(False, True, False, False))
        self.assertEqual(first_table_fill(tbl), "auto")

        tbl = make_shaded_table([3, 3], fill="FCE4D6")
        process_table(
            tbl,
            ProcessOptions(False, True, False, False, table_keep_colors=("FCE4D6",)),
        )
        self.assertEqual(first_table_fill(tbl), "FCE4D6")

    def test_gray_color_list_converts_to_gray_target(self):
        tbl = make_shaded_table([3, 3], fill="BFBFBF")
        process_table(tbl, ProcessOptions(False, True, False, False))
        self.assertEqual(first_table_fill(tbl), "D9D9D9")

        tbl = make_shaded_table([3, 3], fill="FFC000")
        process_table(
            tbl,
            ProcessOptions(
                False,
                True,
                False,
                False,
                table_gray_colors=("FFC000",),
                table_gray_target="CCCCCC",
            ),
        )
        self.assertEqual(first_table_fill(tbl), "CCCCCC")

    def test_special_color_table_is_skipped_entirely(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("特殊顏色表格"))
        special_color_table = make_table([5, 5])
        set_cell_shading(special_color_table, 0, 0, "FFFF00")
        set_cell_shading(special_color_table, 0, 1, "FF0000")
        body.append(special_color_table)

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

            options = ProcessOptions(
                fix_table_layout=True,
                fix_color=True,
                fix_paragraph=False,
                normalize_with_word_com=False,
                skip_special_color_tables=True,
                special_color_skip_colors=("FFFF00",),
            )
            summary = fix_docx_fast(input_docx, output_docx, options)

            record = summary.table_log_records[1]
            self.assertEqual(record["table_type"], "special_color_skipped_table")
            self.assertEqual(record["action"], "skipped_special_color_table")
            self.assertEqual(record["reason"], "matched special color skip list")
            self.assertTrue(record["special_color_skip_matched"])
            self.assertEqual(record["special_color_skip_colors"], ["FFFF00"])
            self.assertEqual(record["special_color_cleared_count"], 0)
            self.assertFalse(record["layout_fixed"])
            self.assertFalse(record["color_fixed"])
            self.assertFalse(record["special_layout_used"])
            self.assertEqual(summary.special_color_skipped_tables, 1)
            self.assertEqual(summary.normal_processed_tables, 0)
            self.assertEqual(summary.special_autofit_right_tables, 0)
            self.assertEqual(summary.word_com_table_autofit_records, [])

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

            # The whole table is skipped: no layout change, no color rules.
            self.assertIsNone(tbl.find("w:tblPr", NS))
            self.assertEqual(cell_fill(tbl, 0, 0), "FFFF00")
            self.assertEqual(cell_fill(tbl, 0, 1), "FF0000")

    def test_special_color_table_clear_after_skip_only_clears_matched_colors(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("特殊顏色表格"))
        special_color_table = make_table([5, 5])
        set_cell_shading(special_color_table, 0, 0, "FFFF00")
        set_cell_shading(special_color_table, 0, 1, "FF0000")
        set_cell_shading(special_color_table, 1, 0, "808080")
        body.append(special_color_table)

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

            options = ProcessOptions(
                fix_table_layout=True,
                fix_color=True,
                fix_paragraph=False,
                normalize_with_word_com=False,
                skip_special_color_tables=True,
                special_color_skip_colors=("FFFF00",),
                clear_special_colors_after_skip=True,
            )
            summary = fix_docx_fast(input_docx, output_docx, options)

            record = summary.table_log_records[1]
            self.assertEqual(record["table_type"], "special_color_skipped_table")
            self.assertEqual(record["special_color_cleared_count"], 1)
            self.assertTrue(record["color_fixed"])

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

            # Only the matched special color is cleared to no color.
            self.assertEqual(cell_fill(tbl, 0, 0), "auto")
            # Other colors are untouched: no clear rule, no gray rule.
            self.assertEqual(cell_fill(tbl, 0, 1), "FF0000")
            self.assertEqual(cell_fill(tbl, 1, 0), "808080")
            self.assertIsNone(tbl.find("w:tblPr", NS))

    def test_special_color_skip_does_not_affect_first_table_rule(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        first_table = make_table([5, 5])
        set_cell_shading(first_table, 0, 0, "FFFF00")
        body.append(first_table)
        body.append(make_paragraph("一般表格"))
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

            options = ProcessOptions(
                fix_table_layout=True,
                fix_color=True,
                fix_paragraph=False,
                normalize_with_word_com=False,
                skip_special_color_tables=True,
                special_color_skip_colors=("FFFF00",),
                clear_special_colors_after_skip=True,
            )
            summary = fix_docx_fast(input_docx, output_docx, options)

            self.assertEqual(summary.table_log_records[0]["table_type"], "skipped_first_table")
            self.assertFalse(summary.table_log_records[0]["special_color_skip_matched"])
            self.assertEqual(summary.table_log_records[1]["table_type"], "normal_table")
            self.assertEqual(summary.special_color_skipped_tables, 0)

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            tables = root.xpath(".//w:tbl", namespaces=NS)

            # The first table stays fully untouched even though it matches.
            self.assertEqual(cell_fill(tables[0], 0, 0), "FFFF00")
            self.assertIsNone(tables[0].find("w:tblPr", NS))

    def test_special_color_skip_blocks_nested_table_color_only(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("第一張表格"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("巢狀表格"))
        outer = make_table([3, 3])
        inner = make_shaded_table([3, 3], fill="FFFF00")
        outer.find("w:tr/w:tc", NS).append(inner)
        body.append(outer)

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

            options = ProcessOptions(
                fix_table_layout=True,
                fix_color=True,
                fix_paragraph=False,
                normalize_with_word_com=False,
                skip_nested_tables=True,
                skip_special_color_tables=True,
                special_color_skip_colors=("FFFF00",),
                clear_special_colors_after_skip=False,
            )
            summary = fix_docx_fast(input_docx, output_docx, options)

            self.assertEqual(
                [record["table_type"] for record in summary.table_log_records],
                ["skipped_first_table", "nested_table_color_only", "special_color_skipped_table"],
            )
            self.assertEqual(summary.special_color_skipped_tables, 1)
            self.assertEqual(summary.skipped_nested_tables, 0)
            self.assertEqual(summary.nested_table_color_only_tables, 1)
            self.assertEqual(summary.changed_to_gray, 0)
            self.assertEqual(summary.cleared_colors, 0)

            with ZipFile(output_docx) as zin:
                root = etree.fromstring(zin.read("word/document.xml"))
            inner_out = root.xpath(".//w:tbl", namespaces=NS)[2]
            self.assertEqual(first_table_fill(inner_out), "FFFF00")

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


    def test_table_log_records_first_level_heading_for_manual_chapters(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("前言"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("壹、比較標的資料"))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("貳、小表格"))
        body.append(make_table([3, 2]))
        body.append(make_paragraph("參、一般表格"))
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
                ProcessOptions(True, False, False, False, normalize_with_word_com=False),
            )

            headings = [record["first_level_heading"] for record in summary.table_log_records]
            self.assertEqual(headings, ["(none)", "壹、", "貳、", "參、"])

    def test_table_log_records_first_level_heading_for_auto_chapter_three(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        body.append(make_paragraph("序言", num_id="1", ilvl=0))
        body.append(make_table([5, 5]))
        body.append(make_paragraph("第二章", num_id="1", ilvl=0))
        body.append(make_paragraph("第三章", num_id="1", ilvl=0))
        body.append(make_table([3, 3]))

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
                zout.writestr("word/numbering.xml", make_legal_traditional_numbering_xml("1"))

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    True,
                    False,
                    False,
                    False,
                    normalize_with_word_com=False,
                    skip_chapter_three_tables=True,
                ),
            )

            self.assertEqual(summary.table_log_records[0]["first_level_heading"], "壹、")
            self.assertEqual(summary.table_log_records[1]["first_level_heading"], "參、")


if __name__ == "__main__":
    unittest.main()
