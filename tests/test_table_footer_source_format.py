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
from docx_fixer.table_format import (
    FOOTER_NOTE_PREFIX_PATTERN,
    apply_table_footer_source_format,
    normalize_footer_source_cell_text,
    set_border_double_black,
    set_border_nil,
)
from docx_fixer.xml_utils import qn


# ----------------------------------------------------------------------------
# Builders
# ----------------------------------------------------------------------------
def make_paragraph(text: str, *, style: str | None = None):
    p = etree.Element(qn("p"))
    if style is not None:
        p_pr = etree.SubElement(p, qn("pPr"))
        p_style = etree.SubElement(p_pr, qn("pStyle"))
        p_style.set(qn("val"), style)
    r = etree.SubElement(p, qn("r"))
    t = etree.SubElement(r, qn("t"))
    t.text = text
    return p


def make_table(rows: list[list[str]]):
    tbl = etree.Element(qn("tbl"))
    for row in rows:
        tr = etree.SubElement(tbl, qn("tr"))
        for text in row:
            tc = etree.SubElement(tr, qn("tc"))
            p = etree.SubElement(tc, qn("p"))
            r = etree.SubElement(p, qn("r"))
            t = etree.SubElement(r, qn("t"))
            t.text = text
    return tbl


def uniform_table(rows: int, cols: int, fill_text: str = "x"):
    return make_table([[fill_text] * cols for _ in range(rows)])


def footer_table():
    """A title row (single cell) + header + data + a 基期/資料來源 last row."""
    return make_table(
        [
            ["報表標題"],
            ["項目", "金額", "比率", "排名", "備註"],
            ["土地", "100", "50%", "1", "ok"],
            ["基期：民國100年", "其他資料", "資料來源：本所", "", ""],
        ]
    )


def build_document(*nodes):
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    for node in nodes:
        body.append(node)
    return document


def run_fix(document, options: ProcessOptions):
    with tempfile.TemporaryDirectory() as tmp:
        input_docx = Path(tmp) / "input.docx"
        output_docx = Path(tmp) / "output.docx"
        with ZipFile(input_docx, "w", ZIP_DEFLATED) as zout:
            zout.writestr(
                "word/document.xml",
                etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True),
            )
        summary = fix_docx_fast(input_docx, output_docx, options)
        with ZipFile(output_docx) as zin:
            root = etree.fromstring(zin.read("word/document.xml"))
    return summary, root


# ----------------------------------------------------------------------------
# Accessors
# ----------------------------------------------------------------------------
def cell_at(tbl, row: int, col: int):
    return tbl.findall("w:tr", NS)[row].findall("w:tc", NS)[col]


def tbl_border(tbl, side: str):
    return tbl.find(f"w:tblPr/w:tblBorders/w:{side}", NS)


def tc_border(tc, side: str):
    return tc.find(f"w:tcPr/w:tcBorders/w:{side}", NS)


def make_tc_borders(tc):
    tc_pr = tc.find("w:tcPr", NS)
    if tc_pr is None:
        tc_pr = etree.Element(qn("tcPr"))
        tc.insert(0, tc_pr)
    borders = tc_pr.find("w:tcBorders", NS)
    if borders is None:
        borders = etree.SubElement(tc_pr, qn("tcBorders"))
    return borders


def is_double_black(border) -> bool:
    return (
        border is not None
        and border.get(qn("val")) == "double"
        and border.get(qn("sz")) == "4"
        and border.get(qn("color")) == "000000"
    )


def is_nil(border) -> bool:
    return border is not None and border.get(qn("val")) == "nil"


def child_order(element) -> list[str]:
    if element is None:
        return []
    return [etree.QName(child).localname for child in element]


def assert_child_before(testcase: unittest.TestCase, parent, first: str, second: str) -> None:
    order = child_order(parent)
    testcase.assertIn(first, order)
    testcase.assertIn(second, order)
    testcase.assertLess(order.index(first), order.index(second), order)


def cell_run_sizes(tc) -> list[str | None]:
    sizes = []
    for run in tc.xpath(".//w:r", namespaces=NS):
        sz = run.find("w:rPr/w:sz", NS)
        sizes.append(sz.get(qn("val")) if sz is not None else None)
    return sizes


def cell_paragraph_alignments(tc) -> list[str | None]:
    out = []
    for p in tc.findall("w:p", NS):
        jc = p.find("w:pPr/w:jc", NS)
        out.append(jc.get(qn("val")) if jc is not None else None)
    return out


def footer_options(**overrides) -> ProcessOptions:
    base = dict(
        fix_table_layout=True,
        fix_color=False,
        fix_paragraph=False,
        normalize_with_word_com=False,
        enable_table_footer_source_format=True,
    )
    base.update(overrides)
    return ProcessOptions(**base)


# ----------------------------------------------------------------------------
# Border primitives: local update must not disturb the other sides.
# ----------------------------------------------------------------------------
class BorderPrimitiveTests(unittest.TestCase):
    def _borders(self):
        return etree.Element(qn("tcBorders"))

    def test_set_double_black_only_touches_named_side(self):
        borders = self._borders()
        set_border_double_black(borders, "left")
        set_border_double_black(borders, "right")

        left = borders.find("w:left", NS)
        self.assertTrue(is_double_black(left))
        self.assertEqual(left.get(qn("sz")), "4")
        self.assertEqual(left.get(qn("space")), "0")
        self.assertIsNone(borders.find("w:top", NS))
        self.assertIsNone(borders.find("w:bottom", NS))

    def test_set_nil_preserves_existing_double_black_on_other_sides(self):
        borders = self._borders()
        set_border_double_black(borders, "top")
        set_border_double_black(borders, "bottom")

        set_border_nil(borders, "bottom")

        # top stays double black; only bottom became nil.
        self.assertTrue(is_double_black(borders.find("w:top", NS)))
        self.assertTrue(is_nil(borders.find("w:bottom", NS)))
        # nil drops the size/color attributes.
        bottom = borders.find("w:bottom", NS)
        self.assertIsNone(bottom.get(qn("sz")))
        self.assertIsNone(bottom.get(qn("color")))

    def test_repeated_updates_do_not_duplicate_a_side(self):
        borders = self._borders()
        set_border_nil(borders, "top")
        set_border_double_black(borders, "top")
        self.assertEqual(len(borders.findall("w:top", NS)), 1)
        self.assertTrue(is_double_black(borders.find("w:top", NS)))

    def test_children_kept_in_schema_order(self):
        borders = self._borders()
        # Add out of order; the helper must insert in top/left/bottom/right order.
        set_border_double_black(borders, "right")
        set_border_double_black(borders, "top")
        set_border_nil(borders, "bottom")
        set_border_nil(borders, "left")
        order = [etree.QName(child).localname for child in borders]
        self.assertEqual(order, ["top", "left", "bottom", "right"])


# ----------------------------------------------------------------------------
# apply_table_footer_source_format unit behaviour (no pipeline).
# ----------------------------------------------------------------------------
class FooterFormatUnitTests(unittest.TestCase):
    def test_whole_table_font_set_to_11pt_then_footer_cells_to_10pt(self):
        tbl = footer_table()
        apply_table_footer_source_format(tbl)

        # Header/data cells -> 11 pt (22 half-points).
        self.assertEqual(cell_run_sizes(cell_at(tbl, 1, 0)), ["22"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 2, 0)), ["22"])
        # 基期/資料來源 cells -> 10 pt (20 half-points), not overwritten by 11 pt.
        self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 0)), ["20"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 2)), ["20"])

    def test_table_outer_frame_is_double_black(self):
        tbl = footer_table()
        result = apply_table_footer_source_format(tbl)
        for side in ("top", "left", "right"):
            self.assertTrue(is_double_black(tbl_border(tbl, side)), side)
        self.assertTrue(is_nil(tbl_border(tbl, "bottom")))
        self.assertEqual(result["table_bottom_border_mode"], "footer_none")

    def test_last_row_cell_bottom_nil_is_overridden_and_xml_verified(self):
        tbl = make_table(
            [
                ["項目", "金額", "比率", "排名", "備註"],
                ["土地", "100", "50%", "1", "ok"],
                ["建物", "200", "80%", "2", "done"],
            ]
        )
        for tc in tbl.findall("w:tr", NS)[-1].findall("w:tc", NS):
            set_border_nil(make_tc_borders(tc), "bottom")

        result = apply_table_footer_source_format(tbl)

        self.assertTrue(is_double_black(tbl_border(tbl, "bottom")))
        for tc in tbl.findall("w:tr", NS)[-1].findall("w:tc", NS):
            bottom = tc_border(tc, "bottom")
            self.assertTrue(is_double_black(bottom))
            self.assertEqual(bottom.get(qn("space")), "0")
        self.assertFalse(result["footer_rows_detected"])
        self.assertEqual(result["table_bottom_border_mode"], "data_double")
        self.assertEqual(result["table_bottom_border_cell_count"], 5)
        self.assertTrue(result["table_bottom_border_xml_verified"])
        self.assertTrue(result["table_border_schema_order_valid"])
        self.assertTrue(result["table_bottom_double_border_applied"])
        self.assertEqual(result["table_bottom_double_border_cell_count"], 5)
        self.assertTrue(result["table_bottom_double_border_xml_verified"])
        self.assertIn(
            "tbl_bottom=double/4/000000",
            result["table_bottom_double_border_verify_detail"],
        )

    def test_normal_table_bottom_border_uses_schema_order_with_existing_layout(self):
        tbl = make_table(
            [
                ["項目", "金額", "比率", "排名", "備註"],
                ["土地", "100", "50%", "1", "ok"],
                ["建物", "200", "80%", "2", "done"],
            ]
        )
        tbl_pr = etree.Element(qn("tblPr"))
        tbl.insert(0, tbl_pr)
        etree.SubElement(tbl_pr, qn("tblW"))
        etree.SubElement(tbl_pr, qn("jc"))
        etree.SubElement(tbl_pr, qn("tblLayout"))

        for tc in tbl.findall("w:tr", NS)[-1].findall("w:tc", NS):
            tc_pr = etree.Element(qn("tcPr"))
            tc.insert(0, tc_pr)
            etree.SubElement(tc_pr, qn("tcW"))
            etree.SubElement(tc_pr, qn("vAlign"))
            set_border_nil(make_tc_borders(tc), "bottom")

        result = apply_table_footer_source_format(tbl)

        self.assertEqual(result["table_bottom_border_mode"], "data_double")
        self.assertTrue(result["table_bottom_border_xml_verified"])
        self.assertTrue(result["table_border_schema_order_valid"])
        assert_child_before(self, tbl.find("w:tblPr", NS), "tblBorders", "tblLayout")
        for tc in tbl.findall("w:tr", NS)[-1].findall("w:tc", NS):
            tc_pr = tc.find("w:tcPr", NS)
            assert_child_before(self, tc_pr, "tcBorders", "vAlign")
            self.assertEqual(len(tc_pr.findall("w:tcBorders", NS)), 1)
            self.assertTrue(is_double_black(tc_border(tc, "bottom")))

    def test_existing_wrong_border_container_order_is_relocated(self):
        tbl = make_table(
            [
                ["項目", "金額", "比率", "排名", "備註"],
                ["土地", "100", "50%", "1", "ok"],
            ]
        )
        tbl_pr = etree.Element(qn("tblPr"))
        tbl.insert(0, tbl_pr)
        etree.SubElement(tbl_pr, qn("tblLayout"))
        etree.SubElement(tbl_pr, qn("tblBorders"))
        tc = cell_at(tbl, 1, 0)
        tc_pr = etree.Element(qn("tcPr"))
        tc.insert(0, tc_pr)
        etree.SubElement(tc_pr, qn("vAlign"))
        etree.SubElement(tc_pr, qn("tcBorders"))

        result = apply_table_footer_source_format(tbl)

        self.assertTrue(result["table_border_schema_order_valid"])
        assert_child_before(self, tbl.find("w:tblPr", NS), "tblBorders", "tblLayout")
        assert_child_before(self, tc_pr, "tcBorders", "vAlign")
        self.assertEqual(len(tbl_pr.findall("w:tblBorders", NS)), 1)
        self.assertEqual(len(tc_pr.findall("w:tcBorders", NS)), 1)

    def test_gridspan_last_row_diagnostics_keep_full_bottom_edge(self):
        tbl = make_table(
            [
                ["h1", "h2", "h3", "h4", "h5", "h6", "h7", "h8"],
                ["d1", "d2", "d3", "d4", "d5", "d6", "d7", "d8"],
                ["span", "c3", "c4", "c5", "c6", "c7", "c8"],
            ]
        )
        first_last_cell = cell_at(tbl, 2, 0)
        tc_pr = etree.Element(qn("tcPr"))
        first_last_cell.insert(0, tc_pr)
        grid_span = etree.SubElement(tc_pr, qn("gridSpan"))
        grid_span.set(qn("val"), "2")

        result = apply_table_footer_source_format(tbl)

        self.assertFalse(result["footer_rows_detected"])
        self.assertEqual(result["table_bottom_border_mode"], "data_double")
        self.assertEqual(result["last_row_physical_cell_count"], 7)
        self.assertEqual(result["last_row_grid_span_sum"], 8)
        self.assertEqual(result["last_row_bottom_edge_target_count"], 7)
        self.assertTrue(result["table_bottom_border_xml_verified"])
        self.assertTrue(result["table_border_schema_order_valid"])
        for tc in tbl.findall("w:tr", NS)[-1].findall("w:tc", NS):
            self.assertTrue(is_double_black(tc_border(tc, "bottom")))

    def test_vmerge_data_bottom_edge_touches_continuation_and_owner(self):
        tbl = make_vmerge_table(last_row_text="value")
        result = apply_table_footer_source_format(tbl)

        owner = cell_at(tbl, 0, 0)
        continuation = cell_at(tbl, 2, 0)
        self.assertEqual(result["table_bottom_border_mode"], "data_double")
        self.assertIn("continue", result["last_row_vmerge_states"])
        self.assertEqual(result["last_row_bottom_edge_target_count"], 4)
        self.assertTrue(is_double_black(tc_border(owner, "bottom")))
        self.assertTrue(is_double_black(tc_border(continuation, "bottom")))
        self.assertEqual(owner.find("w:tcPr/w:vMerge", NS).get(qn("val")), "restart")
        self.assertIsNone(continuation.find("w:tcPr/w:vMerge", NS).get(qn("val")))

    def test_vmerge_footer_terminal_bottom_none_preserves_merge_xml(self):
        tbl = make_vmerge_table(last_row_text="註：垂直合併說明")
        result = apply_table_footer_source_format(tbl)

        continuation = cell_at(tbl, 2, 0)
        note_cell = cell_at(tbl, 2, 1)
        self.assertEqual(result["table_bottom_border_mode"], "footer_none")
        self.assertTrue(result["table_bottom_border_xml_verified"])
        self.assertEqual(result["footer_terminal_bottom_none_cell_count"], 3)
        self.assertTrue(is_nil(tbl_border(tbl, "bottom")))
        self.assertTrue(is_nil(tc_border(continuation, "bottom")))
        self.assertTrue(is_nil(tc_border(note_cell, "bottom")))
        self.assertIsNone(continuation.find("w:tcPr/w:vMerge", NS).get(qn("val")))

    def test_first_row_single_cell_borders(self):
        tbl = footer_table()
        apply_table_footer_source_format(tbl)
        title = cell_at(tbl, 0, 0)
        self.assertTrue(is_nil(tc_border(title, "top")))
        self.assertTrue(is_nil(tc_border(title, "left")))
        self.assertTrue(is_nil(tc_border(title, "right")))
        self.assertTrue(is_double_black(tc_border(title, "bottom")))

    def test_base_period_cell_formatting(self):
        tbl = footer_table()
        apply_table_footer_source_format(tbl)
        cell = cell_at(tbl, 3, 0)
        self.assertEqual(cell_run_sizes(cell), ["20"])
        self.assertEqual(cell_paragraph_alignments(cell), ["left"])
        self.assertTrue(is_double_black(tc_border(cell, "top")))
        self.assertTrue(is_nil(tc_border(cell, "left")))
        self.assertTrue(is_nil(tc_border(cell, "right")))
        self.assertTrue(is_nil(tc_border(cell, "bottom")))
        self.assertTrue(is_nil(tbl_border(tbl, "bottom")))

    def test_data_source_cell_formatting(self):
        tbl = footer_table()
        apply_table_footer_source_format(tbl)
        cell = cell_at(tbl, 3, 2)
        self.assertEqual(cell_run_sizes(cell), ["20"])
        self.assertEqual(cell_paragraph_alignments(cell), ["right"])
        self.assertTrue(is_double_black(tc_border(cell, "top")))
        self.assertTrue(is_nil(tc_border(cell, "left")))
        self.assertTrue(is_nil(tc_border(cell, "right")))
        self.assertTrue(is_nil(tc_border(cell, "bottom")))
        self.assertTrue(is_nil(tbl_border(tbl, "bottom")))

    def test_non_matching_last_row_cell_is_left_alone(self):
        tbl = footer_table()
        apply_table_footer_source_format(tbl)
        other = cell_at(tbl, 3, 1)  # "其他資料" — unmatched, in the top footer row
        # Font/alignment are NOT changed for an unmatched cell...
        self.assertEqual(cell_run_sizes(other), ["22"])
        # ...but the separator top border spans the WHOLE top footer row.
        self.assertTrue(is_double_black(tc_border(other, "top")))
        # No left/right override on the unmatched cell; bottom is still cleared
        # because every physical cell in the terminal footer row must be borderless.
        self.assertIsNone(tc_border(other, "left"))
        self.assertIsNone(tc_border(other, "right"))
        self.assertTrue(is_nil(tc_border(other, "bottom")))

    def test_empty_paragraph_without_run_does_not_crash(self):
        tbl = make_table(
            [
                ["標題"],
                ["a", "b", "c", "d", "e"],
                ["基期：x", "", "", "", ""],
            ]
        )
        # Remove the run from the 基期 paragraph to simulate an empty run.
        base_cell = cell_at(tbl, 2, 0)
        # Add a second, run-less paragraph to the cell.
        etree.SubElement(base_cell, qn("p"))
        apply_table_footer_source_format(tbl)
        # Both paragraphs are left-aligned, no exception raised.
        self.assertEqual(cell_paragraph_alignments(base_cell), ["left", "left"])

    def test_whitespace_and_newlines_do_not_break_prefix_match(self):
        tbl = make_table(
            [
                ["標題"],
                ["a", "b", "c", "d", "e"],
                ["  基期：　民國\n100年 ", "資料來源： 本所 ", "", "", ""],
            ]
        )
        result = apply_table_footer_source_format(tbl)
        self.assertEqual(result["footer_base_period_cells_adjusted"], 1)
        self.assertEqual(result["footer_source_cells_adjusted"], 1)
        self.assertEqual(cell_run_sizes(cell_at(tbl, 2, 0)), ["20"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 2, 1)), ["20"])

    def test_no_single_cell_first_row_means_no_first_row_override(self):
        tbl = make_table(
            [
                ["a", "b", "c", "d", "e"],
                ["1", "2", "3", "4", "5"],
                ["基期：x", "", "", "", ""],
            ]
        )
        result = apply_table_footer_source_format(tbl)
        self.assertFalse(result["first_row_single_cell_border_adjusted"])
        self.assertIsNone(cell_at(tbl, 0, 0).find("w:tcPr/w:tcBorders", NS))


class FooterFormatNormalizeTests(unittest.TestCase):
    def test_normalize_merges_paragraphs_and_strips_control_chars(self):
        tc = etree.Element(qn("tc"))
        for chunk in ("﻿基期：", "民國​100年 "):
            p = etree.SubElement(tc, qn("p"))
            r = etree.SubElement(p, qn("r"))
            t = etree.SubElement(r, qn("t"))
            t.text = chunk
        text = normalize_footer_source_cell_text(tc)
        self.assertTrue(text.startswith("基期："))
        self.assertNotIn("﻿", text)
        self.assertNotIn("​", text)


# ----------------------------------------------------------------------------
# Border-preservation guarantees (only the rule-cleared edges may disappear).
# ----------------------------------------------------------------------------
class FooterFormatBorderPreservationTests(unittest.TestCase):
    def test_pre_existing_outer_double_black_survives_on_non_cleared_edges(self):
        tbl = footer_table()
        # Give the table a full pre-existing double-black frame + grid.
        tbl_pr = etree.SubElement(tbl, qn("tblPr"))
        borders = etree.SubElement(tbl_pr, qn("tblBorders"))
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b = etree.SubElement(borders, qn(side))
            b.set(qn("val"), "double")
            b.set(qn("sz"), "4")
            b.set(qn("color"), "000000")

        apply_table_footer_source_format(tbl)

        # Outer frame keeps top/left/right double; footer mode clears bottom.
        for side in ("top", "left", "right"):
            self.assertTrue(is_double_black(tbl_border(tbl, side)), side)
        self.assertTrue(is_nil(tbl_border(tbl, "bottom")))
        # Inside gridlines are not a rule target, so they are preserved.
        self.assertTrue(is_double_black(tbl_border(tbl, "insideH")))
        self.assertTrue(is_double_black(tbl_border(tbl, "insideV")))

    def test_non_matching_last_row_cell_keeps_its_existing_borders(self):
        tbl = footer_table()
        # Give the non-matching last-row cell a full double-black border.
        other = cell_at(tbl, 3, 1)
        tc_pr = etree.SubElement(other, qn("tcPr"))
        borders = etree.SubElement(tc_pr, qn("tcBorders"))
        for side in ("top", "left", "bottom", "right"):
            b = etree.SubElement(borders, qn(side))
            b.set(qn("val"), "double")
            b.set(qn("sz"), "4")
            b.set(qn("color"), "000000")

        apply_table_footer_source_format(tbl)

        for side in ("top", "left", "right"):
            self.assertTrue(is_double_black(tc_border(other, side)), side)
        self.assertTrue(is_nil(tc_border(other, "bottom")))

    def test_matched_cell_pre_existing_top_not_destroyed_by_nil_edges(self):
        # A matched 基期 cell with a pre-existing double-black top must keep a
        # double-black top (rule sets top double) while left/right/bottom are
        # cleared for the terminal footer row.
        tbl = footer_table()
        base_cell = cell_at(tbl, 3, 0)
        tc_pr = etree.SubElement(base_cell, qn("tcPr"))
        borders = etree.SubElement(tc_pr, qn("tcBorders"))
        for side in ("top", "left", "bottom", "right"):
            b = etree.SubElement(borders, qn(side))
            b.set(qn("val"), "double")
            b.set(qn("sz"), "4")
            b.set(qn("color"), "000000")

        apply_table_footer_source_format(tbl)

        self.assertTrue(is_double_black(tc_border(base_cell, "top")))
        self.assertTrue(is_nil(tc_border(base_cell, "left")))
        self.assertTrue(is_nil(tc_border(base_cell, "right")))
        self.assertTrue(is_nil(tc_border(base_cell, "bottom")))


# ----------------------------------------------------------------------------
# Priority / order checks via the full pipeline.
# ----------------------------------------------------------------------------
class FooterFormatOrderTests(unittest.TestCase):
    def _doc(self):
        return build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),  # first table -> skipped
            make_paragraph("一、報表"),
            footer_table(),
        )

    def test_pipeline_applies_full_format_in_priority_order(self):
        summary, root = run_fix(self._doc(), footer_options())
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

        # 11 pt over the body, 10 pt for footer cells (10 pt not lost to 11 pt).
        self.assertEqual(cell_run_sizes(cell_at(tbl, 1, 0)), ["22"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 0)), ["20"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 2)), ["20"])

        # Outer frame keeps top/left/right double; footer terminal bottom is nil.
        for side in ("top", "left", "right"):
            self.assertTrue(is_double_black(tbl_border(tbl, side)), side)
        self.assertTrue(is_nil(tbl_border(tbl, "bottom")))

        # First-row single cell top/left/right nil survive the outer frame.
        title = cell_at(tbl, 0, 0)
        self.assertTrue(is_nil(tc_border(title, "top")))
        self.assertTrue(is_nil(tc_border(title, "left")))
        self.assertTrue(is_nil(tc_border(title, "right")))
        self.assertTrue(is_double_black(tc_border(title, "bottom")))

        # Last-row footer cells: left/right/bottom nil survive the outer frame.
        # Alignment overrides the centered content from layout formatting.
        base_cell = cell_at(tbl, 3, 0)
        self.assertEqual(cell_paragraph_alignments(base_cell), ["left"])
        self.assertTrue(is_nil(tc_border(base_cell, "left")))
        self.assertTrue(is_nil(tc_border(base_cell, "right")))
        self.assertTrue(is_nil(tc_border(base_cell, "bottom")))
        self.assertTrue(is_double_black(tc_border(base_cell, "top")))

        source_cell = cell_at(tbl, 3, 2)
        self.assertEqual(cell_paragraph_alignments(source_cell), ["right"])

    def test_log_records_footer_fields(self):
        summary, _ = run_fix(self._doc(), footer_options())
        record = summary.table_log_records[1]
        self.assertTrue(record["table_footer_note_source_format_enabled"])
        self.assertTrue(record["table_footer_note_source_format_applied"])
        self.assertTrue(record["outer_double_border_applied_by_footer_source_format"])
        self.assertEqual(record["table_bottom_border_mode"], "footer_none")
        self.assertEqual(record["table_bottom_border_cell_count"], 5)
        self.assertTrue(record["table_bottom_border_xml_verified"])
        self.assertTrue(record["footer_terminal_bottom_none_applied"])
        self.assertEqual(record["footer_terminal_bottom_none_cell_count"], 5)
        self.assertFalse(record["table_bottom_double_border_applied"])
        self.assertEqual(record["table_bottom_double_border_cell_count"], 0)
        self.assertFalse(record["table_bottom_double_border_xml_verified"])
        self.assertEqual(record["last_row_physical_cell_count"], 5)
        self.assertEqual(record["last_row_grid_span_sum"], 5)
        self.assertEqual(record["last_row_bottom_edge_target_count"], 5)
        self.assertTrue(record["table_border_schema_order_valid"])
        self.assertIn(
            "last_row_tc_bottoms=nil/missing/missing",
            record["table_bottom_border_verify_detail"],
        )
        self.assertTrue(record["first_row_single_cell_border_adjusted"])
        self.assertEqual(record["footer_base_period_cells_adjusted"], 1)
        self.assertEqual(record["footer_source_cells_adjusted"], 1)
        self.assertEqual(record["footer_note_cells_adjusted"], 0)
        self.assertCountEqual(
            record["footer_note_cell_matches"], ["base_period", "source"]
        )
        self.assertTrue(record["footer_block_top_border_applied"])
        self.assertEqual(record["footer_row_count"], 1)
        self.assertTrue(any("基期" in line for line in record["footer_note_cell_debug"]))
        self.assertEqual(record["table_footer_note_source_format_skipped_reason"], "none")
        self.assertEqual(summary.table_footer_source_format_tables, 1)


# ----------------------------------------------------------------------------
# Default off + independence from existing flows.
# ----------------------------------------------------------------------------
class FooterFormatDefaultTests(unittest.TestCase):
    def test_option_defaults_to_false(self):
        self.assertFalse(ProcessOptions(True, True, False).enable_table_footer_source_format)

    def test_disabled_does_not_change_table(self):
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),
            make_paragraph("一、報表"),
            footer_table(),
        )
        options = footer_options(enable_table_footer_source_format=False)
        summary, root = run_fix(document, options)
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

        # No outer double frame, no footer cell overrides.
        self.assertIsNone(tbl_border(tbl, "top"))
        self.assertIsNone(cell_at(tbl, 0, 0).find("w:tcPr/w:tcBorders", NS))
        self.assertIsNone(cell_at(tbl, 3, 0).find("w:tcPr/w:tcBorders", NS))
        # 基期 cell stays at the normal 11 pt table body size.
        self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 0)), ["22"])
        self.assertEqual(summary.table_footer_source_format_tables, 0)
        record = summary.table_log_records[1]
        self.assertFalse(record["table_footer_note_source_format_enabled"])
        self.assertFalse(record["table_footer_note_source_format_applied"])
        self.assertEqual(
            record["table_footer_note_source_format_skipped_reason"], "feature_disabled"
        )


# ----------------------------------------------------------------------------
# Relationship with existing skip logic (section 10).
# ----------------------------------------------------------------------------
CHAPTER_THREE_TITLE = "參、價格形成之主要因素分析"


class FooterFormatSkipLogicTests(unittest.TestCase):
    def test_chapter_three_layout_skip_still_allows_footer_when_layout_requested(self):
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),
            make_paragraph(CHAPTER_THREE_TITLE),
            footer_table(),
        )
        options = footer_options(
            fix_color=True,
            skip_chapter_three_table_layout=True,
            skip_chapter_three_table_color=False,
        )
        summary, root = run_fix(document, options)
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

        # Layout is protected, but the footer post-process still runs because
        # table layout was requested globally.
        self.assertTrue(is_double_black(tbl_border(tbl, "top")))
        self.assertEqual(summary.table_footer_source_format_tables, 1)
        record = summary.table_log_records[1]
        self.assertTrue(record["chapter_three_table_layout_skipped"])
        self.assertFalse(record["chapter_three_table_color_skipped"])
        self.assertFalse(record["layout_fixed"])
        self.assertTrue(record["color_fixed"])
        self.assertTrue(record["table_footer_note_source_format_should_apply"])
        self.assertTrue(record["table_footer_note_source_format_applied"])
        self.assertEqual(record["table_footer_note_source_format_skipped_reason"], "none")

    def test_global_layout_disabled_still_blocks_footer_format(self):
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),
            make_paragraph("一、報表"),
            footer_table(),
        )
        options = footer_options(fix_table_layout=False, fix_color=True)
        summary, root = run_fix(document, options)
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

        self.assertIsNone(tbl_border(tbl, "top"))
        self.assertEqual(summary.table_footer_source_format_tables, 0)
        record = summary.table_log_records[1]
        self.assertFalse(record["layout_fixed"])
        self.assertTrue(record["color_fixed"])
        self.assertFalse(record["table_footer_note_source_format_should_apply"])
        self.assertFalse(record["table_footer_note_source_format_applied"])
        self.assertEqual(
            record["table_footer_note_source_format_skipped_reason"],
            "layout not adjusted for this table",
        )

    def test_color_only_skip_does_not_block_footer_format(self):
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),
            make_paragraph(CHAPTER_THREE_TITLE),
            footer_table(),
        )
        options = footer_options(
            fix_color=True,
            skip_chapter_three_table_layout=False,
            skip_chapter_three_table_color=True,
        )
        summary, root = run_fix(document, options)
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

        # Only color is protected; layout (and thus the footer format) still runs.
        self.assertTrue(is_double_black(tbl_border(tbl, "top")))
        self.assertTrue(record_applied(summary))

    def test_chapter_three_layout_and_color_protection_still_records_footer(self):
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),
            make_paragraph("參、受保護章節"),
            footer_table(),
        )
        options = footer_options(skip_chapter_three_adjustments=True)
        summary, root = run_fix(document, options)
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

        self.assertTrue(is_double_black(tbl_border(tbl, "top")))
        self.assertEqual(summary.table_footer_source_format_tables, 1)
        record = summary.table_log_records[1]
        self.assertEqual(record["table_type"], "skipped_chapter_three_table")
        self.assertFalse(record["layout_fixed"])
        self.assertFalse(record["color_fixed"])
        self.assertTrue(record["chapter_three_table_layout_skipped"])
        self.assertTrue(record["chapter_three_table_color_skipped"])
        self.assertTrue(record["table_footer_note_source_format_should_apply"])
        self.assertTrue(record["table_footer_note_source_format_applied"])
        self.assertEqual(record["table_bottom_border_mode"], "footer_none")
        self.assertTrue(record["table_bottom_border_xml_verified"])
        self.assertFalse(record["table_bottom_double_border_applied"])
        self.assertFalse(record["table_bottom_double_border_xml_verified"])
        self.assertEqual(record["table_footer_note_source_format_skipped_reason"], "none")
        self.assertIsNone(tbl.find("w:tblPr/w:tblW", NS))
        self.assertIsNone(tbl.find("w:tblPr/w:tblLayout", NS))

    def test_explicit_chapter_three_layout_and_color_skips_do_not_block_footer(self):
        protected_table = footer_table()
        shaded_cell = cell_at(protected_table, 2, 0)
        tc_pr = shaded_cell.find("w:tcPr", NS)
        if tc_pr is None:
            tc_pr = etree.Element(qn("tcPr"))
            shaded_cell.insert(0, tc_pr)
        shd = etree.SubElement(tc_pr, qn("shd"))
        shd.set(qn("fill"), "BFBFBF")
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),
            make_paragraph(CHAPTER_THREE_TITLE),
            protected_table,
        )
        options = footer_options(
            fix_color=True,
            skip_chapter_three_table_layout=True,
            skip_chapter_three_table_color=True,
        )

        summary, root = run_fix(document, options)
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]
        record = summary.table_log_records[1]

        self.assertEqual(record["table_type"], "skipped_chapter_three_table")
        self.assertFalse(record["layout_fixed"])
        self.assertFalse(record["color_fixed"])
        self.assertTrue(record["chapter_three_table_layout_skipped"])
        self.assertTrue(record["chapter_three_table_color_skipped"])
        self.assertEqual(len(summary.table_footer_source_format_records), 1)
        self.assertTrue(record["table_footer_note_source_format_applied"])
        self.assertEqual(record["table_bottom_border_mode"], "footer_none")
        self.assertEqual(record["table_bottom_border_cell_count"], 5)
        self.assertTrue(record["table_bottom_border_xml_verified"])
        self.assertFalse(record["table_bottom_double_border_applied"])
        self.assertEqual(record["table_bottom_double_border_cell_count"], 0)
        self.assertFalse(record["table_bottom_double_border_xml_verified"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 0)), ["20"])
        self.assertEqual(cell_paragraph_alignments(cell_at(tbl, 3, 0)), ["left"])
        self.assertEqual(cell_paragraph_alignments(cell_at(tbl, 3, 2)), ["right"])
        self.assertTrue(is_nil(tc_border(cell_at(tbl, 3, 0), "left")))
        self.assertTrue(is_nil(tc_border(cell_at(tbl, 3, 0), "right")))
        self.assertTrue(is_nil(tc_border(cell_at(tbl, 3, 0), "bottom")))
        self.assertIsNone(tbl.find("w:tblPr/w:tblW", NS))
        self.assertIsNone(tbl.find("w:tblPr/w:tblLayout", NS))
        self.assertEqual(
            cell_at(tbl, 2, 0).find("w:tcPr/w:shd", NS).get(qn("fill")),
            "BFBFBF",
        )
        self.assertEqual(summary.changed_to_gray, 0)

    def test_small_table_is_not_footer_formatted(self):
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),
            make_paragraph("一、小表格"),
            make_table([["基期：x", "y"]]),  # 2 cells -> small table skip
        )
        summary, root = run_fix(document, footer_options())
        self.assertEqual(summary.table_footer_source_format_tables, 0)
        record = summary.table_log_records[1]
        self.assertEqual(record["table_type"], "skipped_small_table")
        self.assertFalse(record["table_footer_note_source_format_applied"])

    def test_first_document_table_is_not_footer_formatted(self):
        document = build_document(
            make_paragraph("封面"),
            footer_table(),  # first table -> always skipped
        )
        summary, root = run_fix(document, footer_options())
        record = summary.table_log_records[0]
        self.assertEqual(record["table_type"], "skipped_first_table")
        self.assertFalse(record["table_footer_note_source_format_applied"])
        self.assertEqual(summary.table_footer_source_format_tables, 0)


def record_applied(summary) -> bool:
    return summary.table_log_records[1]["table_footer_note_source_format_applied"]


def note_last_row_table(note_text: str):
    """Title row (single cell) + header + data + a last row whose first cell is
    a note and the rest are non-matching."""
    return make_table(
        [
            ["報表標題"],
            ["項目", "金額", "比率", "排名", "備註"],
            ["土地", "100", "50%", "1", "ok"],
            [note_text, "其他", "說明", "", ""],
        ]
    )


def make_gridspan_note_table(note_text: str, span: int = 5):
    """A table whose last row is a single horizontally-merged note cell."""
    tbl = etree.Element(qn("tbl"))
    for row in (
        ["項目", "金額", "比率", "排名", "備註"],
        ["土地", "100", "50%", "1", "ok"],
    ):
        tr = etree.SubElement(tbl, qn("tr"))
        for text in row:
            tc = etree.SubElement(tr, qn("tc"))
            p = etree.SubElement(tc, qn("p"))
            r = etree.SubElement(p, qn("r"))
            t = etree.SubElement(r, qn("t"))
            t.text = text
    # Last row: one merged cell spanning all columns.
    tr = etree.SubElement(tbl, qn("tr"))
    tc = etree.SubElement(tr, qn("tc"))
    tc_pr = etree.SubElement(tc, qn("tcPr"))
    grid_span = etree.SubElement(tc_pr, qn("gridSpan"))
    grid_span.set(qn("val"), str(span))
    p = etree.SubElement(tc, qn("p"))
    r = etree.SubElement(p, qn("r"))
    t = etree.SubElement(r, qn("t"))
    t.text = note_text
    return tbl


def make_vmerge_table(last_row_text: str):
    """A 3-column table whose first column vertically merges into the last row."""
    tbl = etree.Element(qn("tbl"))

    def add_cell(tr, text: str, *, vmerge: str | None = None):
        tc = etree.SubElement(tr, qn("tc"))
        if vmerge is not None:
            tc_pr = etree.SubElement(tc, qn("tcPr"))
            merge = etree.SubElement(tc_pr, qn("vMerge"))
            if vmerge == "restart":
                merge.set(qn("val"), "restart")
        p = etree.SubElement(tc, qn("p"))
        r = etree.SubElement(p, qn("r"))
        t = etree.SubElement(r, qn("t"))
        t.text = text
        return tc

    tr = etree.SubElement(tbl, qn("tr"))
    add_cell(tr, "owner", vmerge="restart")
    add_cell(tr, "h2")
    add_cell(tr, "h3")

    tr = etree.SubElement(tbl, qn("tr"))
    add_cell(tr, "", vmerge="continue")
    add_cell(tr, "mid")
    add_cell(tr, "row")

    tr = etree.SubElement(tbl, qn("tr"))
    add_cell(tr, "", vmerge="continue")
    add_cell(tr, last_row_text)
    add_cell(tr, "tail")
    return tbl


# ----------------------------------------------------------------------------
# Last-row note cell regex and formatting.
# ----------------------------------------------------------------------------
class FooterNoteRegexTests(unittest.TestCase):
    def test_matches_note_with_and_without_number(self):
        for text in ("註：", "註:", "註1：", "註1:", "註2：", "註3:", "註10：", "註10:"):
            self.assertIsNotNone(FOOTER_NOTE_PREFIX_PATTERN.match(text), text)

    def test_matches_note_with_spaces(self):
        # Optional half-width / full-width / tab spaces around the number.
        for text in ("註 1：", "註　1：", "註 2：", "註　2：", "註 10:", "註\t3："):
            self.assertIsNotNone(FOOTER_NOTE_PREFIX_PATTERN.match(text), text)

    def test_rejects_lookalikes(self):
        for text in ("備註：", "註記：", "本註：", "說明註：", "註冊：", "註明事項"):
            self.assertIsNone(FOOTER_NOTE_PREFIX_PATTERN.match(text), text)


class FooterNoteCellUnitTests(unittest.TestCase):
    def _assert_note_cell_formatted(self, tbl, row, col):
        cell = cell_at(tbl, row, col)
        self.assertEqual(cell_run_sizes(cell), ["20"])  # 10 pt, not 11 pt
        self.assertEqual(cell_paragraph_alignments(cell), ["left"])
        self.assertTrue(is_double_black(tc_border(cell, "top")))
        self.assertTrue(is_nil(tc_border(cell, "left")))
        self.assertTrue(is_nil(tc_border(cell, "right")))
        self.assertTrue(is_nil(tc_border(cell, "bottom")))

    def test_note_colon_cell_is_formatted(self):
        tbl = note_last_row_table("註：本表為新臺幣元")
        result = apply_table_footer_source_format(tbl)
        self._assert_note_cell_formatted(tbl, 3, 0)
        self.assertEqual(result["footer_note_cell_matches"], ["note"])

    def test_note_halfwidth_colon_with_number_is_formatted(self):
        tbl = note_last_row_table("註1: 資料估計")
        apply_table_footer_source_format(tbl)
        self._assert_note_cell_formatted(tbl, 3, 0)

    def test_note_fullwidth_colon_with_number_is_formatted(self):
        tbl = note_last_row_table("註2：資料估計")
        apply_table_footer_source_format(tbl)
        self._assert_note_cell_formatted(tbl, 3, 0)

    def test_note_two_digit_number_is_formatted(self):
        tbl = note_last_row_table("註10：詳附表")
        apply_table_footer_source_format(tbl)
        self._assert_note_cell_formatted(tbl, 3, 0)

    def test_note_lookalikes_are_not_formatted(self):
        for text in ("備註：說明", "註記：說明", "本註：說明"):
            tbl = note_last_row_table(text)
            result = apply_table_footer_source_format(tbl)
            self.assertEqual(result["footer_note_cells_adjusted"], 0, text)
            # The cell keeps the whole-table 11 pt and gets only the data-table
            # bottom edge, not footer-cell-specific formatting.
            self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 0)), ["22"], text)
            self.assertIsNone(tc_border(cell_at(tbl, 3, 0), "top"), text)
            self.assertIsNone(tc_border(cell_at(tbl, 3, 0), "left"), text)
            self.assertIsNone(tc_border(cell_at(tbl, 3, 0), "right"), text)
            self.assertTrue(is_double_black(tc_border(cell_at(tbl, 3, 0), "bottom")), text)

    def test_other_last_row_cells_are_not_changed_to_10pt(self):
        tbl = note_last_row_table("註：說明")
        apply_table_footer_source_format(tbl)
        # "其他" / "說明" cells stay at 11 pt and keep their alignment; only the
        # separator top border is added across the whole top footer row.
        for col in (1, 2):
            self.assertEqual(cell_run_sizes(cell_at(tbl, 3, col)), ["22"])
            self.assertTrue(is_double_black(tc_border(cell_at(tbl, 3, col), "top")))
            self.assertIsNone(tc_border(cell_at(tbl, 3, col), "left"))
            self.assertTrue(is_nil(tc_border(cell_at(tbl, 3, col), "bottom")))

    def test_note_in_non_last_row_is_ignored(self):
        tbl = make_table(
            [
                ["報表標題"],
                ["註1：這在第二列", "x", "y", "z", "w"],  # not the last row
                ["土地", "100", "50%", "1", "ok"],
                ["基期：民國100年", "", "", "", ""],
            ]
        )
        result = apply_table_footer_source_format(tbl)
        # Only the last-row 基期 cell matched; the 2nd-row note was untouched.
        self.assertEqual(result["footer_note_cell_matches"], ["base_period"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 1, 0)), ["22"])

    def test_note_base_period_and_source_coexist_in_last_row(self):
        tbl = make_table(
            [
                ["報表標題"],
                ["項目", "金額", "比率", "排名", "備註"],
                ["註：單位元", "基期：100年", "資料來源：本所", "其他", ""],
            ]
        )
        result = apply_table_footer_source_format(tbl)
        self.assertCountEqual(
            result["footer_note_cell_matches"], ["note", "base_period", "source"]
        )
        self.assertEqual(cell_paragraph_alignments(cell_at(tbl, 2, 0)), ["left"])
        self.assertEqual(cell_paragraph_alignments(cell_at(tbl, 2, 1)), ["left"])
        self.assertEqual(cell_paragraph_alignments(cell_at(tbl, 2, 2)), ["right"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 2, 3)), ["22"])  # 其他 unchanged

    def test_whitespace_and_newline_note_still_matches(self):
        tbl = note_last_row_table("  註1：　本表\n說明 ")
        result = apply_table_footer_source_format(tbl)
        self.assertEqual(result["footer_note_cell_matches"], ["note"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 0)), ["20"])

    def test_merged_note_cell_processed_once(self):
        tbl = make_gridspan_note_table("註：本表單位為新臺幣元", span=5)
        result = apply_table_footer_source_format(tbl)
        # The single merged last-row cell is formatted and logged exactly once.
        self.assertEqual(result["footer_note_cells_adjusted"], 1)
        self.assertEqual(result["footer_note_cell_matches"], ["note"])

    def test_note_cell_pre_existing_double_black_top_preserved(self):
        tbl = note_last_row_table("註：說明")
        note_cell = cell_at(tbl, 3, 0)
        tc_pr = etree.SubElement(note_cell, qn("tcPr"))
        borders = etree.SubElement(tc_pr, qn("tcBorders"))
        for side in ("top", "left", "bottom", "right"):
            b = etree.SubElement(borders, qn(side))
            b.set(qn("val"), "double")
            b.set(qn("sz"), "4")
            b.set(qn("color"), "000000")

        apply_table_footer_source_format(tbl)

        self.assertTrue(is_double_black(tc_border(note_cell, "top")))
        self.assertTrue(is_nil(tc_border(note_cell, "left")))
        self.assertTrue(is_nil(tc_border(note_cell, "right")))
        self.assertTrue(is_nil(tc_border(note_cell, "bottom")))


def special_footer_table(note_text="註：說明"):
    """A 4-column (special) table: title row + header + data + footer row."""
    return make_table(
        [
            ["報表標題"],
            ["項目", "金額", "比率", "備註"],
            ["土地", "100", "50%", "ok"],
            [note_text, "基期：100年", "資料來源：本所", ""],
        ]
    )


def reset_document_tables_to_11pt_center(docx_path) -> None:
    """Simulate Word COM / fallback clobbering: force every run to 11 pt and
    every paragraph to centered in word/document.xml, then re-save the docx."""
    with ZipFile(docx_path) as zin:
        items = zin.infolist()
        data = {item.filename: zin.read(item.filename) for item in items}

    root = etree.fromstring(data["word/document.xml"])
    for run in root.xpath(".//w:r", namespaces=NS):
        rpr = run.find("w:rPr", NS)
        if rpr is None:
            rpr = etree.SubElement(run, qn("rPr"))
        sz = rpr.find("w:sz", NS)
        if sz is None:
            sz = etree.SubElement(rpr, qn("sz"))
        sz.set(qn("val"), "22")
    for p in root.xpath(".//w:p", namespaces=NS):
        ppr = p.find("w:pPr", NS)
        if ppr is None:
            ppr = etree.Element(qn("pPr"))
            p.insert(0, ppr)
        jc = ppr.find("w:jc", NS)
        if jc is None:
            jc = etree.SubElement(ppr, qn("jc"))
        jc.set(qn("val"), "center")

    data["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    with ZipFile(docx_path, "w", ZIP_DEFLATED) as zout:
        for item in items:
            zout.writestr(item, data[item.filename])


class FooterFormatSpecialTableTests(unittest.TestCase):
    """Special tables (column_count <= 4) must also go through the final footer
    formatting, exactly like normal tables."""

    def _doc(self, note_text="註：說明"):
        return build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),  # first table -> skipped
            make_paragraph("一、特殊表格"),
            special_footer_table(note_text),
        )

    def test_special_table_note_cell_is_10pt_left(self):
        summary, root = run_fix(self._doc("註：說明"), footer_options())
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]
        self.assertEqual(summary.table_log_records[1]["table_type"], "special_table")
        note_cell = cell_at(tbl, 3, 0)
        self.assertEqual(cell_run_sizes(note_cell), ["20"])
        self.assertEqual(cell_paragraph_alignments(note_cell), ["left"])
        self.assertTrue(is_double_black(tc_border(note_cell, "top")))
        self.assertTrue(is_nil(tc_border(note_cell, "bottom")))

    def test_special_table_data_source_cell_is_10pt_right(self):
        summary, root = run_fix(self._doc(), footer_options())
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]
        source_cell = cell_at(tbl, 3, 2)
        self.assertEqual(cell_run_sizes(source_cell), ["20"])
        self.assertEqual(cell_paragraph_alignments(source_cell), ["right"])
        record = summary.table_log_records[1]
        self.assertTrue(record["table_footer_note_source_format_applied"])
        self.assertCountEqual(
            record["footer_note_cell_matches"], ["note", "base_period", "source"]
        )


class FooterFormatWordComOrderTests(unittest.TestCase):
    """The footer format must be the FINAL table step: it must survive both a
    Word COM AutoFit success (which re-saves the doc) and the XML fallback
    (which re-applies apply_table_format -> 11 pt + center)."""

    def _doc(self, note_text="基期：民國100年"):
        # The footer table is a 5-column normal table, so it is queued for Word
        # COM AutoFit (global_table_index = 2).
        return build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),  # first table -> skipped
            make_paragraph("一、報表"),
            note_last_row_table(note_text),
        )

    def _assert_footer_cell_ok(self, root):
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]
        note_cell = cell_at(tbl, 3, 0)
        # Must NOT be left at 11 pt / center.
        self.assertEqual(cell_run_sizes(note_cell), ["20"])
        self.assertEqual(cell_paragraph_alignments(note_cell), ["left"])
        self.assertTrue(is_nil(tc_border(note_cell, "left")))
        self.assertTrue(is_nil(tc_border(note_cell, "right")))
        self.assertTrue(is_nil(tc_border(note_cell, "bottom")))
        self.assertTrue(is_double_black(tc_border(note_cell, "top")))
        return tbl

    def test_no_word_com_footer_applies(self):
        summary, root = run_fix(self._doc("註：說明"), footer_options())
        self._assert_footer_cell_ok(root)
        self.assertTrue(summary.table_log_records[1]["table_footer_note_source_format_applied"])

    def test_word_com_success_does_not_clobber_footer(self):
        def fake_autofit(docx_path, records, stop=None):
            # Word COM "succeeds" but its re-save reset fonts/alignment.
            reset_document_tables_to_11pt_center(docx_path)
            return (
                ["WORD_COM_TABLE_AUTOFIT_APPLIED global_table_index=2 sequence=content_then_window"],
                {2},
                set(),
            )

        with patch(
            "docx_fixer.docx_processor.apply_table_autofit_with_word_com",
            side_effect=fake_autofit,
        ) as autofit:
            summary, root = run_fix(
                self._doc("基期：民國100年"),
                footer_options(normalize_with_word_com=True),
            )

        autofit.assert_called_once()
        # Even though Word COM clobbered to 11 pt/center, the final footer pass
        # restored 10 pt/left and the borders.
        self._assert_footer_cell_ok(root)
        record = summary.table_log_records[1]
        self.assertTrue(record["table_footer_note_source_format_applied"])
        self.assertEqual(summary.table_footer_source_format_tables, 1)

    def test_word_com_failure_fallback_does_not_clobber_footer(self):
        def fake_autofit(docx_path, records, stop=None):
            # Word COM fails -> the real XML fallback (apply_table_format) runs
            # and resets the table to 11 pt / centered before the footer pass.
            failed = {int(r["global_table_index"]) for r in records}
            return (
                ["WORD_COM_TABLE_AUTOFIT_EXCEPTION type=Test message=boom"],
                set(),
                failed,
            )

        with patch(
            "docx_fixer.docx_processor.apply_table_autofit_with_word_com",
            side_effect=fake_autofit,
        ):
            summary, root = run_fix(
                self._doc("資料來源：本所"),
                footer_options(normalize_with_word_com=True),
            )

        # The fallback ran (apply_table_format), yet the footer survived.
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]
        source_cell = cell_at(tbl, 3, 0)
        self.assertEqual(cell_run_sizes(source_cell), ["20"])
        self.assertEqual(cell_paragraph_alignments(source_cell), ["right"])
        self.assertTrue(is_double_black(tc_border(source_cell, "top")))
        record = summary.table_log_records[1]
        self.assertTrue(record["table_footer_note_source_format_applied"])
        self.assertTrue(record["word_com_autofit_fallback_applied"])

    def test_table_log_marks_should_apply_and_final_applied(self):
        summary, _ = run_fix(self._doc("註：說明"), footer_options())
        record = summary.table_log_records[1]
        self.assertTrue(record["table_footer_note_source_format_enabled"])
        self.assertTrue(record["table_footer_note_source_format_should_apply"])
        self.assertTrue(record["table_footer_note_source_format_applied"])
        # The reapply ran in the final post-process.
        logs = "\n".join(summary.table_footer_source_format_logs)
        self.assertIn("FOOTER_SOURCE_FORMAT_REAPPLY_APPLIED", logs)
        self.assertIn("table_bottom_border_mode=footer_none", logs)
        self.assertIn("table_bottom_border_xml_verified=True", logs)
        self.assertIn("table_bottom_double_border_xml_verified=False", logs)
        self.assertIn("FOOTER_SOURCE_FORMAT_REAPPLY_DONE applied=1", logs)


class FooterNotePipelineTests(unittest.TestCase):
    def _doc(self, note_text="註：本表單位為新臺幣元"):
        return build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),  # first table -> skipped
            make_paragraph("一、報表"),
            note_last_row_table(note_text),
        )

    def test_pipeline_formats_note_and_preserves_priority_order(self):
        summary, root = run_fix(self._doc(), footer_options())
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

        # 11 pt body, 10 pt note (10 pt not lost to 11 pt).
        self.assertEqual(cell_run_sizes(cell_at(tbl, 1, 0)), ["22"])
        note_cell = cell_at(tbl, 3, 0)
        self.assertEqual(cell_run_sizes(note_cell), ["20"])
        self.assertEqual(cell_paragraph_alignments(note_cell), ["left"])

        # Outer frame keeps top/left/right double; note cell nil edges survive it.
        for side in ("top", "left", "right"):
            self.assertTrue(is_double_black(tbl_border(tbl, side)), side)
        self.assertTrue(is_nil(tbl_border(tbl, "bottom")))
        self.assertTrue(is_nil(tc_border(note_cell, "left")))
        self.assertTrue(is_nil(tc_border(note_cell, "right")))
        self.assertTrue(is_nil(tc_border(note_cell, "bottom")))
        self.assertTrue(is_double_black(tc_border(note_cell, "top")))

        record = summary.table_log_records[1]
        self.assertTrue(record["table_footer_note_source_format_applied"])
        self.assertEqual(record["footer_note_cells_adjusted"], 1)
        self.assertEqual(record["footer_note_cell_matches"], ["note"])
        self.assertTrue(any("note:" in line for line in record["footer_note_cell_debug"]))

    def test_table_log_reports_note_move_force_disabled(self):
        summary, _ = run_fix(self._doc(), footer_options())
        record = summary.table_log_records[1]
        # The legacy note-move feature is hidden and off for this default run.
        self.assertTrue(record["table_note_move_gui_hidden"])
        self.assertTrue(record["table_note_move_forced_false"])
        # No note cell was moved out of the table.
        self.assertFalse(record["note_cells_moved"])
        self.assertEqual(summary.note_cells_moved_tables, 0)
        # The note paragraph was NOT added below the table (structure intact).
        tbl_count = len(summary.table_log_records)
        self.assertGreaterEqual(tbl_count, 2)


# ----------------------------------------------------------------------------
# Footer rows are scanned UPWARD from the bottom, stopping at the first
# non-footer row (only the contiguous bottom block is processed).
# ----------------------------------------------------------------------------
class FooterContiguousBottomScanTests(unittest.TestCase):
    def test_multiple_contiguous_footer_rows_are_all_processed(self):
        tbl = make_table(
            [
                ["項目", "金額", "比率", "排名", "備註"],  # row 0 header
                ["土地", "100", "50%", "1", "ok"],  # row 1 data
                ["註：本表單位元", "", "", "", ""],  # row 2 footer (note)
                ["基期：民國100年", "", "資料來源：本所", "", ""],  # row 3 footer
            ]
        )
        result = apply_table_footer_source_format(tbl)

        # Rows 3 and 2 are the contiguous footer block; row 1 (data) stops it.
        self.assertEqual(result["footer_row_count"], 2)
        self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 0)), ["20"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 2)), ["20"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 2, 0)), ["20"])
        self.assertEqual(cell_paragraph_alignments(cell_at(tbl, 2, 0)), ["left"])
        # Borders applied to the footer cells above the last row too.
        self.assertTrue(is_double_black(tc_border(cell_at(tbl, 2, 0), "top")))
        self.assertTrue(is_nil(tc_border(cell_at(tbl, 2, 0), "bottom")))
        # Row 1 (data) untouched.
        self.assertEqual(cell_run_sizes(cell_at(tbl, 1, 0)), ["22"])
        self.assertIsNone(cell_at(tbl, 1, 0).find("w:tcPr/w:tcBorders", NS))

        # footer_cell_matches is top-to-bottom: row 2 (note) first, then row 3.
        self.assertEqual(len(result["footer_cell_matches"]), 2)
        self.assertEqual(result["footer_cell_matches"][0], "note")
        self.assertCountEqual(
            result["footer_cell_matches"][1].split(","), ["base_period", "source"]
        )
        # Per-type counts: 1 note (row 2) + 1 base_period + 1 source (row 3).
        self.assertEqual(result["footer_note_cells_adjusted"], 1)
        self.assertEqual(result["footer_base_period_cells_adjusted"], 1)
        self.assertEqual(result["footer_source_cells_adjusted"], 1)
        # Only the top footer row (row 2) draws the separator; row 3's top is cleared.
        self.assertTrue(result["footer_block_top_border_applied"])
        self.assertEqual(result["footer_internal_top_borders_cleared"], 1)
        self.assertTrue(is_double_black(tc_border(cell_at(tbl, 2, 0), "top")))
        self.assertTrue(is_nil(tc_border(cell_at(tbl, 3, 0), "top")))

    def test_data_row_between_footers_stops_the_scan(self):
        tbl = make_table(
            [
                ["項目", "金額", "比率", "排名", "備註"],
                ["註：上方非連續註記", "", "", "", ""],  # row 1: note ABOVE a data row
                ["土地", "100", "50%", "1", "ok"],  # row 2: data
                ["基期：民國100年", "", "", "", ""],  # row 3: footer (bottom)
            ]
        )
        result = apply_table_footer_source_format(tbl)

        # Only the bottom row is processed; row 2 (data) ends the scan, so the
        # non-contiguous note in row 1 is left alone.
        self.assertEqual(result["footer_row_count"], 1)
        self.assertEqual(result["footer_note_cell_matches"], ["base_period"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 0)), ["20"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 1, 0)), ["22"])
        self.assertIsNone(cell_at(tbl, 1, 0).find("w:tcPr/w:tcBorders", NS))

    def test_blank_row_stops_the_scan(self):
        tbl = make_table(
            [
                ["註：頂部註記", "", "", "", ""],  # row 0: note (top)
                ["", "", "", "", ""],  # row 1: blank
                ["基期：100年", "", "", "", ""],  # row 2: footer (bottom)
            ]
        )
        result = apply_table_footer_source_format(tbl)

        self.assertEqual(result["footer_row_count"], 1)
        self.assertEqual(cell_run_sizes(cell_at(tbl, 2, 0)), ["20"])
        # Blank row stops the scan before the top note row.
        self.assertEqual(cell_run_sizes(cell_at(tbl, 0, 0)), ["22"])

    def test_no_footer_row_at_bottom_processes_nothing(self):
        tbl = make_table(
            [
                ["註：上方註記", "", "", "", ""],  # footer-like, but not at bottom
                ["項目", "金額", "比率", "排名", "備註"],
                ["土地", "100", "50%", "1", "ok"],  # bottom row is data
            ]
        )
        result = apply_table_footer_source_format(tbl)

        self.assertEqual(result["footer_row_count"], 0)
        self.assertEqual(result["footer_note_cells_adjusted"], 0)
        self.assertEqual(cell_run_sizes(cell_at(tbl, 0, 0)), ["22"])
        self.assertFalse(result["footer_rows_detected"])
        self.assertTrue(result["table_bottom_double_border_applied"])
        self.assertTrue(result["table_bottom_double_border_xml_verified"])
        for col in range(5):
            self.assertTrue(is_double_black(tc_border(cell_at(tbl, 2, col), "bottom")))

    def test_merged_cell_in_footer_row_processed_once_per_row(self):
        tbl = make_gridspan_note_table("註：本表單位為新臺幣元", span=5)
        result = apply_table_footer_source_format(tbl)
        # The single merged bottom-row cell is one footer row, one cell.
        self.assertEqual(result["footer_row_count"], 1)
        self.assertEqual(result["footer_note_cells_adjusted"], 1)
        self.assertEqual(result["footer_cell_matches"], ["note"])

    def test_pipeline_log_reports_footer_rows_processed(self):
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),  # first table -> skipped
            make_paragraph("一、報表"),
            make_table(
                [
                    ["項目", "金額", "比率", "排名", "備註"],
                    ["土地", "100", "50%", "1", "ok"],
                    ["註：本表單位元", "", "", "", ""],
                    ["基期：100年", "", "資料來源：本所", "", ""],
                ]
            ),
        )
        summary, _ = run_fix(document, footer_options())
        record = summary.table_log_records[1]
        self.assertTrue(record["table_footer_note_source_format_applied"])
        self.assertEqual(record["footer_row_count"], 2)
        self.assertEqual(len(record["footer_cell_matches"]), 2)
        self.assertEqual(record["footer_note_cells_adjusted"], 1)
        self.assertEqual(record["footer_base_period_cells_adjusted"], 1)
        self.assertEqual(record["footer_source_cells_adjusted"], 1)
        self.assertEqual(record["footer_internal_top_borders_cleared"], 1)
        self.assertTrue(record["footer_block_top_border_applied"])


# ----------------------------------------------------------------------------
# The data/footer separator is a SINGLE double line above the top footer row;
# there must be no horizontal line between consecutive 註1/註2/註3 rows.
# ----------------------------------------------------------------------------
class FooterBlockSeparatorBorderTests(unittest.TestCase):
    def test_single_note_row_has_top_double_border(self):
        tbl = make_table(
            [
                ["最後推定比較租金", "", "", "", ""],  # data row
                ["註1：價格日期調整……", "", "", "", ""],  # only footer row
            ]
        )
        result = apply_table_footer_source_format(tbl)
        self.assertEqual(result["footer_row_count"], 1)
        self.assertTrue(is_double_black(tc_border(cell_at(tbl, 1, 0), "top")))
        self.assertEqual(result["footer_internal_top_borders_cleared"], 0)

    def test_two_consecutive_notes_only_first_row_has_top_border(self):
        tbl = make_table(
            [
                ["最後推定比較租金", "", "", "", ""],  # data
                ["註1：價格日期調整……", "", "", "", ""],  # top footer row
                ["註2：比較標的間試算……", "", "", "", ""],  # second footer row
            ]
        )
        result = apply_table_footer_source_format(tbl)
        self.assertEqual(result["footer_row_count"], 2)
        # 註1 row -> separator above it; 註2 row -> top cleared (no line between).
        self.assertTrue(is_double_black(tc_border(cell_at(tbl, 1, 0), "top")))
        self.assertTrue(is_nil(tc_border(cell_at(tbl, 2, 0), "top")))
        self.assertEqual(result["footer_internal_top_borders_cleared"], 1)
        # The data row above the block is untouched.
        self.assertIsNone(cell_at(tbl, 0, 0).find("w:tcPr/w:tcBorders", NS))

    def test_three_consecutive_notes_only_first_row_has_top_border(self):
        tbl = make_table(
            [
                ["資料", "", "", "", ""],
                ["註1：a", "", "", "", ""],  # top footer
                ["註2：b", "", "", "", ""],
                ["註3：c", "", "", "", ""],
            ]
        )
        result = apply_table_footer_source_format(tbl)
        self.assertEqual(result["footer_row_count"], 3)
        self.assertTrue(is_double_black(tc_border(cell_at(tbl, 1, 0), "top")))
        self.assertTrue(is_nil(tc_border(cell_at(tbl, 2, 0), "top")))
        self.assertTrue(is_nil(tc_border(cell_at(tbl, 3, 0), "top")))
        self.assertEqual(result["footer_internal_top_borders_cleared"], 2)
        self.assertEqual(result["footer_note_cells_adjusted"], 3)

    def test_base_source_note_block_only_top_row_has_separator(self):
        tbl = make_table(
            [
                ["資料", "", "", "", ""],
                ["基期：100年", "", "", "", ""],  # top footer
                ["資料來源：本所", "", "", "", ""],
                ["註1：a", "", "", "", ""],
                ["註2：b", "", "", "", ""],
            ]
        )
        result = apply_table_footer_source_format(tbl)
        # All four bottom rows form one contiguous footer block.
        self.assertEqual(result["footer_row_count"], 4)
        self.assertTrue(is_double_black(tc_border(cell_at(tbl, 1, 0), "top")))
        for row in (2, 3, 4):
            self.assertTrue(is_nil(tc_border(cell_at(tbl, row, 0), "top")), row)
        self.assertEqual(result["footer_base_period_cells_adjusted"], 1)
        self.assertEqual(result["footer_source_cells_adjusted"], 1)
        self.assertEqual(result["footer_note_cells_adjusted"], 2)

    def test_top_border_spans_all_cells_of_top_footer_row(self):
        tbl = make_table(
            [
                ["項目", "金額", "比率", "排名", "備註"],
                ["土地", "100", "50%", "1", "ok"],
                ["基期：100年", "其他", "資料來源：本所", "", ""],  # footer row
            ]
        )
        apply_table_footer_source_format(tbl)
        # The separator double line spans EVERY cell of the footer row, including
        # the unmatched ones, so the line is not broken.
        for col in range(5):
            self.assertTrue(
                is_double_black(tc_border(cell_at(tbl, 2, col), "top")), col
            )

    def test_blank_last_row_stops_scan_upward(self):
        tbl = make_table(
            [
                ["註1：a", "", "", "", ""],  # footer-like, but not at the bottom
                ["", "", "", "", ""],  # bottom row is blank -> stops the scan
            ]
        )
        result = apply_table_footer_source_format(tbl)
        self.assertEqual(result["footer_row_count"], 0)
        # The note row above the blank bottom row is NOT processed.
        self.assertEqual(cell_run_sizes(cell_at(tbl, 0, 0)), ["22"])

    def test_borders_outside_footer_block_are_preserved(self):
        tbl = make_table(
            [
                ["資料", "", "", "", ""],  # data row with a pre-existing border
                ["註1：a", "", "", "", ""],  # footer
            ]
        )
        data_cell = cell_at(tbl, 0, 0)
        tc_pr = etree.SubElement(data_cell, qn("tcPr"))
        borders = etree.SubElement(tc_pr, qn("tcBorders"))
        bottom = etree.SubElement(borders, qn("bottom"))
        bottom.set(qn("val"), "double")
        bottom.set(qn("color"), "000000")
        bottom.set(qn("sz"), "4")

        apply_table_footer_source_format(tbl)

        # The data row's own border is untouched (footer block only formats the
        # footer rows; it never rebuilds tcBorders of cells outside the block).
        self.assertTrue(is_double_black(tc_border(data_cell, "bottom")))


if __name__ == "__main__":
    unittest.main()
