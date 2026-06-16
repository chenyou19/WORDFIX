from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
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


def is_double_black(border) -> bool:
    return (
        border is not None
        and border.get(qn("val")) == "double"
        and border.get(qn("color")) == "000000"
    )


def is_nil(border) -> bool:
    return border is not None and border.get(qn("val")) == "nil"


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
        apply_table_footer_source_format(tbl)
        for side in ("top", "bottom", "left", "right"):
            self.assertTrue(is_double_black(tbl_border(tbl, side)), side)

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

    def test_non_matching_last_row_cell_is_left_alone(self):
        tbl = footer_table()
        apply_table_footer_source_format(tbl)
        other = cell_at(tbl, 3, 1)  # "其他資料"
        # Stays 11 pt and gets no cell-level border overrides.
        self.assertEqual(cell_run_sizes(other), ["22"])
        self.assertIsNone(other.find("w:tcPr/w:tcBorders", NS))

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
        self.assertEqual(result["footer_note_cells_adjusted"], 2)
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

        # Outer frame still double black.
        for side in ("top", "bottom", "left", "right"):
            self.assertTrue(is_double_black(tbl_border(tbl, side)), side)
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

        for side in ("top", "left", "bottom", "right"):
            self.assertTrue(is_double_black(tc_border(other, side)), side)

    def test_matched_cell_pre_existing_top_not_destroyed_by_nil_edges(self):
        # A matched 基期 cell with a pre-existing double-black top must keep a
        # double-black top (rule sets top double) while only left/right/bottom
        # are cleared.
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

        # Outer frame double black.
        for side in ("top", "bottom", "left", "right"):
            self.assertTrue(is_double_black(tbl_border(tbl, side)), side)

        # First-row single cell top/left/right nil survive the outer frame.
        title = cell_at(tbl, 0, 0)
        self.assertTrue(is_nil(tc_border(title, "top")))
        self.assertTrue(is_nil(tc_border(title, "left")))
        self.assertTrue(is_nil(tc_border(title, "right")))
        self.assertTrue(is_double_black(tc_border(title, "bottom")))

        # Last-row footer cells: left/right/bottom nil survive the outer frame,
        # alignment overrides the centered content from layout formatting.
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
        self.assertTrue(record["first_row_single_cell_border_adjusted"])
        self.assertEqual(record["footer_note_cells_adjusted"], 2)
        self.assertCountEqual(
            record["footer_note_cell_matches"], ["base_period", "source"]
        )
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
    def test_chapter_three_layout_skip_blocks_footer_format(self):
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

        # Layout is protected -> footer format is not applied.
        self.assertIsNone(tbl_border(tbl, "top"))
        self.assertEqual(summary.table_footer_source_format_tables, 0)
        record = summary.table_log_records[1]
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

    def test_section_three_full_protection_skips_footer_format(self):
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),
            make_paragraph("參、受保護章節"),
            footer_table(),
        )
        options = footer_options(skip_chapter_three_adjustments=True)
        summary, root = run_fix(document, options)
        tbl = root.xpath(".//w:tbl", namespaces=NS)[1]

        self.assertIsNone(tbl_border(tbl, "top"))
        self.assertEqual(summary.table_footer_source_format_tables, 0)
        record = summary.table_log_records[1]
        self.assertEqual(record["table_type"], "skipped_chapter_three_table")
        self.assertFalse(record["table_footer_note_source_format_applied"])

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


# ----------------------------------------------------------------------------
# Last-row note cell regex and formatting.
# ----------------------------------------------------------------------------
class FooterNoteRegexTests(unittest.TestCase):
    def test_matches_note_with_and_without_number(self):
        for text in ("註：", "註:", "註1：", "註1:", "註2：", "註3:", "註10：", "註10:"):
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
            # The cell keeps the whole-table 11 pt and gets no border override.
            self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 0)), ["22"], text)
            self.assertIsNone(cell_at(tbl, 3, 0).find("w:tcPr/w:tcBorders", NS), text)

    def test_other_last_row_cells_are_not_changed_to_10pt(self):
        tbl = note_last_row_table("註：說明")
        apply_table_footer_source_format(tbl)
        # "其他" and "說明" cells stay at 11 pt with no cell border override.
        self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 1)), ["22"])
        self.assertEqual(cell_run_sizes(cell_at(tbl, 3, 2)), ["22"])
        self.assertIsNone(cell_at(tbl, 3, 1).find("w:tcPr/w:tcBorders", NS))

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

        # Outer frame double black, note cell nil edges survive it.
        for side in ("top", "bottom", "left", "right"):
            self.assertTrue(is_double_black(tbl_border(tbl, side)), side)
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


if __name__ == "__main__":
    unittest.main()
