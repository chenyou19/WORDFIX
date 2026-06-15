from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from docx_fixer.constants import NS, W_NS
from docx_fixer.docx_processor import fix_docx_fast
from docx_fixer.models import ProcessOptions
from docx_fixer.table_format import apply_double_black_table_borders
from docx_fixer.xml_utils import qn

BORDER_TAGS = ("top", "left", "bottom", "right", "insideH", "insideV")


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


def has_double_black_borders(tbl) -> bool:
    tbl_borders = tbl.find("./w:tblPr/w:tblBorders", NS)
    if tbl_borders is None:
        return False
    for tag in BORDER_TAGS:
        border = tbl_borders.find(f"w:{tag}", NS)
        if border is None:
            return False
        if border.get(qn("val")) != "double" or border.get(qn("color")) != "000000":
            return False
    return True


def paragraph_text_of(p) -> str:
    return "".join(p.xpath(".//w:t/text()", namespaces=NS))


class DoubleBlackBorderTests(unittest.TestCase):
    def test_apply_double_black_table_borders_sets_all_edges(self):
        tbl = uniform_table(2, 2)
        apply_double_black_table_borders(tbl)

        tbl_borders = tbl.find("./w:tblPr/w:tblBorders", NS)
        self.assertIsNotNone(tbl_borders)
        for tag in BORDER_TAGS:
            border = tbl_borders.find(f"w:{tag}", NS)
            self.assertIsNotNone(border, tag)
            self.assertEqual(border.get(qn("val")), "double")
            self.assertEqual(border.get(qn("color")), "000000")
            self.assertEqual(border.get(qn("space")), "0")
            self.assertIsNotNone(border.get(qn("sz")))

    def test_replaces_existing_borders(self):
        tbl = uniform_table(2, 2)
        tbl_pr = etree.SubElement(tbl, qn("tblPr"))
        old = etree.SubElement(tbl_pr, qn("tblBorders"))
        top = etree.SubElement(old, qn("top"))
        top.set(qn("val"), "single")
        top.set(qn("color"), "FF0000")

        apply_double_black_table_borders(tbl)

        borders = tbl.findall("./w:tblPr/w:tblBorders", NS)
        self.assertEqual(len(borders), 1)
        self.assertEqual(borders[0].find("w:top", NS).get(qn("val")), "double")
        self.assertEqual(borders[0].find("w:top", NS).get(qn("color")), "000000")

    def test_normal_and_special_tables_get_double_black_borders(self):
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),  # first table -> skipped
            make_paragraph("一、一般表格"),
            uniform_table(2, 5),  # normal (col 5)
            make_paragraph("二、特殊表格"),
            uniform_table(3, 4),  # special (col 4)
        )
        options = ProcessOptions(
            fix_table_layout=True,
            fix_color=False,
            fix_paragraph=False,
            normalize_with_word_com=False,
        )
        summary, root = run_fix(document, options)
        tables = root.xpath(".//w:tbl", namespaces=NS)

        self.assertFalse(has_double_black_borders(tables[0]))  # skipped first table
        self.assertTrue(has_double_black_borders(tables[1]))  # normal
        self.assertTrue(has_double_black_borders(tables[2]))  # special
        self.assertEqual(summary.table_log_records[1]["table_type"], "normal_table")
        self.assertTrue(summary.table_log_records[1]["double_border_applied"])
        self.assertEqual(summary.table_log_records[2]["table_type"], "special_table")
        self.assertTrue(summary.table_log_records[2]["double_border_applied"])
        self.assertEqual(summary.double_border_tables, 2)


class MoveNotesIntegrationTests(unittest.TestCase):
    def test_note_moved_below_and_table_keeps_double_borders(self):
        note_table = make_table(
            [
                ["項目", "金額", "備註", "x", "y"],
                ["土地", "100", "", "", ""],
                ["註：本表單位為新臺幣元", "", "", "", ""],
            ]
        )
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),  # first table -> skipped
            make_paragraph("一、含註記的表格"),
            note_table,
        )
        options = ProcessOptions(
            fix_table_layout=True,
            fix_color=False,
            fix_paragraph=False,
            normalize_with_word_com=False,
            move_table_notes_below=True,
        )
        summary, root = run_fix(document, options)
        tables = root.xpath(".//w:tbl", namespaces=NS)
        processed = tables[1]

        # Note row removed, table still has borders.
        self.assertEqual(len(processed.findall("w:tr", NS)), 2)
        self.assertTrue(has_double_black_borders(processed))

        # Note paragraph inserted right after the table.
        sibling = processed.getnext()
        self.assertEqual(sibling.tag, qn("p"))
        self.assertEqual(paragraph_text_of(sibling), "註：本表單位為新臺幣元")

        record = summary.table_log_records[1]
        self.assertTrue(record["move_table_notes_below_enabled"])
        self.assertTrue(record["note_cells_moved"])
        self.assertEqual(record["moved_note_count"], 1)
        self.assertEqual(record["deleted_note_rows"], 1)
        self.assertEqual(record["inserted_note_paragraphs"], 1)
        self.assertTrue(record["double_border_applied"])
        self.assertEqual(summary.note_cells_moved_tables, 1)
        self.assertEqual(summary.inserted_note_paragraphs, 1)


class SectionThreeProtectionTests(unittest.TestCase):
    def _options(self, **overrides):
        base = dict(
            fix_table_layout=True,
            fix_color=True,
            fix_paragraph=True,
            normalize_with_word_com=False,
            move_table_notes_below=True,
            skip_chapter_three_adjustments=True,
        )
        base.update(overrides)
        return ProcessOptions(**base)

    def test_body_section_three_table_is_fully_protected(self):
        note_table = make_table(
            [
                ["甲", "乙", "丙", "丁", "戊"],
                ["1", "2", "3", "4", "5"],
                ["註：受保護", "", "", "", ""],
            ]
        )
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),  # first table -> skipped
            make_paragraph("參、其他章節標題"),
            note_table,
        )
        summary, root = run_fix(document, self._options())
        tables = root.xpath(".//w:tbl", namespaces=NS)
        protected = tables[1]

        # No layout, no borders, no note move inside the 參 section.
        self.assertFalse(has_double_black_borders(protected))
        self.assertEqual(len(protected.findall("w:tr", NS)), 3)
        self.assertEqual(protected.getnext(), None)

        record = summary.table_log_records[1]
        self.assertEqual(record["table_type"], "skipped_chapter_three_table")
        self.assertTrue(record["in_section_three_protected"])
        self.assertTrue(record["skipped_by_section_three_protection"])
        self.assertEqual(
            record["section_three_detection_source"], "generic_section_three_chapter_參"
        )
        self.assertFalse(record["note_cells_moved"])
        self.assertFalse(record["double_border_applied"])
        self.assertEqual(summary.section_three_protected_tables, 1)
        # Word COM must not be asked to touch a protected table.
        self.assertEqual(summary.word_com_table_autofit_records, [])

    def test_toc_section_three_marker_does_not_trigger_protection(self):
        document = build_document(
            make_paragraph("參、目錄項目", style="TOC1"),
            make_paragraph("封面"),
            uniform_table(2, 5),  # first table -> skipped
            make_paragraph("一、一般表格"),
            uniform_table(2, 5),  # normal, not under any body 參 chapter
        )
        summary, root = run_fix(document, self._options())
        tables = root.xpath(".//w:tbl", namespaces=NS)

        # The only 參 is in the TOC, so nothing is protected and the table is processed.
        self.assertTrue(has_double_black_borders(tables[1]))
        record = summary.table_log_records[1]
        self.assertEqual(record["table_type"], "normal_table")
        self.assertFalse(record["in_section_three_protected"])
        self.assertFalse(record["skipped_by_section_three_protection"])
        self.assertEqual(summary.section_three_protected_tables, 0)

    def test_table_before_section_three_is_not_protected(self):
        document = build_document(
            make_paragraph("封面"),
            uniform_table(2, 5),  # first table -> skipped
            make_paragraph("一、一般表格"),
            uniform_table(2, 5),  # before 參 -> processed
            make_paragraph("參、受保護章節"),
            uniform_table(2, 5),  # under 參 -> protected
            make_paragraph("肆、之後章節"),
            uniform_table(2, 5),  # after 參 -> processed again
        )
        summary, root = run_fix(document, self._options())
        tables = root.xpath(".//w:tbl", namespaces=NS)

        self.assertTrue(has_double_black_borders(tables[1]))  # before 參
        self.assertFalse(has_double_black_borders(tables[2]))  # under 參
        self.assertTrue(has_double_black_borders(tables[3]))  # after 參 (肆)
        self.assertEqual(summary.table_log_records[2]["table_type"], "skipped_chapter_three_table")
        self.assertTrue(summary.table_log_records[2]["skipped_by_section_three_protection"])


if __name__ == "__main__":
    unittest.main()
