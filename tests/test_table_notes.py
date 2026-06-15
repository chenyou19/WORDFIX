from __future__ import annotations

import unittest

from lxml import etree

from docx_fixer.constants import NS, W_NS
from docx_fixer.note_detection import is_note_cell_text
from docx_fixer.table_notes import move_table_note_cells_below
from docx_fixer.xml_utils import qn


def make_cell(text: str):
    tc = etree.Element(qn("tc"))
    p = etree.SubElement(tc, qn("p"))
    r = etree.SubElement(p, qn("r"))
    t = etree.SubElement(r, qn("t"))
    t.text = text
    return tc


def make_row(*texts: str):
    tr = etree.Element(qn("tr"))
    for text in texts:
        tr.append(make_cell(text))
    return tr


def make_table_with_rows(rows: list[list[str]]):
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    tbl = etree.SubElement(body, qn("tbl"))
    for row in rows:
        tbl.append(make_row(*row))
    return document, body, tbl


def paragraphs_after_table(tbl) -> list:
    result = []
    sibling = tbl.getnext()
    while sibling is not None and sibling.tag == qn("p"):
        result.append(sibling)
        sibling = sibling.getnext()
    return result


def paragraph_text_of(p) -> str:
    return "".join(p.xpath(".//w:t/text()", namespaces=NS))


class IsNoteCellTextTests(unittest.TestCase):
    def test_matches_note_markers(self):
        for text in (
            "註：本表單位為新臺幣元",
            "註1：價格日期調整係參考地價指數",
            "註2. 總價差異率係四捨五入後計算",
            "註一：本表資料來源如下",
            "註十、其他",
        ):
            with self.subTest(text=text):
                self.assertTrue(is_note_cell_text(text))

    def test_does_not_match_non_notes(self):
        for text in (
            "本表註1如下",
            "價格調整註記",
            "註冊資料",
            "註銷登記",
            "註明事項",
            "一般內容",
            "",
        ):
            with self.subTest(text=text):
                self.assertFalse(is_note_cell_text(text))


class MoveTableNoteCellsTests(unittest.TestCase):
    def test_note_colon_cell_in_blank_row_deletes_row(self):
        document, body, tbl = make_table_with_rows(
            [
                ["項目", "金額"],
                ["土地", "100"],
                ["註：本表單位為新臺幣元", "", ""],
            ]
        )

        result = move_table_note_cells_below(tbl)

        self.assertTrue(result.note_cells_moved)
        self.assertEqual(result.moved_note_count, 1)
        self.assertEqual(result.deleted_note_rows, 1)
        self.assertEqual(result.deleted_note_cells, 0)
        self.assertEqual(result.inserted_note_paragraphs, 1)
        self.assertEqual(len(tbl.findall("w:tr", NS)), 2)

        moved = paragraphs_after_table(tbl)
        self.assertEqual(len(moved), 1)
        self.assertEqual(paragraph_text_of(moved[0]), "註：本表單位為新臺幣元")
        self.assertEqual(result.moved_notes[0].delete_action, "delete_row")

    def test_note_number_and_chinese_number_markers_move(self):
        for marker in ("註1：價格日期調整", "註一：資料來源"):
            with self.subTest(marker=marker):
                document, body, tbl = make_table_with_rows(
                    [["欄位", "值"], [marker, ""]]
                )
                result = move_table_note_cells_below(tbl)
                self.assertTrue(result.note_cells_moved)
                self.assertEqual(paragraph_text_of(paragraphs_after_table(tbl)[0]), marker)

    def test_non_note_cells_are_not_moved(self):
        document, body, tbl = make_table_with_rows(
            [
                ["註冊資料", "註銷登記"],
                ["註明事項", "本表註1如下"],
            ]
        )

        result = move_table_note_cells_below(tbl)

        self.assertFalse(result.note_cells_moved)
        self.assertEqual(result.moved_note_count, 0)
        self.assertEqual(len(tbl.findall("w:tr", NS)), 2)
        self.assertEqual(paragraphs_after_table(tbl), [])

    def test_note_cell_with_other_content_deletes_cell_only(self):
        document, body, tbl = make_table_with_rows(
            [
                ["項目", "金額"],
                ["註：含稅", "1000"],
            ]
        )

        result = move_table_note_cells_below(tbl)

        self.assertTrue(result.note_cells_moved)
        self.assertEqual(result.deleted_note_cells, 1)
        self.assertEqual(result.deleted_note_rows, 0)
        # The row stays but the note cell is gone, leaving the "1000" cell.
        rows = tbl.findall("w:tr", NS)
        self.assertEqual(len(rows), 2)
        second_row_cells = rows[1].findall("w:tc", NS)
        self.assertEqual(len(second_row_cells), 1)
        self.assertEqual(result.moved_notes[0].delete_action, "delete_cell")

    def test_multiple_notes_preserve_scan_order(self):
        document, body, tbl = make_table_with_rows(
            [
                ["a", "b"],
                ["註1：第一", ""],
                ["註2：第二", ""],
            ]
        )

        result = move_table_note_cells_below(tbl)

        moved = [paragraph_text_of(p) for p in paragraphs_after_table(tbl)]
        self.assertEqual(moved, ["註1：第一", "註2：第二"])
        self.assertEqual(result.deleted_note_rows, 2)

    def test_inserted_paragraph_is_kaiti_10pt_body_level(self):
        document, body, tbl = make_table_with_rows([["x", "y"], ["註：abc", ""]])

        move_table_note_cells_below(tbl)
        paragraph = paragraphs_after_table(tbl)[0]

        r_fonts = paragraph.find("./w:r/w:rPr/w:rFonts", NS)
        self.assertIsNotNone(r_fonts)
        self.assertEqual(r_fonts.get(qn("eastAsia")), "標楷體")
        self.assertEqual(paragraph.find("./w:r/w:rPr/w:sz", NS).get(qn("val")), "20")
        self.assertEqual(paragraph.find("./w:r/w:rPr/w:szCs", NS).get(qn("val")), "20")
        self.assertEqual(paragraph.find("./w:pPr/w:outlineLvl", NS).get(qn("val")), "9")
        self.assertIsNone(paragraph.find("./w:pPr/w:numPr", NS))
        # Paragraph-mark run properties also carry the font for safe inheritance.
        self.assertEqual(
            paragraph.find("./w:pPr/w:rPr/w:rFonts", NS).get(qn("eastAsia")),
            "標楷體",
        )


if __name__ == "__main__":
    unittest.main()
