from __future__ import annotations

import unittest

from lxml import etree

from docx_fixer.constants import (
    NS,
    PREFACE_OUTLINE_INDENTS,
    TEMPLATE_OUTLINE_INDENTS,
    W_NS,
    validate_template_outline_indents,
)
from docx_fixer.models import ProcessSummary
from docx_fixer.numbering import apply_numbering_outline_format
from docx_fixer.outline import (
    detect_outline_level,
    fix_outline_paragraphs,
    remove_all_outline_levels_from_root,
)
from docx_fixer.xml_utils import qn


def make_root(*children):
    root = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(root, qn("body"))
    for child in children:
        body.append(child)
    return root


def make_paragraph(
    text: str,
    *,
    style: str | None = None,
    outline: int | None = None,
    num_id: str | None = None,
    ilvl: int | None = None,
):
    p = etree.Element(qn("p"))
    pPr = etree.SubElement(p, qn("pPr"))

    if style is not None:
        p_style = etree.SubElement(pPr, qn("pStyle"))
        p_style.set(qn("val"), style)

    if outline is not None:
        outline_lvl = etree.SubElement(pPr, qn("outlineLvl"))
        outline_lvl.set(qn("val"), str(outline))

    if num_id is not None:
        num_pr = etree.SubElement(pPr, qn("numPr"))
        if ilvl is not None:
            ilvl_el = etree.SubElement(num_pr, qn("ilvl"))
            ilvl_el.set(qn("val"), str(ilvl))
        num_id_el = etree.SubElement(num_pr, qn("numId"))
        num_id_el.set(qn("val"), num_id)

    r = etree.SubElement(p, qn("r"))
    t = etree.SubElement(r, qn("t"))
    t.text = text
    return p


def make_table_paragraph(text: str):
    tbl = etree.Element(qn("tbl"))
    tr = etree.SubElement(tbl, qn("tr"))
    tc = etree.SubElement(tr, qn("tc"))
    p = make_paragraph(text)
    tc.append(p)
    return tbl, p


def paragraph_indent(p):
    ind = p.find("./w:pPr/w:ind", NS)
    if ind is None:
        return None
    return ind.get(qn("left")), ind.get(qn("hanging"))


def paragraph_left_indent(p):
    ind = p.find("./w:pPr/w:ind", NS)
    if ind is None:
        return None
    return ind.get(qn("left"))


def paragraph_outline(p):
    outline = p.find("./w:pPr/w:outlineLvl", NS)
    if outline is None:
        return None
    return outline.get(qn("val"))


def expected_indent(level: int):
    spec = TEMPLATE_OUTLINE_INDENTS[level]
    return spec["left"], spec["hanging"]


def expected_preface_indent(level: int):
    spec = PREFACE_OUTLINE_INDENTS[level]
    return spec["left"], spec["hanging"]


def make_numbering_xml():
    root = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    abstract = etree.SubElement(root, qn("abstractNum"))
    abstract.set(qn("abstractNumId"), "1")

    decimal_lvl = etree.SubElement(abstract, qn("lvl"))
    decimal_lvl.set(qn("ilvl"), "3")
    num_fmt = etree.SubElement(decimal_lvl, qn("numFmt"))
    num_fmt.set(qn("val"), "decimal")
    lvl_text = etree.SubElement(decimal_lvl, qn("lvlText"))
    lvl_text.set(qn("val"), "%1.")

    bullet_lvl = etree.SubElement(abstract, qn("lvl"))
    bullet_lvl.set(qn("ilvl"), "4")
    bullet_fmt = etree.SubElement(bullet_lvl, qn("numFmt"))
    bullet_fmt.set(qn("val"), "bullet")
    bullet_pPr = etree.SubElement(bullet_lvl, qn("pPr"))
    bullet_outline = etree.SubElement(bullet_pPr, qn("outlineLvl"))
    bullet_outline.set(qn("val"), "2")

    return etree.tostring(root)


class OutlineFixTests(unittest.TestCase):
    def test_template_indents_match_requested_number_start(self):
        self.assertTrue(validate_template_outline_indents())

    def test_manual_numbering_gets_outline_and_indent(self):
        paragraphs = [
            make_paragraph("\u4e09\u3001\u52d8\u4f30\u6a19\u7684\u57fa\u672c\u8cc7\u6599\uff1a"),
            make_paragraph("\uff08\u4e00\uff09\u66f4\u65b0\u524d\u57fa\u672c\u8cc7\u6599\uff1a"),
            make_paragraph("1. \u52d8\u4f30\u6a19\u7684\u5167\u5bb9\uff1a"),
            make_paragraph("\uff081\uff09\u571f\u5730\u6a19\u793a\uff1a"),
            make_paragraph("\u58f9\u3001\u5e8f\u8a00"),
            make_paragraph("\u4e00\u3001\u5e8f\u8a00\u5167\u5c64"),
            make_paragraph("\uff08\u4e00\uff09\u5e8f\u8a00\u5167\u5c64"),
            make_paragraph("1. \u5e8f\u8a00\u5167\u5c64"),
            make_paragraph("\uff081\uff09\u5e8f\u8a00\u5167\u5c64"),
        ]
        root = make_root(*paragraphs)

        fix_outline_paragraphs(root, include_tables=True)

        for paragraph, level in zip(paragraphs, [1, 2, 3, 4, 0, 1, 2, 3, 4]):
            self.assertEqual(paragraph_indent(paragraph), expected_indent(level))
            self.assertEqual(paragraph_outline(paragraph), str(level))

    def test_detects_all_manual_outline_numbering_shapes(self):
        samples = [
            ("\u58f9\u3001\u5e8f\u8a00", 0),
            ("\u4e00\u3001\u6a19\u984c", 1),
            ("\uff08\u4e00\uff09\u6a19\u984c", 2),
            ("(\u4e00)\u6a19\u984c", 2),
            ("1. \u6a19\u984c", 3),
            ("\uff081\uff09\u6a19\u984c", 4),
            ("(1)\u6a19\u984c", 4),
            ("A. \u6a19\u984c", 5),
            ("\uff08A\uff09\u6a19\u984c", 6),
            ("(A)\u6a19\u984c", 6),
            ("a. \u6a19\u984c", 7),
            ("\uff08a\uff09\u6a19\u984c", 8),
            ("(a)\u6a19\u984c", 8),
        ]

        for text, level in samples:
            with self.subTest(text=text):
                self.assertEqual(detect_outline_level(text), level)

    def test_toc_paragraph_does_not_start_outline_processing(self):
        toc = make_paragraph("\u58f9\u3001\u5e8f\u8a00", style="TOC1")
        preface = make_paragraph("\u4e00\u3001\u76ee\u9304\u5f8c\u7684\u524d\u7f6e\u9805")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(toc, preface, marker)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertIsNone(paragraph_indent(toc))
        self.assertIsNone(paragraph_outline(toc))
        self.assertEqual(paragraph_indent(preface), expected_indent(1))
        self.assertEqual(paragraph_outline(preface), "1")
        self.assertEqual(paragraph_outline(marker), "0")

    def test_plain_toc_range_is_skipped_until_body_start(self):
        toc_heading = make_paragraph("\u76ee\u9304")
        toc_entry = make_paragraph("\u4e00\u3001\u5e8f\u8a00")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(toc_heading, toc_entry, marker)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertIsNone(paragraph_indent(toc_heading))
        self.assertIsNone(paragraph_outline(toc_heading))
        self.assertIsNone(paragraph_indent(toc_entry))
        self.assertIsNone(paragraph_outline(toc_entry))
        self.assertEqual(paragraph_indent(marker), expected_indent(0))
        self.assertEqual(paragraph_outline(marker), "0")

    def test_style_numbering_skip_keeps_notes_and_long_body_out_of_outline(self):
        note = make_paragraph("\u203b \u9019\u662f\u8a3b\u8a18", style="NumberedStyle", outline=2)
        long_body = make_paragraph(
            "\u9019\u662f\u4e00\u6bb5\u5f88\u9577\u7684\u6b63\u6587\u5167\u5bb9"
            "\u4e0d\u61c9\u8a72\u56e0\u70ba\u6a23\u5f0f\u5e36\u7de8\u865f\u88ab\u7576\u6210\u7ae0\u7bc0"
            "\u4e5f\u4e0d\u61c9\u8a72\u51fa\u73fe\u5728\u5c0e\u89bd\u7a97\u683c\u6216\u76ee\u9304\u88e1",
            style="NumberedStyle",
            outline=3,
        )
        root = make_root(note, long_body)

        fix_outline_paragraphs(
            root,
            include_tables=True,
            numbering_level_lookup={("1", 1): 1},
            style_numbering_lookup={"NumberedStyle": ("1", 1)},
        )

        self.assertEqual(paragraph_outline(note), "2")
        self.assertEqual(paragraph_outline(long_body), "3")
        self.assertIsNone(paragraph_indent(note))
        self.assertIsNone(paragraph_indent(long_body))

    def test_auto_numbering_uses_same_preface_shift_and_outline_rules(self):
        auto_before = make_paragraph("\u81ea\u52d5\u7de8\u865f\u524d", num_id="1", ilvl=1)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        auto_after = make_paragraph("\u81ea\u52d5\u7de8\u865f\u5f8c", num_id="1", ilvl=1)
        root = make_root(auto_before, marker, auto_after)

        fix_outline_paragraphs(
            root,
            include_tables=True,
            numbering_level_lookup={("1", 1): 1},
        )

        self.assertEqual(paragraph_indent(auto_before), expected_indent(1))
        self.assertEqual(paragraph_outline(auto_before), "1")
        self.assertEqual(paragraph_outline(marker), "0")
        self.assertEqual(paragraph_indent(auto_after), expected_indent(1))
        self.assertEqual(paragraph_outline(auto_after), "1")

    def test_remove_preface_outline_shifts_indent_and_starts_outline_at_marker(self):
        before_one = make_paragraph("\u4e00\u3001\u76ee\u9304\u5f8c\u7684\u524d\u7f6e\u9805", outline=1)
        before_nested = make_paragraph("\uff08\u4e00\uff09\u524d\u7f6e\u5167\u5c64", outline=2)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        after = make_paragraph("\u4e00\u3001\u5e8f\u8a00\u5167\u5c64")
        root = make_root(before_one, before_nested, marker, after)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            remove_preface_outline=True,
        )

        self.assertEqual(paragraph_indent(before_one), expected_preface_indent(0))
        self.assertIsNone(paragraph_outline(before_one))
        self.assertEqual(paragraph_indent(before_nested), expected_preface_indent(1))
        self.assertIsNone(paragraph_outline(before_nested))
        self.assertEqual(paragraph_outline(marker), "0")
        self.assertEqual(paragraph_outline(after), "1")
        self.assertEqual(summary.removed_preface_outline_paragraphs, 2)

    def test_remove_preface_outline_uses_preface_number_start_indents(self):
        before_decimal = make_paragraph("1. \u524d\u7f6e\u7b2c\u4e09\u968e", outline=3)
        before_parenthesized_decimal = make_paragraph("\uff081\uff09\u524d\u7f6e\u7b2c\u56db\u968e", outline=4)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(before_decimal, before_parenthesized_decimal, marker)

        fix_outline_paragraphs(
            root,
            include_tables=True,
            remove_preface_outline=True,
        )

        self.assertEqual(paragraph_indent(before_decimal), expected_preface_indent(2))
        self.assertIsNone(paragraph_outline(before_decimal))
        self.assertEqual(paragraph_indent(before_parenthesized_decimal), expected_preface_indent(3))
        self.assertIsNone(paragraph_outline(before_parenthesized_decimal))

    def test_body_paragraph_after_heading_aligns_to_heading_text_start(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("\u4e00\u3001\u5e8f\u8a00\u5167\u5c64")
        body = make_paragraph("\u9019\u662f\u6a19\u984c\u4e0b\u65b9\u5167\u6587")
        nested_heading = make_paragraph("\uff08\u4e00\uff09\u5167\u5c64\u6a19\u984c")
        nested_body = make_paragraph("\u9019\u662f\u5167\u5c64\u6a19\u984c\u4e0b\u65b9\u5167\u6587")
        root = make_root(marker, heading, body, nested_heading, nested_body)
        summary = ProcessSummary()

        fix_outline_paragraphs(root, include_tables=True, summary=summary)

        self.assertEqual(paragraph_left_indent(body), expected_indent(1)[0])
        self.assertIsNone(paragraph_indent(body)[1])
        self.assertEqual(paragraph_left_indent(nested_body), expected_indent(2)[0])
        self.assertIsNone(paragraph_indent(nested_body)[1])
        self.assertIsNone(paragraph_outline(body))
        self.assertIsNone(paragraph_outline(nested_body))
        self.assertTrue(
            any(record["prefix"] == "\u4e00\u3001" for record in summary.numbering_measurements.values())
        )
        self.assertTrue(
            all(float(record["number_size_cm"]) > 0 for record in summary.numbering_measurements.values())
        )

    def test_preface_body_paragraph_aligns_to_preface_heading_text_start_and_loses_outline(self):
        heading = make_paragraph("\u4e00\u3001\u524d\u7f6e\u9805")
        body = make_paragraph("\u9019\u662f\u5e8f\u8a00\u524d\u7684\u5167\u6587", outline=2)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(heading, body, marker)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            remove_preface_outline=True,
        )

        self.assertEqual(paragraph_left_indent(body), expected_preface_indent(0)[0])
        self.assertIsNone(paragraph_indent(body)[1])
        self.assertIsNone(paragraph_outline(body))
        self.assertEqual(summary.removed_preface_outline_paragraphs, 1)

    def test_remove_preface_outline_option_can_run_without_paragraph_fixing(self):
        before = make_paragraph("\u666e\u901a\u524d\u7f6e\u6bb5\u843d", outline=3)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00", outline=4)
        after = make_paragraph("\u4e00\u3001\u5e8f\u8a00\u5167\u5c64")
        root = make_root(before, marker, after)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            remove_preface_outline=True,
            fix_numbered_paragraphs=False,
        )

        self.assertIsNone(paragraph_outline(before))
        self.assertIsNone(paragraph_indent(before))
        self.assertEqual(paragraph_outline(marker), "4")
        self.assertIsNone(paragraph_indent(after))
        self.assertEqual(summary.removed_preface_outline_paragraphs, 1)

    def test_remove_preface_outline_also_cleans_toc_before_marker(self):
        toc = make_paragraph("\u58f9\u3001\u5e8f\u8a00", style="TOC1", outline=0)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(toc, marker)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            remove_preface_outline=True,
            fix_numbered_paragraphs=False,
        )

        self.assertIsNone(paragraph_outline(toc))
        self.assertIsNone(paragraph_outline(marker))
        self.assertEqual(summary.skipped_toc_paragraphs, 1)
        self.assertEqual(summary.removed_preface_outline_paragraphs, 1)

    def test_table_paragraph_is_skipped_even_when_include_tables_is_true(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        table, table_p = make_table_paragraph("\uff08\u4e00\uff09\u8868\u683c\u5167\u9805\u76ee")
        root = make_root(marker, table)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertIsNone(paragraph_indent(table_p))
        self.assertIsNone(paragraph_outline(table_p))

    def test_existing_outline_on_non_numbered_body_is_left_unchanged(self):
        body = make_paragraph("\u666e\u901a\u6b63\u6587", outline=2)
        root = make_root(body)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertEqual(paragraph_outline(body), "2")

    def test_remove_all_outline_levels_from_root_removes_every_existing_outline(self):
        first = make_paragraph("\u666e\u901a\u6b63\u6587", outline=2)
        second = make_paragraph("\u58f9\u3001\u5e8f\u8a00", outline=0)
        third = make_paragraph("\u7121\u5927\u7db1")
        root = make_root(first, second, third)
        summary = ProcessSummary()

        removed = remove_all_outline_levels_from_root(root, summary=summary)

        self.assertEqual(removed, 2)
        self.assertEqual(summary.removed_all_outline_paragraphs, 2)
        self.assertIsNone(paragraph_outline(first))
        self.assertIsNone(paragraph_outline(second))
        self.assertIsNone(paragraph_outline(third))

    def test_paragraph_summary_counts_levels_and_skips(self):
        toc = make_paragraph("\u76ee\u9304")
        toc_entry = make_paragraph("\u4e00\u3001\u5e8f\u8a00")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        normal = make_paragraph("\u666e\u901a\u6b63\u6587")
        table, _ = make_table_paragraph("1. \u8868\u683c\u5167\u9805\u76ee")
        root = make_root(toc, toc_entry, marker, normal, table)
        summary = ProcessSummary()

        fix_outline_paragraphs(root, include_tables=True, summary=summary)

        self.assertEqual(summary.total_paragraphs, 5)
        self.assertEqual(summary.skipped_toc_paragraphs, 2)
        self.assertEqual(summary.skipped_table_paragraphs, 1)
        self.assertEqual(summary.paragraph_level_counts[0], 1)
        self.assertEqual(summary.unknown_paragraphs, 0)

    def test_numbering_xml_uses_template_indents_and_keeps_bullets_out_of_outline(self):
        updated = apply_numbering_outline_format(make_numbering_xml())
        root = etree.fromstring(updated)

        decimal_lvl = root.xpath("./w:abstractNum/w:lvl[@w:ilvl='3']", namespaces=NS)[0]
        decimal_ind = decimal_lvl.find("./w:pPr/w:ind", NS)
        decimal_tab = decimal_lvl.find("./w:pPr/w:tabs/w:tab", NS)
        decimal_suff = decimal_lvl.find("./w:suff", NS)
        self.assertEqual(
            (decimal_ind.get(qn("left")), decimal_ind.get(qn("hanging"))),
            expected_indent(3),
        )
        self.assertEqual(decimal_tab.get(qn("pos")), TEMPLATE_OUTLINE_INDENTS[3]["left"])
        self.assertEqual(decimal_suff.get(qn("val")), "nothing")

        bullet_lvl = root.xpath("./w:abstractNum/w:lvl[@w:ilvl='4']", namespaces=NS)[0]
        bullet_outline = bullet_lvl.find("./w:pPr/w:outlineLvl", NS)
        self.assertEqual(bullet_outline.get(qn("val")), "9")


if __name__ == "__main__":
    unittest.main()
