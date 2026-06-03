from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from docx_fixer.constants import NS, TEMPLATE_OUTLINE_INDENTS, W_NS
from docx_fixer.docx_processor import fix_docx_fast
from docx_fixer.models import ProcessOptions
from docx_fixer.xml_utils import qn

FORBIDDEN_ATTRS = [
    "leftChars",
    "startChars",
    "rightChars",
    "endChars",
    "firstLineChars",
    "hangingChars",
]


def make_docx(
    path: Path,
    document_xml: bytes,
    styles_xml: bytes | None = None,
    numbering_xml: bytes | None = None,
    extra_parts: dict[str, bytes] | None = None,
) -> None:
    with ZipFile(path, "w") as zf:
        zf.writestr("word/document.xml", document_xml)
        if styles_xml is not None:
            zf.writestr("word/styles.xml", styles_xml)
        if numbering_xml is not None:
            zf.writestr("word/numbering.xml", numbering_xml)
        for name, data in (extra_parts or {}).items():
            zf.writestr(name, data)


def make_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))

    for text, outline, style in [
        ("\u666e\u901a\u6b63\u6587", 5, "Heading1"),
        ("\u58f9\u3001\u5e8f\u8a00", 2, None),
    ]:
        p = etree.SubElement(body, qn("p"))
        pPr = etree.SubElement(p, qn("pPr"))
        if style is not None:
            p_style = etree.SubElement(pPr, qn("pStyle"))
            p_style.set(qn("val"), style)
        outline_lvl = etree.SubElement(pPr, qn("outlineLvl"))
        outline_lvl.set(qn("val"), str(outline))
        r = etree.SubElement(p, qn("r"))
        t = etree.SubElement(r, qn("t"))
        t.text = text

    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_styles_xml() -> bytes:
    styles = etree.Element(qn("styles"), nsmap={"w": W_NS})
    style = etree.SubElement(styles, qn("style"))
    style.set(qn("type"), "paragraph")
    style.set(qn("styleId"), "Heading1")
    pPr = etree.SubElement(style, qn("pPr"))
    outline_lvl = etree.SubElement(pPr, qn("outlineLvl"))
    outline_lvl.set(qn("val"), "0")
    return etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_numbering_xml() -> bytes:
    numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    abstract = etree.SubElement(numbering, qn("abstractNum"))
    abstract.set(qn("abstractNumId"), "1")
    lvl = etree.SubElement(abstract, qn("lvl"))
    lvl.set(qn("ilvl"), "0")
    pPr = etree.SubElement(lvl, qn("pPr"))
    outline_lvl = etree.SubElement(pPr, qn("outlineLvl"))
    outline_lvl.set(qn("val"), "0")
    return etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_ind(parent, **attrs):
    ind = etree.SubElement(parent, qn("ind"))
    for name, value in attrs.items():
        ind.set(qn(name), value)
    return ind


def make_document_with_character_indent(text: str = "\u666e\u901a\u6b63\u6587", font_size_pt: int | None = None) -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    p = etree.SubElement(body, qn("p"))
    pPr = etree.SubElement(p, qn("pPr"))
    make_ind(
        pPr,
        left="1440",
        leftChars="200",
        hangingChars="100",
        startChars="50",
        rightChars="20",
        endChars="30",
        firstLineChars="40",
    )
    r = etree.SubElement(p, qn("r"))
    if font_size_pt is not None:
        rPr = etree.SubElement(r, qn("rPr"))
        sz = etree.SubElement(rPr, qn("sz"))
        sz.set(qn("val"), str(font_size_pt * 2))
    t = etree.SubElement(r, qn("t"))
    t.text = text
    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_styles_with_character_indent() -> bytes:
    styles = etree.Element(qn("styles"), nsmap={"w": W_NS})
    style = etree.SubElement(styles, qn("style"))
    style.set(qn("type"), "paragraph")
    style.set(qn("styleId"), "BodyText")
    pPr = etree.SubElement(style, qn("pPr"))
    make_ind(pPr, left="720", leftChars="200", firstLineChars="100")
    return etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_styles_with_default_text_font(font_size_pt: float = 14) -> bytes:
    styles = etree.Element(qn("styles"), nsmap={"w": W_NS})
    style = etree.SubElement(styles, qn("style"))
    style.set(qn("type"), "paragraph")
    style.set(qn("styleId"), "DefaultText")
    rPr = etree.SubElement(style, qn("rPr"))
    sz = etree.SubElement(rPr, qn("sz"))
    sz.set(qn("val"), str(round(font_size_pt * 2)))
    return etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_document_with_styled_level_four_body() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))

    for text, style in [
        ("\u58f9\u3001\u5e8f\u8a00", None),
        ("1. \u7b2c\u56db\u968e\u6a19\u984c", None),
        ("\u4f7f\u7528 DefaultText 14 pt \u7684\u5167\u6587", "DefaultText"),
    ]:
        p = etree.SubElement(body, qn("p"))
        pPr = etree.SubElement(p, qn("pPr"))
        if style is not None:
            p_style = etree.SubElement(pPr, qn("pStyle"))
            p_style.set(qn("val"), style)
        if style == "DefaultText":
            tabs = etree.SubElement(pPr, qn("tabs"))
            tab = etree.SubElement(tabs, qn("tab"))
            tab.set(qn("val"), "left")
            tab.set(qn("pos"), "1990")
        r = etree.SubElement(p, qn("r"))
        t = etree.SubElement(r, qn("t"))
        t.text = text

    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_unrecognized_numbering_with_character_indent() -> bytes:
    numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    abstract = etree.SubElement(numbering, qn("abstractNum"))
    abstract.set(qn("abstractNumId"), "1")
    lvl = etree.SubElement(abstract, qn("lvl"))
    lvl.set(qn("ilvl"), "12")
    num_fmt = etree.SubElement(lvl, qn("numFmt"))
    num_fmt.set(qn("val"), "customFormat")
    lvl_text = etree.SubElement(lvl, qn("lvlText"))
    lvl_text.set(qn("val"), "custom")
    pPr = etree.SubElement(lvl, qn("pPr"))
    make_ind(pPr, left="360", leftChars="88", hangingChars="44")
    return etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)


def read_part_root(path: Path, part_name: str):
    with ZipFile(path, "r") as zf:
        return etree.fromstring(zf.read(part_name))


def read_document_root(path: Path):
    return read_part_root(path, "word/document.xml")


def assert_no_character_indent_attrs(test_case: unittest.TestCase, root) -> None:
    for ind in root.xpath(".//w:ind", namespaces=NS):
        for attr in FORBIDDEN_ATTRS:
            test_case.assertIsNone(ind.get(qn(attr)), attr)


def assert_docx_has_no_character_indent_attrs(test_case: unittest.TestCase, path: Path) -> None:
    with ZipFile(path, "r") as zf:
        for name in zf.namelist():
            if not name.endswith(".xml"):
                continue
            root = etree.fromstring(zf.read(name))
            for ind in root.xpath(".//w:ind", namespaces=NS):
                for attr in FORBIDDEN_ATTRS:
                    test_case.assertIsNone(ind.get(qn(attr)), f"{name}: {attr}")


def part_outline_count(path: Path, part_name: str) -> int:
    root = read_part_root(path, part_name)
    return len(root.xpath(".//w:outlineLvl", namespaces=NS))


def paragraph_outlines(path: Path) -> list[str | None]:
    root = read_document_root(path)
    values: list[str | None] = []
    for p in root.xpath(".//w:p", namespaces=NS):
        outline = p.find("./w:pPr/w:outlineLvl", NS)
        values.append(None if outline is None else outline.get(qn("val")))
    return values


def assert_all_document_outlines_are_body(test_case: unittest.TestCase, path: Path) -> None:
    root = read_document_root(path)
    paragraphs = root.xpath(".//w:p", namespaces=NS)
    test_case.assertGreater(len(paragraphs), 0)
    for p in paragraphs:
        outline = p.find("./w:pPr/w:outlineLvl", NS)
        test_case.assertIsNotNone(outline)
        test_case.assertEqual(outline.get(qn("val")), "9")


class DocxProcessorTests(unittest.TestCase):
    def test_body_indent_uses_styles_xml_font_size_lookup(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_document_with_styled_level_four_body(),
                styles_xml=make_styles_with_default_text_font(14),
            )

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=True,
                    include_tables_in_paragraph=False,
                ),
            )

            root = read_document_root(output_docx)
            body_paragraph = root.xpath(".//w:p", namespaces=NS)[2]
            ind = body_paragraph.find("./w:pPr/w:ind", NS)
            self.assertEqual(ind.get(qn("left")), TEMPLATE_OUTLINE_INDENTS[3]["body_left"])
            self.assertIsNone(ind.get(qn("hanging")))
            self.assertIsNone(body_paragraph.find("./w:pPr/w:tabs", NS))
            debug = "\n".join(summary.body_indent_debug_logs)
            self.assertIn("font_size_source=paragraph_style:DefaultText", debug)

    def test_document_xml_character_indent_attrs_are_removed_but_twips_remain(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(input_docx, make_document_with_character_indent())

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=False,
                    include_tables_in_paragraph=False,
                ),
            )

            root = read_document_root(output_docx)
            ind = root.find(".//w:ind", NS)
            self.assertEqual(ind.get(qn("left")), "1440")
            self.assertIsNone(ind.get(qn("leftChars")))
            self.assertIsNone(ind.get(qn("hangingChars")))
            assert_no_character_indent_attrs(self, root)
            self.assertEqual(summary.character_indent_attrs_removed, 6)

    def test_non_14_pt_body_keeps_left_indent_but_character_attrs_are_removed(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(input_docx, make_document_with_character_indent(font_size_pt=12))

            fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=True,
                    include_tables_in_paragraph=False,
                ),
            )

            root = read_document_root(output_docx)
            ind = root.find(".//w:ind", NS)
            self.assertEqual(ind.get(qn("left")), "1440")
            self.assertIsNone(ind.get(qn("leftChars")))
            self.assertIsNone(ind.get(qn("hangingChars")))
            self.assertIsNone(ind.get(qn("hanging")))

    def test_styles_xml_character_indent_attrs_are_removed(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_document_xml(),
                styles_xml=make_styles_with_character_indent(),
            )

            fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=False,
                    include_tables_in_paragraph=False,
                ),
            )

            root = read_part_root(output_docx, "word/styles.xml")
            ind = root.find(".//w:ind", NS)
            self.assertEqual(ind.get(qn("left")), "720")
            assert_no_character_indent_attrs(self, root)

    def test_unrecognized_numbering_level_character_indent_attrs_are_removed(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_document_xml(),
                numbering_xml=make_unrecognized_numbering_with_character_indent(),
            )

            fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=True,
                    include_tables_in_paragraph=False,
                ),
            )

            root = read_part_root(output_docx, "word/numbering.xml")
            lvl = root.find(".//w:lvl", NS)
            ind = lvl.find("./w:pPr/w:ind", NS)
            self.assertEqual(ind.get(qn("left")), "360")
            self.assertIsNone(ind.get(qn("hanging")))
            self.assertIsNone(lvl.find("./w:suff", NS))
            assert_no_character_indent_attrs(self, root)

    def test_output_docx_has_no_character_indent_attrs_in_any_xml_part(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_document_with_character_indent(),
                styles_xml=make_styles_with_character_indent(),
                numbering_xml=make_unrecognized_numbering_with_character_indent(),
                extra_parts={
                    "word/header1.xml": make_document_with_character_indent("\u9801\u9996"),
                    "word/footer1.xml": make_document_with_character_indent("\u9801\u5c3e"),
                    "word/footnotes.xml": make_document_with_character_indent("\u8173\u8a3b"),
                    "word/endnotes.xml": make_document_with_character_indent("\u5c3e\u8a3b"),
                },
            )

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=False,
                    include_tables_in_paragraph=False,
                ),
            )

            assert_docx_has_no_character_indent_attrs(self, output_docx)
            self.assertGreater(summary.character_indent_attrs_removed, 0)

    def test_remove_all_outline_only_clears_existing_outline_levels_in_all_parts(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_document_xml(),
                styles_xml=make_styles_xml(),
                numbering_xml=make_numbering_xml(),
            )

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=False,
                    include_tables_in_paragraph=False,
                    remove_all_outline_levels=True,
                ),
            )

            self.assertEqual(paragraph_outlines(output_docx), ["9", "9"])
            assert_all_document_outlines_are_body(self, output_docx)
            self.assertEqual(part_outline_count(output_docx, "word/styles.xml"), 0)
            self.assertEqual(part_outline_count(output_docx, "word/numbering.xml"), 0)
            self.assertEqual(summary.removed_all_outline_paragraphs, 4)
            self.assertEqual(summary.paragraphs, 0)

    def test_remove_all_outline_runs_before_paragraph_fixing_reapplies_numbered_outline(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(input_docx, make_document_xml())

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=True,
                    include_tables_in_paragraph=False,
                    remove_all_outline_levels=True,
                ),
            )

            self.assertEqual(paragraph_outlines(output_docx), ["9", "0"])
            self.assertEqual(summary.removed_all_outline_paragraphs, 2)
            self.assertEqual(summary.paragraph_level_counts[0], 1)


if __name__ == "__main__":
    unittest.main()
