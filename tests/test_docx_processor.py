from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from docx_fixer.constants import NS, W_NS
from docx_fixer.docx_processor import fix_docx_fast
from docx_fixer.models import ProcessOptions
from docx_fixer.xml_utils import qn


def make_docx(
    path: Path,
    document_xml: bytes,
    styles_xml: bytes | None = None,
    numbering_xml: bytes | None = None,
) -> None:
    with ZipFile(path, "w") as zf:
        zf.writestr("word/document.xml", document_xml)
        if styles_xml is not None:
            zf.writestr("word/styles.xml", styles_xml)
        if numbering_xml is not None:
            zf.writestr("word/numbering.xml", numbering_xml)


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


def read_part_root(path: Path, part_name: str):
    with ZipFile(path, "r") as zf:
        return etree.fromstring(zf.read(part_name))


def read_document_root(path: Path):
    return read_part_root(path, "word/document.xml")


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
