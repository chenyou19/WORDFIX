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


def make_docx(path: Path, document_xml: bytes) -> None:
    with ZipFile(path, "w") as zf:
        zf.writestr("word/document.xml", document_xml)


def make_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))

    for text, outline in [
        ("\u666e\u901a\u6b63\u6587", 5),
        ("\u58f9\u3001\u5e8f\u8a00", 2),
    ]:
        p = etree.SubElement(body, qn("p"))
        pPr = etree.SubElement(p, qn("pPr"))
        outline_lvl = etree.SubElement(pPr, qn("outlineLvl"))
        outline_lvl.set(qn("val"), str(outline))
        r = etree.SubElement(p, qn("r"))
        t = etree.SubElement(r, qn("t"))
        t.text = text

    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def read_document_root(path: Path):
    with ZipFile(path, "r") as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def paragraph_outlines(path: Path) -> list[str | None]:
    root = read_document_root(path)
    values: list[str | None] = []
    for p in root.xpath(".//w:p", namespaces=NS):
        outline = p.find("./w:pPr/w:outlineLvl", NS)
        values.append(None if outline is None else outline.get(qn("val")))
    return values


class DocxProcessorTests(unittest.TestCase):
    def test_remove_all_outline_only_clears_existing_document_outline_levels(self):
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
                    fix_paragraph=False,
                    include_tables_in_paragraph=False,
                    remove_all_outline_levels=True,
                ),
            )

            self.assertEqual(paragraph_outlines(output_docx), [None, None])
            self.assertEqual(summary.removed_all_outline_paragraphs, 2)
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

            self.assertEqual(paragraph_outlines(output_docx), [None, "0"])
            self.assertEqual(summary.removed_all_outline_paragraphs, 2)
            self.assertEqual(summary.paragraph_level_counts[0], 1)


if __name__ == "__main__":
    unittest.main()
