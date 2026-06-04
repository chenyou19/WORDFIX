from __future__ import annotations

import json
import tempfile
import sys
import types
import unittest
from pathlib import Path
from unittest.mock import Mock, patch
from zipfile import ZipFile

from lxml import etree

from docx_fixer.constants import NS, TEMPLATE_OUTLINE_INDENTS, W_NS
from docx_fixer.docx_processor import (
    _verify_and_fix_body_indents_with_word_com_in_process,
    find_word_paragraph_index_for_record,
    fix_docx_fast,
    verify_and_fix_body_indents_with_word_com,
)
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


def make_table(row_columns: list[int]) -> etree._Element:
    tbl = etree.Element(qn("tbl"))
    for columns in row_columns:
        tr = etree.SubElement(tbl, qn("tr"))
        for _ in range(columns):
            tc = etree.SubElement(tr, qn("tc"))
            p = etree.SubElement(tc, qn("p"))
            r = etree.SubElement(p, qn("r"))
            t = etree.SubElement(r, qn("t"))
            t.text = "cell"
    return tbl


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


def make_styles_with_default_text_font(font_size_pt: float = 14, start_twips: str | None = None) -> bytes:
    styles = etree.Element(qn("styles"), nsmap={"w": W_NS})
    style = etree.SubElement(styles, qn("style"))
    style.set(qn("type"), "paragraph")
    style.set(qn("styleId"), "DefaultText")
    if start_twips is not None:
        pPr = etree.SubElement(style, qn("pPr"))
        ind = etree.SubElement(pPr, qn("ind"))
        ind.set(qn("start"), start_twips)
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


def make_document_with_styled_level_two_body() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))

    for text, style in [
        ("\u58f9\u3001\u5e8f\u8a00", None),
        ("\u4e00\u3001\u7b2c\u4e8c\u968e\u6a19\u984c", None),
        ("\u4f7f\u7528 DefaultText 14 pt \u7684\u7b2c\u4e8c\u968e\u5167\u6587", "DefaultText"),
    ]:
        p = etree.SubElement(body, qn("p"))
        pPr = etree.SubElement(p, qn("pPr"))
        if style is not None:
            p_style = etree.SubElement(pPr, qn("pStyle"))
            p_style.set(qn("val"), style)
            tabs = etree.SubElement(pPr, qn("tabs"))
            tab = etree.SubElement(tabs, qn("tab"))
            tab.set(qn("val"), "left")
            tab.set(qn("pos"), "1480")
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


def make_level_four_numbering_with_old_indents() -> bytes:
    numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    abstract = etree.SubElement(numbering, qn("abstractNum"))
    abstract.set(qn("abstractNumId"), "1")
    lvl = etree.SubElement(abstract, qn("lvl"))
    lvl.set(qn("ilvl"), "3")
    lvl_jc = etree.SubElement(lvl, qn("lvlJc"))
    lvl_jc.set(qn("val"), "right")
    suff = etree.SubElement(lvl, qn("suff"))
    suff.set(qn("val"), "nothing")
    num_fmt = etree.SubElement(lvl, qn("numFmt"))
    num_fmt.set(qn("val"), "decimal")
    lvl_text = etree.SubElement(lvl, qn("lvlText"))
    lvl_text.set(qn("val"), "%1.")
    pPr = etree.SubElement(lvl, qn("pPr"))
    tabs = etree.SubElement(pPr, qn("tabs"))
    tab = etree.SubElement(tabs, qn("tab"))
    tab.set(qn("val"), "left")
    tab.set(qn("pos"), "1990")
    make_ind(
        pPr,
        left="1990",
        start="1990",
        hanging="420",
        leftChars="20",
        startChars="20",
        hangingChars="20",
        firstLineChars="20",
    )
    num = etree.SubElement(numbering, qn("num"))
    num.set(qn("numId"), "1")
    abstract_ref = etree.SubElement(num, qn("abstractNumId"))
    abstract_ref.set(qn("val"), "1")
    return etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_styles_with_level_four_numbered_and_plain_old_indents() -> bytes:
    styles = etree.Element(qn("styles"), nsmap={"w": W_NS})

    numbered = etree.SubElement(styles, qn("style"))
    numbered.set(qn("type"), "paragraph")
    numbered.set(qn("styleId"), "NumberedL4")
    numbered_pPr = etree.SubElement(numbered, qn("pPr"))
    num_pr = etree.SubElement(numbered_pPr, qn("numPr"))
    ilvl = etree.SubElement(num_pr, qn("ilvl"))
    ilvl.set(qn("val"), "3")
    num_id = etree.SubElement(num_pr, qn("numId"))
    num_id.set(qn("val"), "1")
    make_ind(
        numbered_pPr,
        left="1990",
        start="1990",
        hanging="420",
        leftChars="20",
        startChars="20",
        hangingChars="20",
        firstLineChars="20",
    )

    plain = etree.SubElement(styles, qn("style"))
    plain.set(qn("type"), "paragraph")
    plain.set(qn("styleId"), "BodyText")
    plain_pPr = etree.SubElement(plain, qn("pPr"))
    make_ind(
        plain_pPr,
        left="720",
        start="720",
        hanging="360",
        firstLine="240",
        leftChars="20",
        startChars="20",
        hangingChars="20",
        firstLineChars="20",
    )
    tabs = etree.SubElement(plain_pPr, qn("tabs"))
    tab = etree.SubElement(tabs, qn("tab"))
    tab.set(qn("val"), "left")
    tab.set(qn("pos"), "720")

    return etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_document_with_style_level_four_heading_and_body() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))

    for text, style, font_size_pt in [
        ("\u58f9\u3001\u5e8f\u8a00", None, None),
        ("\u7b2c\u56db\u968e\u81ea\u52d5\u6a19\u984c", "NumberedL4", None),
        ("\u7b2c\u56db\u968e\u6a19\u984c\u4e0b\u65b9\u666e\u901a\u5167\u6587", "BodyText", 14),
    ]:
        p = etree.SubElement(body, qn("p"))
        pPr = etree.SubElement(p, qn("pPr"))
        if style is not None:
            p_style = etree.SubElement(pPr, qn("pStyle"))
            p_style.set(qn("val"), style)
        r = etree.SubElement(p, qn("r"))
        if font_size_pt is not None:
            rPr = etree.SubElement(r, qn("rPr"))
            sz = etree.SubElement(rPr, qn("sz"))
            sz.set(qn("val"), str(font_size_pt * 2))
        t = etree.SubElement(r, qn("t"))
        t.text = text

    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


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
    def test_special_table_uses_previous_paragraph_text_start_and_page_right_boundary(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))

        body.append(make_table([5, 5]))

        p = etree.SubElement(body, qn("p"))
        p_pr = etree.SubElement(p, qn("pPr"))
        ind = etree.SubElement(p_pr, qn("ind"))
        ind.set(qn("left"), "1440")
        ind.set(qn("firstLine"), "360")
        r = etree.SubElement(p, qn("r"))
        t = etree.SubElement(r, qn("t"))
        t.text = "Anchor paragraph"

        body.append(make_table([4, 4]))

        sect_pr = etree.SubElement(body, qn("sectPr"))
        pg_sz = etree.SubElement(sect_pr, qn("pgSz"))
        pg_sz.set(qn("w"), "11906")
        pg_mar = etree.SubElement(sect_pr, qn("pgMar"))
        pg_mar.set(qn("left"), "1800")
        pg_mar.set(qn("right"), "1800")

        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
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
                ),
            )

            self.assertEqual(summary.special_autofit_right_tables, 1)
            self.assertEqual(summary.normal_processed_tables, 0)
            self.assertEqual(summary.skipped_first_page_tables, 1)

            root = read_document_root(output_docx)
            tables = root.xpath(".//w:tbl", namespaces=NS)
            self.assertIsNone(tables[0].find("w:tblPr", NS))
            special_tbl = tables[1]

            self.assertEqual(special_tbl.find("./w:tblPr/w:jc", NS).get(qn("val")), "left")
            self.assertEqual(special_tbl.find("./w:tblPr/w:tblLayout", NS).get(qn("type")), "fixed")
            self.assertEqual(special_tbl.find("./w:tblPr/w:tblInd", NS).get(qn("type")), "dxa")
            self.assertEqual(special_tbl.find("./w:tblPr/w:tblInd", NS).get(qn("w")), "1800")
            self.assertEqual(special_tbl.find("./w:tblPr/w:tblW", NS).get(qn("type")), "dxa")
            self.assertEqual(special_tbl.find("./w:tblPr/w:tblW", NS).get(qn("w")), "6506")

            for run in special_tbl.xpath(".//w:r", namespaces=NS):
                r_pr = run.find("w:rPr", NS)
                self.assertIsNotNone(r_pr)
                self.assertEqual(r_pr.find("w:sz", NS).get(qn("val")), "22")
                self.assertEqual(r_pr.find("w:szCs", NS).get(qn("val")), "22")

    def test_find_word_paragraph_index_for_record_prefers_direct_index_then_falls_back_to_preview(self):
        paragraph_texts = [
            "憯嫘?閮\r",
            "嚗?嚗洵銝?璅?\r",
            "?惜3?扳?\r",
        ]
        direct_record = {
            "paragraph_index": 3,
            "text_preview": "?惜3?扳?",
        }
        fallback_record = {
            "paragraph_index": 2,
            "text_preview": "?惜3?扳?",
        }

        self.assertEqual(find_word_paragraph_index_for_record(paragraph_texts, direct_record), 3)
        self.assertEqual(find_word_paragraph_index_for_record(paragraph_texts, fallback_record), 3)

    def test_word_com_gen_py_cache_error_cleans_retries_and_falls_back_to_dynamic_dispatch(self):
        class FakeRange:
            Text = "body text\r"

        class FakeFormat:
            LeftIndent = 0.0
            FirstLineIndent = 0.0
            CharacterUnitLeftIndent = 2
            CharacterUnitFirstLineIndent = 3

        class FakeParagraph:
            def __init__(self):
                self.Range = FakeRange()
                self.Format = FakeFormat()

        class FakeParagraphs:
            Count = 1

            def __init__(self, paragraph):
                self.paragraph = paragraph

            def __call__(self, index):
                if index != 1:
                    raise IndexError(index)
                return self.paragraph

        class FakeDocument:
            def __init__(self):
                self.Paragraphs = FakeParagraphs(FakeParagraph())
                self.saved = False
                self.closed = False

            def Save(self):
                self.saved = True

            def Close(self, save_changes):
                self.closed = True

        class FakeDocuments:
            def __init__(self, document):
                self.document = document

            def Open(self, *args, **kwargs):
                return self.document

        class FakeWord:
            def __init__(self, document):
                self.Documents = FakeDocuments(document)
                self.Visible = True
                self.quit = False

            def Quit(self):
                self.quit = True

        fake_document = FakeDocument()
        fake_word = FakeWord(fake_document)
        win32com_module = types.ModuleType("win32com")
        win32com_module.__path__ = []
        client_module = types.ModuleType("win32com.client")
        client_module.__path__ = []
        gencache_module = types.ModuleType("win32com.client.gencache")
        dynamic_module = types.ModuleType("win32com.client.dynamic")

        client_module.DispatchEx = Mock(
            side_effect=[
                AttributeError("module has no attribute CLSIDToClassMap"),
                RuntimeError("retry still broken"),
            ]
        )
        gencache_module.GetGeneratePath = Mock(return_value="C:\\fake\\gen_py")
        gencache_module.Rebuild = Mock()
        dynamic_module.Dispatch = Mock(return_value=fake_word)
        win32com_module.client = client_module
        client_module.gencache = gencache_module
        client_module.dynamic = dynamic_module

        fake_modules = {
            "win32com": win32com_module,
            "win32com.client": client_module,
            "win32com.client.gencache": gencache_module,
            "win32com.client.dynamic": dynamic_module,
        }
        records = [
            {
                "paragraph_index": 1,
                "text_preview": "body text",
                "expected_left_cm": 1.0,
                "expected_left_points": 28.3464567,
            }
        ]

        with tempfile.TemporaryDirectory() as temp_dir:
            output_docx = Path(temp_dir) / "output.docx"
            with patch.dict(sys.modules, fake_modules), patch("shutil.rmtree") as rmtree:
                logs = _verify_and_fix_body_indents_with_word_com_in_process(output_docx, records)

        joined = "\n".join(logs)
        self.assertEqual(client_module.DispatchEx.call_count, 2)
        rmtree.assert_called_once_with("C:\\fake\\gen_py", ignore_errors=True)
        gencache_module.Rebuild.assert_called_once()
        dynamic_module.Dispatch.assert_called_once_with("Word.Application")
        self.assertTrue(fake_document.saved)
        self.assertTrue(fake_word.quit)
        self.assertIn("WORD_COM_RETRY_CLEAR_GEN_PY", joined)
        self.assertIn("WORD_COM_RETRY_AFTER_CLEAR_GEN_PY status=failed", joined)
        self.assertIn("WORD_COM_DYNAMIC_DISPATCH status=ok", joined)
        self.assertIn("WORD_COM_BODY_INDENT_FIX:", joined)
        self.assertNotIn("WORD_COM_BODY_INDENT_FIX_SKIPPED", joined)

    def test_word_com_body_indent_applies_level_two_first_line_indent(self):
        class FakeRange:
            Text = "body text\r"

        class FakeFormat:
            LeftIndent = 0.0
            FirstLineIndent = 0.0
            CharacterUnitLeftIndent = 2
            CharacterUnitFirstLineIndent = 3

        class FakeParagraph:
            def __init__(self):
                self.Range = FakeRange()
                self.Format = FakeFormat()

        class FakeParagraphs:
            Count = 1

            def __init__(self, paragraph):
                self.paragraph = paragraph

            def __call__(self, index):
                if index != 1:
                    raise IndexError(index)
                return self.paragraph

        class FakeDocument:
            def __init__(self):
                self.paragraph = FakeParagraph()
                self.Paragraphs = FakeParagraphs(self.paragraph)
                self.saved = False
                self.closed = False

            def Save(self):
                self.saved = True

            def Close(self, SaveChanges=False):
                self.closed = True

        class FakeDocuments:
            def __init__(self, document):
                self.document = document

            def Open(self, *args, **kwargs):
                return self.document

        class FakeWord:
            def __init__(self, document):
                self.Visible = False
                self.Documents = FakeDocuments(document)
                self.quit = False

            def Quit(self):
                self.quit = True

        fake_document = FakeDocument()
        fake_word = FakeWord(fake_document)
        win32com_module = types.ModuleType("win32com")
        win32com_module.__path__ = []
        client_module = types.ModuleType("win32com.client")
        client_module.__path__ = []
        client_module.DispatchEx = Mock(return_value=fake_word)
        win32com_module.client = client_module

        fake_modules = {
            "win32com": win32com_module,
            "win32com.client": client_module,
        }
        records = [
            {
                "paragraph_index": 1,
                "text_preview": "body text",
                "kind": "body",
                "level": 1,
                "expected_left_cm": 1.83,
                "expected_left_points": 51.874015761,
                "expected_firstline_cm": 0.987777777,
                "expected_firstline_points": 28.0,
                "expected_first_line_twips": "560",
            }
        ]

        with tempfile.TemporaryDirectory() as temp_dir:
            output_docx = Path(temp_dir) / "output.docx"
            with patch.dict(sys.modules, fake_modules):
                logs = _verify_and_fix_body_indents_with_word_com_in_process(output_docx, records)

        self.assertTrue(fake_document.saved)
        self.assertTrue(fake_word.quit)
        self.assertEqual(fake_document.paragraph.Format.FirstLineIndent, 28.0)
        self.assertEqual(fake_document.paragraph.Format.CharacterUnitFirstLineIndent, 0)
        self.assertIn("expected_first_line_twips=560", "\n".join(logs))

    def test_word_com_body_font_check_skips_when_word_font_is_not_14_pt(self):
        class FakeFont:
            Size = 12.0

        class FakeRange:
            Text = "body text\r"
            Start = 0
            End = 10
            Font = FakeFont()

            @property
            def Duplicate(self):
                return self

        class FakeFormat:
            LeftIndent = 0.0
            FirstLineIndent = 0.0
            CharacterUnitLeftIndent = 2
            CharacterUnitFirstLineIndent = 3

        class FakeParagraph:
            def __init__(self):
                self.Range = FakeRange()
                self.Format = FakeFormat()

        class FakeParagraphs:
            Count = 1

            def __init__(self, paragraph):
                self.paragraph = paragraph

            def __call__(self, index):
                if index != 1:
                    raise IndexError(index)
                return self.paragraph

        class FakeDocument:
            def __init__(self):
                self.paragraph = FakeParagraph()
                self.Paragraphs = FakeParagraphs(self.paragraph)
                self.saved = False
                self.closed = False

            def Save(self):
                self.saved = True

            def Close(self, SaveChanges=False):
                self.closed = True

        class FakeDocuments:
            def __init__(self, document):
                self.document = document

            def Open(self, *args, **kwargs):
                return self.document

        class FakeWord:
            def __init__(self, document):
                self.Visible = False
                self.Documents = FakeDocuments(document)
                self.quit = False

            def Quit(self):
                self.quit = True

        fake_document = FakeDocument()
        fake_word = FakeWord(fake_document)
        win32com_module = types.ModuleType("win32com")
        win32com_module.__path__ = []
        client_module = types.ModuleType("win32com.client")
        client_module.__path__ = []
        client_module.DispatchEx = Mock(return_value=fake_word)
        win32com_module.client = client_module

        fake_modules = {
            "win32com": win32com_module,
            "win32com.client": client_module,
        }
        records = [
            {
                "paragraph_index": 1,
                "text_preview": "body text",
                "kind": "body_font_check",
                "level": 1,
                "expected_left_cm": 1.83,
                "expected_left_points": 51.874015761,
                "expected_firstline_cm": 0.987777777,
                "expected_firstline_points": 28.0,
                "expected_first_line_twips": "560",
                "xml_font_size": 12.0,
                "xml_font_size_source": "dominant_runs",
                "apply_only_if_word_font_size_is_14": True,
            }
        ]

        with tempfile.TemporaryDirectory() as temp_dir:
            output_docx = Path(temp_dir) / "output.docx"
            with patch.dict(sys.modules, fake_modules):
                logs = _verify_and_fix_body_indents_with_word_com_in_process(output_docx, records)

        self.assertTrue(fake_document.saved)
        self.assertTrue(fake_word.quit)
        self.assertEqual(fake_document.paragraph.Format.LeftIndent, 0.0)
        self.assertEqual(fake_document.paragraph.Format.FirstLineIndent, 0.0)
        joined = "\n".join(logs)
        self.assertIn("word_dominant_font_size=12", joined)
        self.assertIn("decision=skipped_word_font_not_14", joined)

    def test_word_com_body_indent_uses_powershell_wrapper_with_logs(self):
        records = [
            {
                "paragraph_index": 1,
                "text_preview": "body text",
                "expected_left_cm": 1.0,
                "expected_left_points": 28.3464567,
            }
        ]
        captured: dict[str, object] = {}

        def fake_run(script_path, *, arguments=None, **kwargs):
            captured["script_path"] = Path(script_path)
            captured["arguments"] = list(arguments or [])
            records_path = Path(arguments[3])
            result_path = Path(arguments[5])
            captured["records"] = json.loads(records_path.read_text(encoding="utf-8"))
            captured["script_text"] = Path(script_path).read_text(encoding="utf-8")
            result_path.write_text(
                "\n".join(
                    [
                        "WORD_COM_PS_STARTED",
                        "WORD_COM_PS_RECORDS_LOADED count=1",
                        "WORD_COM_PS_WORD_CREATED",
                        "WORD_COM_PS_DOC_OPENED",
                        "WORD_COM_PS_BEFORE_LOOP",
                        "WORD_COM_PS_PARAGRAPHS_COUNT count=1",
                        "WORD_COM_RECORD_BEGIN i=1 paragraph_index=1 expected_left_cm=1.0 text=body text",
                        "WORD_COM_RECORD_MATCHED i=1 word_index=1",
                        "WORD_COM_BODY_INDENT_FIX: paragraph_index=1; matched_paragraph_index=1; text=body text; expected_left_cm=1.00; before_left_cm=0.00; after_left_cm=1.00; status=ok",
                        "WORD_COM_PS_BEFORE_SAVE",
                        "WORD_COM_PS_DOC_SAVED",
                        "WORD_COM_PS_DONE",
                        "WORD_COM_BODY_INDENT_FIX_SUMMARY processed=1 ok=1 mismatch=0 not_found=0 errors=0",
                    ]
                ),
                encoding="utf-8",
            )
            return types.SimpleNamespace(
                stdout="",
                stderr="",
                returncode=0,
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            output_docx = Path(temp_dir) / "output.docx"
            output_docx.write_text("docx-bytes-placeholder", encoding="utf-8")
            with patch("docx_fixer.docx_processor.run_powershell_file", side_effect=fake_run) as runner:
                logs = verify_and_fix_body_indents_with_word_com(output_docx, records)

        self.assertIn("WORD_COM_BODY_INDENT_FIX_STARTED", logs)
        self.assertTrue(any("WORD_COM_POWERSHELL_SCRIPT_PATH=" in line for line in logs))
        self.assertTrue(any("WORD_COM_RECORDS_JSON_PATH=" in line for line in logs))
        self.assertTrue(any("WORD_COM_DOCX_WORK_PATH=" in line for line in logs))
        self.assertTrue(any("WORD_COM_RESULT_LOG_PATH=" in line for line in logs))
        self.assertTrue(any("WORD_COM_POWERSHELL_RETURN_CODE=0" in line for line in logs))
        self.assertTrue(any("WORD_COM_BODY_INDENT_FIX:" in line for line in logs))
        self.assertTrue(any("WORD_COM_BODY_INDENT_FIX_SUMMARY" in line for line in logs))
        self.assertTrue(any("WORD_COM_PS_STARTED" in line for line in logs))
        self.assertTrue(any("WORD_COM_PS_DONE" in line for line in logs))
        self.assertTrue(any("WORD_COM_PS_BEFORE_LOOP" in line for line in logs))
        self.assertTrue(any("WORD_COM_PS_PARAGRAPHS_COUNT count=1" in line for line in logs))
        self.assertEqual(captured["records"], records)
        self.assertIn("param(", captured["script_text"])
        self.assertIn("$RecordsPath", captured["script_text"])
        self.assertIn("Add-Log", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_BEGIN", captured["script_text"])
        runner.assert_called_once()

    def test_word_com_body_indent_reads_result_file_when_stdout_is_empty(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_docx = Path(temp_dir) / "output.docx"
            output_docx.write_text("docx-bytes-placeholder", encoding="utf-8")

            def fake_run(script_path, *, arguments=None, **kwargs):
                result_path = Path(arguments[5])
                result_path.write_text(
                    "\n".join(
                        [
                            "WORD_COM_PS_STARTED",
                            "WORD_COM_PS_RECORDS_LOADED count=1",
                            "WORD_COM_PS_DONE",
                            "WORD_COM_BODY_INDENT_FIX_SUMMARY processed=1 ok=0 mismatch=0 not_found=1 errors=0",
                        ]
                    ),
                    encoding="utf-8",
                )
                return types.SimpleNamespace(stdout="", stderr="", returncode=0)

            with patch(
                "docx_fixer.docx_processor.run_powershell_file",
                side_effect=fake_run,
            ):
                logs = verify_and_fix_body_indents_with_word_com(
                    output_docx,
                    [{"paragraph_index": 1, "text_preview": "body text", "expected_left_cm": 1.0, "expected_left_points": 28.3464567}],
                )

        self.assertTrue(any("WORD_COM_PS_STARTED" in line for line in logs))
        self.assertTrue(any("WORD_COM_BODY_INDENT_FIX_SUMMARY" in line for line in logs))
        self.assertFalse(any("powershell_no_logs" in line for line in logs))

    def test_word_com_body_indent_logs_stderr_when_stdout_is_empty(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_docx = Path(temp_dir) / "output.docx"
            output_docx.write_text("docx-bytes-placeholder", encoding="utf-8")
            with patch(
                "docx_fixer.docx_processor.run_powershell_file",
                return_value=types.SimpleNamespace(stdout="", stderr="boom on stderr", returncode=1),
            ):
                logs = verify_and_fix_body_indents_with_word_com(
                    output_docx,
                    [{"paragraph_index": 1, "text_preview": "body text", "expected_left_cm": 1.0, "expected_left_points": 28.3464567}],
                )

        self.assertTrue(any("WORD_COM_POWERSHELL_STDERR=boom on stderr" in line for line in logs))
        self.assertTrue(any("WORD_COM_POWERSHELL_RETURN_CODE=1" in line for line in logs))
        self.assertTrue(any("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_no_logs:boom on stderr" in line for line in logs))

    def test_word_com_body_indent_logs_script_exception(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_docx = Path(temp_dir) / "output.docx"
            output_docx.write_text("docx-bytes-placeholder", encoding="utf-8")

            def fake_run(script_path, *, arguments=None, **kwargs):
                result_path = Path(arguments[5])
                result_path.write_text(
                    "\n".join(
                        [
                            "WORD_COM_PS_STARTED",
                            "WORD_COM_PS_EXCEPTION System.Exception:boom",
                            "WORD_COM_PS_STACK at fix_word_indent",
                            "WORD_COM_PS_FINALLY_BEGIN",
                        ]
                    ),
                    encoding="utf-8",
                )
                return types.SimpleNamespace(stdout="", stderr="", returncode=1)

            with patch(
                "docx_fixer.docx_processor.run_powershell_file",
                side_effect=fake_run,
            ):
                logs = verify_and_fix_body_indents_with_word_com(
                    output_docx,
                    [{"paragraph_index": 1, "text_preview": "body text", "expected_left_cm": 1.0, "expected_left_points": 28.3464567}],
                )

        self.assertTrue(any("WORD_COM_PS_EXCEPTION System.Exception:boom" in line for line in logs))
        self.assertTrue(any("WORD_COM_BODY_INDENT_FIX_FAILED_AFTER_PARTIAL_LOGS" in line for line in logs))
        self.assertTrue(any("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_script_failed" in line for line in logs))

    def test_word_com_body_indent_large_records_use_json_file_and_short_command(self):
        records = [
            {
                "paragraph_index": index + 1,
                "text_preview": f"body text {index}",
                "expected_left_cm": 1.0,
                "expected_left_points": 28.3464567,
            }
            for index in range(5000)
        ]
        observed: dict[str, object] = {}

        def fake_run(script_path, *, arguments=None, **kwargs):
            args = list(arguments or [])
            records_path = Path(args[3])
            docx_path = Path(args[1])
            observed["args"] = args
            observed["script_path"] = Path(script_path)
            observed["records_path"] = records_path
            observed["docx_path"] = docx_path
            observed["records_count"] = len(json.loads(records_path.read_text(encoding="utf-8")))
            observed["command_length"] = sum(
                len(part)
                for part in ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script_path), *args]
            ) + (10 - 1)
            return types.SimpleNamespace(
                stdout='["WORD_COM_BODY_INDENT_FIX_SUMMARY records=5000; ok=0; mismatch=0; not_found=5000"]',
                stderr="",
            )

        with tempfile.TemporaryDirectory() as temp_dir:
            long_dir = Path(temp_dir) / ("nested_" * 12)
            long_dir.mkdir(parents=True, exist_ok=True)
            output_docx = long_dir / "very_long_output_name.docx"
            output_docx.write_text("docx-bytes-placeholder", encoding="utf-8")

            with patch("docx_fixer.docx_processor.run_powershell_file", side_effect=fake_run):
                logs = verify_and_fix_body_indents_with_word_com(output_docx, records)

        command_length_line = next(line for line in logs if line.startswith("command_length="))
        work_path_line = next(line for line in logs if line.startswith("WORD_COM_DOCX_WORK_PATH="))
        self.assertEqual(observed["records_count"], 5000)
        self.assertNotIn("body text 4999", " ".join(observed["args"]))
        self.assertTrue(str(observed["records_path"]).endswith(".json"))
        self.assertTrue(str(observed["script_path"]).endswith(".ps1"))
        self.assertLess(int(command_length_line.split("=", 1)[1]), 400)
        self.assertLess(len(work_path_line.split("=", 1)[1]), len(str(output_docx)))

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
                ),
            )

            root = read_document_root(output_docx)
            body_paragraph = root.xpath(".//w:p", namespaces=NS)[2]
            ind = body_paragraph.find("./w:pPr/w:ind", NS)
            self.assertEqual(ind.get(qn("left")), TEMPLATE_OUTLINE_INDENTS[3]["body_left"])
            self.assertIsNone(ind.get(qn("start")))
            self.assertIsNone(ind.get(qn("hanging")))
            self.assertIsNone(body_paragraph.find("./w:pPr/w:tabs", NS))
            debug = "\n".join(summary.body_indent_debug_logs)
            self.assertIn("font_size_source=paragraph_style:DefaultText", debug)
            self.assertTrue(any(int(record["paragraph_index"]) == 3 for record in summary.body_indent_records))

    def test_level_two_body_indent_sets_first_line_twips_and_record(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_document_with_styled_level_two_body(),
                styles_xml=make_styles_with_default_text_font(14),
            )

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=True,
                    normalize_with_word_com=False,
                    enable_level2_body_first_line_indent=True,
                ),
            )

            root = read_document_root(output_docx)
            body_paragraph = root.xpath(".//w:p", namespaces=NS)[2]
            ind = body_paragraph.find("./w:pPr/w:ind", NS)
            self.assertEqual(ind.get(qn("left")), TEMPLATE_OUTLINE_INDENTS[1]["body_left"])
            self.assertEqual(ind.get(qn("firstLine")), "560")
            self.assertIsNone(ind.get(qn("hanging")))
            self.assertIsNone(ind.get(qn("start")))
            self.assertIsNone(body_paragraph.find("./w:pPr/w:tabs", NS))
            debug = "\n".join(summary.body_indent_debug_logs)
            self.assertIn("spec_firstLine_twips=560", debug)
            self.assertIn("written_firstLine=560", debug)
            records_by_kind = {record["kind"]: record for record in summary.body_indent_records}
            self.assertEqual(records_by_kind["body"]["expected_first_line_twips"], "560")
            self.assertEqual(records_by_kind["body"]["expected_firstline_points"], 28.0)

    def test_body_indent_direct_format_overrides_styles_xml_start_indent(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_document_with_styled_level_four_body(),
                styles_xml=make_styles_with_default_text_font(14, start_twips="1480"),
            )

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=True,
                ),
            )

            root = read_document_root(output_docx)
            body_paragraph = root.xpath(".//w:p", namespaces=NS)[2]
            ind = body_paragraph.find("./w:pPr/w:ind", NS)
            self.assertEqual(ind.get(qn("left")), TEMPLATE_OUTLINE_INDENTS[3]["body_left"])
            self.assertIsNone(ind.get(qn("start")))
            self.assertIsNone(body_paragraph.find("./w:pPr/w:tabs", NS))
            debug = "\n".join(summary.body_indent_debug_logs)
            self.assertIn("written_start_twips=None", debug)
            self.assertIn("validation=ok", debug)

    def test_level_four_indents_are_normalized_in_document_numbering_and_styles(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_document_with_style_level_four_heading_and_body(),
                styles_xml=make_styles_with_level_four_numbered_and_plain_old_indents(),
                numbering_xml=make_level_four_numbering_with_old_indents(),
            )

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=True,
                    normalize_with_word_com=False,
                ),
            )

            spec = TEMPLATE_OUTLINE_INDENTS[3]
            document_root = read_part_root(output_docx, "word/document.xml")
            paragraphs = document_root.xpath(".//w:p", namespaces=NS)
            heading_ind = paragraphs[1].find("./w:pPr/w:ind", NS)
            body_ind = paragraphs[2].find("./w:pPr/w:ind", NS)
            self.assertEqual(heading_ind.get(qn("left")), spec["left"])
            self.assertEqual(heading_ind.get(qn("hanging")), spec["hanging"])
            self.assertAlmostEqual(int(heading_ind.get(qn("left"))) / 20 / 28.3464567, 3.46, places=2)
            self.assertAlmostEqual(int(heading_ind.get(qn("hanging"))) / 20 / 28.3464567, 0.50, places=2)
            self.assertEqual(body_ind.get(qn("left")), spec["body_left"])
            self.assertIsNone(body_ind.get(qn("hanging")))
            self.assertIsNone(body_ind.get(qn("firstLine")))
            self.assertIsNone(body_ind.get(qn("start")))
            self.assertIsNone(paragraphs[2].find("./w:pPr/w:tabs", NS))

            numbering_root = read_part_root(output_docx, "word/numbering.xml")
            numbering_lvl = numbering_root.find(".//w:lvl", NS)
            numbering_ind = numbering_lvl.find("./w:pPr/w:ind", NS)
            self.assertEqual(numbering_ind.get(qn("left")), spec["left"])
            self.assertEqual(numbering_ind.get(qn("hanging")), spec["hanging"])
            self.assertIsNone(numbering_ind.get(qn("start")))
            self.assertEqual(numbering_lvl.find("./w:suff", NS).get(qn("val")), "nothing")
            self.assertIsNone(numbering_lvl.find("./w:pPr/w:tabs", NS))

            styles_root = read_part_root(output_docx, "word/styles.xml")
            numbered_style = styles_root.xpath("./w:style[@w:styleId='NumberedL4']", namespaces=NS)[0]
            numbered_ind = numbered_style.find("./w:pPr/w:ind", NS)
            self.assertEqual(numbered_ind.get(qn("left")), spec["left"])
            self.assertEqual(numbered_ind.get(qn("hanging")), spec["hanging"])
            self.assertIsNone(numbered_ind.get(qn("start")))
            self.assertIsNone(numbered_style.find("./w:pPr/w:tabs", NS))
            plain_style = styles_root.xpath("./w:style[@w:styleId='BodyText']", namespaces=NS)[0]
            self.assertIsNone(plain_style.find("./w:pPr/w:ind", NS))
            self.assertIsNone(plain_style.find("./w:pPr/w:tabs", NS))

            assert_docx_has_no_character_indent_attrs(self, output_docx)
            logs = "\n".join(summary.numbering_xml_logs)
            self.assertIn("NUMBERING_XML_LEVEL_INDENT", logs)
            self.assertIn("STYLES_XML_NUMBERED_STYLE_INDENT", logs)
            self.assertIn("expected_number_start_cm=2.96", logs)
            self.assertIn("suff=nothing", logs)
            self.assertIn("tab_pos_cm=None", logs)
            records_by_kind = {record["kind"]: record for record in summary.body_indent_records}
            self.assertAlmostEqual(records_by_kind["auto(style)"]["expected_heading_left_cm"], 3.46, places=2)
            self.assertAlmostEqual(records_by_kind["auto(style)"]["expected_hanging_cm"], 0.50, places=2)
            self.assertAlmostEqual(records_by_kind["body"]["expected_body_left_cm"], 3.70, places=2)
            self.assertEqual(records_by_kind["body"]["expected_firstline_cm"], 0.0)

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
                ),
            )

            root = read_part_root(output_docx, "word/numbering.xml")
            lvl = root.find(".//w:lvl", NS)
            ind = lvl.find("./w:pPr/w:ind", NS)
            self.assertEqual(ind.get(qn("left")), "360")
            self.assertIsNone(ind.get(qn("hanging")))
            self.assertEqual(lvl.find("./w:suff", NS).get(qn("val")), "nothing")
            self.assertIsNone(lvl.find("./w:pPr/w:tabs", NS))
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
                    remove_all_outline_levels=True,
                ),
            )

            self.assertEqual(paragraph_outlines(output_docx), ["9", "0"])
            self.assertEqual(summary.removed_all_outline_paragraphs, 2)
            self.assertEqual(summary.paragraph_level_counts[0], 1)


if __name__ == "__main__":
    unittest.main()
