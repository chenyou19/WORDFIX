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
from docx_fixer.docx_processor import collect_heading_suffix_records_from_docx, fix_docx_fast
from docx_fixer.indent_settings import twips_to_cm
from docx_fixer.note_debug_log import collect_note_debug_records_from_docx
from docx_fixer.outline import collect_all_toc_paragraph_ids
from docx_fixer.protected_region import ProtectedRegionContext, collect_chapter_three_paragraph_ids
from docx_fixer.numbering import (
    build_numbering_format_lookup,
    build_numbering_level_lookup,
    build_style_numbering_lookup,
    numbering_suffix_for_level,
)
from docx_fixer.word_com_indent import (
    WORD_COM_TIMEOUT_SECONDS,
    _filter_word_com_body_indent_records,
    _verify_and_fix_body_indents_with_word_com_in_process,
    apply_word_com_approved_body_indents_to_docx_xml,
    find_word_paragraph_index_for_record,
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


def assert_numbering_level_follows_suffix_rule(test_case, lvl, level):
    """Assert one numbering w:lvl matches the central suffix rule for its level.

    Levels 3/5/7 use w:suff="space"; all other levels use
    w:suff="nothing". No level keeps numbering tab stops.
    """
    suff = lvl.find("./w:suff", NS)
    tabs = lvl.find("./w:pPr/w:tabs", NS)
    test_case.assertEqual(suff.get(qn("val")), numbering_suffix_for_level(level))
    test_case.assertIsNone(tabs)


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


def make_settings_xml(*, do_not_use_indent_as_numbering_tab_stop: str | None | bool = None) -> bytes:
    settings = etree.Element(qn("settings"), nsmap={"w": W_NS})
    compat = etree.SubElement(settings, qn("compat"))
    if do_not_use_indent_as_numbering_tab_stop is not None:
        flag = etree.SubElement(compat, qn("doNotUseIndentAsNumberingTabStop"))
        if isinstance(do_not_use_indent_as_numbering_tab_stop, str):
            flag.set(qn("val"), do_not_use_indent_as_numbering_tab_stop)
    return etree.tostring(settings, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))

    for text, outline, style in [
        ("\u666e\u901a\u6b63\u6587", 5, "Heading1"),
        ("\u58f9\u3001\u5e8f\u8a00", 2, None),
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


def make_heading_suffix_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))

    for text in ["壹、序言", "壹、序言", "一、 手動標題"]:
        p = etree.SubElement(body, qn("p"))
        r = etree.SubElement(p, qn("r"))
        t = etree.SubElement(r, qn("t"))
        t.text = text

    p = etree.SubElement(body, qn("p"))
    pPr = etree.SubElement(p, qn("pPr"))
    numPr = etree.SubElement(pPr, qn("numPr"))
    ilvl = etree.SubElement(numPr, qn("ilvl"))
    ilvl.set(qn("val"), "0")
    numId = etree.SubElement(numPr, qn("numId"))
    numId.set(qn("val"), "42")
    r = etree.SubElement(p, qn("r"))
    t = etree.SubElement(r, qn("t"))
    t.text = "自動標題"

    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_heading_suffix_numbering_xml(
    suffix: str | None = "tab",
    *,
    include_tabs: bool = True,
    lvl_text_value: str = "%1.",
    tab_val: str = "left",
    tab_pos: str = "2279",
) -> bytes:
    numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    abstract = etree.SubElement(numbering, qn("abstractNum"))
    abstract.set(qn("abstractNumId"), "42")
    lvl = etree.SubElement(abstract, qn("lvl"))
    lvl.set(qn("ilvl"), "0")
    num_fmt = etree.SubElement(lvl, qn("numFmt"))
    num_fmt.set(qn("val"), "decimal")
    lvl_text = etree.SubElement(lvl, qn("lvlText"))
    lvl_text.set(qn("val"), lvl_text_value)
    if suffix is not None:
        suff = etree.SubElement(lvl, qn("suff"))
        suff.set(qn("val"), suffix)
    pPr = etree.SubElement(lvl, qn("pPr"))
    if include_tabs:
        tabs = etree.SubElement(pPr, qn("tabs"))
        tab = etree.SubElement(tabs, qn("tab"))
        tab.set(qn("val"), tab_val)
        tab.set(qn("pos"), tab_pos)
    ind = etree.SubElement(pPr, qn("ind"))
    ind.set(qn("left"), "2279")
    ind.set(qn("hanging"), "420")

    num = etree.SubElement(numbering, qn("num"))
    num.set(qn("numId"), "42")
    abstract_id = etree.SubElement(num, qn("abstractNumId"))
    abstract_id.set(qn("val"), "42")
    return etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_chapter_three_shared_numbering_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    add_test_paragraph(body, "\u58f9\u3001\u5e8f\u8a00")
    add_test_paragraph(
        body,
        "\u50f9\u683c\u5f62\u6210\u4e4b\u4e3b\u8981\u56e0\u7d20\u5206\u6790",
        num_id="42",
        ilvl=0,
    )
    add_test_paragraph(body, "\u53c3\u7ae0\u5167\u6587")
    add_test_paragraph(body, "\u8086\u3001\u7b2c\u56db\u7ae0", num_id="99", ilvl=0)
    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_shared_chapter_three_numbering_xml() -> bytes:
    numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    abstract = etree.SubElement(numbering, qn("abstractNum"))
    abstract.set(qn("abstractNumId"), "1")
    lvl = etree.SubElement(abstract, qn("lvl"))
    lvl.set(qn("ilvl"), "0")
    num_fmt = etree.SubElement(lvl, qn("numFmt"))
    num_fmt.set(qn("val"), "ideographLegalTraditional")
    lvl_text = etree.SubElement(lvl, qn("lvlText"))
    lvl_text.set(qn("val"), "%1\u3001 ")
    suff = etree.SubElement(lvl, qn("suff"))
    suff.set(qn("val"), "tab")
    pPr = etree.SubElement(lvl, qn("pPr"))
    tabs = etree.SubElement(pPr, qn("tabs"))
    tab = etree.SubElement(tabs, qn("tab"))
    tab.set(qn("val"), "num")
    tab.set(qn("pos"), "2061")

    for num_id in ("42", "99"):
        num = etree.SubElement(numbering, qn("num"))
        num.set(qn("numId"), num_id)
        abstract_ref = etree.SubElement(num, qn("abstractNumId"))
        abstract_ref.set(qn("val"), "1")

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


def make_document_with_style_level_four_heading_and_body(
    body_font_size_pt: int = 14,
    body_text: str = "\u7b2c\u56db\u968e\u6a19\u984c\u4e0b\u65b9\u666e\u901a\u5167\u6587",
) -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))

    for text, style, font_size_pt in [
        ("\u58f9\u3001\u5e8f\u8a00", None, None),
        ("\u58f9\u3001\u5e8f\u8a00", None, None),
        ("\u7b2c\u56db\u968e\u81ea\u52d5\u6a19\u984c", "NumberedL4", None),
        (body_text, "BodyText", body_font_size_pt),
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


def add_test_paragraph(
    body,
    text: str,
    *,
    style: str | None = None,
    num_id: str | None = None,
    ilvl: int | None = None,
    outline: int | None = None,
    ind_attrs: dict[str, str] | None = None,
    tab_pos: str | None = None,
    font_size_pt: int | None = None,
):
    p = etree.SubElement(body, qn("p"))
    pPr = etree.SubElement(p, qn("pPr"))
    if style is not None:
        p_style = etree.SubElement(pPr, qn("pStyle"))
        p_style.set(qn("val"), style)
    if outline is not None:
        outline_el = etree.SubElement(pPr, qn("outlineLvl"))
        outline_el.set(qn("val"), str(outline))
    if num_id is not None:
        num_pr = etree.SubElement(pPr, qn("numPr"))
        ilvl_el = etree.SubElement(num_pr, qn("ilvl"))
        ilvl_el.set(qn("val"), str(ilvl or 0))
        num_id_el = etree.SubElement(num_pr, qn("numId"))
        num_id_el.set(qn("val"), num_id)
    if ind_attrs:
        make_ind(pPr, **ind_attrs)
    if tab_pos is not None:
        tabs = etree.SubElement(pPr, qn("tabs"))
        tab = etree.SubElement(tabs, qn("tab"))
        tab.set(qn("val"), "left")
        tab.set(qn("pos"), tab_pos)
    r = etree.SubElement(p, qn("r"))
    if font_size_pt is not None:
        rPr = etree.SubElement(r, qn("rPr"))
        sz = etree.SubElement(rPr, qn("sz"))
        sz.set(qn("val"), str(font_size_pt * 2))
    t = etree.SubElement(r, qn("t"))
    t.text = text
    return p


def make_toc_immutable_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    add_test_paragraph(body, "\u76ee\u9304")
    add_test_paragraph(
        body,
        "\u58f9\u3001\u5e8f\u8a00",
        style="TOC1",
        num_id="1",
        ilvl=0,
        outline=5,
        ind_attrs={
            "left": "111",
            "hanging": "22",
            "leftChars": "333",
            "startChars": "444",
            "hangingChars": "555",
            "firstLineChars": "666",
        },
        tab_pos="777",
    )
    add_test_paragraph(
        body,
        "\u4e00\u3001\u76ee\u9304\u9805\u76ee",
        outline=6,
        ind_attrs={
            "left": "211",
            "start": "212",
            "leftChars": "313",
            "startChars": "414",
            "hangingChars": "515",
            "firstLineChars": "616",
        },
        tab_pos="878",
    )
    add_test_paragraph(body, "\u58f9\u3001\u5e8f\u8a00")
    add_test_paragraph(
        body,
        "\u9019\u662f\u6b63\u6587 14 pt \u5167\u5bb9",
        font_size_pt=14,
        ind_attrs={"left": "999", "leftChars": "111", "firstLineChars": "222"},
        tab_pos="999",
    )
    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_toc_immutable_numbering_xml() -> bytes:
    numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    for abstract_id, left, chars, tab_pos in [("1", "111", "333", "777"), ("2", "999", "444", "888")]:
        abstract = etree.SubElement(numbering, qn("abstractNum"))
        abstract.set(qn("abstractNumId"), abstract_id)
        lvl = etree.SubElement(abstract, qn("lvl"))
        lvl.set(qn("ilvl"), "0")
        num_fmt = etree.SubElement(lvl, qn("numFmt"))
        num_fmt.set(qn("val"), "decimal")
        lvl_text = etree.SubElement(lvl, qn("lvlText"))
        lvl_text.set(qn("val"), "%1. ")
        pPr = etree.SubElement(lvl, qn("pPr"))
        make_ind(pPr, left=left, hanging="22", leftChars=chars, firstLineChars="555")
        tabs = etree.SubElement(pPr, qn("tabs"))
        tab = etree.SubElement(tabs, qn("tab"))
        tab.set(qn("val"), "left")
        tab.set(qn("pos"), tab_pos)
        suff = etree.SubElement(lvl, qn("suff"))
        suff.set(qn("val"), "tab")

        num = etree.SubElement(numbering, qn("num"))
        num.set(qn("numId"), abstract_id)
        abstract_el = etree.SubElement(num, qn("abstractNumId"))
        abstract_el.set(qn("val"), abstract_id)
    return etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_toc_immutable_styles_xml() -> bytes:
    styles = etree.Element(qn("styles"), nsmap={"w": W_NS})
    toc_style = etree.SubElement(styles, qn("style"))
    toc_style.set(qn("type"), "paragraph")
    toc_style.set(qn("styleId"), "TOC1")
    toc_name = etree.SubElement(toc_style, qn("name"))
    toc_name.set(qn("val"), "Table of Contents 1")
    toc_pPr = etree.SubElement(toc_style, qn("pPr"))
    make_ind(toc_pPr, left="111", leftChars="333", firstLineChars="555")
    toc_tabs = etree.SubElement(toc_pPr, qn("tabs"))
    toc_tab = etree.SubElement(toc_tabs, qn("tab"))
    toc_tab.set(qn("val"), "left")
    toc_tab.set(qn("pos"), "777")

    numbered_style = etree.SubElement(styles, qn("style"))
    numbered_style.set(qn("type"), "paragraph")
    numbered_style.set(qn("styleId"), "BodyNumbered")
    pPr = etree.SubElement(numbered_style, qn("pPr"))
    num_pr = etree.SubElement(pPr, qn("numPr"))
    ilvl_el = etree.SubElement(num_pr, qn("ilvl"))
    ilvl_el.set(qn("val"), "0")
    num_id_el = etree.SubElement(num_pr, qn("numId"))
    num_id_el.set(qn("val"), "2")
    make_ind(pPr, left="999", leftChars="444")
    return etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_chapter_three_skip_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    add_test_paragraph(body, "\u58f9\u3001\u5e8f\u8a00")
    add_test_paragraph(body, "\u58f9\u3001\u5e8f\u8a00")
    add_test_paragraph(
        body,
        "\u53c3\u3001\u50f9\u683c\u5f62\u6210\u4e4b\u4e3b\u8981\u56e0\u7d20\u5206\u6790",
        outline=4,
        ind_attrs={"left": "321", "leftChars": "111", "firstLineChars": "222"},
        font_size_pt=14,
    )
    add_test_paragraph(
        body,
        "\u4e00\u3001\u53c3\u7ae0\u5b50\u6a19",
        outline=5,
        ind_attrs={"left": "654", "leftChars": "333", "firstLineChars": "444"},
        font_size_pt=14,
    )
    add_test_paragraph(
        body,
        "\u53c3\u7ae0\u666e\u901a\u5167\u6587",
        outline=6,
        ind_attrs={"left": "777", "leftChars": "555", "firstLineChars": "666"},
        font_size_pt=14,
    )
    body.append(make_table([3, 3]))
    add_test_paragraph(body, "\u8086\u3001\u7b2c\u56db\u7ae0")
    add_test_paragraph(
        body,
        "\u8086\u7ae0\u666e\u901a\u5167\u6587",
        ind_attrs={"left": "999", "leftChars": "888", "firstLineChars": "777"},
        font_size_pt=14,
    )
    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_chapter_three_options_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    add_test_paragraph(body, "\u58f9\u3001\u5e8f\u8a00")
    add_test_paragraph(body, "\u58f9\u3001\u5e8f\u8a00")
    body.append(make_table([5, 5]))
    add_test_paragraph(
        body,
        "\u53c3\u3001\u50f9\u683c\u5f62\u6210\u4e4b\u4e3b\u8981\u56e0\u7d20\u5206\u6790",
        outline=4,
        ind_attrs={"left": "321", "leftChars": "111", "firstLineChars": "222"},
        font_size_pt=14,
    )
    add_test_paragraph(
        body,
        "\u4e00\u3001\u53c3\u7ae0\u5b50\u6a19",
        outline=5,
        ind_attrs={"left": "654", "leftChars": "333", "firstLineChars": "444"},
        font_size_pt=14,
    )
    add_test_paragraph(
        body,
        "\u53c3\u7ae0\u666e\u901a\u5167\u6587",
        outline=6,
        ind_attrs={"left": "777", "leftChars": "555", "firstLineChars": "666"},
        font_size_pt=14,
    )
    body.append(make_table([3, 3]))
    add_test_paragraph(body, "\u8086\u3001\u7b2c\u56db\u7ae0")
    add_test_paragraph(
        body,
        "\u8086\u7ae0\u666e\u901a\u5167\u6587",
        ind_attrs={"left": "999", "leftChars": "888", "firstLineChars": "777"},
        font_size_pt=14,
    )
    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_shared_chapter_three_indent_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    add_test_paragraph(
        body,
        "\u76ee\u9304\u9805\u76ee",
        style="TOC1",
        num_id="7",
        ilvl=0,
        ind_attrs={"left": "101", "leftChars": "701", "firstLineChars": "702"},
    )
    add_test_paragraph(
        body,
        "\u58f9\u3001\u5e8f\u8a00",
        style="SharedPara",
        ind_attrs={"left": "201", "leftChars": "801", "firstLineChars": "802"},
    )
    add_test_paragraph(
        body,
        "\u53c3\u3001\u50f9\u683c\u5f62\u6210\u4e4b\u4e3b\u8981\u56e0\u7d20\u5206\u6790",
        style="SharedPara",
        num_id="42",
        ilvl=0,
        ind_attrs={"left": "321", "leftChars": "111", "firstLineChars": "222"},
    )
    add_test_paragraph(
        body,
        "\u53c3\u7ae0\u5167\u6587",
        style="SharedPara",
        ind_attrs={"left": "777", "leftChars": "555", "firstLineChars": "666"},
    )
    add_test_paragraph(
        body,
        "\u8086\u3001\u7b2c\u56db\u7ae0",
        style="SharedPara",
        num_id="99",
        ilvl=0,
        ind_attrs={"left": "987", "leftChars": "887", "firstLineChars": "787"},
    )
    add_test_paragraph(
        body,
        "\u8086\u7ae0\u5171\u7528\u6a23\u5f0f\u5167\u6587",
        style="SharedPara",
        ind_attrs={"left": "654", "leftChars": "754", "firstLineChars": "755"},
    )
    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_shared_chapter_three_indent_styles_xml() -> bytes:
    styles = etree.Element(qn("styles"), nsmap={"w": W_NS})

    toc_style = etree.SubElement(styles, qn("style"))
    toc_style.set(qn("type"), "paragraph")
    toc_style.set(qn("styleId"), "TOC1")
    toc_name = etree.SubElement(toc_style, qn("name"))
    toc_name.set(qn("val"), "Table of Contents 1")
    toc_pPr = etree.SubElement(toc_style, qn("pPr"))
    make_ind(toc_pPr, left="111", leftChars="333", firstLineChars="555")

    shared_style = etree.SubElement(styles, qn("style"))
    shared_style.set(qn("type"), "paragraph")
    shared_style.set(qn("styleId"), "SharedPara")
    shared_pPr = etree.SubElement(shared_style, qn("pPr"))
    make_ind(shared_pPr, left="720", leftChars="888", firstLineChars="999", hangingChars="777")

    return etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_shared_chapter_three_indent_numbering_xml() -> bytes:
    numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})

    for abstract_id, num_ids, left, chars in [
        ("1", ("42", "99"), "1440", "444"),
        ("7", ("7",), "360", "777"),
    ]:
        abstract = etree.SubElement(numbering, qn("abstractNum"))
        abstract.set(qn("abstractNumId"), abstract_id)
        lvl = etree.SubElement(abstract, qn("lvl"))
        lvl.set(qn("ilvl"), "0")
        num_fmt = etree.SubElement(lvl, qn("numFmt"))
        num_fmt.set(qn("val"), "ideographLegalTraditional")
        lvl_text = etree.SubElement(lvl, qn("lvlText"))
        lvl_text.set(qn("val"), "%1\u3001 ")
        pPr = etree.SubElement(lvl, qn("pPr"))
        make_ind(pPr, left=left, hanging="120", leftChars=chars, firstLineChars="555")

        for num_id in num_ids:
            num = etree.SubElement(numbering, qn("num"))
            num.set(qn("numId"), num_id)
            abstract_ref = etree.SubElement(num, qn("abstractNumId"))
            abstract_ref.set(qn("val"), abstract_id)

    return etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_note_alignment_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    add_test_paragraph(body, "\u58f9\u3001\u5e8f\u8a00")
    add_test_paragraph(body, "\u58f9\u3001\u5e8f\u8a00")
    add_test_paragraph(body, "\u4e00\u3001\u7814\u7a76\u76ee\u7684")
    add_test_paragraph(
        body,
        "\u8a3b\uff1a\u9019\u662f\u8aaa\u660e",
        style="NoteStyle",
        outline=5,
        ind_attrs={"left": "123", "leftChars": "99", "firstLineChars": "98"},
        tab_pos="1480",
        font_size_pt=14,
    )
    add_test_paragraph(
        body,
        "  \u8a3b1\uff1a\u9019\u662f\u8aaa\u660e",
        style="NoteStyle",
        num_id="9",
        ilvl=0,
        ind_attrs={"left": "223", "leftChars": "88", "firstLineChars": "87"},
        tab_pos="1580",
        font_size_pt=14,
    )
    add_test_paragraph(
        body,
        "\u8a3b\u4e00\uff1a\u9019\u662f\u8aaa\u660e",
        style="NoteStyle",
        ind_attrs={"left": "323", "leftChars": "77", "firstLineChars": "76"},
        tab_pos="1680",
        font_size_pt=14,
    )
    add_test_paragraph(
        body,
        "\u81ea\u52d5\u7de8\u865f\u8a3b\u89e3\u5167\u5bb9",
        num_id="10",
        ilvl=0,
        ind_attrs={"left": "423", "leftChars": "66", "firstLineChars": "65"},
        tab_pos="1780",
        font_size_pt=14,
    )
    add_test_paragraph(
        body,
        "\u6a23\u5f0f\u7de8\u865f\u8a3b\u89e3\u5167\u5bb9",
        style="StyleNoteNumbered",
        ind_attrs={"left": "523", "leftChars": "55", "firstLineChars": "54"},
        tab_pos="1880",
        font_size_pt=14,
    )
    add_test_paragraph(body, "\u9019\u662f\u666e\u901a 14pt \u5167\u6587", font_size_pt=14)
    add_test_paragraph(
        body,
        "\u4e0d\u662f\u8a3b\u89e3\u7684\u7de8\u865f\u5167\u5bb9",
        num_id="12",
        ilvl=0,
        font_size_pt=14,
    )

    tbl = etree.SubElement(body, qn("tbl"))
    tr = etree.SubElement(tbl, qn("tr"))
    tc = etree.SubElement(tr, qn("tc"))
    table_note = add_test_paragraph(tc, "\u8a3b\uff1a\u8868\u683c\u5167\u8aaa\u660e", font_size_pt=14)
    table_note_pPr = table_note.find("./w:pPr", NS)
    table_note_jc = etree.SubElement(table_note_pPr, qn("jc"))
    table_note_jc.set(qn("val"), "center")

    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_note_alignment_styles_xml() -> bytes:
    styles = etree.Element(qn("styles"), nsmap={"w": W_NS})

    centered_note_base = etree.SubElement(styles, qn("style"))
    centered_note_base.set(qn("type"), "paragraph")
    centered_note_base.set(qn("styleId"), "CenteredNoteBase")
    centered_note_name = etree.SubElement(centered_note_base, qn("name"))
    centered_note_name.set(qn("val"), "Centered Note Base")
    centered_note_pPr = etree.SubElement(centered_note_base, qn("pPr"))
    centered_note_jc = etree.SubElement(centered_note_pPr, qn("jc"))
    centered_note_jc.set(qn("val"), "center")

    plain_note = etree.SubElement(styles, qn("style"))
    plain_note.set(qn("type"), "paragraph")
    plain_note.set(qn("styleId"), "NoteStyle")
    plain_note_name = etree.SubElement(plain_note, qn("name"))
    plain_note_name.set(qn("val"), "Note Text")
    plain_note_based_on = etree.SubElement(plain_note, qn("basedOn"))
    plain_note_based_on.set(qn("val"), "CenteredNoteBase")

    numbered_note = etree.SubElement(styles, qn("style"))
    numbered_note.set(qn("type"), "paragraph")
    numbered_note.set(qn("styleId"), "StyleNoteNumbered")
    numbered_note_name = etree.SubElement(numbered_note, qn("name"))
    numbered_note_name.set(qn("val"), "Style Note Numbered")
    pPr = etree.SubElement(numbered_note, qn("pPr"))
    num_pr = etree.SubElement(pPr, qn("numPr"))
    ilvl = etree.SubElement(num_pr, qn("ilvl"))
    ilvl.set(qn("val"), "0")
    num_id = etree.SubElement(num_pr, qn("numId"))
    num_id.set(qn("val"), "11")

    return etree.tostring(styles, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_note_alignment_numbering_xml() -> bytes:
    numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    for num_id, abstract_id, lvl_text_value, lvl_jc_value in [
        ("10", "10", "\u8a3b%1\uff1a", "center"),
        ("11", "11", "\u8a3b %1", None),
        ("12", "12", "%1.", None),
    ]:
        abstract = etree.SubElement(numbering, qn("abstractNum"))
        abstract.set(qn("abstractNumId"), abstract_id)
        lvl = etree.SubElement(abstract, qn("lvl"))
        lvl.set(qn("ilvl"), "0")
        num_fmt = etree.SubElement(lvl, qn("numFmt"))
        num_fmt.set(qn("val"), "decimal")
        lvl_text = etree.SubElement(lvl, qn("lvlText"))
        lvl_text.set(qn("val"), lvl_text_value)
        if lvl_jc_value is not None:
            lvl_jc = etree.SubElement(lvl, qn("lvlJc"))
            lvl_jc.set(qn("val"), lvl_jc_value)
        pPr = etree.SubElement(lvl, qn("pPr"))
        make_ind(pPr, left="720", hanging="120")

        num = etree.SubElement(numbering, qn("num"))
        num.set(qn("numId"), num_id)
        abstract_ref = etree.SubElement(num, qn("abstractNumId"))
        abstract_ref.set(qn("val"), abstract_id)

    return etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_manual_heading_with_numpr_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    add_test_paragraph(body, "\u58f9\u3001\u5e8f\u8a00")
    add_test_paragraph(body, "\u58f9\u3001\u5e8f\u8a00")
    add_test_paragraph(
        body,
        "\u58f9\u3001\u5e8f\u8a0037",
        num_id="42",
        ilvl=0,
        ind_attrs={"left": "2279", "hanging": "420"},
        tab_pos="1200",
    )
    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def read_part_root(path: Path, part_name: str):
    with ZipFile(path, "r") as zf:
        return etree.fromstring(zf.read(part_name))


def read_document_root(path: Path):
    return read_part_root(path, "word/document.xml")


def find_paragraph_by_exact_text(root, text: str):
    for p in root.xpath(".//w:p", namespaces=NS):
        if "".join(p.xpath(".//w:t/text()", namespaces=NS)) == text:
            return p
    raise AssertionError(f"paragraph not found: {text}")


def paragraph_style_value(p) -> str | None:
    style = p.find("./w:pPr/w:pStyle", NS)
    return style.get(qn("val")) if style is not None else None


def paragraph_jc_value(p) -> str | None:
    jc = p.find("./w:pPr/w:jc", NS)
    return jc.get(qn("val")) if jc is not None else None


def paragraph_ppr_child_tags(p) -> list[str]:
    pPr = p.find("./w:pPr", NS)
    return [child.tag for child in pPr] if pPr is not None else []


def dirty_note_alignment_center_after_outline_in_docx(path: Path) -> None:
    temp_path = path.with_suffix(path.suffix + ".dirty_note_alignment.tmp")
    with ZipFile(path, "r") as zin, ZipFile(temp_path, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                root = etree.fromstring(data)
                for p in root.xpath(".//w:p[not(ancestor::w:tbl)]", namespaces=NS):
                    text = "".join(p.xpath(".//w:t/text()", namespaces=NS))
                    if not text.lstrip().startswith("\u8a3b"):
                        continue
                    pPr = p.find("./w:pPr", NS)
                    if pPr is None:
                        pPr = etree.SubElement(p, qn("pPr"))
                    jc = pPr.find("w:jc", NS)
                    if jc is not None:
                        pPr.remove(jc)
                    jc = etree.Element(qn("jc"))
                    jc.set(qn("val"), "center")
                    pPr.append(jc)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    temp_path.replace(path)


def dirty_numbering_suffix_tabs_in_docx(path: Path) -> None:
    temp_path = path.with_suffix(path.suffix + ".dirty_numbering.tmp")
    with ZipFile(path, "r") as zin, ZipFile(temp_path, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/numbering.xml":
                root = etree.fromstring(data)
                for lvl in root.xpath("./w:abstractNum/w:lvl | ./w:num/w:lvlOverride/w:lvl", namespaces=NS):
                    suff = lvl.find("w:suff", NS)
                    if suff is not None:
                        lvl.remove(suff)
                    lvl_text = lvl.find("w:lvlText", NS)
                    if lvl_text is not None:
                        lvl_text.set(qn("val"), (lvl_text.get(qn("val")) or "") + " ")
                    pPr = lvl.find("w:pPr", NS)
                    if pPr is None:
                        pPr = etree.SubElement(lvl, qn("pPr"))
                    tabs = pPr.find("w:tabs", NS)
                    if tabs is not None:
                        pPr.remove(tabs)
                    tabs = etree.SubElement(pPr, qn("tabs"))
                    tab = etree.SubElement(tabs, qn("tab"))
                    tab.set(qn("val"), "num")
                    tab.set(qn("pos"), "2061")
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    temp_path.replace(path)


def assert_no_character_indent_attrs(test_case: unittest.TestCase, root) -> None:
    for ind in root.xpath(".//w:ind", namespaces=NS):
        for attr in FORBIDDEN_ATTRS:
            test_case.assertIsNone(ind.get(qn(attr)), attr)


def assert_ind_has_no_character_indent_attrs(test_case: unittest.TestCase, ind) -> None:
    test_case.assertIsNotNone(ind)
    for attr in FORBIDDEN_ATTRS:
        test_case.assertIsNone(ind.get(qn(attr)), attr)


def assert_body_indent_hard_override(
    test_case: unittest.TestCase,
    paragraph,
    expected_left: str,
    expected_first_line: str = "0",
) -> None:
    ind = paragraph.find("./w:pPr/w:ind", NS)
    test_case.assertIsNotNone(ind)
    test_case.assertEqual(ind.get(qn("left")), expected_left)
    test_case.assertEqual(ind.get(qn("start")), expected_left)
    test_case.assertEqual(ind.get(qn("firstLine")), expected_first_line)
    test_case.assertEqual(ind.get(qn("hanging")), "0")
    test_case.assertEqual(ind.get(qn("leftChars")), "0")
    test_case.assertEqual(ind.get(qn("startChars")), "0")
    test_case.assertEqual(ind.get(qn("firstLineChars")), "0")
    test_case.assertEqual(ind.get(qn("hangingChars")), "0")
    test_case.assertIsNone(paragraph.find("./w:pPr/w:tabs", NS))
    test_case.assertIsNone(paragraph.find("./w:pPr/w:numPr", NS))


def assert_docx_has_no_character_indent_attrs(test_case: unittest.TestCase, path: Path) -> None:
    with ZipFile(path, "r") as zf:
        for name in zf.namelist():
            if not name.endswith(".xml"):
                continue
            root = etree.fromstring(zf.read(name))
            for ind in root.xpath(".//w:ind", namespaces=NS):
                for attr in FORBIDDEN_ATTRS:
                    test_case.assertIn(ind.get(qn(attr)), (None, "0"), f"{name}: {attr}")


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


def assert_toc_outlines_are_body(test_case: unittest.TestCase, path: Path) -> None:
    root = read_document_root(path)
    paragraphs = root.xpath(".//w:p", namespaces=NS)
    toc_ids = collect_all_toc_paragraph_ids(
        root,
        numbering_level_lookup={},
        style_numbering_lookup={},
        paragraphs=paragraphs,
    )
    test_case.assertGreater(len(toc_ids), 0)
    for p in paragraphs:
        if id(p) not in toc_ids:
            continue
        outline = p.find("./w:pPr/w:outlineLvl", NS)
        test_case.assertIsNotNone(outline)
        test_case.assertEqual(outline.get(qn("val")), "9")


class DocxProcessorTests(unittest.TestCase):
    def test_numbering_format_lookup_records_effective_level_source_and_child_order(self):
        numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})
        abstract = etree.SubElement(numbering, qn("abstractNum"))
        abstract.set(qn("abstractNumId"), "1")
        abstract_lvl = etree.SubElement(abstract, qn("lvl"))
        abstract_lvl.set(qn("ilvl"), "0")
        for tag in ("start", "numFmt", "suff", "lvlText", "lvlJc"):
            child = etree.SubElement(abstract_lvl, qn(tag))
            if tag == "numFmt":
                child.set(qn("val"), "decimal")
            elif tag == "suff":
                child.set(qn("val"), "tab")
            elif tag == "lvlText":
                child.set(qn("val"), "%1.")
        abstract_ppr = etree.SubElement(abstract_lvl, qn("pPr"))
        abstract_tabs = etree.SubElement(abstract_ppr, qn("tabs"))
        abstract_tab = etree.SubElement(abstract_tabs, qn("tab"))
        abstract_tab.set(qn("val"), "left")
        abstract_tab.set(qn("pos"), "111")
        etree.SubElement(abstract_ppr, qn("ind"))
        etree.SubElement(abstract_lvl, qn("rPr"))

        num_abstract = etree.SubElement(numbering, qn("num"))
        num_abstract.set(qn("numId"), "42")
        abstract_ref = etree.SubElement(num_abstract, qn("abstractNumId"))
        abstract_ref.set(qn("val"), "1")

        num_override = etree.SubElement(numbering, qn("num"))
        num_override.set(qn("numId"), "43")
        override_ref = etree.SubElement(num_override, qn("abstractNumId"))
        override_ref.set(qn("val"), "1")
        override = etree.SubElement(num_override, qn("lvlOverride"))
        override.set(qn("ilvl"), "0")
        override_lvl = etree.SubElement(override, qn("lvl"))
        override_lvl.set(qn("ilvl"), "0")
        for tag in ("start", "numFmt", "lvlText", "suff", "lvlJc"):
            child = etree.SubElement(override_lvl, qn(tag))
            if tag == "numFmt":
                child.set(qn("val"), "decimal")
            elif tag == "suff":
                child.set(qn("val"), "nothing")
            elif tag == "lvlText":
                child.set(qn("val"), "%9.")
        override_ppr = etree.SubElement(override_lvl, qn("pPr"))
        etree.SubElement(override_ppr, qn("ind"))
        override_tabs = etree.SubElement(override_ppr, qn("tabs"))
        override_tab = etree.SubElement(override_tabs, qn("tab"))
        override_tab.set(qn("val"), "left")
        override_tab.set(qn("pos"), "222")

        abstract_missing = etree.SubElement(numbering, qn("abstractNum"))
        abstract_missing.set(qn("abstractNumId"), "2")
        missing_lvl = etree.SubElement(abstract_missing, qn("lvl"))
        missing_lvl.set(qn("ilvl"), "0")
        missing_fmt = etree.SubElement(missing_lvl, qn("numFmt"))
        missing_fmt.set(qn("val"), "decimal")
        missing_text = etree.SubElement(missing_lvl, qn("lvlText"))
        missing_text.set(qn("val"), "%1.")
        num_missing = etree.SubElement(numbering, qn("num"))
        num_missing.set(qn("numId"), "44")
        missing_ref = etree.SubElement(num_missing, qn("abstractNumId"))
        missing_ref.set(qn("val"), "2")

        lookup = build_numbering_format_lookup(etree.tostring(numbering))

        self.assertEqual(lookup[("42", 0)]["level_source"], "abstractNum")
        self.assertEqual(lookup[("42", 0)]["suff"], "tab")
        self.assertEqual(lookup[("42", 0)]["tab_pos"], "111")
        self.assertEqual(lookup[("42", 0)]["lvl_child_order"], "start,numFmt,suff,lvlText,lvlJc,pPr,rPr")
        self.assertEqual(lookup[("42", 0)]["pPr_child_order"], "tabs,ind")
        self.assertTrue(lookup[("42", 0)]["suffix_before_lvlText"])
        self.assertTrue(lookup[("42", 0)]["tabs_before_ind"])

        self.assertEqual(lookup[("43", 0)]["level_source"], "lvlOverride")
        self.assertEqual(lookup[("43", 0)]["suff"], "nothing")
        self.assertEqual(lookup[("43", 0)]["tab_pos"], "222")
        self.assertEqual(lookup[("43", 0)]["lvl_child_order"], "start,numFmt,lvlText,suff,lvlJc,pPr")
        self.assertEqual(lookup[("43", 0)]["pPr_child_order"], "ind,tabs")
        self.assertFalse(lookup[("43", 0)]["suffix_before_lvlText"])
        self.assertFalse(lookup[("43", 0)]["tabs_before_ind"])

        self.assertEqual(lookup[("44", 0)]["level_source"], "abstractNum")
        self.assertEqual(lookup[("44", 0)]["pPr_child_order"], "none")
        self.assertFalse(lookup[("44", 0)]["tabs_before_ind"])

    def test_heading_suffix_records_include_settings_compat_flag_for_auto_numbering(self):
        cases = [
            ("present_no_val", make_settings_xml(do_not_use_indent_as_numbering_tab_stop=True), True),
            ("missing", None, False),
            ("explicit_zero", make_settings_xml(do_not_use_indent_as_numbering_tab_stop="0"), False),
        ]
        for _name, settings_xml, expected in cases:
            with self.subTest(case=_name):
                with tempfile.TemporaryDirectory() as temp_dir:
                    input_docx = Path(temp_dir) / "input.docx"
                    extra_parts = {"word/settings.xml": settings_xml} if settings_xml is not None else None
                    make_docx(
                        input_docx,
                        make_heading_suffix_document_xml(),
                        numbering_xml=make_heading_suffix_numbering_xml(),
                        extra_parts=extra_parts,
                    )

                    records = collect_heading_suffix_records_from_docx(input_docx)

                auto = next(record for record in records if record.get("source") == "auto_numbering_xml")
                self.assertEqual(auto["compat_doNotUseIndentAsNumberingTabStop"], expected)

    def test_collect_chapter_three_ids_uses_title_text_and_level_zero_when_prefix_unknown(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        before = add_test_paragraph(body, "\u58f9\u3001\u5e8f\u8a00")
        target = add_test_paragraph(
            body,
            "\u50f9\u683c\u5f62\u6210\u4e4b\u4e3b\u8981\u56e0\u7d20\u5206\u6790",
            num_id="9",
            ilvl=0,
        )
        protected_body = add_test_paragraph(body, "\u53c3\u7ae0\u5167\u6587")
        after = add_test_paragraph(body, "\u8086\u3001\u7b2c\u56db\u7ae0")

        skip_ids = collect_chapter_three_paragraph_ids(
            document,
            numbering_level_lookup={("9", 0): 0},
            numbering_format_lookup={},
            style_numbering_lookup={},
        )

        self.assertNotIn(id(before), skip_ids)
        self.assertIn(id(target), skip_ids)
        self.assertIn(id(protected_body), skip_ids)
        self.assertNotIn(id(after), skip_ids)

    def test_protected_region_context_collects_toc_chapter_numbering_and_styles(self):
        document = etree.Element(qn("document"), nsmap={"w": W_NS})
        body = etree.SubElement(document, qn("body"))
        toc = add_test_paragraph(body, "目錄項目", style="TOC1", num_id="1", ilvl=0)
        before = add_test_paragraph(body, "壹、序言")
        chapter = add_test_paragraph(
            body,
            "價格形成之主要因素分析",
            style="ChapterHeading",
        )
        protected_body = add_test_paragraph(body, "參章內文", style="ChapterBody")
        after = add_test_paragraph(body, "肆、第四章")

        styles = etree.Element(qn("styles"), nsmap={"w": W_NS})
        for style_id, name, num_id in (
            ("TOC1", "Table of Contents 1", "1"),
            ("ChapterHeading", "Chapter Heading", "2"),
            ("ChapterBody", "Chapter Body", None),
        ):
            style = etree.SubElement(styles, qn("style"))
            style.set(qn("type"), "paragraph")
            style.set(qn("styleId"), style_id)
            name_el = etree.SubElement(style, qn("name"))
            name_el.set(qn("val"), name)
            if num_id is not None:
                p_pr = etree.SubElement(style, qn("pPr"))
                num_pr = etree.SubElement(p_pr, qn("numPr"))
                ilvl_el = etree.SubElement(num_pr, qn("ilvl"))
                ilvl_el.set(qn("val"), "0")
                num_id_el = etree.SubElement(num_pr, qn("numId"))
                num_id_el.set(qn("val"), num_id)

        numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})
        for abstract_id, num_id in (("1", "1"), ("2", "2")):
            abstract = etree.SubElement(numbering, qn("abstractNum"))
            abstract.set(qn("abstractNumId"), abstract_id)
            lvl = etree.SubElement(abstract, qn("lvl"))
            lvl.set(qn("ilvl"), "0")
            num_fmt = etree.SubElement(lvl, qn("numFmt"))
            num_fmt.set(qn("val"), "ideographLegalTraditional")
            lvl_text = etree.SubElement(lvl, qn("lvlText"))
            lvl_text.set(qn("val"), "%1、")
            num = etree.SubElement(numbering, qn("num"))
            num.set(qn("numId"), num_id)
            abstract_ref = etree.SubElement(num, qn("abstractNumId"))
            abstract_ref.set(qn("val"), abstract_id)

        styles_xml = etree.tostring(styles)
        numbering_xml = etree.tostring(numbering)
        style_numbering_lookup = build_style_numbering_lookup(styles_xml)
        context = ProtectedRegionContext.from_document(
            document,
            protect_chapter_three=True,
            numbering_level_lookup=build_numbering_level_lookup(numbering_xml),
            numbering_format_lookup=build_numbering_format_lookup(numbering_xml),
            style_numbering_lookup=style_numbering_lookup,
            numbering_xml=numbering_xml,
        )

        self.assertIn(id(toc), context.document_toc_paragraph_ids)
        self.assertNotIn(id(before), context.document_chapter_three_paragraph_ids)
        self.assertIn(id(chapter), context.document_chapter_three_paragraph_ids)
        self.assertIn(id(protected_body), context.document_chapter_three_paragraph_ids)
        self.assertNotIn(id(after), context.document_chapter_three_paragraph_ids)
        self.assertIn(("1", 0), context.toc_numbering_pairs)
        self.assertIn("1", context.toc_num_ids)
        self.assertIn("1", context.toc_abstract_ids)
        self.assertIn(("2", 0), context.chapter_three_numbering_pairs)
        self.assertIn("2", context.chapter_three_num_ids)
        self.assertIn("2", context.chapter_three_abstract_ids)
        self.assertEqual(context.chapter_three_style_ids, {"ChapterHeading", "ChapterBody"})

    def test_skip_chapter_three_indents_restores_heading_outline_but_skips_formatting_and_word_com(self):
        with tempfile.TemporaryDirectory() as tmp:
            input_docx = Path(tmp) / "input.docx"
            output_docx = Path(tmp) / "output.docx"
            make_docx(input_docx, make_chapter_three_skip_document_xml())

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=True,
                    fix_color=True,
                    fix_paragraph=True,
                    remove_all_outline_levels=True,
                    normalize_with_word_com=False,
                    skip_chapter_three_tables=True,
                    skip_chapter_three_indents=True,
                ),
            )

            root = read_document_root(output_docx)
            paragraphs = root.xpath(".//w:p[not(ancestor::w:tbl)]", namespaces=NS)
            chapter_heading_ind = paragraphs[2].find("./w:pPr/w:ind", NS)
            chapter_child_ind = paragraphs[3].find("./w:pPr/w:ind", NS)
            chapter_body_ind = paragraphs[4].find("./w:pPr/w:ind", NS)
            after_chapter_body_ind = paragraphs[6].find("./w:pPr/w:ind", NS)

            self.assertEqual(paragraphs[2].find("./w:pPr/w:outlineLvl", NS).get(qn("val")), "0")
            self.assertEqual(paragraphs[3].find("./w:pPr/w:outlineLvl", NS).get(qn("val")), "1")
            self.assertEqual(paragraphs[4].find("./w:pPr/w:outlineLvl", NS).get(qn("val")), "9")
            self.assertEqual(chapter_heading_ind.get(qn("left")), "321")
            self.assertEqual(chapter_child_ind.get(qn("left")), "654")
            self.assertEqual(chapter_body_ind.get(qn("left")), "777")
            self.assertEqual(chapter_heading_ind.get(qn("leftChars")), "111")
            self.assertEqual(chapter_heading_ind.get(qn("firstLineChars")), "222")
            self.assertEqual(chapter_child_ind.get(qn("leftChars")), "333")
            self.assertEqual(chapter_child_ind.get(qn("firstLineChars")), "444")
            self.assertEqual(chapter_body_ind.get(qn("leftChars")), "555")
            self.assertEqual(chapter_body_ind.get(qn("firstLineChars")), "666")

            self.assertEqual(paragraphs[5].find("./w:pPr/w:outlineLvl", NS).get(qn("val")), "0")
            self.assertEqual(after_chapter_body_ind.get(qn("leftChars")), "0")
            self.assertEqual(after_chapter_body_ind.get(qn("firstLineChars")), "0")
            self.assertTrue(any(
                record.get("text_preview") == "\u8086\u7ae0\u666e\u901a\u5167\u6587"
                for record in summary.body_indent_records
            ))
            self.assertFalse(any(
                "\u53c3\u7ae0" in str(record.get("text_preview"))
                for record in summary.body_indent_records
            ))

            tables = root.xpath(".//w:tbl", namespaces=NS)
            self.assertIsNone(tables[0].find("w:tblPr", NS))
            self.assertEqual(summary.table_log_records[0]["table_type"], "skipped_first_table")
            self.assertEqual(summary.table_log_records[0]["action"], "skipped")
            self.assertEqual(
                summary.table_log_records[0]["reason"],
                "first table in word/document.xml",
            )

            joined_logs = "\n".join(summary.paragraph_logs + summary.numbering_xml_logs)
            self.assertIn("CHAPTER_THREE_SKIP_IDS collected=", joined_logs)
            self.assertNotIn("CHAPTER_THREE_SKIP_IDS collected=0", joined_logs)
            self.assertIn(
                "skipped chapter \u53c3\u3001\u50f9\u683c\u5f62\u6210\u4e4b\u4e3b\u8981\u56e0\u7d20\u5206\u6790 content; no formatting applied",
                joined_logs,
            )
            self.assertIn("restored protected chapter outline level only", joined_logs)
            self.assertIn("CHAR_INDENT_SANITIZE_SKIP_EXCLUDED", joined_logs)

    def test_character_indent_sanitize_cleans_non_toc_when_chapter_three_indent_skip_disabled(self):
        with tempfile.TemporaryDirectory() as tmp:
            input_docx = Path(tmp) / "input.docx"
            output_docx = Path(tmp) / "output.docx"
            make_docx(
                input_docx,
                make_shared_chapter_three_indent_document_xml(),
                styles_xml=make_shared_chapter_three_indent_styles_xml(),
                numbering_xml=make_shared_chapter_three_indent_numbering_xml(),
            )

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=False,
                    normalize_with_word_com=False,
                    skip_chapter_three_indents=False,
                ),
            )

            output_document_root = read_document_root(output_docx)
            output_styles_root = read_part_root(output_docx, "word/styles.xml")
            output_numbering_root = read_part_root(output_docx, "word/numbering.xml")
            paragraphs = output_document_root.xpath(".//w:p", namespaces=NS)

            toc_paragraph_ind = paragraphs[0].find("./w:pPr/w:ind", NS)
            self.assertEqual(toc_paragraph_ind.get(qn("leftChars")), "701")
            self.assertEqual(toc_paragraph_ind.get(qn("firstLineChars")), "702")
            for paragraph in paragraphs[1:]:
                assert_ind_has_no_character_indent_attrs(
                    self,
                    paragraph.find("./w:pPr/w:ind", NS),
                )

            toc_style_ind = output_styles_root.xpath(
                "./w:style[@w:styleId='TOC1']/w:pPr/w:ind",
                namespaces=NS,
            )[0]
            self.assertEqual(toc_style_ind.get(qn("leftChars")), "333")
            self.assertEqual(toc_style_ind.get(qn("firstLineChars")), "555")
            shared_style_ind = output_styles_root.xpath(
                "./w:style[@w:styleId='SharedPara']/w:pPr/w:ind",
                namespaces=NS,
            )[0]
            self.assertEqual(shared_style_ind.get(qn("left")), "720")
            assert_ind_has_no_character_indent_attrs(self, shared_style_ind)

            toc_numbering_ind = output_numbering_root.xpath(
                "./w:abstractNum[@w:abstractNumId='7']/w:lvl/w:pPr/w:ind",
                namespaces=NS,
            )[0]
            self.assertEqual(toc_numbering_ind.get(qn("leftChars")), "777")
            self.assertEqual(toc_numbering_ind.get(qn("firstLineChars")), "555")
            shared_numbering_ind = output_numbering_root.xpath(
                "./w:abstractNum[@w:abstractNumId='1']/w:lvl/w:pPr/w:ind",
                namespaces=NS,
            )[0]
            self.assertEqual(shared_numbering_ind.get(qn("left")), "1440")
            assert_ind_has_no_character_indent_attrs(self, shared_numbering_ind)

            self.assertGreater(summary.character_indent_attrs_removed, 0)

    def test_skip_chapter_three_indents_only_excludes_document_chapter_paragraphs_from_character_indent_sanitize(self):
        with tempfile.TemporaryDirectory() as tmp:
            input_docx = Path(tmp) / "input.docx"
            output_docx = Path(tmp) / "output.docx"
            make_docx(
                input_docx,
                make_shared_chapter_three_indent_document_xml(),
                styles_xml=make_shared_chapter_three_indent_styles_xml(),
                numbering_xml=make_shared_chapter_three_indent_numbering_xml(),
            )

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=False,
                    normalize_with_word_com=False,
                    skip_chapter_three_indents=True,
                ),
            )

            output_document_root = read_document_root(output_docx)
            output_styles_root = read_part_root(output_docx, "word/styles.xml")
            output_numbering_root = read_part_root(output_docx, "word/numbering.xml")
            paragraphs = output_document_root.xpath(".//w:p", namespaces=NS)

            toc_paragraph_ind = paragraphs[0].find("./w:pPr/w:ind", NS)
            self.assertEqual(toc_paragraph_ind.get(qn("leftChars")), "701")
            self.assertEqual(toc_paragraph_ind.get(qn("firstLineChars")), "702")

            assert_ind_has_no_character_indent_attrs(self, paragraphs[1].find("./w:pPr/w:ind", NS))

            chapter_heading_ind = paragraphs[2].find("./w:pPr/w:ind", NS)
            chapter_body_ind = paragraphs[3].find("./w:pPr/w:ind", NS)
            self.assertEqual(chapter_heading_ind.get(qn("leftChars")), "111")
            self.assertEqual(chapter_heading_ind.get(qn("firstLineChars")), "222")
            self.assertEqual(chapter_body_ind.get(qn("leftChars")), "555")
            self.assertEqual(chapter_body_ind.get(qn("firstLineChars")), "666")

            assert_ind_has_no_character_indent_attrs(self, paragraphs[4].find("./w:pPr/w:ind", NS))
            assert_ind_has_no_character_indent_attrs(self, paragraphs[5].find("./w:pPr/w:ind", NS))

            shared_style_ind = output_styles_root.xpath(
                "./w:style[@w:styleId='SharedPara']/w:pPr/w:ind",
                namespaces=NS,
            )[0]
            self.assertEqual(shared_style_ind.get(qn("left")), "720")
            assert_ind_has_no_character_indent_attrs(self, shared_style_ind)

            shared_numbering_ind = output_numbering_root.xpath(
                "./w:abstractNum[@w:abstractNumId='1']/w:lvl/w:pPr/w:ind",
                namespaces=NS,
            )[0]
            self.assertEqual(shared_numbering_ind.get(qn("left")), "1440")
            assert_ind_has_no_character_indent_attrs(self, shared_numbering_ind)

            toc_style_ind = output_styles_root.xpath(
                "./w:style[@w:styleId='TOC1']/w:pPr/w:ind",
                namespaces=NS,
            )[0]
            self.assertEqual(toc_style_ind.get(qn("leftChars")), "333")
            self.assertEqual(toc_style_ind.get(qn("firstLineChars")), "555")
            toc_numbering_ind = output_numbering_root.xpath(
                "./w:abstractNum[@w:abstractNumId='7']/w:lvl/w:pPr/w:ind",
                namespaces=NS,
            )[0]
            self.assertEqual(toc_numbering_ind.get(qn("leftChars")), "777")
            self.assertEqual(toc_numbering_ind.get(qn("firstLineChars")), "555")

            joined_logs = "\n".join(summary.numbering_xml_logs)
            self.assertIn("CHAPTER_THREE_SKIP_IDS collected=2", joined_logs)
            self.assertIn("CHAR_INDENT_SANITIZE_SKIP_EXCLUDED", joined_logs)
            self.assertNotIn("CHAR_INDENT_SANITIZE_SKIP_EXCLUDED_STYLE", joined_logs)

    def test_manual_heading_with_existing_numpr_removes_list_numbering_tabs_and_keeps_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            input_docx = Path(tmp) / "input.docx"
            output_docx = Path(tmp) / "output.docx"
            make_docx(
                input_docx,
                make_manual_heading_with_numpr_document_xml(),
                numbering_xml=make_heading_suffix_numbering_xml(
                    suffix="tab",
                    include_tabs=True,
                    lvl_text_value="%1. \t",
                    tab_pos="1200",
                ),
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

            root = read_document_root(output_docx)
            paragraphs = root.xpath(".//w:p[not(ancestor::w:tbl)]", namespaces=NS)
            target = paragraphs[2]
            self.assertEqual("".join(target.xpath(".//w:t/text()", namespaces=NS)), "\u58f9\u3001\u5e8f\u8a0037")
            self.assertIsNone(target.find("./w:pPr/w:numPr", NS))
            self.assertIsNone(target.find("./w:pPr/w:tabs", NS))

            before_record = next(
                record for record in summary.heading_suffix_before_records
                if record.get("heading_text") == "\u58f9\u3001\u5e8f\u8a0037"
            )
            after_record = next(
                record for record in summary.heading_suffix_after_records
                if record.get("heading_text") == "\u58f9\u3001\u5e8f\u8a0037"
            )
            self.assertEqual(before_record["source"], "manual_text")
            self.assertTrue(before_record["paragraph_has_numPr"])
            self.assertNotEqual(before_record["paragraph_tabs"], "none")
            self.assertEqual(before_record["numbering_suff"], "tab")
            self.assertEqual(after_record["source"], "manual_text")
            self.assertFalse(after_record["paragraph_has_numPr"])
            self.assertEqual(after_record["paragraph_tabs"], "none")
            self.assertEqual(after_record["raw_separator_repr"], "''")
            self.assertEqual(after_record["tab_count"], 0)

            numbering_root = read_part_root(output_docx, "word/numbering.xml")
            for lvl in numbering_root.xpath("./w:abstractNum/w:lvl | ./w:num/w:lvlOverride/w:lvl", namespaces=NS):
                # The only numbering level is decimal "%1." -> outline level 3,
                # a space-suffix level.
                assert_numbering_level_follows_suffix_rule(self, lvl, 3)
                self.assertFalse(lvl.find("w:lvlText", NS).get(qn("val")).endswith((" ", "\t", "\u3000")))

    def test_chapter_three_table_and_indent_options_are_independent(self):
        cases = [
            (True, True, True, "skipped_chapter_three_table", False, False, True),
            (True, False, False, "color_only_table", False, True, False),
            (False, True, True, "special_table", True, False, True),
            (False, False, False, "special_table", True, True, False),
        ]

        for (
            skip_table_layout,
            skip_table_color,
            skip_indents,
            expected_table_type,
            expected_layout_fixed,
            expected_color_fixed,
            expect_indent_skipped,
        ) in cases:
            with self.subTest(
                skip_table_layout=skip_table_layout,
                skip_table_color=skip_table_color,
                skip_indents=skip_indents,
            ):
                with tempfile.TemporaryDirectory() as tmp:
                    input_docx = Path(tmp) / "input.docx"
                    output_docx = Path(tmp) / "output.docx"
                    make_docx(input_docx, make_chapter_three_options_document_xml())

                    summary = fix_docx_fast(
                        input_docx,
                        output_docx,
                        ProcessOptions(
                            fix_table_layout=True,
                            fix_color=True,
                            fix_paragraph=True,
                            remove_all_outline_levels=True,
                            normalize_with_word_com=False,
                            skip_chapter_three_table_layout=skip_table_layout,
                            skip_chapter_three_table_color=skip_table_color,
                            skip_chapter_three_indents=skip_indents,
                        ),
                    )

                    root = read_document_root(output_docx)
                    paragraphs = root.xpath(".//w:p[not(ancestor::w:tbl)]", namespaces=NS)
                    tables = root.xpath(".//w:tbl", namespaces=NS)
                    chapter_heading_ind = paragraphs[2].find("./w:pPr/w:ind", NS)
                    chapter_child_ind = paragraphs[3].find("./w:pPr/w:ind", NS)
                    chapter_body_ind = paragraphs[4].find("./w:pPr/w:ind", NS)

                    self.assertEqual(summary.table_log_records[1]["table_type"], expected_table_type)
                    self.assertEqual(summary.table_log_records[1]["layout_fixed"], expected_layout_fixed)
                    self.assertEqual(summary.table_log_records[1]["color_fixed"], expected_color_fixed)
                    self.assertEqual(
                        summary.table_log_records[1]["chapter_three_table_layout_skipped"],
                        skip_table_layout,
                    )
                    self.assertEqual(
                        summary.table_log_records[1]["chapter_three_table_color_skipped"],
                        skip_table_color,
                    )
                    if not expected_layout_fixed:
                        self.assertIsNone(tables[1].find("w:tblPr", NS))
                    else:
                        self.assertIsNotNone(tables[1].find("w:tblPr", NS))

                    self.assertEqual(paragraphs[2].find("./w:pPr/w:outlineLvl", NS).get(qn("val")), "0")
                    self.assertEqual(paragraphs[3].find("./w:pPr/w:outlineLvl", NS).get(qn("val")), "1")

                    if expect_indent_skipped:
                        self.assertEqual(chapter_heading_ind.get(qn("left")), "321")
                        self.assertEqual(chapter_child_ind.get(qn("left")), "654")
                        self.assertEqual(chapter_body_ind.get(qn("left")), "777")
                        self.assertEqual(chapter_heading_ind.get(qn("leftChars")), "111")
                        self.assertEqual(chapter_body_ind.get(qn("firstLineChars")), "666")
                        self.assertFalse(any(
                            "\u53c3\u7ae0" in str(record.get("text_preview"))
                            for record in summary.body_indent_records
                        ))
                    else:
                        self.assertEqual(chapter_heading_ind.get(qn("left")), TEMPLATE_OUTLINE_INDENTS[0]["left"])
                        self.assertEqual(chapter_child_ind.get(qn("left")), TEMPLATE_OUTLINE_INDENTS[1]["left"])
                        self.assertEqual(chapter_body_ind.get(qn("left")), TEMPLATE_OUTLINE_INDENTS[1]["body_left"])
                        self.assertIsNone(chapter_heading_ind.get(qn("leftChars")))
                        self.assertEqual(chapter_body_ind.get(qn("firstLineChars")), "0")
                        self.assertTrue(any(
                            "\u53c3\u7ae0" in str(record.get("text_preview"))
                            for record in summary.body_indent_records
                        ))

    def test_final_numbering_cleanup_protects_chapter_three_shared_level_precisely(self):
        # \u53c3 (numId 42) and \u8086 (numId 99) share abstractNum 1 at the SAME ilvl 0.
        # \u300c\u53c3\u3001\u4e0d\u8981\u6e05\u7406\u7de8\u865f\u5f8c\u7db4 tab/space\u300d defaults on, and chapter \u53c3 protection
        # is expressed as a precise (numId/abstract, ilvl) \u2014 never as the whole
        # abstractNumId. The single shared level \u53c3 actually uses is therefore
        # preserved (\u8086 shares that exact lvl element, so it is collaterally kept).
        with tempfile.TemporaryDirectory() as tmp:
            input_docx = Path(tmp) / "input.docx"
            output_docx = Path(tmp) / "output.docx"
            make_docx(
                input_docx,
                make_chapter_three_shared_numbering_document_xml(),
                numbering_xml=make_shared_chapter_three_numbering_xml(),
            )

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=False,
                    normalize_with_word_com=False,
                    skip_chapter_three_indents=True,
                ),
            )

            numbering_root = read_part_root(output_docx, "word/numbering.xml")
            protected_lvl = numbering_root.xpath(
                "./w:abstractNum[@w:abstractNumId='1']/w:lvl",
                namespaces=NS,
            )[0]
            # The level \u53c3 uses keeps its original suffix / tab / lvlText.
            self.assertEqual(protected_lvl.find("w:suff", NS).get(qn("val")), "tab")
            self.assertIsNotNone(protected_lvl.find("./w:pPr/w:tabs", NS))
            self.assertEqual(protected_lvl.find("w:lvlText", NS).get(qn("val")), "%1\u3001 ")

            logs = "\n".join(summary.numbering_xml_logs)
            self.assertIn("CHAPTER_THREE_SKIP_IDS collected=2", logs)
            # The whole abstractNumId is NOT used to protect chapter \u53c3.
            self.assertIn("CHAPTER_THREE_NUMBERING_SUFFIX_CLEANUP_SKIP enabled=true", logs)
            self.assertIn("protected_abstract_levels=1:0", logs)
            self.assertIn("protected_abstractIds_not_used_for_chapter_three=true", logs)

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
            "?臬??蹓??捍\r",
            "????瘣菟??????\r",
            "???3???\r",
        ]
        direct_record = {
            "paragraph_index": 3,
            "text_preview": "???3???",
        }
        fallback_record = {
            "paragraph_index": 2,
            "text_preview": "???3???",
        }

        self.assertEqual(find_word_paragraph_index_for_record(paragraph_texts, direct_record), 3)
        self.assertEqual(find_word_paragraph_index_for_record(paragraph_texts, fallback_record), 3)

    def test_find_word_paragraph_index_uses_match_prefix_and_strips_preview_ellipsis(self):
        paragraph_texts = [
            "A very long paragraph prefix that continues beyond the display preview and has no ellipsis\r",
        ]
        prefix_record = {
            "paragraph_index": 1,
            "text_preview": "wrong preview...",
            "text_match_prefix": "A very long paragraph prefix that continues",
        }
        legacy_preview_record = {
            "paragraph_index": 1,
            "text_preview": "A very long paragraph prefix...",
        }

        self.assertEqual(find_word_paragraph_index_for_record(paragraph_texts, prefix_record), 1)
        self.assertEqual(find_word_paragraph_index_for_record(paragraph_texts, legacy_preview_record), 1)

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
            with patch("docx_fixer.word_com_indent.run_powershell_file", side_effect=fake_run) as runner:
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
        self.assertTrue(captured["script_text"].lstrip().startswith("param("))
        self.assertLess(
            captured["script_text"].index("param("),
            captured["script_text"].index("$utf8NoBom = New-Object System.Text.UTF8Encoding($false)"),
        )
        self.assertIn("$utf8NoBom = New-Object System.Text.UTF8Encoding($false)", captured["script_text"])
        self.assertIn("[Console]::OutputEncoding = $utf8NoBom", captured["script_text"])
        self.assertIn("$RecordsPath", captured["script_text"])
        self.assertIn("Add-Log", captured["script_text"])
        self.assertIn("Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value $msg", captured["script_text"])
        self.assertNotIn("Write-Output $msg", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_BEGIN", captured["script_text"])
        self.assertIn("$matchPrefix = [string]$record.text_match_prefix", captured["script_text"])
        self.assertIn("if ([string]::IsNullOrWhiteSpace($matchPrefix))", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_MATCH_PREFIX", captured["script_text"])
        self.assertIn("$preview.EndsWith('...')", captured["script_text"])
        self.assertIn("$normalizedPreview.EndsWith('...')", captured["script_text"])
        self.assertIn("Paragraph-TextMatchesPreview $candidateText $matchPrefix", captured["script_text"])
        self.assertNotIn("Paragraph-TextMatchesPreview $candidateText $preview", captured["script_text"])
        self.assertIn("WORD_COM_PS_RECORD_LOOP_BEGIN", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_DIRECT_MATCHED", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_FALLBACK_MATCHED", captured["script_text"])
        self.assertIn("WORD_COM_FALLBACK_SCAN_PROGRESS", captured["script_text"])
        self.assertIn("$paragraphIndexOffset = $null", captured["script_text"])
        self.assertIn("WORD_COM_PARAGRAPH_INDEX_OFFSET_LEARNED", captured["script_text"])
        self.assertIn("WORD_COM_PARAGRAPH_INDEX_OFFSET_CHANGED", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_TRY_OFFSET", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_OFFSET_MATCHED", captured["script_text"])
        self.assertIn("function Find-MatchingParagraphInWindow", captured["script_text"])
        self.assertIn("WORD_COM_LOCAL_SCAN_BEGIN", captured["script_text"])
        self.assertIn("WORD_COM_LOCAL_SCAN_PROGRESS", captured["script_text"])
        self.assertIn("WORD_COM_LOCAL_SCAN_BEGIN record={0} source={1} start={2} end={3}\" -f $recordIndex, $source, $start, $end)) | Out-Null", captured["script_text"])
        self.assertIn("WORD_COM_LOCAL_SCAN_PROGRESS record={0} source={1} current={2} start={3} end={4}\" -f $recordIndex, $source, $j, $start, $end)) | Out-Null", captured["script_text"])
        self.assertIn("return [int]$j", captured["script_text"])
        self.assertIn("return $null", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_LOCAL_MATCHED", captured["script_text"])
        self.assertIn("$localMatch = Find-MatchingParagraphInWindow", captured["script_text"])
        self.assertNotIn("[int]$localMatch = Find-MatchingParagraphInWindow", captured["script_text"])
        self.assertIn("$lastMatchedWordIndex = $null", captured["script_text"])
        self.assertIn("$lastMatchedWordIndex = [int]$matchIndex", captured["script_text"])
        self.assertIn("source=offset_window", captured["script_text"])
        self.assertIn("source=target_window", captured["script_text"])
        self.assertIn("source=last_match_window", captured["script_text"])
        self.assertIn("WORD_COM_FULL_FALLBACK_SCAN_BEGIN", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_MATCHED i={0} word_index={1} source={2}", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_BEFORE_GET_PARAGRAPH", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_AFTER_GET_PARAGRAPH", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_BEFORE_FONT_CHECK", captured["script_text"])
        self.assertIn("WORD_COM_RECORD_AFTER_FONT_CHECK", captured["script_text"])
        self.assertIn("WORD_COM_FONT_CHECK_ONLY_STARTED", captured["script_text"])
        self.assertIn("WORD_COM_FONT_CHECK_APPROVED:", captured["script_text"])
        self.assertIn("WORD_COM_APPROVED_RECORD_JSON", captured["script_text"])
        self.assertIn("WORD_COM_FONT_CHECK_SUMMARY", captured["script_text"])
        self.assertIn("decision=approved_for_xml_body_indent", captured["script_text"])
        self.assertIn("decision=skipped_word_font_not_14", captured["script_text"])
        self.assertIn("$status = 'approved'", captured["script_text"])
        self.assertIn("function Get-ParagraphDiagnostic", captured["script_text"])
        self.assertIn("$diag = Get-ParagraphDiagnostic $paragraph", captured["script_text"])
        self.assertIn("word_com_LeftIndent_cm=", captured["script_text"])
        self.assertIn("word_com_CharacterUnitLeftIndent=", captured["script_text"])
        self.assertIn("word_com_TabStops_Count=", captured["script_text"])
        self.assertIn("word_com_Style_NameLocal=", captured["script_text"])
        self.assertIn("section_left_margin_cm=", captured["script_text"])
        self.assertIn("absolute_text_start_cm=", captured["script_text"])
        self.assertIn("final_left_cm=not_read", captured["script_text"])
        self.assertNotIn("WORD_COM_RECORD_BEFORE_GET_FORMAT", captured["script_text"])
        self.assertNotIn("WORD_COM_RECORD_AFTER_GET_FORMAT", captured["script_text"])
        self.assertNotIn("WORD_COM_RECORD_CHAR_UNIT_CLEAR_SKIPPED", captured["script_text"])
        self.assertNotIn("WORD_COM_RECORD_BEFORE_SET_INDENTS", captured["script_text"])
        self.assertNotIn("WORD_COM_RECORD_AFTER_SET_INDENTS", captured["script_text"])
        self.assertNotIn("$pf.LeftIndent = $expectedLeftPoints", captured["script_text"])
        self.assertNotIn("$pf.FirstLineIndent = $expectedFirstLinePoints", captured["script_text"])
        self.assertNotIn("CharacterUnitLeftIndent = 0", captured["script_text"])
        self.assertNotIn("CharacterUnitFirstLineIndent = 0", captured["script_text"])
        self.assertNotIn("WORD_COM_RECORD_BEFORE_CLEAR_CHAR_INDENTS", captured["script_text"])
        self.assertNotIn("WORD_COM_RECORD_AFTER_CLEAR_CHAR_INDENTS", captured["script_text"])
        self.assertNotIn("WORD_COM_RECORD_CHAR_LEFT_CLEAR_FAILED", captured["script_text"])
        self.assertNotIn("WORD_COM_RECORD_CHAR_FIRST_CLEAR_FAILED", captured["script_text"])
        self.assertNotIn("WORD_COM_RECORD_BEFORE_READ_INDENTS", captured["script_text"])
        self.assertNotIn("WORD_COM_RECORD_AFTER_READ_INDENTS", captured["script_text"])
        self.assertNotIn("$beforeLeftPt = [double]$pf.LeftIndent", captured["script_text"])
        self.assertNotIn("$beforeFirstLinePt = [double]$pf.FirstLineIndent", captured["script_text"])
        self.assertNotIn("$actualLeftPt = [double]$pf.LeftIndent", captured["script_text"])
        self.assertNotIn("$actualFirstLinePt = [double]$pf.FirstLineIndent", captured["script_text"])
        self.assertNotIn("New-Object System.Collections.Generic.List[string]", captured["script_text"])
        self.assertNotIn("$paragraphs", captured["script_text"])
        self.assertEqual(WORD_COM_TIMEOUT_SECONDS, 600)
        self.assertEqual(runner.call_args.kwargs["timeout"], WORD_COM_TIMEOUT_SECONDS)
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
                "docx_fixer.word_com_indent.run_powershell_file",
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
                "docx_fixer.word_com_indent.run_powershell_file",
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
                "docx_fixer.word_com_indent.run_powershell_file",
                side_effect=fake_run,
            ):
                logs = verify_and_fix_body_indents_with_word_com(
                    output_docx,
                    [{"paragraph_index": 1, "text_preview": "body text", "expected_left_cm": 1.0, "expected_left_points": 28.3464567}],
                )

        self.assertTrue(any("WORD_COM_PS_EXCEPTION System.Exception:boom" in line for line in logs))
        self.assertTrue(any("WORD_COM_BODY_INDENT_FIX_FAILED_AFTER_PARTIAL_LOGS" in line for line in logs))
        self.assertTrue(any("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_script_failed" in line for line in logs))

    def test_word_com_body_indent_requires_done_marker_before_copy_back(self):
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
                            "WORD_COM_RECORD_BEGIN i=1 paragraph_index=1 expected_left_cm=1.0 text=body text",
                            "WORD_COM_RECORD_MATCHED i=1 word_index=1",
                            "WORD_COM_INDENT_VERIFY: paragraph_index=1; matched_paragraph_index=1; decision=apply_body_indent; status=ok",
                        ]
                    ),
                    encoding="utf-8",
                )
                return types.SimpleNamespace(stdout="", stderr="partial failure", returncode=0)

            with patch(
                "docx_fixer.word_com_indent.run_powershell_file",
                side_effect=fake_run,
            ):
                logs = verify_and_fix_body_indents_with_word_com(
                    output_docx,
                    [{"paragraph_index": 1, "text_preview": "body text", "expected_left_cm": 1.0, "expected_left_points": 28.3464567}],
                )

        self.assertTrue(any("WORD_COM_RECORD_BEGIN" in line for line in logs))
        self.assertTrue(any("WORD_COM_RECORD_MATCHED" in line for line in logs))
        self.assertTrue(any("WORD_COM_INDENT_VERIFY:" in line for line in logs))
        self.assertTrue(any("WORD_COM_BODY_INDENT_FIX_FAILED_AFTER_PARTIAL_LOGS" in line for line in logs))
        self.assertTrue(any("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_script_failed" in line for line in logs))
        self.assertFalse(any("WORD_COM_PS_DONE" in line for line in logs))

    def test_word_com_body_indent_partial_logs_without_error_are_timeout_or_interrupted(self):
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
                            "WORD_COM_PS_WORD_CREATED",
                            "WORD_COM_PS_DOC_OPENED",
                            "WORD_COM_PS_BEFORE_LOOP",
                            "WORD_COM_PS_PARAGRAPHS_COUNT count=6204",
                            "WORD_COM_PS_RECORD_LOOP_BEGIN count=336",
                        ]
                    ),
                    encoding="utf-8",
                )
                return types.SimpleNamespace(stdout="", stderr="", returncode=1)

            with patch(
                "docx_fixer.word_com_indent.run_powershell_file",
                side_effect=fake_run,
            ):
                logs = verify_and_fix_body_indents_with_word_com(
                    output_docx,
                    [{"paragraph_index": 1, "text_preview": "body text", "expected_left_cm": 1.0, "expected_left_points": 28.3464567}],
                )

        self.assertTrue(any("WORD_COM_BODY_INDENT_FIX_FAILED_AFTER_PARTIAL_LOGS" in line for line in logs))
        self.assertTrue(any("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_interrupted_or_timeout" in line for line in logs))
        self.assertTrue(any("timeout_seconds=" in line for line in logs))
        self.assertTrue(any("last_partial_log=WORD_COM_PS_RECORD_LOOP_BEGIN count=336" in line for line in logs))
        self.assertFalse(any("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_script_failed" in line for line in logs))

    def test_word_com_body_indent_record_filter_keeps_only_font_check_records(self):
        records = [
            {"kind": "auto", "apply_only_if_word_font_size_is_14": False},
            {"kind": "manual"},
            {"kind": "body", "apply_only_if_word_font_size_is_14": False},
            {"kind": "body_font_check", "apply_only_if_word_font_size_is_14": True, "xml_font_size": 10},
            {"kind": "body_font_check", "apply_only_if_word_font_size_is_14": True, "xml_font_size": 11},
            {"kind": "body_font_check", "apply_only_if_word_font_size_is_14": True, "xml_font_size": 12},
            {"kind": "body_font_check", "apply_only_if_word_font_size_is_14": True, "xml_font_size": None},
        ]

        filtered = _filter_word_com_body_indent_records(records)

        self.assertEqual([record.get("xml_font_size") for record in filtered], [12, None])

    def test_apply_word_com_approved_body_indents_updates_document_xml(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(output_docx, make_document_with_character_indent(text="Approved body paragraph"))

            logs = apply_word_com_approved_body_indents_to_docx_xml(
                output_docx,
                [
                    {
                        "paragraph_index": 1,
                        "expected_left_twips": "1440",
                        "expected_first_line_twips": "560",
                        "text_match_prefix": "Approved body paragraph",
                    }
                ],
            )

            root = read_document_root(output_docx)
            ind = root.find(".//w:ind", NS)
            self.assertEqual(ind.get(qn("left")), "1440")
            self.assertEqual(ind.get(qn("start")), "1440")
            self.assertEqual(ind.get(qn("firstLine")), "560")
            self.assertEqual(ind.get(qn("hanging")), "0")
            self.assertEqual(ind.get(qn("leftChars")), "0")
            self.assertEqual(ind.get(qn("startChars")), "0")
            self.assertEqual(ind.get(qn("firstLineChars")), "0")
            self.assertEqual(ind.get(qn("hangingChars")), "0")
            joined_logs = "\n".join(logs)
            self.assertIn("WORD_COM_XML_APPLY_STARTED approved_records=1", joined_logs)
            self.assertIn("WORD_COM_XML_APPLY_RECORD paragraph_index=1 status=applied", joined_logs)
            self.assertIn("WORD_COM_XML_APPLY_DONE applied=1 skipped=0 errors=0", joined_logs)

    def test_verify_word_com_approved_records_are_applied_by_python_xml(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(output_docx, make_document_with_character_indent(text="Approved by fake Word"))

            def fake_run(script_path, *, arguments=None, **kwargs):
                result_path = Path(arguments[5])
                approved = {
                    "record_index": 1,
                    "paragraph_index": 1,
                    "matched_paragraph_index": 1,
                    "word_dominant_font_size": 14,
                    "expected_left_twips": "1720",
                    "expected_first_line_twips": None,
                    "text_match_prefix": "Approved by fake Word",
                }
                result_path.write_text(
                    "\n".join(
                        [
                            "WORD_COM_PS_STARTED",
                            "WORD_COM_FONT_CHECK_ONLY_STARTED",
                            "WORD_COM_PS_RECORDS_LOADED count=1",
                            "WORD_COM_PS_DONE",
                            "WORD_COM_FONT_CHECK_SUMMARY processed=1 approved=1 skipped_not_14=0 not_found=0 errors=0",
                            "WORD_COM_APPROVED_RECORD_JSON " + json.dumps(approved),
                        ]
                    ),
                    encoding="utf-8",
                )
                return types.SimpleNamespace(stdout="", stderr="", returncode=0)

            with patch("docx_fixer.word_com_indent.run_powershell_file", side_effect=fake_run):
                logs = verify_and_fix_body_indents_with_word_com(
                    output_docx,
                    [
                        {
                            "paragraph_index": 1,
                            "text_preview": "Approved by fake Word",
                            "text_match_prefix": "Approved by fake Word",
                            "kind": "body_font_check",
                            "expected_left_cm": 3.0,
                            "expected_left_twips": "1720",
                            "expected_left_points": 86.0,
                            "apply_only_if_word_font_size_is_14": True,
                        }
                    ],
                )

            root = read_document_root(output_docx)
            ind = root.find(".//w:ind", NS)
            self.assertEqual(ind.get(qn("left")), "1720")
            self.assertEqual(ind.get(qn("start")), "1720")
            self.assertEqual(ind.get(qn("firstLine")), "0")
            self.assertEqual(ind.get(qn("hanging")), "0")
            self.assertEqual(ind.get(qn("leftChars")), "0")
            self.assertEqual(ind.get(qn("startChars")), "0")
            self.assertEqual(ind.get(qn("firstLineChars")), "0")
            self.assertEqual(ind.get(qn("hangingChars")), "0")
            joined_logs = "\n".join(logs)
            self.assertIn("WORD_COM_FONT_CHECK_APPROVED_COUNT=1", joined_logs)
            self.assertIn("WORD_COM_XML_APPLY_RECORD paragraph_index=1 status=applied", joined_logs)

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

            with patch("docx_fixer.word_com_indent.run_powershell_file", side_effect=fake_run):
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
            body_paragraph = root.xpath(".//w:p", namespaces=NS)[3]
            assert_body_indent_hard_override(self, body_paragraph, TEMPLATE_OUTLINE_INDENTS[3]["body_left"])
            debug = "\n".join(summary.body_indent_debug_logs)
            self.assertIn("font_size_source=paragraph_style:DefaultText", debug)
            self.assertTrue(any(int(record["paragraph_index"]) == 4 for record in summary.body_indent_records))

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
                    enable_level1_level2_body_first_line_indent=True,
                ),
            )

            root = read_document_root(output_docx)
            body_paragraph = root.xpath(".//w:p", namespaces=NS)[3]
            assert_body_indent_hard_override(
                self,
                body_paragraph,
                TEMPLATE_OUTLINE_INDENTS[1]["body_left"],
                expected_first_line="560",
            )
            debug = "\n".join(summary.body_indent_debug_logs)
            self.assertIn("spec_firstLine_twips=560", debug)
            self.assertIn("written_firstLine_twips=560", debug)
            self.assertIn("written_firstLineChars=0", debug)
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
            body_paragraph = root.xpath(".//w:p", namespaces=NS)[3]
            assert_body_indent_hard_override(self, body_paragraph, TEMPLATE_OUTLINE_INDENTS[3]["body_left"])
            debug = "\n".join(summary.body_indent_debug_logs)
            self.assertIn(f"written_start_twips={TEMPLATE_OUTLINE_INDENTS[3]['body_left']}", debug)
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
            heading_ind = paragraphs[2].find("./w:pPr/w:ind", NS)
            body_ind = paragraphs[3].find("./w:pPr/w:ind", NS)
            self.assertEqual(heading_ind.get(qn("left")), spec["left"])
            self.assertEqual(heading_ind.get(qn("hanging")), spec["hanging"])
            self.assertAlmostEqual(
                int(heading_ind.get(qn("left"))) / 20 / 28.3464567,
                twips_to_cm(spec["left"]),
                places=2,
            )
            self.assertAlmostEqual(
                int(heading_ind.get(qn("hanging"))) / 20 / 28.3464567,
                twips_to_cm(spec["hanging"]),
                places=2,
            )
            assert_body_indent_hard_override(self, paragraphs[3], spec["body_left"])

            numbering_root = read_part_root(output_docx, "word/numbering.xml")
            numbering_lvl = numbering_root.find(".//w:lvl", NS)
            numbering_ind = numbering_lvl.find("./w:pPr/w:ind", NS)
            self.assertEqual(numbering_ind.get(qn("left")), spec["left"])
            self.assertEqual(numbering_ind.get(qn("hanging")), spec["hanging"])
            self.assertIsNone(numbering_ind.get(qn("start")))
            # "1." is outline level 3, a space-suffix level.
            assert_numbering_level_follows_suffix_rule(self, numbering_lvl, 3)

            styles_root = read_part_root(output_docx, "word/styles.xml")
            numbered_style = styles_root.xpath("./w:style[@w:styleId='NumberedL4']", namespaces=NS)[0]
            numbered_ind = numbered_style.find("./w:pPr/w:ind", NS)
            self.assertEqual(numbered_ind.get(qn("left")), spec["left"])
            self.assertEqual(numbered_ind.get(qn("hanging")), spec["hanging"])
            self.assertIsNone(numbered_ind.get(qn("start")))
            # styles.xml mirrors the level geometry without creating tab stops;
            # suffix spacing is carried only by numbering.xml's w:suff.
            self.assertIsNone(numbered_style.find("./w:pPr/w:tabs", NS))
            plain_style = styles_root.xpath("./w:style[@w:styleId='BodyText']", namespaces=NS)[0]
            self.assertIsNone(plain_style.find("./w:pPr/w:ind", NS))
            self.assertIsNone(plain_style.find("./w:pPr/w:tabs", NS))

            assert_docx_has_no_character_indent_attrs(self, output_docx)
            logs = "\n".join(summary.numbering_xml_logs)
            self.assertIn("NUMBERING_XML_LEVEL_INDENT", logs)
            self.assertIn("STYLES_XML_NUMBERED_STYLE_INDENT", logs)
            self.assertIn(f"expected_number_start_cm={twips_to_cm(spec['number_start']):.2f}", logs)
            self.assertIn("expected_heading_text_start_cm=", logs)
            self.assertIn("expected_tab_pos_cm=", logs)
            self.assertIn("suff=space", logs)
            self.assertIn("tab_pos_cm=None", logs)
            records_by_kind = {record["kind"]: record for record in summary.body_indent_records}
            self.assertAlmostEqual(
                records_by_kind["auto(style)"]["expected_heading_left_cm"],
                twips_to_cm(spec["left"]),
                places=2,
            )
            self.assertAlmostEqual(
                records_by_kind["auto(style)"]["expected_heading_text_start_cm"],
                round(int(spec["heading_text_start"]) / 20 / 28.3464567, 2),
                places=2,
            )
            self.assertAlmostEqual(
                records_by_kind["auto(style)"]["expected_hanging_cm"],
                twips_to_cm(spec["hanging"]),
                places=2,
            )
            self.assertAlmostEqual(
                records_by_kind["body"]["expected_body_left_cm"],
                twips_to_cm(spec["body_left"]),
                places=2,
            )
            self.assertEqual(records_by_kind["body"]["expected_firstline_cm"], 0.0)

    def test_fix_docx_fast_passes_only_font_check_records_to_word_com(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            long_body_text = " ".join(f"bodysegment{index:02d}" for index in range(20))
            make_docx(
                input_docx,
                make_document_with_style_level_four_heading_and_body(
                    body_font_size_pt=12,
                    body_text=long_body_text,
                ),
                styles_xml=make_styles_with_level_four_numbered_and_plain_old_indents(),
                numbering_xml=make_level_four_numbering_with_old_indents(),
            )
            captured: dict[str, object] = {}

            def fake_verify(docx_path, records, stop=None):
                captured["docx_path"] = docx_path
                captured["records"] = list(records)
                return ["WORD_COM_FAKE_VERIFIED"]

            with patch(
                "docx_fixer.word_com_indent.verify_and_fix_body_indents_with_word_com",
                side_effect=fake_verify,
            ) as verifier:
                summary = fix_docx_fast(
                    input_docx,
                    output_docx,
                    ProcessOptions(
                        fix_table_layout=False,
                        fix_color=False,
                        fix_paragraph=True,
                        normalize_with_word_com=True,
                        word_com_check_body_font_when_xml_not_14=True,
                    ),
                )

        verifier.assert_called_once()
        records = captured["records"]
        self.assertGreater(len(summary.body_indent_records), len(records))
        self.assertEqual([record["kind"] for record in records], ["body_font_check"])
        self.assertTrue(all(record.get("apply_only_if_word_font_size_is_14") is True for record in records))
        self.assertTrue(records[0]["text_preview"].endswith("..."))
        self.assertFalse(records[0]["text_match_prefix"].endswith("..."))
        self.assertEqual(records[0]["text_match_prefix"], long_body_text[:120])
        self.assertGreater(len(records[0]["text_match_prefix"]), len(records[0]["text_preview"]))
        self.assertTrue(any(record["kind"] == "auto(style)" for record in summary.body_indent_records))
        self.assertFalse(any(record["kind"] == "auto(style)" for record in records))
        joined_logs = "\n".join(summary.word_com_body_indent_logs)
        self.assertIn("WORD_COM_BODY_INDENT_RECORD_FILTER", joined_logs)
        self.assertIn(f"total_records={len(summary.body_indent_records)}", joined_logs)
        self.assertIn("word_com_records=1", joined_logs)
        self.assertIn(f"skipped_records={len(summary.body_indent_records) - 1}", joined_logs)
        self.assertIn("criteria=apply_only_if_word_font_size_is_14_and_xml_font_size_gt_11", joined_logs)
        self.assertIn("WORD_COM_FAKE_VERIFIED", joined_logs)

    def test_fix_docx_fast_skips_word_com_when_filter_has_no_font_check_records(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_document_with_style_level_four_heading_and_body(body_font_size_pt=14),
                styles_xml=make_styles_with_level_four_numbered_and_plain_old_indents(),
                numbering_xml=make_level_four_numbering_with_old_indents(),
            )

            with patch("docx_fixer.word_com_indent.verify_and_fix_body_indents_with_word_com") as verifier:
                summary = fix_docx_fast(
                    input_docx,
                    output_docx,
                    ProcessOptions(
                        fix_table_layout=False,
                        fix_color=False,
                        fix_paragraph=True,
                        normalize_with_word_com=True,
                        word_com_check_body_font_when_xml_not_14=True,
                    ),
                )

        verifier.assert_not_called()
        self.assertTrue(summary.body_indent_records)
        self.assertFalse(
            any(bool(record.get("apply_only_if_word_font_size_is_14")) for record in summary.body_indent_records)
        )
        joined_logs = "\n".join(summary.word_com_body_indent_logs)
        self.assertIn("WORD_COM_BODY_INDENT_RECORD_FILTER", joined_logs)
        self.assertIn("word_com_records=0", joined_logs)
        self.assertIn(f"skipped_records={len(summary.body_indent_records)}", joined_logs)
        self.assertIn("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=no_font_check_records", joined_logs)

    def test_fix_docx_fast_skips_word_com_for_xml_font_size_at_or_below_11(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_document_with_style_level_four_heading_and_body(body_font_size_pt=10),
                styles_xml=make_styles_with_level_four_numbered_and_plain_old_indents(),
                numbering_xml=make_level_four_numbering_with_old_indents(),
            )

            with patch("docx_fixer.word_com_indent.verify_and_fix_body_indents_with_word_com") as verifier:
                summary = fix_docx_fast(
                    input_docx,
                    output_docx,
                    ProcessOptions(
                        fix_table_layout=False,
                        fix_color=False,
                        fix_paragraph=True,
                        normalize_with_word_com=True,
                        word_com_check_body_font_when_xml_not_14=True,
                    ),
                )

        verifier.assert_not_called()
        font_check_records = [
            record
            for record in summary.body_indent_records
            if record.get("apply_only_if_word_font_size_is_14")
        ]
        self.assertEqual(len(font_check_records), 1)
        self.assertEqual(font_check_records[0]["kind"], "body_font_check")
        self.assertEqual(font_check_records[0]["xml_font_size"], 10.0)
        joined_logs = "\n".join(summary.word_com_body_indent_logs)
        self.assertIn("WORD_COM_BODY_INDENT_RECORD_FILTER", joined_logs)
        self.assertIn("word_com_records=0", joined_logs)
        self.assertIn(f"skipped_records={len(summary.body_indent_records)}", joined_logs)
        self.assertIn("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=no_font_check_records", joined_logs)

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

    def test_toc_range_is_immutable_across_document_styles_numbering_and_sanitize(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            document_xml = make_toc_immutable_document_xml()
            styles_xml = make_toc_immutable_styles_xml()
            numbering_xml = make_toc_immutable_numbering_xml()
            make_docx(
                input_docx,
                document_xml,
                styles_xml=styles_xml,
                numbering_xml=numbering_xml,
            )

            input_document_root = etree.fromstring(document_xml)
            input_styles_root = etree.fromstring(styles_xml)
            input_numbering_root = etree.fromstring(numbering_xml)
            input_paragraphs = input_document_root.xpath(".//w:p", namespaces=NS)
            input_toc_marker_xml = etree.tostring(input_paragraphs[1])
            input_toc_entry_xml = etree.tostring(input_paragraphs[2])
            input_toc_style_xml = etree.tostring(
                input_styles_root.xpath("./w:style[@w:styleId='TOC1']", namespaces=NS)[0]
            )
            input_toc_numbering_lvl = input_numbering_root.xpath(
                "./w:abstractNum[@w:abstractNumId='1']/w:lvl",
                namespaces=NS,
            )[0]
            input_toc_numbering_ind = input_toc_numbering_lvl.find("./w:pPr/w:ind", NS)

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=True,
                ),
            )

            output_document_root = read_document_root(output_docx)
            output_styles_root = read_part_root(output_docx, "word/styles.xml")
            output_numbering_root = read_part_root(output_docx, "word/numbering.xml")
            output_paragraphs = output_document_root.xpath(".//w:p", namespaces=NS)

            self.assertEqual(etree.tostring(output_paragraphs[1]), input_toc_marker_xml)
            self.assertEqual(etree.tostring(output_paragraphs[2]), input_toc_entry_xml)
            self.assertEqual(
                etree.tostring(output_styles_root.xpath("./w:style[@w:styleId='TOC1']", namespaces=NS)[0]),
                input_toc_style_xml,
            )
            output_toc_numbering_lvl = output_numbering_root.xpath(
                "./w:abstractNum[@w:abstractNumId='1']/w:lvl",
                namespaces=NS,
            )[0]
            output_toc_numbering_ind = output_toc_numbering_lvl.find("./w:pPr/w:ind", NS)
            self.assertEqual(output_toc_numbering_ind.get(qn("left")), input_toc_numbering_ind.get(qn("left")))
            self.assertEqual(output_toc_numbering_ind.get(qn("hanging")), input_toc_numbering_ind.get(qn("hanging")))
            # The excluded TOC numbering level is now left fully intact: its
            # original suffix and tab stops are preserved (skip is decided before
            # sanitizing), so the TOC numbering definition is truly immutable.
            self.assertEqual(
                output_toc_numbering_lvl.find("./w:suff", NS).get(qn("val")),
                input_toc_numbering_lvl.find("./w:suff", NS).get(qn("val")),
            )
            self.assertIsNotNone(output_toc_numbering_lvl.find("./w:pPr/w:tabs", NS))

            body_marker = output_paragraphs[3]
            body_paragraph = output_paragraphs[4]
            self.assertEqual(body_marker.find("./w:pPr/w:outlineLvl", NS).get(qn("val")), "0")
            self.assertIsNotNone(body_paragraph.find("./w:pPr/w:ind", NS))
            self.assertEqual(body_paragraph.find("./w:pPr/w:ind", NS).get(qn("leftChars")), "0")
            self.assertIsNone(body_paragraph.find("./w:pPr/w:tabs", NS))

            non_toc_numbering_ind = output_numbering_root.xpath(
                "./w:abstractNum[@w:abstractNumId='2']/w:lvl/w:pPr/w:ind",
                namespaces=NS,
            )[0]
            self.assertNotEqual(non_toc_numbering_ind.get(qn("left")), "999")
            self.assertIsNone(non_toc_numbering_ind.get(qn("leftChars")))

            joined_paragraph_logs = "\n".join(summary.paragraph_logs)
            joined_numbering_logs = "\n".join(summary.numbering_xml_logs)
            self.assertIn("skipped TOC paragraph; no formatting applied", joined_paragraph_logs)
            self.assertIn("CHAR_INDENT_SANITIZE_SKIP_EXCLUDED", joined_numbering_logs)
            self.assertIn("STYLES_XML_SKIP_TOC_STYLE: styleId=TOC1", joined_numbering_logs)
            self.assertIn("NUMBERING_XML_SKIP_TOC_NUMBERING", joined_numbering_logs)
            self.assertNotIn("STYLES_XML_NUMBERED_STYLE_INDENT: styleId=TOC1", joined_numbering_logs)

    def test_remove_all_outline_forces_toc_paragraphs_to_body_without_other_toc_changes(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            document_xml = make_toc_immutable_document_xml()
            make_docx(
                input_docx,
                document_xml,
                styles_xml=make_toc_immutable_styles_xml(),
                numbering_xml=make_toc_immutable_numbering_xml(),
            )

            input_root = etree.fromstring(document_xml)
            input_paragraphs = input_root.xpath(".//w:p", namespaces=NS)
            input_toc_texts = [
                "".join(p.xpath(".//w:t/text()", namespaces=NS))
                for p in input_paragraphs[:3]
            ]
            input_toc_ind_attrs = [
                dict(p.find("./w:pPr/w:ind", NS).attrib)
                if p.find("./w:pPr/w:ind", NS) is not None
                else None
                for p in input_paragraphs[:3]
            ]
            input_toc_tabs = [
                etree.tostring(p.find("./w:pPr/w:tabs", NS))
                if p.find("./w:pPr/w:tabs", NS) is not None
                else None
                for p in input_paragraphs[:3]
            ]

            fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=True,
                    remove_all_outline_levels=True,
                    normalize_with_word_com=False,
                ),
            )

            assert_toc_outlines_are_body(self, output_docx)
            output_root = read_document_root(output_docx)
            output_paragraphs = output_root.xpath(".//w:p", namespaces=NS)
            for index in range(3):
                with self.subTest(toc_paragraph=index):
                    self.assertEqual(
                        "".join(output_paragraphs[index].xpath(".//w:t/text()", namespaces=NS)),
                        input_toc_texts[index],
                    )
                    output_ind = output_paragraphs[index].find("./w:pPr/w:ind", NS)
                    output_ind_attrs = dict(output_ind.attrib) if output_ind is not None else None
                    self.assertEqual(output_ind_attrs, input_toc_ind_attrs[index])
                    output_tabs = output_paragraphs[index].find("./w:pPr/w:tabs", NS)
                    output_tabs_xml = etree.tostring(output_tabs) if output_tabs is not None else None
                    self.assertEqual(output_tabs_xml, input_toc_tabs[index])

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

            self.assertEqual(paragraph_outlines(output_docx), ["9", "9", "9"])
            assert_all_document_outlines_are_body(self, output_docx)
            self.assertEqual(part_outline_count(output_docx, "word/styles.xml"), 0)
            self.assertEqual(part_outline_count(output_docx, "word/numbering.xml"), 0)
            self.assertEqual(summary.removed_all_outline_paragraphs, 5)
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

            self.assertEqual(paragraph_outlines(output_docx), ["9", "9", "0"])
            self.assertEqual(summary.removed_all_outline_paragraphs, 3)
            self.assertEqual(summary.paragraph_level_counts[0], 1)

    def test_heading_suffix_summary_records_before_and_after_fix_state(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_heading_suffix_document_xml(),
                numbering_xml=make_heading_suffix_numbering_xml(),
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

            before_by_index = {
                int(record["paragraph_index"]): record
                for record in summary.heading_suffix_before_records
            }
            after_by_index = {
                int(record["paragraph_index"]): record
                for record in summary.heading_suffix_after_records
            }

            self.assertEqual(before_by_index[3]["source"], "manual_text")
            self.assertEqual(before_by_index[3]["suffix"], "space")
            self.assertEqual(before_by_index[3]["space_count"], 1)
            self.assertEqual(after_by_index[3]["suffix"], "nothing")
            self.assertEqual(after_by_index[3]["space_count"], 0)

            self.assertEqual(before_by_index[4]["source"], "auto_numbering_xml")
            self.assertEqual(before_by_index[4]["suffix"], "tab")
            self.assertEqual(before_by_index[4]["has_tab_stop"], True)
            self.assertEqual(before_by_index[4]["tab_pos_twips"], "2279")
            # decimal "%1." is outline level 3, which uses w:suff="space"
            # without any numbering tab stop.
            self.assertEqual(after_by_index[4]["suffix"], "space")
            self.assertEqual(after_by_index[4]["has_tab_stop"], False)
            self.assertIsNone(after_by_index[4]["tab_pos_twips"])
            self.assertIsNone(after_by_index[4]["expected_tab_pos_twips"])
            self.assertEqual(after_by_index[4]["numbering_level_source"], "abstractNum")
            self.assertIn("suff", after_by_index[4]["numbering_lvl_child_order"])
            self.assertIn("pPr", after_by_index[4]["numbering_lvl_child_order"])
            self.assertEqual(after_by_index[4]["numbering_pPr_child_order"], "ind")
            self.assertIn(after_by_index[4]["suffix_before_lvlText"], {True, False})
            self.assertFalse(after_by_index[4]["tabs_before_ind"])
            self.assertFalse(after_by_index[4]["compat_doNotUseIndentAsNumberingTabStop"])

    def test_heading_suffix_normalizes_auto_numbering_missing_suffix_and_tabs(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_heading_suffix_document_xml(),
                numbering_xml=make_heading_suffix_numbering_xml(suffix=None, include_tabs=True),
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

            before_auto = next(
                record for record in summary.heading_suffix_before_records
                if record.get("source") == "auto_numbering_xml"
            )
            after_auto = next(
                record for record in summary.heading_suffix_after_records
                if record.get("source") == "auto_numbering_xml"
            )
            self.assertEqual(before_auto["raw_suffix"], "missing")
            self.assertEqual(before_auto["effective_suffix"], "space")
            # decimal "%1." is outline level 3 -> space suffix, with no tabs.
            self.assertEqual(after_auto["raw_suffix"], "space")
            self.assertEqual(after_auto["effective_suffix"], "space")
            self.assertEqual(after_auto["has_tab_stop"], False)
            self.assertIsNone(after_auto["tab_pos_twips"])
            self.assertIsNone(after_auto["heading_text_start_twips"])

            numbering_root = read_part_root(output_docx, "word/numbering.xml")
            lvl = numbering_root.xpath("./w:abstractNum/w:lvl", namespaces=NS)[0]
            assert_numbering_level_follows_suffix_rule(self, lvl, 3)

    def test_heading_suffix_normalizes_auto_numbering_tab_and_space_suffixes(self):
        for suffix in ("tab", "space"):
            with self.subTest(suffix=suffix):
                with tempfile.TemporaryDirectory() as temp_dir:
                    input_docx = Path(temp_dir) / "input.docx"
                    output_docx = Path(temp_dir) / "output.docx"
                    make_docx(
                        input_docx,
                        make_heading_suffix_document_xml(),
                        numbering_xml=make_heading_suffix_numbering_xml(suffix=suffix, include_tabs=True),
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

                    before_auto = next(
                        record for record in summary.heading_suffix_before_records
                        if record.get("source") == "auto_numbering_xml"
                    )
                    after_auto = next(
                        record for record in summary.heading_suffix_after_records
                        if record.get("source") == "auto_numbering_xml"
                    )
                    self.assertEqual(before_auto["raw_suffix"], suffix)
                    # decimal "%1." is outline level 3 -> space suffix.
                    self.assertEqual(after_auto["raw_suffix"], "space")
                    self.assertEqual(after_auto["effective_suffix"], "space")
                    self.assertEqual(after_auto["has_tab_stop"], False)

                    numbering_root = read_part_root(output_docx, "word/numbering.xml")
                    lvl = numbering_root.xpath("./w:abstractNum/w:lvl", namespaces=NS)[0]
                    assert_numbering_level_follows_suffix_rule(self, lvl, 3)

    def test_heading_suffix_final_cleanup_trims_lvl_text_trailing_space(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_heading_suffix_document_xml(),
                numbering_xml=make_heading_suffix_numbering_xml(
                    suffix=None,
                    include_tabs=True,
                    lvl_text_value="%5. ",
                    tab_val="num",
                    tab_pos="2061",
                ),
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

            before_auto = next(
                record for record in summary.heading_suffix_before_records
                if record.get("source") == "auto_numbering_xml"
            )
            after_auto = next(
                record for record in summary.heading_suffix_after_records
                if record.get("source") == "auto_numbering_xml"
            )
            self.assertEqual(before_auto["lvlText"], "%5. ")
            self.assertEqual(before_auto["lvlText_has_trailing_space"], True)
            self.assertEqual(after_auto["lvlText"], "%5.")
            self.assertEqual(after_auto["lvlText_has_trailing_space"], False)
            # decimal "%5." is outline level 3 -> space suffix; trailing space
            # is trimmed from lvlText and spacing is realized by w:suff="space".
            self.assertEqual(after_auto["raw_suffix"], "space")
            self.assertEqual(after_auto["effective_suffix"], "space")
            self.assertEqual(after_auto["has_tab_stop"], False)

            numbering_root = read_part_root(output_docx, "word/numbering.xml")
            lvl = numbering_root.xpath("./w:abstractNum/w:lvl", namespaces=NS)[0]
            self.assertEqual(lvl.find("w:lvlText", NS).get(qn("val")), "%5.")
            assert_numbering_level_follows_suffix_rule(self, lvl, 3)

    def test_heading_suffix_final_cleanup_runs_after_word_com_changes_numbering(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_heading_suffix_document_xml(),
                numbering_xml=make_heading_suffix_numbering_xml(
                    suffix="nothing",
                    include_tabs=False,
                    lvl_text_value="%5.",
                ),
            )

            def fake_word_com(docx_path, records, stop=None):
                del records, stop
                dirty_numbering_suffix_tabs_in_docx(Path(docx_path))
                return ["WORD_COM_FAKE_DIRTIED_NUMBERING"]

            with patch(
                "docx_fixer.word_com_indent.filter_word_com_body_indent_records",
                return_value=[{"paragraph_index": 99}],
            ), patch(
                "docx_fixer.word_com_indent.verify_and_fix_body_indents_with_word_com",
                side_effect=fake_word_com,
            ):
                summary = fix_docx_fast(
                    input_docx,
                    output_docx,
                    ProcessOptions(
                        fix_table_layout=False,
                        fix_color=False,
                        fix_paragraph=True,
                        normalize_with_word_com=True,
                    ),
                )

            after_auto = next(
                record for record in summary.heading_suffix_after_records
                if record.get("source") == "auto_numbering_xml"
            )
            # decimal "%5." is outline level 3 -> space suffix; the final hard
            # clean after Word COM removes tabs and writes w:suff="space".
            self.assertEqual(after_auto["raw_suffix"], "space")
            self.assertEqual(after_auto["effective_suffix"], "space")
            self.assertEqual(after_auto["has_tab_stop"], False)
            self.assertEqual(after_auto["lvlText"], "%5.")
            self.assertEqual(after_auto["lvlText_has_trailing_space"], False)
            self.assertIn("WORD_COM_FAKE_DIRTIED_NUMBERING", summary.word_com_body_indent_logs)
            self.assertTrue(any("FINAL_NUMBERING_SUFFIX_CLEAN_DOCX changed=true" in log for log in summary.numbering_xml_logs))

            numbering_root = read_part_root(output_docx, "word/numbering.xml")
            lvl = numbering_root.xpath("./w:abstractNum/w:lvl", namespaces=NS)[0]
            self.assertEqual(lvl.find("w:lvlText", NS).get(qn("val")), "%5.")
            assert_numbering_level_follows_suffix_rule(self, lvl, 3)

    def test_note_debug_log_records_note_sources_and_center_inheritance(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            make_docx(
                input_docx,
                make_note_alignment_document_xml(),
                styles_xml=make_note_alignment_styles_xml(),
                numbering_xml=make_note_alignment_numbering_xml(),
            )

            note_debug_log = "\n".join(
                collect_note_debug_records_from_docx(input_docx, "debug_unit")
            )

            self.assertIn("stage=debug_unit", note_debug_log)
            self.assertIn("note_source=text", note_debug_log)
            self.assertIn("note_source=numPr", note_debug_log)
            self.assertIn("note_source=styleNumPr", note_debug_log)
            self.assertIn("style_jc_effective=center", note_debug_log)
            self.assertIn("style_based_on_chain=NoteStyle>CenteredNoteBase", note_debug_log)
            self.assertIn("numbering_lvlJc=center", note_debug_log)
            self.assertIn("WARNING: numbering level alignment is center", note_debug_log)
            self.assertIn("decision=SKIPPED_TABLE", note_debug_log)
            self.assertIn("final_paragraph_jc=center", note_debug_log)
            self.assertIn("WARNING: final paragraph alignment is center", note_debug_log)

    def test_final_note_alignment_is_disabled_by_default(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_note_alignment_document_xml(),
                styles_xml=make_note_alignment_styles_xml(),
                numbering_xml=make_note_alignment_numbering_xml(),
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

            joined_paragraph_logs = "\n".join(summary.paragraph_logs)
            # The dedicated note left-alignment pass must not run by default.
            self.assertIn(
                "FINAL_NOTE_ALIGNMENT_FIX_SKIPPED reason=disabled", joined_paragraph_logs
            )
            self.assertNotIn("FINAL_NOTE_ALIGNMENT_FIX part=", joined_paragraph_logs)
            self.assertNotIn("FINAL_NOTE_ALIGNMENT_SUMMARY", joined_paragraph_logs)

    def test_final_note_alignment_runs_after_word_com_and_keeps_notes_out_of_indent_processing(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_note_alignment_document_xml(),
                styles_xml=make_note_alignment_styles_xml(),
                numbering_xml=make_note_alignment_numbering_xml(),
            )

            def fake_word_com(docx_path, records, stop=None):
                del records, stop
                dirty_note_alignment_center_after_outline_in_docx(Path(docx_path))
                return ["WORD_COM_FAKE_DIRTIED_NOTE_ALIGNMENT"]

            with patch(
                "docx_fixer.word_com_indent.filter_word_com_body_indent_records",
                return_value=[{"paragraph_index": 99}],
            ), patch(
                "docx_fixer.word_com_indent.verify_and_fix_body_indents_with_word_com",
                side_effect=fake_word_com,
            ):
                summary = fix_docx_fast(
                    input_docx,
                    output_docx,
                    ProcessOptions(
                        fix_table_layout=False,
                        fix_color=False,
                        fix_paragraph=True,
                        normalize_with_word_com=True,
                        force_note_paragraph_left_alignment=True,
                        # Developer-only flag: opt in so this test can assert the
                        # note debug log content. It is False by default.
                        write_note_debug_log=True,
                    ),
                )

            root = read_document_root(output_docx)
            note_colon = find_paragraph_by_exact_text(root, "\u8a3b\uff1a\u9019\u662f\u8aaa\u660e")
            note_number = find_paragraph_by_exact_text(root, "  \u8a3b1\uff1a\u9019\u662f\u8aaa\u660e")
            note_chinese = find_paragraph_by_exact_text(root, "\u8a3b\u4e00\uff1a\u9019\u662f\u8aaa\u660e")
            auto_note = find_paragraph_by_exact_text(root, "\u81ea\u52d5\u7de8\u865f\u8a3b\u89e3\u5167\u5bb9")
            style_note = find_paragraph_by_exact_text(root, "\u6a23\u5f0f\u7de8\u865f\u8a3b\u89e3\u5167\u5bb9")
            non_note_numbered = find_paragraph_by_exact_text(root, "\u4e0d\u662f\u8a3b\u89e3\u7684\u7de8\u865f\u5167\u5bb9")
            body = find_paragraph_by_exact_text(root, "\u9019\u662f\u666e\u901a 14pt \u5167\u6587")
            table_note = find_paragraph_by_exact_text(root, "\u8a3b\uff1a\u8868\u683c\u5167\u8aaa\u660e")

            self.assertEqual(paragraph_jc_value(note_colon), "left")
            note_colon_tags = paragraph_ppr_child_tags(note_colon)
            self.assertLess(note_colon_tags.index(qn("jc")), note_colon_tags.index(qn("outlineLvl")))
            self.assertEqual(paragraph_style_value(note_colon), "NoteStyle")
            self.assertEqual(note_colon.find("./w:pPr/w:ind", NS).get(qn("left")), "123")
            self.assertIsNotNone(note_colon.find("./w:pPr/w:tabs", NS))

            self.assertEqual(paragraph_jc_value(note_number), "left")
            self.assertEqual(paragraph_style_value(note_number), "NoteStyle")
            self.assertIsNotNone(note_number.find("./w:pPr/w:numPr", NS))
            self.assertEqual(note_number.find("./w:pPr/w:ind", NS).get(qn("left")), "223")

            self.assertEqual(paragraph_jc_value(note_chinese), "left")
            self.assertEqual(paragraph_style_value(note_chinese), "NoteStyle")
            self.assertIsNone(note_chinese.find("./w:pPr/w:outlineLvl", NS))
            self.assertEqual(note_chinese.find("./w:pPr/w:ind", NS).get(qn("left")), "323")

            self.assertEqual(paragraph_jc_value(auto_note), "left")
            self.assertIsNotNone(auto_note.find("./w:pPr/w:numPr", NS))
            self.assertEqual(auto_note.find("./w:pPr/w:ind", NS).get(qn("left")), "423")

            self.assertEqual(paragraph_jc_value(style_note), "left")
            self.assertEqual(paragraph_style_value(style_note), "StyleNoteNumbered")
            self.assertIsNone(style_note.find("./w:pPr/w:numPr", NS))
            self.assertEqual(style_note.find("./w:pPr/w:ind", NS).get(qn("left")), "523")

            self.assertIsNone(paragraph_jc_value(non_note_numbered))
            assert_body_indent_hard_override(self, body, TEMPLATE_OUTLINE_INDENTS[1]["body_left"])
            self.assertEqual(paragraph_jc_value(table_note), "center")

            joined_paragraph_logs = "\n".join(summary.paragraph_logs)
            self.assertIn("WORD_COM_FAKE_DIRTIED_NOTE_ALIGNMENT", "\n".join(summary.word_com_body_indent_logs))
            self.assertIn("FINAL_NOTE_ALIGNMENT_FIX part=word/document.xml", joined_paragraph_logs)
            self.assertIn("source=text", joined_paragraph_logs)
            self.assertIn("source=numPr", joined_paragraph_logs)
            self.assertIn("source=styleNumPr", joined_paragraph_logs)
            self.assertIn("before_jc=center after_jc=left", joined_paragraph_logs)
            self.assertEqual(joined_paragraph_logs.count("FINAL_NOTE_ALIGNMENT_FIX part=word/document.xml"), 5)
            self.assertIn("FINAL_NOTE_ALIGNMENT_SUMMARY part=word/document.xml", joined_paragraph_logs)
            self.assertIn("matched_text=4", joined_paragraph_logs)
            self.assertIn("matched_numPr=1", joined_paragraph_logs)
            self.assertIn("matched_styleNumPr=1", joined_paragraph_logs)
            self.assertIn("fixed_count=5", joined_paragraph_logs)
            self.assertIn("skipped_table=1", joined_paragraph_logs)
            self.assertIn("center_after_fix_count=0", joined_paragraph_logs)

            note_debug_log_path = output_docx.with_name("output_note_debug_log.txt")
            self.assertTrue(note_debug_log_path.exists())
            note_debug_log = note_debug_log_path.read_text(encoding="utf-8")
            self.assertIn("stage=after_xml_pipeline_before_word_com", note_debug_log)
            self.assertIn("stage=after_final_output", note_debug_log)
            self.assertIn("note_source=text", note_debug_log)
            self.assertIn("note_source=numPr", note_debug_log)
            self.assertIn("note_source=styleNumPr", note_debug_log)
            self.assertIn("style_jc_effective=center", note_debug_log)
            self.assertIn("style_based_on_chain=NoteStyle>CenteredNoteBase", note_debug_log)
            self.assertIn("decision=SKIPPED_TABLE", note_debug_log)
            self.assertIn("final_paragraph_jc=center", note_debug_log)
            self.assertIn("WARNING: final paragraph alignment is center", note_debug_log)

    def test_note_debug_log_is_not_written_by_default(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_note_alignment_document_xml(),
                styles_xml=make_note_alignment_styles_xml(),
                numbering_xml=make_note_alignment_numbering_xml(),
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

            # The default is False, so no *_note_debug_log.txt is produced.
            self.assertFalse(ProcessOptions(True, True, True).write_note_debug_log)
            note_debug_log_path = output_docx.with_name("output_note_debug_log.txt")
            self.assertFalse(note_debug_log_path.exists())
            self.assertEqual(
                list(Path(temp_dir).glob("*_note_debug_log.txt")),
                [],
            )
            joined = "\n".join(summary.paragraph_logs)
            self.assertIn("NOTE_DEBUG_LOG_SKIPPED reason=disabled", joined)
            self.assertNotIn("NOTE_DEBUG_LOG_WRITTEN", joined)

    def test_simulated_gui_temp_output_does_not_create_tmp_note_debug_log(self):
        # Reproduces the GUI path: fix_docx_fast runs on the *.__tmp__.docx temp
        # output, which is what produced the __tmp___note_debug_log.txt file.
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            final_output = Path(temp_dir) / "output.docx"
            temp_output = final_output.with_name(
                f"{final_output.stem}.__tmp__{final_output.suffix}"
            )
            make_docx(input_docx, make_document_xml())

            fix_docx_fast(
                input_docx,
                temp_output,
                ProcessOptions(
                    fix_table_layout=True,
                    fix_color=True,
                    fix_paragraph=True,
                    normalize_with_word_com=False,
                ),
            )

            self.assertEqual(
                list(Path(temp_dir).glob("*__tmp___note_debug_log.txt")),
                [],
            )
            self.assertEqual(list(Path(temp_dir).glob("*_note_debug_log.txt")), [])

    def test_note_debug_log_written_only_when_flag_enabled(self):
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
                    normalize_with_word_com=False,
                    write_note_debug_log=True,
                ),
            )

            note_debug_log_path = output_docx.with_name("output_note_debug_log.txt")
            self.assertTrue(note_debug_log_path.exists())
            joined = "\n".join(summary.paragraph_logs)
            self.assertIn("NOTE_DEBUG_LOG_WRITTEN", joined)
            self.assertNotIn("NOTE_DEBUG_LOG_SKIPPED", joined)

    def test_official_logs_still_written_with_note_debug_disabled(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(input_docx, make_document_xml())

            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=True,
                    fix_color=True,
                    fix_paragraph=True,
                    normalize_with_word_com=False,
                    skip_log_output=False,
                ),
            )

            from docx_fixer.process_log import (
                write_heading_suffix_log_file,
                write_process_log,
                write_table_log_file,
            )

            process_log_path = write_process_log(output_docx, summary)
            table_log_path = write_table_log_file(output_docx, summary)
            heading_log_path = write_heading_suffix_log_file(output_docx, summary)

            # Official logs are unaffected by the disabled note debug log.
            self.assertTrue(process_log_path.exists())
            self.assertTrue(table_log_path.exists())
            self.assertTrue(heading_log_path.exists())
            self.assertFalse(
                output_docx.with_name("output_note_debug_log.txt").exists()
            )


CHAPTER_THREE_NUMBERING_TITLE = "參、價格形成之主要因素分析"


def make_chapter_three_numbering_xml() -> bytes:
    """Two dedicated numbering definitions:

    - numId 40 (BODY): used by a level-0 body heading outside 參, must be cleaned.
    - numId 50 (CH3): used by a paragraph inside 參, protected when the option is on.
    """
    numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    for abstract_id, fmt, lvl_text_val in [
        ("40", "ideographLegalTraditional", "%1、"),  # -> level 0 (body heading)
        ("50", "decimal", "%1.　"),  # -> level 3, trailing ideographic space
    ]:
        abstract = etree.SubElement(numbering, qn("abstractNum"))
        abstract.set(qn("abstractNumId"), abstract_id)
        lvl = etree.SubElement(abstract, qn("lvl"))
        lvl.set(qn("ilvl"), "0")
        suff = etree.SubElement(lvl, qn("suff"))
        suff.set(qn("val"), "tab")
        num_fmt = etree.SubElement(lvl, qn("numFmt"))
        num_fmt.set(qn("val"), fmt)
        lvl_text = etree.SubElement(lvl, qn("lvlText"))
        lvl_text.set(qn("val"), lvl_text_val)
        pPr = etree.SubElement(lvl, qn("pPr"))
        tabs = etree.SubElement(pPr, qn("tabs"))
        tab = etree.SubElement(tabs, qn("tab"))
        tab.set(qn("val"), "left")
        tab.set(qn("pos"), "999")
        ind = etree.SubElement(pPr, qn("ind"))
        ind.set(qn("left"), "700")
        ind.set(qn("hanging"), "400")

        num = etree.SubElement(numbering, qn("num"))
        num.set(qn("numId"), abstract_id)
        ref = etree.SubElement(num, qn("abstractNumId"))
        ref.set(qn("val"), abstract_id)
    return etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_chapter_three_numbering_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))

    def numbered_p(text: str, num_id: str) -> None:
        p = etree.SubElement(body, qn("p"))
        pPr = etree.SubElement(p, qn("pPr"))
        num_pr = etree.SubElement(pPr, qn("numPr"))
        ilvl = etree.SubElement(num_pr, qn("ilvl"))
        ilvl.set(qn("val"), "0")
        nid = etree.SubElement(num_pr, qn("numId"))
        nid.set(qn("val"), num_id)
        r = etree.SubElement(p, qn("r"))
        t = etree.SubElement(r, qn("t"))
        t.text = text

    def plain_p(text: str) -> None:
        p = etree.SubElement(body, qn("p"))
        r = etree.SubElement(p, qn("r"))
        t = etree.SubElement(r, qn("t"))
        t.text = text

    plain_p("封面")
    numbered_p("序言", "40")  # 壹 body heading, auto-numbered (level 0)
    plain_p(CHAPTER_THREE_NUMBERING_TITLE)  # starts the 參 region
    numbered_p("估價方法說明", "50")  # inside 參, auto-numbered (level 3)
    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_shared_abstract_multi_ilvl_numbering_xml() -> bytes:
    """One abstractNumId (70) referenced by one numId (70) with two levels.

    參 uses ilvl 0; another section uses ilvl 1. Only ilvl 0 should be protected.
    """
    numbering = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    abstract = etree.SubElement(numbering, qn("abstractNum"))
    abstract.set(qn("abstractNumId"), "70")
    for ilvl, tab_pos in (("0", "901"), ("1", "902")):
        lvl = etree.SubElement(abstract, qn("lvl"))
        lvl.set(qn("ilvl"), ilvl)
        suff = etree.SubElement(lvl, qn("suff"))
        suff.set(qn("val"), "tab")
        num_fmt = etree.SubElement(lvl, qn("numFmt"))
        num_fmt.set(qn("val"), "decimal")
        lvl_text = etree.SubElement(lvl, qn("lvlText"))
        lvl_text.set(qn("val"), "%1.　")  # trailing ideographic space
        pPr = etree.SubElement(lvl, qn("pPr"))
        tabs = etree.SubElement(pPr, qn("tabs"))
        tab = etree.SubElement(tabs, qn("tab"))
        tab.set(qn("val"), "left")
        tab.set(qn("pos"), tab_pos)
        ind = etree.SubElement(pPr, qn("ind"))
        ind.set(qn("left"), "700")
        ind.set(qn("hanging"), "400")
    num = etree.SubElement(numbering, qn("num"))
    num.set(qn("numId"), "70")
    ref = etree.SubElement(num, qn("abstractNumId"))
    ref.set(qn("val"), "70")
    return etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_shared_abstract_multi_ilvl_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))

    def numbered_p(text: str, num_id: str, ilvl: int) -> None:
        p = etree.SubElement(body, qn("p"))
        pPr = etree.SubElement(p, qn("pPr"))
        num_pr = etree.SubElement(pPr, qn("numPr"))
        ilvl_el = etree.SubElement(num_pr, qn("ilvl"))
        ilvl_el.set(qn("val"), str(ilvl))
        nid = etree.SubElement(num_pr, qn("numId"))
        nid.set(qn("val"), num_id)
        r = etree.SubElement(p, qn("r"))
        t = etree.SubElement(r, qn("t"))
        t.text = text

    def plain_p(text: str) -> None:
        p = etree.SubElement(body, qn("p"))
        r = etree.SubElement(p, qn("r"))
        t = etree.SubElement(r, qn("t"))
        t.text = text

    plain_p(CHAPTER_THREE_NUMBERING_TITLE)  # starts 參
    numbered_p("參內子項", "70", 0)  # inside 參, uses (70, ilvl 0)
    plain_p("肆、第四章")  # first-level heading -> ends 參
    numbered_p("肆內子項", "70", 1)  # outside 參, uses (70, ilvl 1)
    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


class ChapterThreeNumberingSuffixCleanupTests(unittest.TestCase):
    def _run(self, *, skip_cleanup: bool):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_chapter_three_numbering_document_xml(),
                numbering_xml=make_chapter_three_numbering_xml(),
            )
            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=True,
                    normalize_with_word_com=False,
                    skip_chapter_three_numbering_suffix_cleanup=skip_cleanup,
                ),
            )
            numbering_root = read_part_root(output_docx, "word/numbering.xml")
        return summary, numbering_root

    def _lvl(self, numbering_root, abstract_id: str):
        return numbering_root.xpath(
            f"./w:abstractNum[@w:abstractNumId='{abstract_id}']/w:lvl",
            namespaces=NS,
        )[0]

    def test_option_default_is_true(self):
        self.assertTrue(ProcessOptions(True, True, True).skip_chapter_three_numbering_suffix_cleanup)

    def test_enabled_protects_chapter_three_numbering(self):
        summary, numbering_root = self._run(skip_cleanup=True)
        ch3 = self._lvl(numbering_root, "50")

        # 參 numbering keeps its original suffix / tab / lvlText trailing space.
        self.assertEqual(ch3.find("./w:suff", NS).get(qn("val")), "tab")
        self.assertIsNotNone(ch3.find("./w:pPr/w:tabs", NS))
        self.assertEqual(ch3.find("./w:lvlText", NS).get(qn("val")), "%1.　")

        logs = "\n".join(summary.numbering_xml_logs)
        self.assertIn("CHAPTER_THREE_NUMBERING_SUFFIX_CLEANUP_SKIP enabled=true", logs)
        # Protection is expressed as precise pairs / abstract levels, never as a
        # whole abstractNumId.
        self.assertIn("protected_pairs=50:0", logs)
        self.assertIn("protected_abstract_levels=50:0", logs)
        self.assertIn("protected_abstractIds_not_used_for_chapter_three=true", logs)

    def test_enabled_still_cleans_other_body_heading_numbering(self):
        _summary, numbering_root = self._run(skip_cleanup=True)
        body = self._lvl(numbering_root, "40")

        # A body heading outside 參 is still cleaned.
        self.assertEqual(body.find("./w:suff", NS).get(qn("val")), "nothing")
        self.assertIsNone(body.find("./w:pPr/w:tabs", NS))

    def test_disabled_cleans_chapter_three_numbering(self):
        summary, numbering_root = self._run(skip_cleanup=False)
        ch3 = self._lvl(numbering_root, "50")

        # With the protection off, 參 numbering is cleaned to the central rule;
        # decimal "%1." is outline level 3, a space-suffix level.
        assert_numbering_level_follows_suffix_rule(self, ch3, 3)
        self.assertEqual(ch3.find("./w:lvlText", NS).get(qn("val")), "%1.")

    def _run_shared_abstract(self, *, skip_cleanup: bool):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            make_docx(
                input_docx,
                make_shared_abstract_multi_ilvl_document_xml(),
                numbering_xml=make_shared_abstract_multi_ilvl_numbering_xml(),
            )
            summary = fix_docx_fast(
                input_docx,
                output_docx,
                ProcessOptions(
                    fix_table_layout=False,
                    fix_color=False,
                    fix_paragraph=True,
                    normalize_with_word_com=False,
                    skip_chapter_three_numbering_suffix_cleanup=skip_cleanup,
                ),
            )
            numbering_root = read_part_root(output_docx, "word/numbering.xml")
        return summary, numbering_root

    def _shared_lvl(self, numbering_root, ilvl: str):
        return numbering_root.xpath(
            f"./w:abstractNum[@w:abstractNumId='70']/w:lvl[@w:ilvl='{ilvl}']",
            namespaces=NS,
        )[0]

    def test_shared_abstract_protects_only_chapter_three_ilvl(self):
        # 參 uses (70, ilvl 0); another section uses (70, ilvl 1) of the SAME
        # abstractNumId. Only the 參 level (ilvl 0) is protected.
        summary, numbering_root = self._run_shared_abstract(skip_cleanup=True)

        protected = self._shared_lvl(numbering_root, "0")
        self.assertEqual(protected.find("./w:suff", NS).get(qn("val")), "tab")
        self.assertIsNotNone(protected.find("./w:pPr/w:tabs", NS))
        self.assertEqual(protected.find("./w:lvlText", NS).get(qn("val")), "%1.　")

        # The other level in the same abstractNumId is still cleaned to the rule;
        # decimal "%1." is outline level 3, a space-suffix level.
        cleaned = self._shared_lvl(numbering_root, "1")
        assert_numbering_level_follows_suffix_rule(self, cleaned, 3)
        self.assertEqual(cleaned.find("./w:lvlText", NS).get(qn("val")), "%1.")

        logs = "\n".join(summary.numbering_xml_logs)
        self.assertIn("protected_abstract_levels=70:0", logs)
        self.assertIn("protected_abstractIds_not_used_for_chapter_three=true", logs)

    def test_shared_abstract_disabled_cleans_both_ilvls(self):
        summary, numbering_root = self._run_shared_abstract(skip_cleanup=False)
        for ilvl in ("0", "1"):
            lvl = self._shared_lvl(numbering_root, ilvl)
            # decimal "%1." is outline level 3, a space-suffix level.
            assert_numbering_level_follows_suffix_rule(self, lvl, 3)

        logs = "\n".join(summary.numbering_xml_logs)
        self.assertIn("CHAPTER_THREE_NUMBERING_SUFFIX_CLEANUP_SKIP enabled=false", logs)


if __name__ == "__main__":
    unittest.main()


