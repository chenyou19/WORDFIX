from __future__ import annotations

import unittest

from lxml import etree

from docx_fixer.constants import (
    NS,
    OUTLINE_LEVEL_FONT_SIZE_PT,
    PREFACE_OUTLINE_INDENTS,
    TEMPLATE_OUTLINE_INDENTS,
    W_NS,
    make_outline_indent_spec,
    validate_template_outline_indents,
)
from docx_fixer.models import ProcessSummary
from docx_fixer.numbering import (
    TAB_SUFFIX_OUTLINE_LEVELS,
    apply_numbering_outline_format,
    apply_styles_outline_format_to_root,
    build_numbering_format_lookup,
    build_numbering_level_lookup,
    detect_valid_auto_heading_level,
    force_clean_numbering_suffix_tabs,
    uses_tab_suffix,
)
from docx_fixer.style_resolver import build_style_font_size_lookup
from docx_fixer.indent_settings import twips_to_cm
from docx_fixer.outline import (
    apply_indent_spec_to_pPr,
    detect_outline_level,
    fix_outline_paragraphs,
    force_all_paragraphs_to_body_outline_level,
    is_note_paragraph,
    preface_indent_level_from_detected_level,
    remove_all_outline_levels_from_any_root,
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
    font_size_pt: float | None = None,
    run_style: str | None = None,
    runs: list[dict[str, object]] | None = None,
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

    run_specs = runs or [
        {
            "text": text,
            "font_size_pt": font_size_pt,
            "run_style": run_style,
        }
    ]
    for run_spec in run_specs:
        r = etree.SubElement(p, qn("r"))
        run_font_size_pt = run_spec.get("font_size_pt")
        run_style_value = run_spec.get("run_style")
        if run_font_size_pt is not None or run_style_value is not None:
            rPr = etree.SubElement(r, qn("rPr"))
            if run_style_value is not None:
                r_style = etree.SubElement(rPr, qn("rStyle"))
                r_style.set(qn("val"), str(run_style_value))
            if run_font_size_pt is not None:
                for tag in ("sz", "szCs"):
                    size = etree.SubElement(rPr, qn(tag))
                    size.set(qn("val"), str(round(float(run_font_size_pt) * 2)))
        t = etree.SubElement(r, qn("t"))
        t.text = str(run_spec.get("text", ""))
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


def paragraph_start_indent(p):
    ind = p.find("./w:pPr/w:ind", NS)
    if ind is None:
        return None
    return ind.get(qn("start"))


def paragraph_first_line_indent(p):
    ind = p.find("./w:pPr/w:ind", NS)
    if ind is None:
        return None
    return ind.get(qn("firstLine"))


def assert_body_indent_hard_override(testcase, p, expected_left: str, expected_first_line: str = "0"):
    ind = p.find("./w:pPr/w:ind", NS)
    testcase.assertIsNotNone(ind)
    testcase.assertEqual(ind.get(qn("left")), expected_left)
    testcase.assertEqual(ind.get(qn("start")), expected_left)
    testcase.assertEqual(ind.get(qn("firstLine")), expected_first_line)
    testcase.assertEqual(ind.get(qn("hanging")), "0")
    testcase.assertEqual(ind.get(qn("leftChars")), "0")
    testcase.assertEqual(ind.get(qn("startChars")), "0")
    testcase.assertEqual(ind.get(qn("firstLineChars")), "0")
    testcase.assertEqual(ind.get(qn("hangingChars")), "0")
    testcase.assertIsNone(p.find("./w:pPr/w:tabs", NS))
    testcase.assertIsNone(p.find("./w:pPr/w:numPr", NS))


def expected_body_left(level: int, set_outline: bool = True):
    spec = TEMPLATE_OUTLINE_INDENTS[level] if set_outline else PREFACE_OUTLINE_INDENTS[level]
    return spec["body_left"]


def paragraph_outline(p):
    outline = p.find("./w:pPr/w:outlineLvl", NS)
    if outline is None:
        return None
    return outline.get(qn("val"))


def paragraph_style(p):
    style = p.find("./w:pPr/w:pStyle", NS)
    if style is None:
        return None
    return style.get(qn("val"))


def paragraph_jc(p):
    jc = p.find("./w:pPr/w:jc", NS)
    if jc is None:
        return None
    return jc.get(qn("val"))


def paragraph_text_run_sizes(p):
    sizes = []
    for run in p.findall("./w:r", NS):
        if not "".join(run.xpath(".//w:t/text()", namespaces=NS)):
            continue
        size = run.find("./w:rPr/w:sz", NS)
        sizes.append(size.get(qn("val")) if size is not None else None)
    return sizes


def paragraph_text_run_szcs(p):
    sizes = []
    for run in p.findall("./w:r", NS):
        if not "".join(run.xpath(".//w:t/text()", namespaces=NS)):
            continue
        size = run.find("./w:rPr/w:szCs", NS)
        sizes.append(size.get(qn("val")) if size is not None else None)
    return sizes


def add_ind_with_char_attrs(p):
    pPr = p.find("./w:pPr", NS)
    if pPr is None:
        pPr = etree.SubElement(p, qn("pPr"))
    ind = pPr.find("w:ind", NS)
    if ind is None:
        ind = etree.SubElement(pPr, qn("ind"))
    ind.set(qn("left"), "123")
    ind.set(qn("hanging"), "45")
    for attr in ("leftChars", "startChars", "hangingChars", "firstLineChars"):
        ind.set(qn(attr), "99")
    return ind


def add_num_pr(p, num_id: str = "99", ilvl: str = "0"):
    pPr = p.find("./w:pPr", NS)
    if pPr is None:
        pPr = etree.SubElement(p, qn("pPr"))
    num_pr = pPr.find("w:numPr", NS)
    if num_pr is None:
        num_pr = etree.SubElement(pPr, qn("numPr"))
    ilvl_el = etree.SubElement(num_pr, qn("ilvl"))
    ilvl_el.set(qn("val"), ilvl)
    num_id_el = etree.SubElement(num_pr, qn("numId"))
    num_id_el.set(qn("val"), num_id)
    return num_pr


def add_tab_stop(p, pos="999"):
    pPr = p.find("./w:pPr", NS)
    if pPr is None:
        pPr = etree.SubElement(p, qn("pPr"))
    tabs = etree.SubElement(pPr, qn("tabs"))
    tab = etree.SubElement(tabs, qn("tab"))
    tab.set(qn("val"), "left")
    tab.set(qn("pos"), pos)
    return tabs


def assert_no_char_indent_attrs(testcase, element):
    for attr in ("leftChars", "startChars", "hangingChars", "firstLineChars"):
        testcase.assertIsNone(element.get(qn(attr)), attr)


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


def make_heading_validation_numbering_xml():
    """numbering.xml with one supported heading format, one bullet, and one
    unsupported format that must not fall back to ilvl."""
    root = etree.Element(qn("numbering"), nsmap={"w": W_NS})

    def add_abstract(abstract_id: str, ilvl: str, num_fmt: str, lvl_text: str):
        abstract = etree.SubElement(root, qn("abstractNum"))
        abstract.set(qn("abstractNumId"), abstract_id)
        lvl = etree.SubElement(abstract, qn("lvl"))
        lvl.set(qn("ilvl"), ilvl)
        fmt = etree.SubElement(lvl, qn("numFmt"))
        fmt.set(qn("val"), num_fmt)
        text_el = etree.SubElement(lvl, qn("lvlText"))
        text_el.set(qn("val"), lvl_text)

    def add_num(num_id: str, abstract_id: str):
        num = etree.SubElement(root, qn("num"))
        num.set(qn("numId"), num_id)
        abstract_el = etree.SubElement(num, qn("abstractNumId"))
        abstract_el.set(qn("val"), abstract_id)

    add_abstract("64", "0", "decimal", "（%1）")
    add_abstract("9", "0", "bullet", "")
    add_abstract("77", "2", "decimal", "%1")
    add_num("64", "64")
    add_num("9", "9")
    add_num("77", "77")
    return etree.tostring(root)


# Recognizable (numFmt, lvlText) per detected outline level 0-8.
RECOGNIZABLE_LEVEL_FORMATS = {
    0: ("ideographLegalTraditional", "%1、"),
    1: ("taiwaneseCountingThousand", "%1、"),
    2: ("taiwaneseCountingThousand", "（%1）"),
    3: ("decimal", "%1."),
    4: ("decimal", "（%1）"),
    5: ("upperLetter", "%1."),
    6: ("upperLetter", "（%1）"),
    7: ("lowerLetter", "%1."),
    8: ("lowerLetter", "（%1）"),
}


def build_recognizable_nine_level_numbering(*, pollute: bool = False):
    """Numbering root with one abstractNum carrying all 9 recognizable levels.

    When pollute=True, every level gets a wrong suffix, wrong/duplicated tab
    stops, and trailing whitespace in lvlText, so cleanup behaviour can be
    verified against the central rule.
    """
    root = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    abstract = etree.SubElement(root, qn("abstractNum"))
    abstract.set(qn("abstractNumId"), "1")
    for level, (num_fmt, lvl_text) in RECOGNIZABLE_LEVEL_FORMATS.items():
        lvl = etree.SubElement(abstract, qn("lvl"))
        lvl.set(qn("ilvl"), str(level))
        fmt = etree.SubElement(lvl, qn("numFmt"))
        fmt.set(qn("val"), num_fmt)
        text_el = etree.SubElement(lvl, qn("lvlText"))
        text_el.set(qn("val"), lvl_text + (" \t　" if pollute else ""))
        if pollute:
            suff = etree.SubElement(lvl, qn("suff"))
            suff.set(qn("val"), "space")  # deliberately wrong for every level
            pPr = etree.SubElement(lvl, qn("pPr"))
            tabs = etree.SubElement(pPr, qn("tabs"))
            for pos in ("111", "222"):  # wrong type, wrong position, duplicated
                tab = etree.SubElement(tabs, qn("tab"))
                tab.set(qn("val"), "num")
                tab.set(qn("pos"), pos)
    num = etree.SubElement(root, qn("num"))
    num.set(qn("numId"), "1")
    abstract_ref = etree.SubElement(num, qn("abstractNumId"))
    abstract_ref.set(qn("val"), "1")
    return root


def assert_level_suffix_rule(testcase, lvl, level):
    """Assert one numbering w:lvl matches the central tab-suffix rule."""
    suff = lvl.find("./w:suff", NS)
    tabs = lvl.find("./w:pPr/w:tabs", NS)
    if uses_tab_suffix(level):
        testcase.assertEqual(suff.get(qn("val")), "tab")
        testcase.assertIsNotNone(tabs)
        tab_list = tabs.findall("./w:tab", NS)
        testcase.assertEqual(len(tab_list), 1)
        testcase.assertEqual(tab_list[0].get(qn("val")), "left")
        testcase.assertEqual(tab_list[0].get(qn("pos")), TEMPLATE_OUTLINE_INDENTS[level]["heading_text_start"])
    else:
        testcase.assertEqual(suff.get(qn("val")), "nothing")
        testcase.assertIsNone(tabs)
    lvl_text = lvl.find("./w:lvlText", NS)
    if lvl_text is not None and lvl_text.get(qn("val")) is not None:
        testcase.assertFalse(lvl_text.get(qn("val")).endswith((" ", "\t", "　")))


def make_styles_font_xml(
    *,
    doc_default_pt: float | None = None,
    paragraph_styles: dict[str, tuple[float | None, str | None]] | None = None,
    character_styles: dict[str, tuple[float | None, str | None]] | None = None,
    paragraph_style_indents: dict[str, dict[str, str]] | None = None,
):
    root = etree.Element(qn("styles"), nsmap={"w": W_NS})
    if doc_default_pt is not None:
        doc_defaults = etree.SubElement(root, qn("docDefaults"))
        rpr_default = etree.SubElement(doc_defaults, qn("rPrDefault"))
        rPr = etree.SubElement(rpr_default, qn("rPr"))
        sz = etree.SubElement(rPr, qn("sz"))
        sz.set(qn("val"), str(round(doc_default_pt * 2)))

    for style_type, styles in (
        ("paragraph", paragraph_styles or {}),
        ("character", character_styles or {}),
    ):
        for style_id, (font_size_pt, based_on) in styles.items():
            style = etree.SubElement(root, qn("style"))
            style.set(qn("type"), style_type)
            style.set(qn("styleId"), style_id)
            if based_on is not None:
                based = etree.SubElement(style, qn("basedOn"))
                based.set(qn("val"), based_on)
            if font_size_pt is not None:
                rPr = etree.SubElement(style, qn("rPr"))
                sz = etree.SubElement(rPr, qn("sz"))
                sz.set(qn("val"), str(round(font_size_pt * 2)))
            indent_attrs = (paragraph_style_indents or {}).get(style_id) if style_type == "paragraph" else None
            if indent_attrs:
                pPr = etree.SubElement(style, qn("pPr"))
                ind = etree.SubElement(pPr, qn("ind"))
                for attr_name, attr_value in indent_attrs.items():
                    ind.set(qn(attr_name), attr_value)

    return etree.tostring(root)


class OutlineFixTests(unittest.TestCase):
    def test_template_indents_match_requested_number_start(self):
        self.assertTrue(validate_template_outline_indents())

    def test_template_indents_match_requested_cm_values(self):
        expected = {
            0: (1.23, -0.04, 1.27, 1.23),
            1: (1.86, 0.73, 1.13, 1.86),
            2: (2.99, 1.51, 1.48, 2.99),
            3: (3.99, 3.49, 0.50, 3.99),
            4: (4.97, 3.74, 1.23, 4.96),
            5: (5.95, 5.45, 0.50, 5.95),
            6: (5.93, 4.70, 1.23, 5.94),
            7: (6.43, 5.94, 0.49, 6.85),
            8: (8.96, 7.72, 1.24, 8.96),
        }

        for level, (left_cm, number_start_cm, hanging_cm, body_left_cm) in expected.items():
            with self.subTest(level=level):
                spec = TEMPLATE_OUTLINE_INDENTS[level]
                self.assertAlmostEqual(twips_to_cm(spec["left"]), left_cm, places=2)
                self.assertAlmostEqual(twips_to_cm(spec["number_start"]), number_start_cm, places=2)
                self.assertAlmostEqual(twips_to_cm(spec["hanging"]), hanging_cm, places=2)
                self.assertAlmostEqual(
                    twips_to_cm(spec["heading_text_start"]),
                    body_left_cm + 0.85,
                    places=2,
                )
                self.assertAlmostEqual(twips_to_cm(spec["body_left"]), body_left_cm, places=2)

    def test_custom_heading_text_start_controls_tab_without_changing_left_hanging_or_body(self):
        original = dict(TEMPLATE_OUTLINE_INDENTS[3])
        TEMPLATE_OUTLINE_INDENTS[3] = make_outline_indent_spec(
            number_start_cm=1.00,
            text_indent_cm=1.50,
            tab_stop_cm=3.75,
            body_left_cm=2.25,
        )
        spec = TEMPLATE_OUTLINE_INDENTS[3]
        try:
            p = make_paragraph("1. 自訂標題")
            pPr = p.find("./w:pPr", NS)
            written = apply_indent_spec_to_pPr(pPr, spec, "heading_numbered", use_tab_stop=True)

            ind = pPr.find("w:ind", NS)
            tab = pPr.find("./w:tabs/w:tab", NS)
            self.assertEqual(ind.get(qn("left")), spec["left"])
            self.assertEqual(ind.get(qn("hanging")), spec["hanging"])
            self.assertEqual(written["tab_pos"], spec["heading_text_start"])
            self.assertEqual(tab.get(qn("pos")), spec["heading_text_start"])
            self.assertNotEqual(tab.get(qn("pos")), spec["left"])

            body = make_paragraph("標題下方普通內文")
            body_pPr = body.find("./w:pPr", NS)
            add_tab_stop(body, "999")
            add_num_pr(body)
            apply_indent_spec_to_pPr(body_pPr, spec, "body_plain")
            assert_body_indent_hard_override(self, body, spec["body_left"])
        finally:
            TEMPLATE_OUTLINE_INDENTS[3] = original

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
        first_marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root.find("./w:body", NS).insert(4, first_marker)

        fix_outline_paragraphs(root, include_tables=True)

        for paragraph in paragraphs[:4]:
            self.assertIsNone(paragraph_indent(paragraph))
            self.assertIsNone(paragraph_outline(paragraph))
        self.assertIsNone(paragraph_indent(first_marker))
        self.assertIsNone(paragraph_outline(first_marker))

        for paragraph, level in zip(paragraphs[4:], [0, 1, 2, 3, 4]):
            self.assertEqual(paragraph_indent(paragraph), expected_indent(level))
            self.assertEqual(paragraph_outline(paragraph), str(level))

    def test_user_visible_level_two_text_is_set_to_16_pt(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00", font_size_pt=12)
        level_two = make_paragraph("\u4e00\u3001\u6a19\u984c", font_size_pt=12)
        level_three = make_paragraph("\uff08\u4e00\uff09\u6a19\u984c", font_size_pt=12)
        add_ind_with_char_attrs(level_two)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, level_two, level_three)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertEqual(OUTLINE_LEVEL_FONT_SIZE_PT[1], 16.0)
        self.assertEqual(paragraph_text_run_sizes(level_two), ["32"])
        self.assertEqual(paragraph_text_run_sizes(marker), ["24"])
        # Level 0 (壹、) is never forced; level 1 (一、) keeps 16 pt; level 2
        # (（一）) and deeper are forced to 14 pt = 28 half-points.
        self.assertEqual(paragraph_text_run_sizes(level_three), ["28"])
        assert_no_char_indent_attrs(self, level_two.find("./w:pPr/w:ind", NS))

    def test_auto_numbered_level_two_paragraph_text_is_set_to_16_pt(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00", font_size_pt=12)
        auto_level_two = make_paragraph(
            "\u81ea\u52d5\u7de8\u865f\u6a19\u984c",
            num_id="1",
            ilvl=1,
            font_size_pt=12,
        )
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, auto_level_two)

        fix_outline_paragraphs(
            root,
            include_tables=True,
            numbering_level_lookup={("1", 1): 1},
        )

        self.assertEqual(paragraph_text_run_sizes(auto_level_two), ["32"])
        self.assertEqual(paragraph_outline(auto_level_two), "1")

    def test_level_zero_financial_heading_keeps_original_font_size(self):
        # 壹、/貳、 (level 0) must never be forced; its original size survives.
        marker = make_paragraph("壹、序言", font_size_pt=12)
        level_zero = make_paragraph("貳、價格分析", font_size_pt=18)
        root = make_root(make_paragraph("壹、序言"), marker, level_zero)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertNotIn(0, OUTLINE_LEVEL_FONT_SIZE_PT)
        self.assertEqual(paragraph_outline(level_zero), "0")
        # Original 18 pt = 36 half-points stays; not rewritten to 28 (14 pt).
        self.assertEqual(paragraph_text_run_sizes(level_zero), ["36"])
        self.assertEqual(paragraph_text_run_szcs(level_zero), ["36"])

    def test_level_one_simple_heading_keeps_existing_16pt(self):
        # 一、 (level 1) keeps the pre-existing 16 pt rule, never 14 pt.
        marker = make_paragraph("壹、序言", font_size_pt=12)
        level_one = make_paragraph("一、標題", font_size_pt=12)
        root = make_root(make_paragraph("壹、序言"), marker, level_one)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertEqual(OUTLINE_LEVEL_FONT_SIZE_PT[1], 16.0)
        self.assertEqual(paragraph_outline(level_one), "1")
        self.assertEqual(paragraph_text_run_sizes(level_one), ["32"])
        self.assertEqual(paragraph_text_run_szcs(level_one), ["32"])

    def test_levels_two_through_eight_manual_headings_are_forced_to_14pt(self):
        # Every supported hierarchy level from 2 to the deepest (8) must end up
        # at 14 pt on the visible run, with both w:sz and w:szCs = 28.
        prefixes = {
            2: "（一）標題二",  # （一）標題二
            3: "1.標題三",                  # 1.標題三
            4: "（1）標題四",        # （1）標題四
            5: "A.標題五",                  # A.標題五
            6: "（A）標題六",        # （A）標題六
            7: "a.標題七",                  # a.標題七
            8: "（a）標題八",        # （a）標題八
        }
        headings = {
            level: make_paragraph(text, font_size_pt=12)
            for level, text in prefixes.items()
        }
        root = make_root(
            make_paragraph("壹、序言"),
            make_paragraph("壹、序言", font_size_pt=12),
            *[headings[level] for level in sorted(headings)],
        )

        fix_outline_paragraphs(root, include_tables=True)

        for level in sorted(headings):
            with self.subTest(level=level):
                self.assertEqual(OUTLINE_LEVEL_FONT_SIZE_PT[level], 14.0)
                p = headings[level]
                self.assertEqual(paragraph_outline(p), str(level))
                self.assertEqual(paragraph_text_run_sizes(p), ["28"])
                self.assertEqual(paragraph_text_run_szcs(p), ["28"])

    def test_numbering_xml_levels_two_through_eight_get_14pt_run_properties(self):
        # Auto-numbered headings get their number font from numbering.xml level
        # run properties: levels 2-8 -> 14 pt, levels 0/1 untouched by 14 pt.
        root = etree.Element(qn("numbering"), nsmap={"w": W_NS})
        level_formats = {
            0: ("ideographLegalTraditional", "%1、"),       # 壹、
            1: ("taiwaneseCountingThousand", "%1、"),       # 一、
            2: ("taiwaneseCountingThousand", "（%1）"),  # （一）
            3: ("decimal", "%1."),                              # 1.
            4: ("decimal", "（%1）"),                    # （1）
            5: ("upperLetter", "%1."),                          # A.
            6: ("upperLetter", "（%1）"),                # （A）
            7: ("lowerLetter", "%1."),                          # a.
            8: ("lowerLetter", "（%1）"),                # （a）
        }
        for level, (num_fmt, lvl_text) in level_formats.items():
            abstract_id = str(100 + level)
            abstract = etree.SubElement(root, qn("abstractNum"))
            abstract.set(qn("abstractNumId"), abstract_id)
            lvl = etree.SubElement(abstract, qn("lvl"))
            lvl.set(qn("ilvl"), "0")
            fmt = etree.SubElement(lvl, qn("numFmt"))
            fmt.set(qn("val"), num_fmt)
            text_el = etree.SubElement(lvl, qn("lvlText"))
            text_el.set(qn("val"), lvl_text)
            num = etree.SubElement(root, qn("num"))
            num.set(qn("numId"), abstract_id)
            abstract_ref = etree.SubElement(num, qn("abstractNumId"))
            abstract_ref.set(qn("val"), abstract_id)

        updated_root = etree.fromstring(apply_numbering_outline_format(etree.tostring(root)))

        def level_rpr(level):
            abstract_id = str(100 + level)
            lvl = updated_root.xpath(
                f"./w:abstractNum[@w:abstractNumId='{abstract_id}']/w:lvl",
                namespaces=NS,
            )[0]
            return lvl.find("./w:rPr", NS)

        # Level 0 (壹、) receives no forced numbering font size at all.
        self.assertIsNone(level_rpr(0))
        # Level 1 (一、) keeps the existing 16 pt = 32 half-points rule.
        self.assertEqual(level_rpr(1).find("./w:sz", NS).get(qn("val")), "32")
        self.assertEqual(level_rpr(1).find("./w:szCs", NS).get(qn("val")), "32")
        # Levels 2-8 are forced to 14 pt with both w:sz and w:szCs = 28.
        for level in range(2, 9):
            with self.subTest(level=level):
                rPr = level_rpr(level)
                self.assertIsNotNone(rPr)
                self.assertEqual(rPr.find("./w:sz", NS).get(qn("val")), "28")
                self.assertEqual(rPr.find("./w:szCs", NS).get(qn("val")), "28")

    def test_auto_numbering_debug_log_records_paragraph_and_level_positioning(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00", font_size_pt=12)
        auto_level_two = make_paragraph(
            "\u81ea\u52d5\u7de8\u865f\u6a19\u984c",
            num_id="1",
            ilvl=1,
            font_size_pt=12,
        )
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, auto_level_two)
        summary = ProcessSummary()
        spec = TEMPLATE_OUTLINE_INDENTS[1]
        numbering_format_lookup = {
            ("1", 1): {
                "left": spec["left"],
                "hanging": spec["hanging"],
                "number_start": str(int(spec["left"]) - int(spec["hanging"])),
                "heading_text_start": spec["heading_text_start"],
                "lvlJc": "left",
                "suff": "nothing",
                "tab_pos": None,
            }
        }

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            numbering_level_lookup={("1", 1): 1},
            numbering_format_lookup=numbering_format_lookup,
        )

        debug = "\n".join(summary.numbering_debug_logs)
        self.assertIn("kind=auto", debug)
        self.assertIn("numId=1", debug)
        self.assertIn("ilvl=1", debug)
        self.assertIn(f"p_left={spec['left']}", debug)
        self.assertIn(f"p_hanging={spec['hanging']}", debug)
        self.assertIn(f"lvl_left={spec['left']}", debug)
        self.assertIn(f"lvl_hanging={spec['hanging']}", debug)
        self.assertIn(f"lvl_heading_text_start={spec['heading_text_start']}", debug)
        self.assertIn("lvlJc=left", debug)
        self.assertIn("suff=nothing", debug)

    def test_build_numbering_level_lookup_has_no_ilvl_fallback(self):
        lookup = build_numbering_level_lookup(make_heading_validation_numbering_xml())

        self.assertEqual(lookup.get(("64", 0)), 4)
        # The unsupported decimal "%1" pattern must not fall back to ilvl=2.
        self.assertNotIn(("77", 2), lookup)

    def test_detect_valid_auto_heading_level_supported_format(self):
        numbering_xml = make_heading_validation_numbering_xml()
        p = make_paragraph("433-2地號土地開發分析法評估結果", num_id="64", ilvl=0)

        level, details = detect_valid_auto_heading_level(
            p,
            numbering_level_lookup=build_numbering_level_lookup(numbering_xml),
            numbering_format_lookup=build_numbering_format_lookup(numbering_xml),
        )

        self.assertEqual(level, 4)
        self.assertEqual(details["num_id"], "64")
        self.assertEqual(details["ilvl"], 0)
        self.assertEqual(details["num_fmt"], "decimal")
        self.assertEqual(details["lvl_text"], "（%1）")

    def test_detect_valid_auto_heading_level_rejects_unknown_bullet_and_unsupported(self):
        numbering_xml = make_heading_validation_numbering_xml()
        numbering_level_lookup = build_numbering_level_lookup(numbering_xml)
        numbering_format_lookup = build_numbering_format_lookup(numbering_xml)

        unknown = make_paragraph("殘留編號內文", num_id="0", ilvl=0)
        bullet = make_paragraph("項目符號內容", num_id="9", ilvl=0)
        unsupported = make_paragraph("不支援格式", num_id="77", ilvl=2)

        for paragraph, expected_num_id in ((unknown, "0"), (bullet, "9"), (unsupported, "77")):
            with self.subTest(num_id=expected_num_id):
                level, details = detect_valid_auto_heading_level(
                    paragraph,
                    numbering_level_lookup=numbering_level_lookup,
                    numbering_format_lookup=numbering_format_lookup,
                )
                self.assertIsNone(level)
                self.assertEqual(details["num_id"], expected_num_id)

    def test_leftover_numpr_body_paragraph_is_not_treated_as_heading(self):
        numbering_xml = make_heading_validation_numbering_xml()
        first_marker = make_paragraph("壹、序言")
        marker = make_paragraph("壹、序言")
        body = make_paragraph(
            "依上述計算之總銷售金額及相關開發費用後，依照土地開發分析法，推算土地開發分析價格如下。",
            num_id="0",
            ilvl=0,
            font_size_pt=14,
        )
        root = make_root(first_marker, marker, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            numbering_level_lookup=build_numbering_level_lookup(numbering_xml),
            numbering_format_lookup=build_numbering_format_lookup(numbering_xml),
        )

        # No heading outline level 0-8 and no heading indent.
        self.assertNotIn(paragraph_outline(body), [str(i) for i in range(9)])
        self.assertNotEqual(paragraph_indent(body), expected_indent(0))
        # The paragraph follows the current heading's body indent instead.
        self.assertEqual(paragraph_left_indent(body), TEMPLATE_OUTLINE_INDENTS[0]["body_left"])

        joined = "\n".join(summary.paragraph_logs)
        self.assertIn("auto numbering skipped: missing or unsupported numbering format", joined)
        self.assertIn("auto_heading_valid=False", joined)
        self.assertIn("numId=0", joined)
        self.assertIn("Body indent applied", joined)

    def test_valid_auto_numbering_heading_without_text_marker_is_level_four(self):
        numbering_xml = make_heading_validation_numbering_xml()
        first_marker = make_paragraph("壹、序言")
        marker = make_paragraph("壹、序言")
        heading = make_paragraph("433-2地號土地開發分析法評估結果", num_id="64", ilvl=0)
        root = make_root(first_marker, marker, heading)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            numbering_level_lookup=build_numbering_level_lookup(numbering_xml),
            numbering_format_lookup=build_numbering_format_lookup(numbering_xml),
        )

        # The marker "（6）" only exists in Word auto numbering, not in w:t
        # text, and the heading must still resolve to level 4 (階層 5).
        self.assertEqual(paragraph_outline(heading), "4")
        self.assertEqual(paragraph_indent(heading), expected_indent(4))
        self.assertIsNotNone(heading.find("./w:pPr/w:numPr", NS))

        joined = "\n".join(summary.paragraph_logs)
        self.assertIn("auto numbering valid", joined)
        self.assertIn("numId=64", joined)
        self.assertIn("numFmt=decimal", joined)
        self.assertIn("resolved_level=4", joined)
        self.assertIn("auto_heading_valid=True", joined)
        self.assertIn("apply level 5 outline and indent", joined)

    def test_invalid_auto_numbering_falls_back_to_manual_text_prefix(self):
        numbering_xml = make_heading_validation_numbering_xml()
        first_marker = make_paragraph("壹、序言")
        marker = make_paragraph("壹、序言")
        heading = make_paragraph(
            "（6）433-2地號土地開發分析法評估結果",
            num_id="0",
            ilvl=0,
        )
        root = make_root(first_marker, marker, heading)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            numbering_level_lookup=build_numbering_level_lookup(numbering_xml),
            numbering_format_lookup=build_numbering_format_lookup(numbering_xml),
        )

        self.assertEqual(paragraph_outline(heading), "4")
        self.assertEqual(paragraph_indent(heading), expected_indent(4))
        # The leftover numPr is removed so Word cannot renumber the heading.
        self.assertIsNone(heading.find("./w:pPr/w:numPr", NS))

        # The manual heading must not enter the invalid-auto-numbering body
        # cleanup, so no forced 標楷體 font rewrite happens.
        self.assertIsNone(heading.find("./w:r/w:rPr/w:rFonts", NS))
        self.assertIsNone(heading.find("./w:pPr/w:rPr/w:rFonts", NS))

        joined = "\n".join(summary.paragraph_logs)
        self.assertIn("Manual heading paragraph numPr removed", joined)
        self.assertIn("manual numbering; apply level 5 outline and indent", joined)
        self.assertNotIn("invalid auto numbering body paragraph", joined)
        self.assertNotIn("font normalized", joined)

    def test_bullet_numbering_paragraph_is_never_a_heading(self):
        numbering_xml = make_heading_validation_numbering_xml()
        first_marker = make_paragraph("壹、序言")
        marker = make_paragraph("壹、序言")
        bullet = make_paragraph("項目符號內容", num_id="9", ilvl=0, font_size_pt=14)
        root = make_root(first_marker, marker, bullet)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            numbering_level_lookup=build_numbering_level_lookup(numbering_xml),
            numbering_format_lookup=build_numbering_format_lookup(numbering_xml),
        )

        self.assertNotIn(paragraph_outline(bullet), [str(i) for i in range(9)])
        self.assertNotEqual(paragraph_indent(bullet), expected_indent(0))
        joined = "\n".join(summary.paragraph_logs)
        self.assertIn("auto numbering skipped: missing or unsupported numbering format", joined)
        self.assertIn("numFmt=bullet", joined)

    def test_invalid_auto_numbering_body_paragraph_removes_numbering_style(self):
        numbering_xml = make_heading_validation_numbering_xml()
        first_marker = make_paragraph("壹、序言")
        marker = make_paragraph("壹、序言")
        body = make_paragraph(
            "依上述計算之總銷售金額及相關開發費用後，依照土地開發分析法，推算土地開發分析價格如下。",
            style="51",
            outline=3,
            num_id="0",
            ilvl=0,
            font_size_pt=14,
        )
        # Existing run properties such as bold must survive the cleanup.
        etree.SubElement(body.find("./w:r/w:rPr", NS), qn("b"))
        root = make_root(first_marker, marker, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            numbering_level_lookup=build_numbering_level_lookup(numbering_xml),
            numbering_format_lookup=build_numbering_format_lookup(numbering_xml),
            # Style 51 carries numbering, so Word would re-apply "1." from the
            # style even after the paragraph-level numPr is removed.
            style_numbering_lookup={"51": ("0", 0)},
        )

        self.assertIsNone(body.find("./w:pPr/w:numPr", NS))
        self.assertIsNone(body.find("./w:pPr/w:pStyle", NS))
        self.assertEqual(paragraph_outline(body), "9")
        self.assertEqual(paragraph_left_indent(body), TEMPLATE_OUTLINE_INDENTS[0]["body_left"])

        # The removed style carried the fonts, so 標楷體 14pt is written
        # directly on the paragraph mark and every visible text run.
        p_rpr_fonts = body.find("./w:pPr/w:rPr/w:rFonts", NS)
        self.assertIsNotNone(p_rpr_fonts)
        for attr in ("ascii", "eastAsia", "hAnsi", "cs"):
            self.assertEqual(p_rpr_fonts.get(qn(attr)), "標楷體")
        self.assertEqual(body.find("./w:pPr/w:rPr/w:sz", NS).get(qn("val")), "28")
        self.assertEqual(body.find("./w:pPr/w:rPr/w:szCs", NS).get(qn("val")), "28")

        run_rpr = body.find("./w:r/w:rPr", NS)
        run_fonts = run_rpr.find("w:rFonts", NS)
        self.assertIsNotNone(run_fonts)
        for attr in ("ascii", "eastAsia", "hAnsi", "cs"):
            self.assertEqual(run_fonts.get(qn(attr)), "標楷體")
        self.assertEqual(run_rpr.find("w:sz", NS).get(qn("val")), "28")
        self.assertEqual(run_rpr.find("w:szCs", NS).get(qn("val")), "28")
        self.assertIsNotNone(run_rpr.find("w:b", NS))

        record = summary.body_indent_records[-1]
        self.assertTrue(record["font_normalized_after_numbering_cleanup"])
        self.assertEqual(record["normalized_font_name"], "標楷體")
        self.assertEqual(record["normalized_font_size_pt"], 14.0)

        joined = "\n".join(summary.paragraph_logs)
        self.assertIn(
            "invalid auto numbering body paragraph: removed paragraph numPr and numbering style",
            joined,
        )
        self.assertIn("paragraph_style_id_before=51", joined)
        self.assertIn("paragraph_style_id_after=none", joined)
        self.assertIn("style_numbering_removed=True", joined)
        self.assertIn("reason=style-level numbering would reappear in Word", joined)
        self.assertIn(
            "invalid auto numbering body paragraph font normalized: font=標楷體; size=14pt",
            joined,
        )
        self.assertIn("reason=numbering/style cleanup removed inherited formatting", joined)

    def test_invalid_auto_numbering_body_paragraph_keeps_plain_style(self):
        numbering_xml = make_heading_validation_numbering_xml()
        first_marker = make_paragraph("壹、序言")
        marker = make_paragraph("壹、序言")
        body = make_paragraph(
            "依上述計算之總銷售金額及相關開發費用後，依照土地開發分析法，推算土地開發分析價格如下。",
            style="BodyText",
            num_id="0",
            ilvl=0,
            font_size_pt=14,
        )
        root = make_root(first_marker, marker, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            numbering_level_lookup=build_numbering_level_lookup(numbering_xml),
            numbering_format_lookup=build_numbering_format_lookup(numbering_xml),
            # No style in the lookup carries numbering.
            style_numbering_lookup={},
        )

        self.assertIsNone(body.find("./w:pPr/w:numPr", NS))
        self.assertEqual(paragraph_style(body), "BodyText")
        self.assertEqual(paragraph_left_indent(body), TEMPLATE_OUTLINE_INDENTS[0]["body_left"])

        # The style still provides the fonts, so no forced font rewrite: the
        # run keeps its original properties and gains no rFonts.
        self.assertIsNone(body.find("./w:r/w:rPr/w:rFonts", NS))
        self.assertIsNone(body.find("./w:pPr/w:rPr/w:rFonts", NS))
        self.assertEqual(body.find("./w:r/w:rPr/w:sz", NS).get(qn("val")), "28")
        record = summary.body_indent_records[-1]
        self.assertFalse(record["font_normalized_after_numbering_cleanup"])
        self.assertIsNone(record["normalized_font_name"])

        joined = "\n".join(summary.paragraph_logs)
        self.assertNotIn("invalid auto numbering body paragraph", joined)
        self.assertNotIn("font normalized", joined)

    def test_valid_auto_numbering_heading_keeps_style_and_numbering(self):
        numbering_xml = make_heading_validation_numbering_xml()
        first_marker = make_paragraph("壹、序言")
        marker = make_paragraph("壹、序言")
        heading = make_paragraph(
            "433-2地號土地開發分析法評估結果",
            style="51",
            num_id="64",
            ilvl=0,
        )
        root = make_root(first_marker, marker, heading)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            numbering_level_lookup=build_numbering_level_lookup(numbering_xml),
            numbering_format_lookup=build_numbering_format_lookup(numbering_xml),
            style_numbering_lookup={"51": ("0", 0)},
        )

        # The valid auto numbering heading keeps its style and numbering.
        self.assertEqual(paragraph_outline(heading), "4")
        self.assertIsNotNone(heading.find("./w:pPr/w:numPr", NS))
        self.assertEqual(paragraph_style(heading), "51")

        # The body-font cleanup flow must not touch a valid heading.
        self.assertIsNone(heading.find("./w:r/w:rPr/w:rFonts", NS))
        self.assertIsNone(heading.find("./w:pPr/w:rPr/w:rFonts", NS))

        joined = "\n".join(summary.paragraph_logs)
        self.assertIn("auto numbering valid", joined)
        self.assertNotIn("invalid auto numbering body paragraph", joined)
        self.assertNotIn("font normalized", joined)

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
        first_marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(toc, preface, first_marker, marker)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertIsNone(paragraph_indent(toc))
        self.assertIsNone(paragraph_outline(toc))
        self.assertIsNone(paragraph_indent(preface))
        self.assertIsNone(paragraph_outline(preface))
        self.assertIsNone(paragraph_indent(first_marker))
        self.assertIsNone(paragraph_outline(first_marker))
        self.assertEqual(paragraph_outline(marker), "0")

    def test_plain_toc_range_is_skipped_until_body_start(self):
        toc_heading = make_paragraph("\u76ee\u9304")
        toc_entry = make_paragraph("\u4e00\u3001\u5e8f\u8a00")
        first_marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(toc_heading, toc_entry, first_marker, marker)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertIsNone(paragraph_indent(toc_heading))
        self.assertIsNone(paragraph_outline(toc_heading))
        self.assertIsNone(paragraph_indent(toc_entry))
        self.assertIsNone(paragraph_outline(toc_entry))
        self.assertIsNone(paragraph_indent(first_marker))
        self.assertIsNone(paragraph_outline(first_marker))
        self.assertEqual(paragraph_indent(marker), expected_indent(0))
        self.assertEqual(paragraph_outline(marker), "0")

    def test_main_outline_starts_only_on_second_processing_start_marker(self):
        first_marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        preface_after_first = make_paragraph("\u4e00\u3001\u524d\u7f6e\u5167\u5bb9", outline=2)
        second_marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        body_after_second = make_paragraph("\u9019\u662f\u7b2c\u4e8c\u6b21\u5e8f\u8a00\u5f8c\u7684 14 pt \u5167\u6587", font_size_pt=14)
        root = make_root(first_marker, preface_after_first, second_marker, body_after_second)
        summary = ProcessSummary()

        fix_outline_paragraphs(root, include_tables=True, summary=summary)

        self.assertIsNone(paragraph_indent(first_marker))
        self.assertIsNone(paragraph_outline(first_marker))
        self.assertIsNone(paragraph_indent(preface_after_first))
        self.assertEqual(paragraph_outline(preface_after_first), "2")
        self.assertEqual(paragraph_indent(second_marker), expected_indent(0))
        self.assertEqual(paragraph_outline(second_marker), "0")
        self.assertEqual(paragraph_left_indent(body_after_second), TEMPLATE_OUTLINE_INDENTS[0]["body_left"])
        joined_logs = "\n".join(summary.paragraph_logs)
        self.assertIn("processing start marker seen count=1/2; still before main body", joined_logs)
        self.assertIn("processing start marker seen count=2/2; main outline started", joined_logs)

    def test_single_processing_start_marker_does_not_start_main_outline(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        after_marker = make_paragraph("\u4e00\u3001\u4ecd\u662f\u524d\u7f6e\u5167\u5bb9")
        root = make_root(marker, after_marker)
        summary = ProcessSummary()

        fix_outline_paragraphs(root, include_tables=True, summary=summary)

        self.assertIsNone(paragraph_indent(marker))
        self.assertIsNone(paragraph_outline(marker))
        self.assertIsNone(paragraph_indent(after_marker))
        self.assertIsNone(paragraph_outline(after_marker))
        joined_logs = "\n".join(summary.paragraph_logs)
        self.assertIn("processing start marker seen count=1/2; still before main body", joined_logs)
        self.assertNotIn("main outline started", joined_logs)

    def test_styles_xml_toc_styles_are_not_modified(self):
        root = etree.Element(qn("styles"), nsmap={"w": W_NS})
        originals: dict[str, bytes] = {}
        for style_id, name in [
            ("TOC1", "Table of Contents 1"),
            ("TOC9", "Contents level 9"),
            ("TOCHeading", "TOC Heading"),
            ("\u76ee\u9304", "\u76ee\u9304"),
        ]:
            style = etree.SubElement(root, qn("style"))
            style.set(qn("type"), "paragraph")
            style.set(qn("styleId"), style_id)
            name_el = etree.SubElement(style, qn("name"))
            name_el.set(qn("val"), name)
            pPr = etree.SubElement(style, qn("pPr"))
            num_pr = etree.SubElement(pPr, qn("numPr"))
            ilvl = etree.SubElement(num_pr, qn("ilvl"))
            ilvl.set(qn("val"), "0")
            num_id = etree.SubElement(num_pr, qn("numId"))
            num_id.set(qn("val"), "1")
            ind = etree.SubElement(pPr, qn("ind"))
            ind.set(qn("left"), "111")
            ind.set(qn("leftChars"), "222")
            originals[style_id] = etree.tostring(style)
        logs: list[str] = []

        apply_styles_outline_format_to_root(
            root,
            numbering_level_lookup={("1", 0): 0},
            style_numbering_lookup={style_id: ("1", 0) for style_id in originals},
            change_logs=logs,
        )

        for style_id, original_xml in originals.items():
            with self.subTest(style_id=style_id):
                style = root.xpath(f"./w:style[@w:styleId='{style_id}']", namespaces=NS)[0]
                self.assertEqual(etree.tostring(style), original_xml)
        joined_logs = "\n".join(logs)
        self.assertIn("STYLES_XML_SKIP_TOC_STYLE: styleId=TOC1", joined_logs)
        self.assertIn("STYLES_XML_SKIP_TOC_STYLE: styleId=TOC9", joined_logs)
        self.assertIn("STYLES_XML_SKIP_TOC_STYLE: styleId=TOCHeading", joined_logs)
        self.assertNotIn("STYLES_XML_NUMBERED_STYLE_INDENT", joined_logs)

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

    def test_auto_numbering_before_preface_is_unchanged_by_default(self):
        auto_before = make_paragraph("\u81ea\u52d5\u7de8\u865f\u524d", num_id="1", ilvl=1)
        first_marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        auto_after = make_paragraph("\u81ea\u52d5\u7de8\u865f\u5f8c", num_id="1", ilvl=1)
        root = make_root(auto_before, first_marker, marker, auto_after)

        fix_outline_paragraphs(
            root,
            include_tables=True,
            numbering_level_lookup={("1", 1): 1},
        )

        self.assertIsNone(paragraph_indent(auto_before))
        self.assertIsNone(paragraph_outline(auto_before))
        self.assertIsNone(paragraph_indent(first_marker))
        self.assertIsNone(paragraph_outline(first_marker))
        self.assertEqual(paragraph_outline(marker), "0")
        self.assertEqual(paragraph_indent(auto_after), expected_indent(1))
        self.assertEqual(paragraph_outline(auto_after), "1")

    def test_preface_numbering_is_untouched_when_new_preface_options_are_off(self):
        before_one = make_paragraph("\u4e00\u3001\u76ee\u9304\u5f8c\u7684\u524d\u7f6e\u9805", outline=1)
        before_nested = make_paragraph("\uff08\u4e00\uff09\u524d\u7f6e\u5167\u5c64", outline=2)
        first_marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        after = make_paragraph("\u4e00\u3001\u5e8f\u8a00\u5167\u5c64")
        root = make_root(before_one, before_nested, first_marker, marker, after)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
        )

        self.assertIsNone(paragraph_indent(before_one))
        self.assertEqual(paragraph_outline(before_one), "1")
        self.assertIsNone(paragraph_indent(before_nested))
        self.assertEqual(paragraph_outline(before_nested), "2")
        self.assertIsNone(paragraph_indent(first_marker))
        self.assertIsNone(paragraph_outline(first_marker))
        self.assertEqual(paragraph_outline(marker), "0")
        self.assertEqual(paragraph_outline(after), "1")
        self.assertEqual(summary.indented_preface_paragraphs, 0)
        self.assertEqual(summary.outlined_preface_paragraphs, 0)

    def test_indent_preface_only_uses_preface_indents_without_adding_outline(self):
        before_decimal = make_paragraph("1. \u524d\u7f6e\u7b2c\u4e09\u968e", outline=3)
        before_parenthesized_decimal = make_paragraph("\uff081\uff09\u524d\u7f6e\u7b2c\u56db\u968e", outline=4)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(before_decimal, before_parenthesized_decimal, marker)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            indent_preface_paragraphs=True,
        )

        self.assertEqual(paragraph_indent(before_decimal), expected_preface_indent(2))
        self.assertEqual(paragraph_outline(before_decimal), "3")
        self.assertEqual(paragraph_indent(before_parenthesized_decimal), expected_preface_indent(3))
        self.assertEqual(paragraph_outline(before_parenthesized_decimal), "4")
        self.assertEqual(summary.indented_preface_paragraphs, 2)
        self.assertEqual(summary.outlined_preface_paragraphs, 0)

    def test_outline_preface_only_adds_outline_without_changing_indent(self):
        before_one = make_paragraph("\u4e00\u3001\u524d\u7f6e\u9805")
        before_nested = make_paragraph("\uff08\u4e00\uff09\u524d\u7f6e\u5167\u5c64")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(before_one, before_nested, marker)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            outline_preface_paragraphs=True,
        )

        self.assertIsNone(paragraph_indent(before_one))
        self.assertEqual(paragraph_outline(before_one), "0")
        self.assertIsNone(paragraph_indent(before_nested))
        self.assertEqual(paragraph_outline(before_nested), "1")
        self.assertEqual(summary.indented_preface_paragraphs, 0)
        self.assertEqual(summary.outlined_preface_paragraphs, 2)

    def test_indent_and_outline_preface_apply_both_preface_rules(self):
        before_one = make_paragraph("\u4e00\u3001\u524d\u7f6e\u9805")
        before_nested = make_paragraph("\uff08\u4e00\uff09\u524d\u7f6e\u5167\u5c64")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(before_one, before_nested, marker)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            indent_preface_paragraphs=True,
            outline_preface_paragraphs=True,
        )

        self.assertEqual(paragraph_indent(before_one), expected_preface_indent(0))
        self.assertEqual(paragraph_outline(before_one), "0")
        self.assertEqual(paragraph_indent(before_nested), expected_preface_indent(1))
        self.assertEqual(paragraph_outline(before_nested), "1")
        self.assertEqual(summary.indented_preface_paragraphs, 2)
        self.assertEqual(summary.outlined_preface_paragraphs, 2)

    def test_body_paragraph_after_heading_aligns_to_heading_text_start(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("\u4e00\u3001\u5e8f\u8a00\u5167\u5c64")
        body = make_paragraph("\u9019\u662f\u6a19\u984c\u4e0b\u65b9\u5167\u6587", font_size_pt=14)
        nested_heading = make_paragraph("\uff08\u4e00\uff09\u5167\u5c64\u6a19\u984c")
        nested_body = make_paragraph("\u9019\u662f\u5167\u5c64\u6a19\u984c\u4e0b\u65b9\u5167\u6587", font_size_pt=14)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body, nested_heading, nested_body)
        summary = ProcessSummary()

        fix_outline_paragraphs(root, include_tables=True, summary=summary)

        assert_body_indent_hard_override(self, body, expected_body_left(1))
        assert_body_indent_hard_override(self, nested_body, expected_body_left(2))
        self.assertIsNone(paragraph_outline(body))
        self.assertIsNone(paragraph_outline(nested_body))
        self.assertTrue(
            any(record["prefix"] == "\u4e00\u3001" for record in summary.numbering_measurements.values())
        )
        self.assertTrue(
            all(float(record["number_size_cm"]) > 0 for record in summary.numbering_measurements.values())
        )

    def test_note_paragraphs_skip_body_indent_without_affecting_body_or_manual_headings(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("\u4e00\u3001\u7814\u7a76\u76ee\u7684")
        note_colon = make_paragraph(
            "\u8a3b\uff1a\u9019\u662f\u8aaa\u660e",
            style="NoteStyle",
            outline=5,
            font_size_pt=14,
        )
        note_number = make_paragraph(
            "  \u8a3b1\uff1a\u9019\u662f\u8aaa\u660e",
            style="NoteStyle",
            num_id="9",
            ilvl=0,
            font_size_pt=14,
        )
        note_chinese = make_paragraph("\u8a3b\u4e00\uff1a\u9019\u662f\u8aaa\u660e", style="NoteStyle", font_size_pt=14)
        body = make_paragraph("\u9019\u662f\u666e\u901a 14pt \u5167\u6587", font_size_pt=14)
        manual_heading = make_paragraph("\uff08\u4e00\uff09\u4e00\u822c\u624b\u52d5\u7de8\u865f\u6a19\u984c")
        table, table_note = make_table_paragraph("\u8a3b\uff1a\u8868\u683c\u5167\u8aaa\u660e")

        for note in (note_colon, note_number, note_chinese):
            add_ind_with_char_attrs(note)
            add_tab_stop(note, pos="1480")

        self.assertIsNone(paragraph_jc(note_colon))
        self.assertIsNone(paragraph_jc(note_number))
        self.assertIsNone(paragraph_jc(note_chinese))
        self.assertIsNone(paragraph_jc(table_note))

        root = make_root(
            make_paragraph("\u58f9\u3001\u5e8f\u8a00"),
            marker,
            heading,
            note_colon,
            note_number,
            note_chinese,
            body,
            manual_heading,
            table,
        )
        summary = ProcessSummary()

        fix_outline_paragraphs(root, include_tables=True, summary=summary)

        self.assertTrue(is_note_paragraph("\u8a3b"))
        self.assertTrue(is_note_paragraph("\u8a3b\uff1a"))
        self.assertTrue(is_note_paragraph("\u8a3b1"))
        self.assertTrue(is_note_paragraph("\u8a3b1\uff1a"))
        self.assertTrue(is_note_paragraph("\u8a3b\u4e00"))
        self.assertTrue(is_note_paragraph("\u8a3b\u4e00\uff1a"))
        self.assertTrue(is_note_paragraph("\u8a3b\u3001\u9019\u662f\u8aaa\u660e"))
        self.assertTrue(is_note_paragraph("  \u8a3b \u9019\u662f\u8aaa\u660e"))

        for note in (note_colon, note_number, note_chinese):
            ind = note.find("./w:pPr/w:ind", NS)
            self.assertEqual(ind.get(qn("left")), "123")
            self.assertEqual(ind.get(qn("hanging")), "45")
            self.assertEqual(ind.get(qn("leftChars")), "99")
            self.assertEqual(ind.get(qn("firstLineChars")), "99")
            self.assertIsNotNone(note.find("./w:pPr/w:tabs", NS))
            self.assertEqual(paragraph_style(note), "NoteStyle")
            self.assertEqual(paragraph_jc(note), "left")

        note_colon_ppr_tags = [child.tag for child in note_colon.find("./w:pPr", NS)]
        self.assertLess(note_colon_ppr_tags.index(qn("jc")), note_colon_ppr_tags.index(qn("outlineLvl")))
        self.assertEqual(paragraph_outline(note_colon), "5")
        self.assertIsNotNone(note_number.find("./w:pPr/w:numPr", NS))
        self.assertIsNone(paragraph_jc(table_note))

        assert_body_indent_hard_override(self, body, TEMPLATE_OUTLINE_INDENTS[1]["body_left"])
        self.assertEqual(paragraph_outline(heading), "1")
        self.assertEqual(paragraph_left_indent(heading), TEMPLATE_OUTLINE_INDENTS[1]["left"])
        self.assertEqual(paragraph_outline(manual_heading), "2")
        self.assertEqual(paragraph_left_indent(manual_heading), TEMPLATE_OUTLINE_INDENTS[2]["left"])

        body_indent_texts = {record["text_preview"] for record in summary.body_indent_records}
        self.assertIn("\u9019\u662f\u666e\u901a 14pt \u5167\u6587", body_indent_texts)
        self.assertFalse(any(text.lstrip().startswith("\u8a3b") for text in body_indent_texts))
        joined_logs = "\n".join(summary.paragraph_logs)
        self.assertEqual(joined_logs.count("skipped note paragraph; forced left alignment"), 3)
        self.assertIn("before_jc=none; after_jc=left; paragraph_style_id=NoteStyle; has_numPr=False", joined_logs)
        self.assertIn("before_jc=none; after_jc=left; paragraph_style_id=NoteStyle; has_numPr=True", joined_logs)

    def test_body_after_level_one_heading_aligns_to_heading_text_start_only(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("\u4e00\u3001\u7814\u7a76\u76ee\u7684")
        body = make_paragraph("\u9019\u662f\u666e\u901a\u5167\u6587", font_size_pt=14)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertEqual(paragraph_left_indent(heading), TEMPLATE_OUTLINE_INDENTS[1]["left"])
        assert_body_indent_hard_override(self, body, TEMPLATE_OUTLINE_INDENTS[1]["body_left"])
        self.assertIsNone(paragraph_outline(body))

    def test_body_after_level_two_heading_aligns_to_heading_text_start_only(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("\uff08\u4e00\uff09\u7814\u7a76\u65b9\u6cd5")
        body = make_paragraph("\u9019\u662f\u666e\u901a\u5167\u6587", font_size_pt=14)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertEqual(paragraph_left_indent(heading), TEMPLATE_OUTLINE_INDENTS[2]["left"])
        self.assertIsNone(paragraph_start_indent(heading))
        assert_body_indent_hard_override(self, body, TEMPLATE_OUTLINE_INDENTS[2]["body_left"])
        self.assertIsNone(paragraph_outline(body))

    def test_level_three_body_indent_hard_overrides_old_indent_attrs(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("\uff08\u4e00\uff09\u7b2c\u4e09\u968e\u6a19\u984c")
        body = make_paragraph("\u968e\u5c643\u5167\u6587", font_size_pt=14)
        ind = add_ind_with_char_attrs(body)
        ind.set(qn("left"), "1480")
        ind.set(qn("start"), "1480")
        ind.set(qn("hanging"), "300")
        ind.set(qn("firstLine"), "200")
        ind.set(qn("leftChars"), "100")
        ind.set(qn("startChars"), "100")
        ind.set(qn("hangingChars"), "100")
        ind.set(qn("firstLineChars"), "100")
        add_tab_stop(body, pos="1480")
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(root, include_tables=True, summary=summary)

        expected = TEMPLATE_OUTLINE_INDENTS[2]["body_left"]
        body_ind = body.find("./w:pPr/w:ind", NS)
        assert_body_indent_hard_override(self, body, expected)
        self.assertAlmostEqual(twips_to_cm(expected), 2.99, places=2)
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("heading_level=2", debug)
        self.assertIn("spec_body_left_cm=2.99", debug)
        self.assertIn(f"written_left_twips={expected}", debug)
        self.assertIn(f"written_start_twips={expected}", debug)
        self.assertIn("written_firstLine_twips=0", debug)
        self.assertIn("written_hanging_twips=0", debug)
        self.assertIn("written_leftChars=0", debug)
        self.assertIn("written_startChars=0", debug)
        self.assertIn("written_firstLineChars=0", debug)
        self.assertIn("written_hangingChars=0", debug)
        self.assertIn("force_body_indent_hard_override=True", debug)
        self.assertIn("body_first_line_twips_applied=False", debug)
        self.assertIn("removed_tabs=True", debug)
        self.assertIn("validation=ok", debug)

    def test_level_two_body_indent_uses_body_left_and_first_line_560_twips(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("\u4e00\u3001\u7b2c\u4e8c\u968e\u6a19\u984c")
        body = make_paragraph("\u9019\u662f\u7b2c\u4e8c\u968e\u6a19\u984c\u4e0b\u65b9 14 pt \u5167\u6587", font_size_pt=14)
        ind = add_ind_with_char_attrs(body)
        ind.set(qn("left"), "1480")
        ind.set(qn("firstLine"), "111")
        ind.set(qn("firstLineChars"), "100")
        add_tab_stop(body, pos="1480")
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            change_logs=summary.paragraph_logs,
            enable_level1_level2_body_first_line_indent=True,
        )

        expected_left = TEMPLATE_OUTLINE_INDENTS[1]["body_left"]
        assert_body_indent_hard_override(self, body, expected_left, expected_first_line="560")
        self.assertAlmostEqual(twips_to_cm(expected_left), 1.86, places=2)
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("heading_level=1", debug)
        self.assertIn("spec_body_left_cm=1.86", debug)
        self.assertIn("spec_firstLine_twips=560", debug)
        self.assertIn("written_firstLine_twips=560", debug)
        self.assertIn("written_firstLineChars=0", debug)
        self.assertIn("body_first_line_twips_applied=True", debug)
        self.assertIn("validation=ok", debug)
        self.assertTrue(any("Body indent applied: left=" in line and "firstLine=560 twips" in line for line in summary.paragraph_logs))

    def test_level_one_body_indent_uses_body_left_and_first_line_560_twips(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("\u8cb3\u3001\u7b2c\u4e00\u968e\u6a19\u984c")
        body = make_paragraph("\u9019\u662f\u7b2c\u4e00\u968e\u6a19\u984c\u4e0b\u65b9 14 pt \u5167\u6587", font_size_pt=14)
        ind = add_ind_with_char_attrs(body)
        ind.set(qn("left"), "1480")
        ind.set(qn("firstLine"), "111")
        ind.set(qn("firstLineChars"), "100")
        add_tab_stop(body, pos="1480")
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            change_logs=summary.paragraph_logs,
            enable_level1_level2_body_first_line_indent=True,
        )

        expected_left = TEMPLATE_OUTLINE_INDENTS[0]["body_left"]
        assert_body_indent_hard_override(self, body, expected_left, expected_first_line="560")
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("heading_level=0", debug)
        self.assertIn("spec_firstLine_twips=560", debug)
        self.assertIn("written_firstLine_twips=560", debug)

    def test_level_four_body_indent_uses_body_left_and_removes_hanging_and_tabs(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph("\u9019\u662f\u7b2c\u56db\u968e\u4e0b\u65b9 14 pt \u5167\u6587", font_size_pt=14)
        ind = add_ind_with_char_attrs(body)
        ind.set(qn("firstLine"), "111")
        add_tab_stop(body, pos="1990")
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(root, include_tables=True, summary=summary)

        expected_left = TEMPLATE_OUTLINE_INDENTS[3]["body_left"]
        assert_body_indent_hard_override(self, body, expected_left)
        self.assertAlmostEqual(twips_to_cm(expected_left), 3.99, places=2)
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("heading_level=3", debug)
        self.assertIn("spec_body_left_cm=3.99", debug)
        self.assertIn("spec_firstLine_twips=None", debug)
        self.assertIn(f"written_left_twips={expected_left}", debug)
        self.assertIn("tab_pos=None", debug)
        self.assertTrue(any("Body indent applied: left=" in line and "firstLine cleared" in line for line in summary.paragraph_logs))

    def test_body_indent_hard_override_removes_numPr_and_makes_char_indent_paragraphs_identical(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body_a = make_paragraph("\u6709\u5b57\u5143\u55ae\u4f4d\u7e2e\u6392\u7684 14 pt \u5167\u6587", font_size_pt=14)
        body_b = make_paragraph("\u6c92\u6709\u5b57\u5143\u55ae\u4f4d\u7e2e\u6392\u7684 14 pt \u5167\u6587", font_size_pt=14)
        ind_a = add_ind_with_char_attrs(body_a)
        ind_a.set(qn("leftChars"), "100")
        ind_a.set(qn("firstLineChars"), "100")
        add_tab_stop(body_a, pos="1990")
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body_a, body_b)
        summary = ProcessSummary()

        fix_outline_paragraphs(root, include_tables=True, summary=summary)

        expected_left = TEMPLATE_OUTLINE_INDENTS[3]["body_left"]
        assert_body_indent_hard_override(self, body_a, expected_left)
        assert_body_indent_hard_override(self, body_b, expected_left)
        attrs = [
            "left",
            "start",
            "firstLine",
            "hanging",
            "leftChars",
            "startChars",
            "firstLineChars",
            "hangingChars",
        ]
        ind_a = body_a.find("./w:pPr/w:ind", NS)
        ind_b = body_b.find("./w:pPr/w:ind", NS)
        self.assertEqual(
            {attr: ind_a.get(qn(attr)) for attr in attrs},
            {attr: ind_b.get(qn(attr)) for attr in attrs},
        )
        self.assertIsNone(body_a.find("./w:pPr/w:tabs", NS))
        self.assertIsNone(body_b.find("./w:pPr/w:tabs", NS))
        self.assertIsNone(body_a.find("./w:pPr/w:numPr", NS))
        self.assertIsNone(body_b.find("./w:pPr/w:numPr", NS))
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("removed_tabs=True", debug)
        self.assertIn("removed_tabs=True", debug)

    def test_body_plain_indent_spec_removes_numPr_when_hard_overriding(self):
        p = make_paragraph("\u666e\u901a\u5167\u6587", font_size_pt=14)
        pPr = p.find("./w:pPr", NS)
        add_num_pr(p)

        result = apply_indent_spec_to_pPr(pPr, TEMPLATE_OUTLINE_INDENTS[3], "body_plain")

        expected_left = TEMPLATE_OUTLINE_INDENTS[3]["body_left"]
        assert_body_indent_hard_override(self, p, expected_left)
        self.assertEqual(result["removed_numPr"], "True")

    def test_body_indent_keeps_matching_body_style_by_default(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph("\u9019\u662f affe \u6a23\u5f0f 14 pt \u5167\u6587", style="affe", font_size_pt=14)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(root, include_tables=True, summary=summary)

        self.assertEqual(paragraph_style(body), "affe")
        self.assertEqual(paragraph_left_indent(body), TEMPLATE_OUTLINE_INDENTS[3]["body_left"])
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("normalize_body_style_to_none=False", debug)
        self.assertIn("paragraph_style_id_before=affe", debug)
        self.assertIn("paragraph_style_id_after=affe", debug)
        self.assertIn("body_style_normalized=False", debug)

    def test_body_indent_can_normalize_body_style_to_none_when_enabled(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph("\u9019\u662f affe \u6a23\u5f0f 14 pt \u5167\u6587", style="affe", font_size_pt=14)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            normalize_body_style_to_none=True,
        )

        self.assertIsNone(paragraph_style(body))
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("normalize_body_style_to_none=True", debug)
        self.assertIn("paragraph_style_id_after=none", debug)
        self.assertIn("body_style_normalized=True", debug)

    def test_body_indent_can_keep_original_body_style_when_normalization_is_disabled(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph("\u9019\u662f affe \u6a23\u5f0f 14 pt \u5167\u6587", style="affe", font_size_pt=14)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            normalize_body_style_to_none=False,
        )

        self.assertEqual(paragraph_style(body), "affe")
        self.assertEqual(paragraph_left_indent(body), TEMPLATE_OUTLINE_INDENTS[3]["body_left"])
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("normalize_body_style_to_none=False", debug)
        self.assertIn("paragraph_style_id_before=affe", debug)
        self.assertIn("paragraph_style_id_after=affe", debug)
        self.assertIn("body_style_normalized=False", debug)

    def test_heading_is_not_normalized_to_default_text(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c", style="affe")
        body = make_paragraph("\u9019\u662f 14 pt \u5167\u6587", style="affe", font_size_pt=14)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)

        fix_outline_paragraphs(root, include_tables=True, normalize_body_style_to_none=True)

        self.assertEqual(paragraph_style(heading), "affe")
        self.assertIsNone(paragraph_style(body))

    def test_toc_and_table_paragraphs_are_not_normalized_to_default_text(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        toc = make_paragraph("\u76ee\u9304\u9805\u76ee", style="TOC1", font_size_pt=14)
        tbl, table_body = make_table_paragraph("\u8868\u683c\u5167 affe \u5167\u6587")
        pPr = table_body.find("./w:pPr", NS)
        p_style = etree.SubElement(pPr, qn("pStyle"))
        p_style.set(qn("val"), "affe")
        body = make_paragraph("\u9019\u662f 14 pt \u5167\u6587", style="affe", font_size_pt=14)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, toc, tbl, body)

        fix_outline_paragraphs(root, include_tables=True, normalize_body_style_to_none=True)

        self.assertEqual(paragraph_style(toc), "TOC1")
        self.assertEqual(paragraph_style(table_body), "affe")
        self.assertIsNone(paragraph_style(body))

    def test_skipped_chapter_three_body_indent_is_not_normalized_to_default_text(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph("\u7ae0\u53c3\u5167 affe \u5167\u6587", style="affe", font_size_pt=14)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)

        fix_outline_paragraphs(
            root,
            include_tables=True,
            normalize_body_style_to_none=True,
            skip_paragraph_ids={id(body)},
        )

        self.assertEqual(paragraph_style(body), "affe")
        self.assertIsNone(paragraph_left_indent(body))

    def test_manual_numbering_prefix_suffix_spaces_are_removed(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        samples = [
            make_paragraph("1. \u6a19\u984c"),
            make_paragraph("A.\t\u6a19\u984c"),
            make_paragraph("a.\u3000\u6a19\u984c"),
            make_paragraph("\uff08\u4e00\uff09 \u6a19\u984c"),
        ]
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, *samples)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertEqual("".join(samples[0].xpath(".//w:t/text()", namespaces=NS)), "1.\u6a19\u984c")
        self.assertEqual("".join(samples[1].xpath(".//w:t/text()", namespaces=NS)), "A.\u6a19\u984c")
        self.assertEqual("".join(samples[2].xpath(".//w:t/text()", namespaces=NS)), "a.\u6a19\u984c")
        self.assertEqual("".join(samples[3].xpath(".//w:t/text()", namespaces=NS)), "\uff08\u4e00\uff09\u6a19\u984c")

    def test_body_indent_uses_paragraph_style_font_size(self):
        styles = make_styles_font_xml(
            paragraph_styles={"DefaultText": (14, None)}
        )
        style_lookup = build_style_font_size_lookup(styles)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph(
            "\u4f7f\u7528\u6bb5\u843d\u6a23\u5f0f 14 pt \u7684\u5167\u6587",
            style="DefaultText",
        )
        add_tab_stop(body, pos="1990")
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            style_font_size_lookup=style_lookup,
        )

        expected_left = TEMPLATE_OUTLINE_INDENTS[3]["body_left"]
        assert_body_indent_hard_override(self, body, expected_left)
        self.assertIn("font_size_source=paragraph_style:DefaultText", "\n".join(summary.body_indent_debug_logs))

    def test_body_indent_direct_format_overrides_style_start_indent(self):
        styles = make_styles_font_xml(
            paragraph_styles={"DefaultText": (14, None)},
            paragraph_style_indents={"DefaultText": {"start": "1480"}},
        )
        style_lookup = build_style_font_size_lookup(styles)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("\uff08\u4e00\uff09\u7b2c\u4e09\u968e\u6a19\u984c")
        body = make_paragraph(
            "\u53d7 style start \u5e72\u64fe\u7684\u5167\u6587",
            style="DefaultText",
        )
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            style_font_size_lookup=style_lookup,
        )

        expected = TEMPLATE_OUTLINE_INDENTS[2]["body_left"]
        assert_body_indent_hard_override(self, body, expected)
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("paragraph_style_id=DefaultText", debug)
        self.assertIn(f"written_start_twips={expected}", debug)
        self.assertIn("validation=ok", debug)

    def test_body_indent_uses_based_on_paragraph_style_font_size(self):
        styles = make_styles_font_xml(
            paragraph_styles={
                "Normal": (14, None),
                "DefaultText": (None, "Normal"),
            }
        )
        style_lookup = build_style_font_size_lookup(styles)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph(
            "\u4f7f\u7528 basedOn 14 pt \u7684\u5167\u6587",
            style="DefaultText",
        )
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            style_font_size_lookup=style_lookup,
        )

        self.assertEqual(paragraph_left_indent(body), TEMPLATE_OUTLINE_INDENTS[3]["body_left"])
        self.assertIn("font_size_source=paragraph_style:DefaultText", "\n".join(summary.body_indent_debug_logs))

    def test_body_indent_uses_doc_defaults_font_size(self):
        styles = make_styles_font_xml(doc_default_pt=14)
        style_lookup = build_style_font_size_lookup(styles)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph("\u4f7f\u7528 docDefaults 14 pt \u7684\u5167\u6587")
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            style_font_size_lookup=style_lookup,
        )

        self.assertEqual(paragraph_left_indent(body), TEMPLATE_OUTLINE_INDENTS[3]["body_left"])
        self.assertIn("font_size_source=docDefaults", "\n".join(summary.body_indent_debug_logs))

    def test_body_indent_uses_character_style_font_size(self):
        styles = make_styles_font_xml(
            character_styles={"EmphasisBody": (14, None)}
        )
        style_lookup = build_style_font_size_lookup(styles)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph(
            "\u4f7f\u7528\u5b57\u5143\u6a23\u5f0f 14 pt \u7684\u5167\u6587",
            run_style="EmphasisBody",
        )
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            style_font_size_lookup=style_lookup,
        )

        self.assertEqual(paragraph_left_indent(body), TEMPLATE_OUTLINE_INDENTS[3]["body_left"])
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("font_size_source=character_style:EmphasisBody", debug)
        self.assertIn("run_style_id=EmphasisBody", debug)

    def test_direct_run_font_size_overrides_paragraph_style_font_size(self):
        styles = make_styles_font_xml(
            paragraph_styles={"DefaultText": (14, None)}
        )
        style_lookup = build_style_font_size_lookup(styles)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph(
            "\u76f4\u63a5\u5b57\u865f 10 pt \u61c9\u8a72\u8df3\u904e",
            style="DefaultText",
            font_size_pt=10,
        )
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            change_logs=summary.paragraph_logs,
            style_font_size_lookup=style_lookup,
        )

        self.assertIsNone(paragraph_left_indent(body))
        self.assertFalse(summary.body_indent_debug_logs)
        self.assertTrue(any("source=direct_run" in line for line in summary.paragraph_logs))

    def test_body_indent_uses_dominant_font_size_when_first_run_differs(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph(
            "",
            runs=[
                {"text": "\u958b\u982d", "font_size_pt": 12},
                {"text": "\u5f8c\u9762\u5927\u90e8\u5206\u5167\u6587\u90fd\u662f\u5341\u56db\u9ede", "font_size_pt": 14},
            ],
        )
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            style_font_size_lookup=None,
        )

        self.assertEqual(paragraph_left_indent(body), TEMPLATE_OUTLINE_INDENTS[3]["body_left"])
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("first_font_size=12 pt", debug)
        self.assertIn("dominant_font_size=14 pt", debug)
        self.assertIn("dominant_font_size_source=dominant_runs", debug)

    def test_body_indent_skips_when_dominant_font_size_is_not_14_pt(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph(
            "",
            runs=[
                {"text": "\u958b\u982d", "font_size_pt": 14},
                {"text": "\u5f8c\u9762\u6bd4\u8f03\u9577\u7684\u6bb5\u843d\u5167\u6587\u5176\u5be6\u662f\u5341\u4e8c\u9ede", "font_size_pt": 12},
            ],
        )
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            change_logs=summary.paragraph_logs,
        )

        self.assertIsNone(paragraph_left_indent(body))
        skip_log = "\n".join(summary.paragraph_logs)
        self.assertIn("first_font_size=14 pt", skip_log)
        self.assertIn("dominant_font_size=12 pt", skip_log)

    def test_body_indent_queues_word_com_font_check_when_xml_font_is_not_14_pt(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("\u4e00\u3001\u7b2c\u4e8c\u968e\u6a19\u984c")
        body = make_paragraph("\u9019\u662f XML 12 pt \u4f46 Word \u986f\u793a\u70ba 14 pt \u7684\u5167\u6587", font_size_pt=12)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            change_logs=summary.paragraph_logs,
            enable_level1_level2_body_first_line_indent=True,
            word_com_check_body_font_when_xml_not_14=True,
        )

        self.assertIsNone(paragraph_left_indent(body))
        record = summary.body_indent_records[-1]
        self.assertEqual(record["kind"], "body_font_check")
        self.assertEqual(record["expected_left_twips"], TEMPLATE_OUTLINE_INDENTS[1]["body_left"])
        self.assertEqual(record["expected_first_line_twips"], "560")
        self.assertEqual(record["xml_font_size"], 12.0)
        self.assertEqual(record["apply_only_if_word_font_size_is_14"], True)
        log = "\n".join(summary.paragraph_logs)
        self.assertIn("XML font is not 14pt; queued for Word COM font check", log)

    def test_body_indent_pure_14_pt_paragraph_behavior_is_unchanged(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph("\u7d14 14 pt \u5167\u6587", font_size_pt=14)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
        )

        self.assertEqual(paragraph_left_indent(body), TEMPLATE_OUTLINE_INDENTS[3]["body_left"])
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("first_font_size=14 pt", debug)
        self.assertIn("dominant_font_size=14 pt", debug)

    def test_unknown_body_font_size_is_skipped_and_logged(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("1. \u7b2c\u56db\u968e\u6a19\u984c")
        body = make_paragraph("\u7121\u6cd5\u5224\u65b7\u5b57\u865f\u7684\u5167\u6587")
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            change_logs=summary.paragraph_logs,
        )

        self.assertIsNone(paragraph_left_indent(body))
        self.assertTrue(any("source=unknown" in line for line in summary.paragraph_logs))

    def test_body_indent_skips_non_14_pt_paragraph_and_logs_reason(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("\u4e00\u3001\u7814\u7a76\u76ee\u7684")
        body = make_paragraph("\u9019\u662f 12 pt \u5167\u6587", font_size_pt=12)
        original_ind = add_ind_with_char_attrs(body)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body)
        summary = ProcessSummary()

        fix_outline_paragraphs(root, include_tables=True, summary=summary, change_logs=summary.paragraph_logs)

        ind = body.find("./w:pPr/w:ind", NS)
        self.assertIs(ind, original_ind)
        self.assertEqual(ind.get(qn("left")), "123")
        self.assertEqual(ind.get(qn("hanging")), "45")
        self.assertEqual(ind.get(qn("leftChars")), "99")
        self.assertTrue(any("Body indent skipped:" in line for line in summary.paragraph_logs))

    def test_body_indent_applies_only_to_14_pt_and_clears_char_indent_attrs(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        heading = make_paragraph("\u4e00\u3001\u7814\u7a76\u76ee\u7684")
        body_14 = make_paragraph("\u9019\u662f 14 pt \u5167\u6587", font_size_pt=14)
        body_12 = make_paragraph("\u9019\u662f 12 pt \u5167\u6587", font_size_pt=12)
        add_ind_with_char_attrs(body_14)
        add_ind_with_char_attrs(body_12)
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, heading, body_14, body_12)

        fix_outline_paragraphs(root, include_tables=True)

        body_14_ind = body_14.find("./w:pPr/w:ind", NS)
        body_12_ind = body_12.find("./w:pPr/w:ind", NS)
        assert_body_indent_hard_override(self, body_14, TEMPLATE_OUTLINE_INDENTS[1]["body_left"])
        self.assertEqual(body_12_ind.get(qn("left")), "123")
        self.assertEqual(body_12_ind.get(qn("leftChars")), "99")

    def test_preface_indent_level_maps_detected_levels_to_preface_table(self):
        self.assertIsNone(preface_indent_level_from_detected_level(0))
        for detected_level in range(1, 9):
            with self.subTest(detected_level=detected_level):
                self.assertEqual(
                    preface_indent_level_from_detected_level(detected_level),
                    detected_level - 1,
                )

    def test_preface_heading_indents_use_preface_table_without_level_shift(self):
        headings = [
            make_paragraph("\u4e00\u3001\u524d\u7f6e\u9805"),
            make_paragraph("\uff08\u4e00\uff09\u524d\u7f6e\u5167\u5c64"),
            make_paragraph("1. \u524d\u7f6e\u7b2c\u4e09\u968e"),
            make_paragraph("\uff081\uff09\u524d\u7f6e\u7b2c\u56db\u968e"),
            make_paragraph("A. \u524d\u7f6e\u7b2c\u4e94\u968e"),
            make_paragraph("\uff08A\uff09\u524d\u7f6e\u7b2c\u516d\u968e"),
            make_paragraph("a. \u524d\u7f6e\u7b2c\u4e03\u968e"),
            make_paragraph("\uff08a\uff09\u524d\u7f6e\u7b2c\u516b\u968e"),
        ]
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(*headings, marker)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            indent_preface_paragraphs=True,
        )

        for preface_level, paragraph in enumerate(headings):
            with self.subTest(preface_level=preface_level):
                self.assertEqual(paragraph_indent(paragraph), expected_preface_indent(preface_level))
                self.assertIsNone(paragraph_outline(paragraph))

        self.assertIsNone(paragraph_indent(marker))
        self.assertIsNone(paragraph_outline(marker))
        self.assertEqual(summary.indented_preface_paragraphs, 8)

    def test_preface_body_indent_uses_preface_heading_body_left_and_common_cleanup(self):
        heading = make_paragraph("\u4e00\u3001\u524d\u7f6e\u9805")
        body = make_paragraph("\u9019\u662f\u5e8f\u8a00\u524d 14pt \u5167\u6587", font_size_pt=14)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        add_ind_with_char_attrs(body)
        add_tab_stop(body)
        root = make_root(heading, body, marker)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            indent_preface_paragraphs=True,
        )

        assert_body_indent_hard_override(self, body, PREFACE_OUTLINE_INDENTS[0]["body_left"])
        debug = "\n".join(summary.body_indent_debug_logs)
        self.assertIn("heading_uses_outline=False", debug)
        self.assertIn(
            f"spec_body_left_cm={twips_to_cm(PREFACE_OUTLINE_INDENTS[0]['body_left']):.2f}",
            debug,
        )

    def test_preface_nested_body_indent_uses_matching_preface_body_left(self):
        heading = make_paragraph("\uff08\u4e00\uff09\u524d\u7f6e\u5167\u5c64")
        body = make_paragraph("\u9019\u662f\u5e8f\u8a00\u524d\u5167\u5c64 14pt \u5167\u6587", font_size_pt=14)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(heading, body, marker)

        fix_outline_paragraphs(
            root,
            include_tables=True,
            indent_preface_paragraphs=True,
        )

        assert_body_indent_hard_override(self, body, PREFACE_OUTLINE_INDENTS[1]["body_left"])
        self.assertNotEqual(paragraph_left_indent(body), PREFACE_OUTLINE_INDENTS[0]["body_left"])
        self.assertNotEqual(paragraph_left_indent(body), TEMPLATE_OUTLINE_INDENTS[2]["body_left"])

    def test_preface_outline_uses_preface_indent_table_and_preface_measurements(self):
        first = make_paragraph("\u4e00\u3001\u524d\u7f6e\u9805")
        second = make_paragraph("\uff08\u4e00\uff09\u524d\u7f6e\u5167\u5c64")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(first, second, marker)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            indent_preface_paragraphs=True,
            outline_preface_paragraphs=True,
        )

        self.assertEqual(paragraph_outline(first), "0")
        self.assertEqual(paragraph_outline(second), "1")
        self.assertEqual(paragraph_indent(first), expected_preface_indent(0))
        self.assertEqual(paragraph_indent(second), expected_preface_indent(1))
        self.assertIsNone(paragraph_indent(marker))
        measurements = list(summary.numbering_measurements.values())
        self.assertTrue(any(
            measurement["section"] == "preface"
            and measurement["level"] == 1
            and measurement["indent_level"] == 0
            and measurement["text_start_cm"] == twips_to_cm(PREFACE_OUTLINE_INDENTS[0]["left"])
            for measurement in measurements
        ))
        self.assertTrue(any(
            measurement["section"] == "preface"
            and measurement["level"] == 2
            and measurement["indent_level"] == 1
            and measurement["text_start_cm"] == twips_to_cm(PREFACE_OUTLINE_INDENTS[1]["left"])
            for measurement in measurements
        ))

    def test_preface_body_paragraph_aligns_to_preface_heading_when_indent_enabled(self):
        heading = make_paragraph("\u4e00\u3001\u524d\u7f6e\u9805")
        body = make_paragraph("\u9019\u662f\u5e8f\u8a00\u524d\u7684\u5167\u6587", outline=2, font_size_pt=14)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(heading, body, marker)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            indent_preface_paragraphs=True,
        )

        assert_body_indent_hard_override(self, body, PREFACE_OUTLINE_INDENTS[0]["body_left"])
        self.assertEqual(paragraph_outline(body), "2")
        self.assertEqual(summary.indented_preface_paragraphs, 1)

    def test_preface_body_aligns_when_preface_indent_is_enabled(self):
        heading = make_paragraph("\u4e00\u3001\u524d\u7f6e\u9805")
        body = make_paragraph("\u9019\u662f\u5e8f\u8a00\u524d\u7684\u5167\u6587", font_size_pt=14)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(heading, body, marker)

        fix_outline_paragraphs(
            root,
            include_tables=True,
            indent_preface_paragraphs=True,
        )

        assert_body_indent_hard_override(self, body, PREFACE_OUTLINE_INDENTS[0]["body_left"])
        self.assertIsNone(paragraph_outline(body))

    def test_preface_body_does_not_align_when_only_preface_outline_is_enabled(self):
        heading = make_paragraph("\u4e00\u3001\u524d\u7f6e\u9805")
        body = make_paragraph("\u9019\u662f\u5e8f\u8a00\u524d\u7684\u5167\u6587")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(heading, body, marker)

        fix_outline_paragraphs(
            root,
            include_tables=True,
            outline_preface_paragraphs=True,
        )

        self.assertIsNone(paragraph_indent(body))
        self.assertIsNone(paragraph_outline(body))

    def test_preface_options_can_run_without_main_paragraph_fixing(self):
        before = make_paragraph("\u4e00\u3001\u524d\u7f6e\u9805")
        first_marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00", outline=4)
        after = make_paragraph("\u4e00\u3001\u5e8f\u8a00\u5167\u5c64")
        root = make_root(before, first_marker, marker, after)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            fix_numbered_paragraphs=False,
            indent_preface_paragraphs=True,
            outline_preface_paragraphs=True,
        )

        self.assertEqual(paragraph_indent(before), expected_preface_indent(0))
        self.assertEqual(paragraph_outline(before), "0")
        self.assertIsNone(paragraph_indent(first_marker))
        self.assertIsNone(paragraph_outline(first_marker))
        self.assertEqual(paragraph_outline(marker), "4")
        self.assertIsNone(paragraph_indent(after))
        self.assertEqual(summary.indented_preface_paragraphs, 1)
        self.assertEqual(summary.outlined_preface_paragraphs, 1)

    def test_preface_options_skip_toc_before_marker(self):
        toc = make_paragraph("\u58f9\u3001\u5e8f\u8a00", style="TOC1", outline=0)
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        root = make_root(toc, marker)
        summary = ProcessSummary()

        fix_outline_paragraphs(
            root,
            include_tables=True,
            summary=summary,
            fix_numbered_paragraphs=False,
            indent_preface_paragraphs=True,
            outline_preface_paragraphs=True,
        )

        self.assertEqual(paragraph_outline(toc), "0")
        self.assertIsNone(paragraph_outline(marker))
        self.assertEqual(summary.skipped_toc_paragraphs, 1)
        self.assertEqual(summary.indented_preface_paragraphs, 0)
        self.assertEqual(summary.outlined_preface_paragraphs, 0)

    def test_table_paragraph_is_skipped_even_when_include_tables_is_true(self):
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        table, table_p = make_table_paragraph("\uff08\u4e00\uff09\u8868\u683c\u5167\u9805\u76ee")
        root = make_root(make_paragraph("\u58f9\u3001\u5e8f\u8a00"), marker, table)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertIsNone(paragraph_indent(table_p))
        self.assertIsNone(paragraph_outline(table_p))

    def test_existing_outline_on_non_numbered_body_is_left_unchanged(self):
        body = make_paragraph("\u666e\u901a\u6b63\u6587", outline=2)
        root = make_root(body)

        fix_outline_paragraphs(root, include_tables=True)

        self.assertEqual(paragraph_outline(body), "2")

    def test_remove_all_outline_levels_from_root_forces_every_paragraph_to_body_outline(self):
        first = make_paragraph("\u666e\u901a\u6b63\u6587", outline=2)
        second = make_paragraph("\u58f9\u3001\u5e8f\u8a00", outline=0)
        third = make_paragraph("\u7121\u5927\u7db1")
        root = make_root(first, second, third)
        summary = ProcessSummary()

        removed = remove_all_outline_levels_from_root(root, summary=summary)

        self.assertEqual(removed, 3)
        self.assertEqual(summary.removed_all_outline_paragraphs, 3)
        self.assertEqual(paragraph_outline(first), "9")
        self.assertEqual(paragraph_outline(second), "9")
        self.assertEqual(paragraph_outline(third), "9")

    def test_force_all_paragraphs_to_body_outline_level_overrides_heading_styles(self):
        heading_one = make_paragraph("Heading 1 text", style="Heading1")
        heading_two = make_paragraph("Heading 2 text", style="Heading2", outline=1)
        chinese_heading = make_paragraph("\u6a19\u984c\u6a23\u5f0f", style="1")
        root = make_root(heading_one, heading_two, chinese_heading)

        changed = force_all_paragraphs_to_body_outline_level(root)

        self.assertEqual(changed, 3)
        self.assertEqual(paragraph_outline(heading_one), "9")
        self.assertEqual(paragraph_outline(heading_two), "9")
        self.assertEqual(paragraph_outline(chinese_heading), "9")

    def test_remove_all_outline_levels_from_any_root_cleans_style_ppr_outline(self):
        root = etree.Element(qn("styles"), nsmap={"w": W_NS})
        style = etree.SubElement(root, qn("style"))
        style.set(qn("type"), "paragraph")
        pPr = etree.SubElement(style, qn("pPr"))
        outline = etree.SubElement(pPr, qn("outlineLvl"))
        outline.set(qn("val"), "2")
        summary = ProcessSummary()

        removed = remove_all_outline_levels_from_any_root(root, summary=summary)

        self.assertEqual(removed, 1)
        self.assertEqual(summary.removed_all_outline_paragraphs, 1)
        self.assertFalse(root.xpath(".//w:outlineLvl", namespaces=NS))

    def test_paragraph_summary_counts_levels_and_skips(self):
        toc = make_paragraph("\u76ee\u9304")
        toc_entry = make_paragraph("\u4e00\u3001\u5e8f\u8a00")
        first_marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        marker = make_paragraph("\u58f9\u3001\u5e8f\u8a00")
        normal = make_paragraph("\u666e\u901a\u6b63\u6587")
        table, _ = make_table_paragraph("1. \u8868\u683c\u5167\u9805\u76ee")
        root = make_root(toc, toc_entry, first_marker, marker, normal, table)
        summary = ProcessSummary()

        fix_outline_paragraphs(root, include_tables=True, summary=summary)

        self.assertEqual(summary.total_paragraphs, 6)
        self.assertEqual(summary.skipped_toc_paragraphs, 3)
        self.assertEqual(summary.skipped_table_paragraphs, 1)
        self.assertEqual(summary.paragraph_level_counts[0], 1)
        self.assertEqual(summary.unknown_paragraphs, 0)

    def test_numbering_xml_uses_template_indents_and_keeps_bullets_out_of_outline(self):
        updated = apply_numbering_outline_format(make_numbering_xml())
        root = etree.fromstring(updated)

        decimal_lvl = root.xpath("./w:abstractNum/w:lvl[@w:ilvl='3']", namespaces=NS)[0]
        decimal_ind = decimal_lvl.find("./w:pPr/w:ind", NS)
        decimal_lvl_jc = decimal_lvl.find("./w:lvlJc", NS)
        self.assertEqual(
            (decimal_ind.get(qn("left")), decimal_ind.get(qn("hanging"))),
            expected_indent(3),
        )
        # decimal "%1." is outline level 3, a tab-suffix level.
        assert_level_suffix_rule(self, decimal_lvl, 3)
        self.assertEqual(decimal_lvl_jc.get(qn("val")), "left")

        bullet_lvl = root.xpath("./w:abstractNum/w:lvl[@w:ilvl='4']", namespaces=NS)[0]
        bullet_outline = bullet_lvl.find("./w:pPr/w:outlineLvl", NS)
        self.assertEqual(bullet_outline.get(qn("val")), "9")

    def test_numbering_xml_level_two_font_and_indents_use_twips_without_char_attrs(self):
        root = etree.Element(qn("numbering"), nsmap={"w": W_NS})
        abstract = etree.SubElement(root, qn("abstractNum"))
        abstract.set(qn("abstractNumId"), "1")

        level_two = etree.SubElement(abstract, qn("lvl"))
        level_two.set(qn("ilvl"), "1")
        level_two_fmt = etree.SubElement(level_two, qn("numFmt"))
        level_two_fmt.set(qn("val"), "custom")
        level_two_text = etree.SubElement(level_two, qn("lvlText"))
        level_two_text.set(qn("val"), "%1")
        level_two_pPr = etree.SubElement(level_two, qn("pPr"))
        level_two_ind = etree.SubElement(level_two_pPr, qn("ind"))
        level_two_ind.set(qn("left"), "123")
        for attr in ("leftChars", "startChars", "hangingChars", "firstLineChars"):
            level_two_ind.set(qn(attr), "99")

        level_three = etree.SubElement(abstract, qn("lvl"))
        level_three.set(qn("ilvl"), "2")
        level_three_fmt = etree.SubElement(level_three, qn("numFmt"))
        level_three_fmt.set(qn("val"), "custom")
        level_three_text = etree.SubElement(level_three, qn("lvlText"))
        level_three_text.set(qn("val"), "%1")

        updated = apply_numbering_outline_format(etree.tostring(root))
        updated_root = etree.fromstring(updated)
        updated_level_two = updated_root.xpath("./w:abstractNum/w:lvl[@w:ilvl='1']", namespaces=NS)[0]
        updated_level_three = updated_root.xpath("./w:abstractNum/w:lvl[@w:ilvl='2']", namespaces=NS)[0]

        rPr = updated_level_two.find("./w:rPr", NS)
        self.assertEqual(rPr.find("./w:sz", NS).get(qn("val")), "32")
        self.assertEqual(rPr.find("./w:szCs", NS).get(qn("val")), "32")

        ind = updated_level_two.find("./w:pPr/w:ind", NS)
        suff = updated_level_two.find("./w:suff", NS)
        lvl_jc = updated_level_two.find("./w:lvlJc", NS)
        self.assertEqual(
            (ind.get(qn("left")), ind.get(qn("hanging"))),
            expected_indent(1),
        )
        self.assertIsNone(ind.get(qn("start")))
        self.assertIsNone(updated_level_two.find("./w:pPr/w:tabs", NS))
        self.assertEqual(suff.get(qn("val")), "nothing")
        self.assertEqual(lvl_jc.get(qn("val")), "left")
        assert_no_char_indent_attrs(self, ind)
        # ilvl=2 resolves to outline level 2, which is now forced to 14 pt on the
        # numbering level run properties (both w:sz and w:szCs = 28 half-points).
        level_three_rPr = updated_level_three.find("./w:rPr", NS)
        self.assertEqual(level_three_rPr.find("./w:sz", NS).get(qn("val")), "28")
        self.assertEqual(level_three_rPr.find("./w:szCs", NS).get(qn("val")), "28")

    def test_numbering_xml_normalizes_lvljc_suffix_and_removes_tabs(self):
        root = etree.Element(qn("numbering"), nsmap={"w": W_NS})

        for abstract_id, num_id, jc in (("1", "10", "right"), ("2", "20", "left")):
            abstract = etree.SubElement(root, qn("abstractNum"))
            abstract.set(qn("abstractNumId"), abstract_id)
            lvl = etree.SubElement(abstract, qn("lvl"))
            lvl.set(qn("ilvl"), "1")
            lvl_jc = etree.SubElement(lvl, qn("lvlJc"))
            lvl_jc.set(qn("val"), jc)
            suff = etree.SubElement(lvl, qn("suff"))
            suff.set(qn("val"), "nothing")
            num_fmt = etree.SubElement(lvl, qn("numFmt"))
            num_fmt.set(qn("val"), "taiwaneseCountingThousand")
            lvl_text = etree.SubElement(lvl, qn("lvlText"))
            lvl_text.set(qn("val"), "%1\u3001")
            pPr = etree.SubElement(lvl, qn("pPr"))
            tabs = etree.SubElement(pPr, qn("tabs"))
            tab = etree.SubElement(tabs, qn("tab"))
            tab.set(qn("val"), "left")
            tab.set(qn("pos"), "999")
            ind = etree.SubElement(pPr, qn("ind"))
            ind.set(qn("left"), "999")
            ind.set(qn("hanging"), "111")

            num = etree.SubElement(root, qn("num"))
            num.set(qn("numId"), num_id)
            abstract_ref = etree.SubElement(num, qn("abstractNumId"))
            abstract_ref.set(qn("val"), abstract_id)

        updated = apply_numbering_outline_format(etree.tostring(root))
        updated_root = etree.fromstring(updated)
        levels = updated_root.xpath("./w:abstractNum/w:lvl[@w:ilvl='1']", namespaces=NS)
        expected_left, expected_hanging = expected_indent(1)
        expected_number_start = str(int(expected_left) - int(expected_hanging))

        for lvl in levels:
            with self.subTest(abstract=lvl.getparent().get(qn("abstractNumId"))):
                ind = lvl.find("./w:pPr/w:ind", NS)
                self.assertEqual(lvl.find("./w:lvlJc", NS).get(qn("val")), "left")
                self.assertEqual(lvl.find("./w:suff", NS).get(qn("val")), "nothing")
                self.assertEqual(ind.get(qn("left")), expected_left)
                self.assertEqual(ind.get(qn("hanging")), expected_hanging)
                self.assertEqual(str(int(ind.get(qn("left"))) - int(ind.get(qn("hanging")))), expected_number_start)
                self.assertIsNone(lvl.find("./w:pPr/w:tabs", NS))

        lookup = build_numbering_format_lookup(updated)
        self.assertEqual(lookup[("10", 1)]["lvlJc"], "left")
        self.assertEqual(lookup[("20", 1)]["lvlJc"], "left")
        self.assertEqual(lookup[("10", 1)]["suff"], "nothing")
        self.assertEqual(lookup[("20", 1)]["suff"], "nothing")
        self.assertEqual(lookup[("10", 1)]["left"], lookup[("20", 1)]["left"])
        self.assertEqual(lookup[("10", 1)]["hanging"], lookup[("20", 1)]["hanging"])
        self.assertEqual(lookup[("10", 1)]["number_start"], lookup[("20", 1)]["number_start"])
        self.assertIsNone(lookup[("10", 1)]["tab_pos"])
        self.assertIsNone(lookup[("20", 1)]["tab_pos"])

    def test_numbering_xml_normalizes_lvl_override_missing_suffix_and_tabs(self):
        root = etree.Element(qn("numbering"), nsmap={"w": W_NS})
        abstract = etree.SubElement(root, qn("abstractNum"))
        abstract.set(qn("abstractNumId"), "1")
        abstract_lvl = etree.SubElement(abstract, qn("lvl"))
        abstract_lvl.set(qn("ilvl"), "0")
        abstract_fmt = etree.SubElement(abstract_lvl, qn("numFmt"))
        abstract_fmt.set(qn("val"), "decimal")
        abstract_text = etree.SubElement(abstract_lvl, qn("lvlText"))
        abstract_text.set(qn("val"), "%1.")

        num = etree.SubElement(root, qn("num"))
        num.set(qn("numId"), "42")
        abstract_ref = etree.SubElement(num, qn("abstractNumId"))
        abstract_ref.set(qn("val"), "1")
        override = etree.SubElement(num, qn("lvlOverride"))
        override.set(qn("ilvl"), "0")
        override_lvl = etree.SubElement(override, qn("lvl"))
        override_lvl.set(qn("ilvl"), "0")
        override_fmt = etree.SubElement(override_lvl, qn("numFmt"))
        override_fmt.set(qn("val"), "decimal")
        override_text = etree.SubElement(override_lvl, qn("lvlText"))
        override_text.set(qn("val"), "%1.")
        override_pPr = etree.SubElement(override_lvl, qn("pPr"))
        override_tabs = etree.SubElement(override_pPr, qn("tabs"))
        override_tab = etree.SubElement(override_tabs, qn("tab"))
        override_tab.set(qn("val"), "left")
        override_tab.set(qn("pos"), "999")

        updated = apply_numbering_outline_format(etree.tostring(root))
        updated_root = etree.fromstring(updated)
        updated_override_lvl = updated_root.xpath("./w:num/w:lvlOverride/w:lvl", namespaces=NS)[0]

        # The override's own decimal "%1." is outline level 3, a tab-suffix level.
        assert_level_suffix_rule(self, updated_override_lvl, 3)

    def test_force_clean_numbering_suffix_tabs_cleans_all_levels(self):
        root = etree.Element(qn("numbering"), nsmap={"w": W_NS})
        abstract = etree.SubElement(root, qn("abstractNum"))
        abstract.set(qn("abstractNumId"), "1")
        for ilvl, suffix, lvl_text_value in (
            ("0", None, "%1. "),
            ("1", "tab", "（%1）\t"),
            ("2", "space", "%3.\u3000"),
        ):
            lvl = etree.SubElement(abstract, qn("lvl"))
            lvl.set(qn("ilvl"), ilvl)
            lvl_text = etree.SubElement(lvl, qn("lvlText"))
            lvl_text.set(qn("val"), lvl_text_value)
            if suffix is not None:
                suff = etree.SubElement(lvl, qn("suff"))
                suff.set(qn("val"), suffix)
            pPr = etree.SubElement(lvl, qn("pPr"))
            tabs = etree.SubElement(pPr, qn("tabs"))
            tab = etree.SubElement(tabs, qn("tab"))
            tab.set(qn("val"), "num")
            tab.set(qn("pos"), "2061")
            ind = etree.SubElement(pPr, qn("ind"))
            ind.set(qn("left"), "2279")
            ind.set(qn("hanging"), "420")

        num = etree.SubElement(root, qn("num"))
        num.set(qn("numId"), "42")
        abstract_ref = etree.SubElement(num, qn("abstractNumId"))
        abstract_ref.set(qn("val"), "1")
        override = etree.SubElement(num, qn("lvlOverride"))
        override.set(qn("ilvl"), "0")
        override_lvl = etree.SubElement(override, qn("lvl"))
        override_lvl.set(qn("ilvl"), "0")
        override_text = etree.SubElement(override_lvl, qn("lvlText"))
        override_text.set(qn("val"), "%5. ")
        override_pPr = etree.SubElement(override_lvl, qn("pPr"))
        override_tabs = etree.SubElement(override_pPr, qn("tabs"))
        override_tab = etree.SubElement(override_tabs, qn("tab"))
        override_tab.set(qn("val"), "num")
        override_tab.set(qn("pos"), "2061")

        logs = []
        updated = force_clean_numbering_suffix_tabs(etree.tostring(root), logs=logs)
        updated_root = etree.fromstring(updated)

        for lvl in updated_root.xpath("./w:abstractNum/w:lvl | ./w:num/w:lvlOverride/w:lvl", namespaces=NS):
            self.assertEqual(lvl.find("w:suff", NS).get(qn("val")), "nothing")
            self.assertIsNone(lvl.find("./w:pPr/w:tabs", NS))
            self.assertFalse(lvl.find("w:lvlText", NS).get(qn("val")).endswith((" ", "\t", "\u3000")))
        self.assertEqual(
            [lvl.find("w:lvlText", NS).get(qn("val")) for lvl in updated_root.xpath("./w:abstractNum/w:lvl", namespaces=NS)],
            ["%1.", "（%1）", "%3."],
        )
        # Every level here resolves to outline 0/1/2 (no numFmt -> ilvl fallback),
        # so all are nothing-suffix and all four list tabs are removed.
        self.assertTrue(any("suffixes_set_to_nothing=4" in log for log in logs))
        self.assertTrue(any("suffixes_set_to_tab=0" in log for log in logs))
        self.assertTrue(any("tab_stops_rebuilt=0" in log for log in logs))
        self.assertTrue(any("tab_stops_removed=4" in log for log in logs))
        self.assertTrue(any("lvl_text_trimmed=4" in log for log in logs))

    def test_force_clean_numbering_suffix_tabs_skips_protected_definitions(self):
        root = etree.Element(qn("numbering"), nsmap={"w": W_NS})
        for abstract_id, num_id in (("1", "42"), ("2", "99")):
            abstract = etree.SubElement(root, qn("abstractNum"))
            abstract.set(qn("abstractNumId"), abstract_id)
            lvl = etree.SubElement(abstract, qn("lvl"))
            lvl.set(qn("ilvl"), "0")
            lvl_text = etree.SubElement(lvl, qn("lvlText"))
            lvl_text.set(qn("val"), f"%{abstract_id}. ")
            suff = etree.SubElement(lvl, qn("suff"))
            suff.set(qn("val"), "tab")
            pPr = etree.SubElement(lvl, qn("pPr"))
            tabs = etree.SubElement(pPr, qn("tabs"))
            tab = etree.SubElement(tabs, qn("tab"))
            tab.set(qn("val"), "num")
            tab.set(qn("pos"), "2061")

            num = etree.SubElement(root, qn("num"))
            num.set(qn("numId"), num_id)
            abstract_ref = etree.SubElement(num, qn("abstractNumId"))
            abstract_ref.set(qn("val"), abstract_id)

        logs = []
        updated = force_clean_numbering_suffix_tabs(
            etree.tostring(root),
            logs=logs,
            excluded_num_ids={"42"},
            excluded_abstract_ids={"1"},
        )
        updated_root = etree.fromstring(updated)
        protected_lvl = updated_root.xpath("./w:abstractNum[@w:abstractNumId='1']/w:lvl", namespaces=NS)[0]
        cleaned_lvl = updated_root.xpath("./w:abstractNum[@w:abstractNumId='2']/w:lvl", namespaces=NS)[0]

        self.assertEqual(protected_lvl.find("w:suff", NS).get(qn("val")), "tab")
        self.assertIsNotNone(protected_lvl.find("./w:pPr/w:tabs", NS))
        self.assertEqual(protected_lvl.find("w:lvlText", NS).get(qn("val")), "%1. ")
        self.assertEqual(cleaned_lvl.find("w:suff", NS).get(qn("val")), "nothing")
        self.assertIsNone(cleaned_lvl.find("./w:pPr/w:tabs", NS))
        self.assertEqual(cleaned_lvl.find("w:lvlText", NS).get(qn("val")), "%2.")
        self.assertTrue(any("levels_total=2" in log for log in logs))
        self.assertTrue(any("levels_cleaned=1" in log for log in logs))
        self.assertTrue(any("levels_skipped_protected=1" in log for log in logs))

    def test_force_clean_numbering_suffix_tabs_warns_and_protects_shared_definition(self):
        root = etree.Element(qn("numbering"), nsmap={"w": W_NS})
        abstract = etree.SubElement(root, qn("abstractNum"))
        abstract.set(qn("abstractNumId"), "1")
        lvl = etree.SubElement(abstract, qn("lvl"))
        lvl.set(qn("ilvl"), "0")
        lvl_text = etree.SubElement(lvl, qn("lvlText"))
        lvl_text.set(qn("val"), "%1. ")
        suff = etree.SubElement(lvl, qn("suff"))
        suff.set(qn("val"), "tab")
        pPr = etree.SubElement(lvl, qn("pPr"))
        tabs = etree.SubElement(pPr, qn("tabs"))
        tab = etree.SubElement(tabs, qn("tab"))
        tab.set(qn("val"), "num")
        tab.set(qn("pos"), "2061")

        for num_id in ("42", "99"):
            num = etree.SubElement(root, qn("num"))
            num.set(qn("numId"), num_id)
            abstract_ref = etree.SubElement(num, qn("abstractNumId"))
            abstract_ref.set(qn("val"), "1")

        logs = []
        updated = force_clean_numbering_suffix_tabs(
            etree.tostring(root),
            logs=logs,
            excluded_num_ids={"42"},
        )
        updated_root = etree.fromstring(updated)
        protected_lvl = updated_root.xpath("./w:abstractNum[@w:abstractNumId='1']/w:lvl", namespaces=NS)[0]

        self.assertEqual(protected_lvl.find("w:suff", NS).get(qn("val")), "tab")
        self.assertIsNotNone(protected_lvl.find("./w:pPr/w:tabs", NS))
        self.assertTrue(any(
            "FINAL_NUMBERING_SUFFIX_CLEAN_SKIP_PROTECTED_SHARED_DEFINITION" in log
            and "protected_numIds=42" in log
            and "shared_numIds=99" in log
            for log in logs
        ))

    def test_numbering_xml_does_not_sanitize_excluded_level(self):
        # An excluded numbering level (TOC or chapter 參 protection) must be
        # decided as skipped BEFORE sanitizing, so its suffix / tab stops /
        # lvlText trailing whitespace are all left untouched.
        root = etree.Element(qn("numbering"), nsmap={"w": W_NS})
        abstract = etree.SubElement(root, qn("abstractNum"))
        abstract.set(qn("abstractNumId"), "99")
        lvl = etree.SubElement(abstract, qn("lvl"))
        lvl.set(qn("ilvl"), "0")
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
        tab.set(qn("pos"), "999")
        ind = etree.SubElement(pPr, qn("ind"))
        ind.set(qn("left"), "2279")
        ind.set(qn("hanging"), "420")

        num = etree.SubElement(root, qn("num"))
        num.set(qn("numId"), "99")
        abstract_ref = etree.SubElement(num, qn("abstractNumId"))
        abstract_ref.set(qn("val"), "99")

        original_xml = etree.tostring(root)
        updated = apply_numbering_outline_format(
            original_xml,
            excluded_abstract_ids={"99"},
        )
        updated_root = etree.fromstring(updated)
        updated_lvl = updated_root.xpath("./w:abstractNum/w:lvl", namespaces=NS)[0]
        updated_ind = updated_lvl.find("./w:pPr/w:ind", NS)

        # Suffix / tabs / lvlText trailing whitespace are preserved (not sanitized).
        self.assertEqual(updated_lvl.find("./w:suff", NS).get(qn("val")), "tab")
        self.assertIsNotNone(updated_lvl.find("./w:pPr/w:tabs", NS))
        self.assertEqual(updated_lvl.find("./w:lvlText", NS).get(qn("val")), "%1.　")
        # Indent is also untouched.
        self.assertEqual(updated_ind.get(qn("left")), "2279")
        self.assertEqual(updated_ind.get(qn("hanging")), "420")

    def test_numbering_xml_suffix_and_tabs_follow_outline_level_rule(self):
        # A. Build recognizable 0-8 numbering with trailing lvlText whitespace and
        # verify the normal numbering.xml format pass applies the tab-suffix rule.
        root = build_recognizable_nine_level_numbering()
        for lvl in root.xpath("./w:abstractNum/w:lvl", namespaces=NS):
            lvl_text = lvl.find("w:lvlText", NS)
            lvl_text.set(qn("val"), lvl_text.get(qn("val")) + " \t\u3000")

        updated = apply_numbering_outline_format(etree.tostring(root))
        updated_root = etree.fromstring(updated)

        for level in range(9):
            with self.subTest(level=level):
                lvl = updated_root.xpath(f"./w:abstractNum/w:lvl[@w:ilvl='{level}']", namespaces=NS)[0]
                # Suffix, tab stop, and trailing lvlText whitespace per the rule.
                assert_level_suffix_rule(self, lvl, level)
                ind = lvl.find("./w:pPr/w:ind", NS)
                expected_left, expected_hanging = expected_indent(level)
                self.assertEqual(ind.get(qn("left")), expected_left)
                self.assertEqual(ind.get(qn("hanging")), expected_hanging)

    def test_numbering_xml_lvl_override_resolves_base_format_for_tab_suffix(self):
        # B. A lvlOverride that carries only pPr (no numFmt/lvlText) must inherit
        # the base abstract level's format, so A./a./1. are not misread from ilvl.
        root = etree.Element(qn("numbering"), nsmap={"w": W_NS})
        abstract = etree.SubElement(root, qn("abstractNum"))
        abstract.set(qn("abstractNumId"), "1")
        base_levels = {0: ("upperLetter", "%1."), 1: ("lowerLetter", "%1."), 2: ("decimal", "%1.")}
        for ilvl, (num_fmt, lvl_text) in base_levels.items():
            lvl = etree.SubElement(abstract, qn("lvl"))
            lvl.set(qn("ilvl"), str(ilvl))
            fmt = etree.SubElement(lvl, qn("numFmt"))
            fmt.set(qn("val"), num_fmt)
            text_el = etree.SubElement(lvl, qn("lvlText"))
            text_el.set(qn("val"), lvl_text)

        num = etree.SubElement(root, qn("num"))
        num.set(qn("numId"), "42")
        abstract_ref = etree.SubElement(num, qn("abstractNumId"))
        abstract_ref.set(qn("val"), "1")
        for ilvl in base_levels:
            override = etree.SubElement(num, qn("lvlOverride"))
            override.set(qn("ilvl"), str(ilvl))
            override_lvl = etree.SubElement(override, qn("lvl"))
            override_lvl.set(qn("ilvl"), str(ilvl))
            # Only pPr, no numFmt/lvlText: the level must be recovered from base.
            etree.SubElement(override_lvl, qn("pPr"))

        updated = apply_numbering_outline_format(etree.tostring(root))
        updated_root = etree.fromstring(updated)

        # Base formats A.(level 5), a.(level 7), 1.(level 3) are all tab-suffix.
        expected_levels = {0: 5, 1: 7, 2: 3}
        for ilvl, level in expected_levels.items():
            with self.subTest(ilvl=ilvl):
                override_lvl = updated_root.xpath(
                    f"./w:num/w:lvlOverride[@w:ilvl='{ilvl}']/w:lvl",
                    namespaces=NS,
                )[0]
                assert_level_suffix_rule(self, override_lvl, level)

    def test_force_clean_numbering_suffix_tabs_applies_nine_level_rule(self):
        # C. Pollute every recognizable level (wrong suffix, wrong/duplicated
        # tabs, trailing lvlText whitespace); the final hard clean must rebuild
        # the exact nine-level rule.
        root = build_recognizable_nine_level_numbering(pollute=True)
        logs: list[str] = []
        updated = force_clean_numbering_suffix_tabs(etree.tostring(root), logs=logs)
        updated_root = etree.fromstring(updated)

        for level in range(9):
            with self.subTest(level=level):
                lvl = updated_root.xpath(f"./w:abstractNum/w:lvl[@w:ilvl='{level}']", namespaces=NS)[0]
                assert_level_suffix_rule(self, lvl, level)

        # Five tab-suffix levels (3/5/6/7/8) and four nothing-suffix levels.
        self.assertTrue(any("suffixes_set_to_tab=5" in log for log in logs))
        self.assertTrue(any("suffixes_set_to_nothing=4" in log for log in logs))
        self.assertTrue(any("tab_stops_rebuilt=5" in log for log in logs))
        self.assertTrue(any("tab_stops_removed=4" in log for log in logs))
        self.assertTrue(any("lvl_text_trimmed=9" in log for log in logs))


if __name__ == "__main__":
    unittest.main()



