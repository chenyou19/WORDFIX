"""Word COM acceptance test for the per-level number-suffix tab rule.

This is the real acceptance check requested for the schema-order fix: it is not
enough that numbering.xml contains the string w:suff="tab" - Microsoft Word must
actually ADOPT it. The test writes a full .docx with auto-numbering levels 0-8,
runs it through WORDFIX, opens the result in Word via COM, reads each list
level's TrailingCharacter through the Word object model, and finally lets Word
re-save the document and re-validates the numbering.xml.

It is Windows-only and skips cleanly when Word / win32com is unavailable instead
of pretending to pass.
"""
from __future__ import annotations

import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from docx_fixer.constants import NS, TEMPLATE_OUTLINE_INDENTS, W_NS
from docx_fixer.docx_processor import fix_docx_fast
from docx_fixer.models import ProcessOptions
from docx_fixer.numbering import uses_tab_suffix
from docx_fixer.word_com_numbering_suffix import (
    TRAILING_NONE,
    TRAILING_TAB,
    build_numbering_suffix_word_com_records,
    partition_records_by_template_conflict,
)
from docx_fixer.xml_utils import (
    LEVEL_CHILD_ORDER,
    PPR_CHILD_ORDER,
    children_in_schema_order,
    qn,
)

# (numFmt, lvlText) per detected outline level 0-8 (壹、一、（一）1.（1）A.（A）a.（a）).
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

CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    "</Types>"
).encode("utf-8")

ROOT_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
).encode("utf-8")

DOCUMENT_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
    "</Relationships>"
).encode("utf-8")


def _build_nine_level_numbering_xml() -> bytes:
    root = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    abstract = etree.SubElement(root, qn("abstractNum"))
    abstract.set(qn("abstractNumId"), "1")
    for level, (num_fmt, lvl_text) in RECOGNIZABLE_LEVEL_FORMATS.items():
        lvl = etree.SubElement(abstract, qn("lvl"))
        lvl.set(qn("ilvl"), str(level))
        etree.SubElement(lvl, qn("numFmt")).set(qn("val"), num_fmt)
        etree.SubElement(lvl, qn("lvlText")).set(qn("val"), lvl_text)
        etree.SubElement(lvl, qn("lvlJc")).set(qn("val"), "left")
    num = etree.SubElement(root, qn("num"))
    num.set(qn("numId"), "1")
    etree.SubElement(num, qn("abstractNumId")).set(qn("val"), "1")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _build_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    # Two 壹、序言 start markers so the body-outline processing kicks in, then one
    # auto-numbered paragraph per level 0-8.
    for marker in ("壹、序言", "壹、序言"):
        p = etree.SubElement(body, qn("p"))
        etree.SubElement(etree.SubElement(p, qn("r")), qn("t")).text = marker
    for level in range(9):
        p = etree.SubElement(body, qn("p"))
        pPr = etree.SubElement(p, qn("pPr"))
        num_pr = etree.SubElement(pPr, qn("numPr"))
        etree.SubElement(num_pr, qn("ilvl")).set(qn("val"), str(level))
        etree.SubElement(num_pr, qn("numId")).set(qn("val"), "1")
        etree.SubElement(etree.SubElement(p, qn("r")), qn("t")).text = f"自動編號第{level}階"
    etree.SubElement(body, qn("sectPr"))
    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


def _write_full_docx(path: Path) -> None:
    with ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        zf.writestr("_rels/.rels", ROOT_RELS_XML)
        zf.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS_XML)
        zf.writestr("word/document.xml", _build_document_xml())
        zf.writestr("word/numbering.xml", _build_nine_level_numbering_xml())


_RPC_E_CALL_REJECTED = -2147418111


def _com_call_with_retry(func, attempts: int = 5, delay: float = 0.5):
    """Retry a Word COM call that transiently fails with "call rejected by callee".

    Word frequently returns RPC_E_CALL_REJECTED while it is still starting up or
    busy; the call is fine on a short retry. Other COM errors propagate.
    """
    import time

    last_exc = None
    for _ in range(attempts):
        try:
            return func()
        except Exception as exc:  # pragma: no cover - timing dependent
            hresult = getattr(exc, "hresult", None) or getattr(exc, "args", [None])[0]
            if hresult != _RPC_E_CALL_REJECTED:
                raise
            last_exc = exc
            time.sleep(delay)
    raise last_exc


def _read_numbering_levels(docx_path: Path) -> dict[int, object]:
    with ZipFile(docx_path, "r") as zf:
        root = etree.fromstring(zf.read("word/numbering.xml"))
    levels: dict[int, object] = {}
    for lvl in root.xpath("./w:abstractNum/w:lvl", namespaces=NS):
        try:
            levels[int(lvl.get(qn("ilvl")))] = lvl
        except (TypeError, ValueError):
            continue
    return levels


def _write_docx(path: Path, document_xml: bytes, numbering_xml: bytes) -> None:
    with ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        zf.writestr("_rels/.rels", ROOT_RELS_XML)
        zf.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS_XML)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/numbering.xml", numbering_xml)


def _records_numbering_xml() -> bytes:
    root = etree.Element(qn("numbering"), nsmap={"w": W_NS})
    abstract1 = etree.SubElement(root, qn("abstractNum"))
    abstract1.set(qn("abstractNumId"), "1")
    for level, (num_fmt, lvl_text) in RECOGNIZABLE_LEVEL_FORMATS.items():
        lvl = etree.SubElement(abstract1, qn("lvl"))
        lvl.set(qn("ilvl"), str(level))
        etree.SubElement(lvl, qn("numFmt")).set(qn("val"), num_fmt)
        etree.SubElement(lvl, qn("lvlText")).set(qn("val"), lvl_text)
    # abstract 2 / num 2: an unrecognized format that must never become a record.
    abstract2 = etree.SubElement(root, qn("abstractNum"))
    abstract2.set(qn("abstractNumId"), "2")
    lvl = etree.SubElement(abstract2, qn("lvl"))
    lvl.set(qn("ilvl"), "0")
    etree.SubElement(lvl, qn("numFmt")).set(qn("val"), "custom")
    etree.SubElement(lvl, qn("lvlText")).set(qn("val"), "%1")
    for num_id, abstract_id in (("1", "1"), ("2", "2")):
        num = etree.SubElement(root, qn("num"))
        num.set(qn("numId"), num_id)
        etree.SubElement(num, qn("abstractNumId")).set(qn("val"), abstract_id)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _auto_paragraph(body, num_id: str, ilvl: int, text: str) -> None:
    p = etree.SubElement(body, qn("p"))
    pPr = etree.SubElement(p, qn("pPr"))
    num_pr = etree.SubElement(pPr, qn("numPr"))
    etree.SubElement(num_pr, qn("ilvl")).set(qn("val"), str(ilvl))
    etree.SubElement(num_pr, qn("numId")).set(qn("val"), num_id)
    etree.SubElement(etree.SubElement(p, qn("r")), qn("t")).text = text


def _records_document_xml() -> bytes:
    document = etree.Element(qn("document"), nsmap={"w": W_NS})
    body = etree.SubElement(document, qn("body"))
    # 1: manual numbering (no numPr) -> never a record.
    p = etree.SubElement(body, qn("p"))
    etree.SubElement(etree.SubElement(p, qn("r")), qn("t")).text = "1. 手動標題"
    _auto_paragraph(body, "1", 3, "自動三階")  # 2 -> level 3 (tab)
    _auto_paragraph(body, "2", 0, "未知格式")  # 3 -> unrecognized -> skipped
    _auto_paragraph(body, "1", 5, "自動五階")  # 4 -> level 5 (tab)
    _auto_paragraph(body, "1", 0, "自動零階")  # 5 -> level 0 (nothing)
    etree.SubElement(body, qn("sectPr"))
    return etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)


class NumberingSuffixRecordTests(unittest.TestCase):
    """Python-level tests for record building and conflict detection (no Word)."""

    def _build_records(self, protected_pairs=None):
        with tempfile.TemporaryDirectory() as temp_dir:
            docx_path = Path(temp_dir) / "doc.docx"
            _write_docx(docx_path, _records_document_xml(), _records_numbering_xml())
            return build_numbering_suffix_word_com_records(
                docx_path, protected_numbering_pairs=protected_pairs or set()
            )

    def test_builds_records_only_for_recognized_auto_numbering(self):
        records = self._build_records()
        by_index = {r["paragraph_index"]: r for r in records}
        # Manual (1) and unrecognized auto (3) must be absent.
        self.assertEqual(set(by_index), {2, 4, 5})
        self.assertEqual(by_index[2]["outline_level"], 3)
        self.assertEqual(by_index[2]["expected_trailing"], TRAILING_TAB)
        self.assertEqual(by_index[2]["expected_tab_pos_twips"], int(TEMPLATE_OUTLINE_INDENTS[3]["left"]))
        self.assertEqual(by_index[4]["outline_level"], 5)
        self.assertEqual(by_index[4]["expected_trailing"], TRAILING_TAB)
        self.assertEqual(by_index[5]["outline_level"], 0)
        self.assertEqual(by_index[5]["expected_trailing"], TRAILING_NONE)
        self.assertIsNone(by_index[5]["expected_tab_pos_twips"])
        self.assertFalse(any(r["is_protected"] for r in records))

    def test_chapter_three_protected_pair_is_flagged_not_dropped(self):
        # 參 protection on (numId=1, ilvl=3): the record exists but is protected.
        records = self._build_records(protected_pairs={("1", 3)})
        by_index = {r["paragraph_index"]: r for r in records}
        self.assertTrue(by_index[3 - 1]["is_protected"])  # paragraph index 2 is ilvl 3
        self.assertFalse(by_index[4]["is_protected"])

    def test_partition_applies_clean_targets(self):
        records = self._build_records()
        apply_records, protected_records, conflicts = partition_records_by_template_conflict(records)
        self.assertEqual({r["paragraph_index"] for r in apply_records}, {2, 4, 5})
        self.assertEqual(protected_records, [])
        self.assertEqual(conflicts, [])

    def test_partition_flags_template_conflict_with_protected(self):
        # A target and a protected record sharing (abstractNumId, ilvl) -> conflict;
        # the target is NOT applied and a conflict is logged.
        records = [
            {
                "paragraph_index": 10,
                "abstract_id": "1",
                "ilvl": 3,
                "num_id": "1",
                "outline_level": 3,
                "expected_trailing": TRAILING_TAB,
                "expected_tab_pos_twips": 2262,
                "is_protected": False,
            },
            {
                "paragraph_index": 11,
                "abstract_id": "1",
                "ilvl": 3,
                "num_id": "5",
                "outline_level": 3,
                "expected_trailing": TRAILING_TAB,
                "expected_tab_pos_twips": 2262,
                "is_protected": True,
            },
        ]
        apply_records, protected_records, conflicts = partition_records_by_template_conflict(records)
        self.assertEqual(apply_records, [])
        self.assertEqual(len(protected_records), 1)
        self.assertEqual(len(conflicts), 1)
        self.assertIn("WORD_COM_NUMBERING_SUFFIX_TEMPLATE_CONFLICT", conflicts[0])
        self.assertIn("shared_with_protected", conflicts[0])


@unittest.skipUnless(sys.platform == "win32", "Word COM is only available on Windows")
class WordComNumberingSuffixIntegrationTests(unittest.TestCase):
    def test_word_adopts_per_level_trailing_character(self):
        try:
            import pythoncom  # type: ignore[import-not-found]
            import win32com.client  # type: ignore[import-not-found]
            from win32com.client import constants, gencache  # type: ignore[import-not-found]
        except ImportError as exc:  # pragma: no cover - env dependent
            self.skipTest(f"win32com unavailable: {exc!r}")

        # Manual temp dir: Word can hold a brief lock on the .docx after Quit, so
        # cleanup must tolerate that instead of failing the test.
        temp_dir = tempfile.mkdtemp()
        try:
            input_docx = Path(temp_dir) / "input.docx"
            output_docx = Path(temp_dir) / "output.docx"
            resaved_docx = Path(temp_dir) / "resaved.docx"
            _write_full_docx(input_docx)

            # Full production pipeline WITH Word COM, so the real
            # apply_and_verify_numbering_suffixes_with_word_com step runs.
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

            com_result = summary.word_com_numbering_suffix_result
            self.assertIsNotNone(com_result, "Word COM numbering-suffix step did not run")
            if not getattr(com_result, "word_com_available", False):
                self.skipTest("Word COM not available during pipeline run")
            # Word itself applied and re-verified every level (no XML-only claim).
            self.assertEqual(com_result.template_conflicts, 0)
            self.assertEqual(com_result.com_apply_failed, 0)
            self.assertEqual(com_result.com_verify_failed, 0)
            self.assertEqual(com_result.format_identity_mismatch, 0)
            self.assertEqual(com_result.tab_position_mismatch, 0)
            self.assertEqual(com_result.expected_tab, 5)
            self.assertEqual(com_result.expected_nothing, 4)
            self.assertEqual(com_result.actual_tab, 5)
            self.assertEqual(com_result.actual_nothing, 4)
            self.assertTrue(com_result.verified)

            pythoncom.CoInitialize()
            word = None
            doc = None
            try:
                try:
                    # EnsureDispatch generates the typelib so win32com.client.constants
                    # is populated (no hard-coded Word enum numbers in this test).
                    word = _com_call_with_retry(lambda: gencache.EnsureDispatch("Word.Application"))
                except Exception as exc:  # pragma: no cover - env dependent
                    self.skipTest(f"Word COM unavailable: {exc!r}")

                word.Visible = False
                try:
                    word.DisplayAlerts = constants.wdAlertsNone
                except Exception:
                    pass

                doc = _com_call_with_retry(
                    lambda: word.Documents.Open(
                        str(output_docx.resolve()),
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        Visible=False,
                    )
                )

                # Read each list level's trailing character through the Word OM.
                trailing_by_level: dict[int, int] = {}
                for paragraph in doc.Paragraphs:
                    list_format = paragraph.Range.ListFormat
                    list_template = list_format.ListTemplate
                    if list_template is None:
                        continue
                    level_number = int(list_format.ListLevelNumber)  # 1-based
                    trailing = int(list_template.ListLevels(level_number).TrailingCharacter)
                    trailing_by_level[level_number - 1] = trailing

                self.assertEqual(
                    set(trailing_by_level),
                    set(range(9)),
                    f"Word did not expose all nine list levels: {sorted(trailing_by_level)}",
                )
                for level in range(9):
                    expected = constants.wdTrailingTab if uses_tab_suffix(level) else constants.wdTrailingNone
                    self.assertEqual(
                        trailing_by_level[level],
                        int(expected),
                        f"level {level}: Word TrailingCharacter {trailing_by_level[level]} != expected {int(expected)}",
                    )

                # Let Word re-serialize the document, then re-validate numbering.xml.
                doc.SaveAs2(str(resaved_docx.resolve()), FileFormat=constants.wdFormatXMLDocument)
                doc.Close(SaveChanges=False)
                doc = None
            finally:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass
                if word is not None:
                    word.Quit()
                pythoncom.CoUninitialize()

            # After Word's own round-trip the rule and child order must still hold.
            levels = _read_numbering_levels(resaved_docx)
            self.assertEqual(set(levels), set(range(9)))
            for level, lvl in levels.items():
                suff = lvl.find("./w:suff", NS)
                tabs = lvl.find("./w:pPr/w:tabs", NS)
                self.assertLessEqual(
                    len(lvl.findall("./w:suff", NS)), 1, f"level {level}: duplicate w:suff"
                )
                self.assertTrue(
                    children_in_schema_order(lvl, LEVEL_CHILD_ORDER),
                    f"level {level}: w:lvl children out of schema order after Word save",
                )
                pPr = lvl.find("./w:pPr", NS)
                if pPr is not None:
                    self.assertTrue(
                        children_in_schema_order(pPr, PPR_CHILD_ORDER),
                        f"level {level}: w:pPr children out of schema order after Word save",
                    )
                suff_val = suff.get(qn("val")) if suff is not None else None
                if uses_tab_suffix(level):
                    # WordprocessingML's default w:suff is "tab", so on re-save Word
                    # drops the explicit element; an absent w:suff here still means
                    # Tab. That Word kept it as default is itself proof of adoption.
                    self.assertIn(
                        suff_val, (None, "tab"), f"level {level}: expected tab (or default) suffix, got {suff_val}"
                    )
                    self.assertIsNotNone(tabs, f"level {level}: expected a tab stop")
                else:
                    self.assertEqual(suff_val, "nothing", f"level {level}: expected nothing suffix")
                    self.assertIsNone(tabs, f"level {level}: unexpected tab stop")
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
