from __future__ import annotations

import json
import shutil
import tempfile
import time
import re
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from .constants import NS
from .exceptions import ProcessStopped
from .process_runner import run_powershell_file, run_powershell_script
from .stop_controller import StopController
from .xml_utils import paragraph_text, qn

POINTS_PER_CM = 28.3464567
WORD_COM_TIMEOUT_SECONDS = 600
WORD_COM_TEMP_DIR_NAME = "wfix"

def create_word_application_for_com_fix(logs: list[str]):
    try:
        import win32com.client  # type: ignore[import-not-found]

        return win32com.client.DispatchEx("Word.Application")
    except AttributeError as exc:
        if "CLSIDToClassMap" not in str(exc):
            raise

        logs.append(f"WORD_COM_RETRY_CLEAR_GEN_PY reason={exc!r}")

        try:
            import shutil
            import win32com.client.gencache as gencache  # type: ignore[import-not-found]

            gen_path = gencache.GetGeneratePath()
            logs.append(f"WORD_COM_GEN_PY_PATH path={gen_path}")
            shutil.rmtree(gen_path, ignore_errors=True)
            try:
                gencache.Rebuild()
            except Exception as rebuild_exc:
                logs.append(f"WORD_COM_GEN_PY_REBUILD_FAILED reason={rebuild_exc!r}")
        except Exception as cleanup_exc:
            logs.append(f"WORD_COM_GEN_PY_CLEANUP_FAILED reason={cleanup_exc!r}")

        try:
            import win32com.client  # type: ignore[import-not-found]

            word = win32com.client.DispatchEx("Word.Application")
            logs.append("WORD_COM_RETRY_AFTER_CLEAR_GEN_PY status=ok")
            return word
        except Exception as retry_exc:
            logs.append(f"WORD_COM_RETRY_AFTER_CLEAR_GEN_PY status=failed reason={retry_exc!r}")

        try:
            from win32com.client import dynamic  # type: ignore[import-not-found]

            word = dynamic.Dispatch("Word.Application")
            logs.append("WORD_COM_DYNAMIC_DISPATCH status=ok")
            return word
        except Exception as dynamic_exc:
            logs.append(f"WORD_COM_DYNAMIC_DISPATCH status=failed reason={dynamic_exc!r}")
            raise

def get_word_table_start_pages(input_docx: Path, stop: StopController | None = None) -> list[int | None]:
    wd_collapse_start = 1
    wd_active_end_page_number = 3

    word = None
    doc = None
    pages: list[int | None] = []
    com_failed = False

    try:
        word = create_word_application_for_com_fix([])
        word.Visible = False
        doc = word.Documents.Open(
            str(input_docx.resolve()),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        doc.Repaginate()

        for table in doc.Tables:
            if stop:
                stop.check()

            try:
                table_range = table.Range.Duplicate
                table_range.Collapse(wd_collapse_start)
                pages.append(int(table_range.Information(wd_active_end_page_number)))
            except Exception:
                pages.append(None)
    except Exception:
        com_failed = True
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

    if com_failed:
        return get_word_table_start_pages_with_powershell(input_docx, stop=stop)

    return pages


def get_word_table_start_pages_with_powershell(
    input_docx: Path,
    stop: StopController | None = None,
) -> list[int | None]:
    path_literal = "'" + str(input_docx.resolve()).replace("'", "''") + "'"
    script = f"""
$ErrorActionPreference = 'Stop'
$path = {path_literal}
$word = $null
$doc = $null
$pages = New-Object System.Collections.Generic.List[object]
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($path, $false, $true, $false)
    $doc.Repaginate()
    foreach ($table in $doc.Tables) {{
        if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
        $range = $table.Range.Duplicate
        $range.Collapse(1)
        $pages.Add([int]$range.Information(3))
    }}
    $pages | ConvertTo-Json -Compress
}} catch {{
    "[]"
}} finally {{
    if ($doc -ne $null) {{ $doc.Close($false) | Out-Null }}
    if ($word -ne $null) {{ $word.Quit() | Out-Null }}
}}
"""

    try:
        completed = run_powershell_script(script, stop=stop, timeout=120)
    except ProcessStopped:
        raise
    except Exception:
        return []

    output = completed.stdout.strip()
    if not output:
        return []

    try:
        parsed = json.loads(output)
    except json.JSONDecodeError:
        return []

    if isinstance(parsed, int):
        return [parsed]
    if not isinstance(parsed, list):
        return []

    pages: list[int | None] = []
    for page in parsed:
        try:
            pages.append(int(page))
        except (TypeError, ValueError):
            pages.append(None)
    return pages


def get_rendered_table_start_pages(root) -> list[int]:
    pages: list[int] = []
    current_page = 1

    for element in root.iter():
        if element.tag == qn("tbl"):
            pages.append(current_page)
        elif element.tag == qn("lastRenderedPageBreak"):
            current_page += 1

    return pages


def normalize_word_paragraph_text(text: str) -> str:
    cleaned = (text or "").replace("\r", " ").replace("\x07", " ").replace("\x0b", " ")
    return re.sub(r"\s+", " ", cleaned).strip()


def paragraph_text_matches_preview(actual_text: str, preview: str) -> bool:
    normalized_actual = normalize_word_paragraph_text(actual_text)
    normalized_preview = normalize_word_paragraph_text(preview)
    if normalized_preview.endswith("..."):
        normalized_preview = normalized_preview[:-3].rstrip()
    if not normalized_preview:
        return False
    return normalized_actual.startswith(normalized_preview)


def find_word_paragraph_index_for_record(paragraph_texts: list[str], record: dict[str, object]) -> int | None:
    target_index = int(record.get("paragraph_index") or 0)
    preview = str(record.get("text_match_prefix") or record.get("text_preview") or "")

    if 1 <= target_index <= len(paragraph_texts):
        if paragraph_text_matches_preview(paragraph_texts[target_index - 1], preview):
            return target_index

    for index, text in enumerate(paragraph_texts, start=1):
        if paragraph_text_matches_preview(text, preview):
            return index

    return None


def _safe_com_attr(obj, name: str):
    try:
        return getattr(obj, name)
    except Exception:
        return None


def _points_to_cm(value) -> float | None:
    try:
        return float(value) / POINTS_PER_CM
    except (TypeError, ValueError):
        return None


def _format_optional_cm(value: float | None) -> str:
    if value is None:
        return "None"
    return f"{value:.2f}"


def _safe_tab_stops_count(pf):
    try:
        return pf.TabStops.Count
    except Exception:
        return None


def _safe_clear_tab_stops(pf) -> bool:
    try:
        pf.TabStops.ClearAll()
        return True
    except Exception:
        return False


def _safe_paragraph_style_name(paragraph):
    try:
        style = paragraph.Style
        return getattr(style, "NameLocal", str(style))
    except Exception:
        return None


def _safe_section_diagnostics(paragraph):
    try:
        section = paragraph.Range.Sections(1)
    except Exception:
        try:
            section = paragraph.Range.Sections.Item(1)
        except Exception:
            return None, None, None

    try:
        section_number = section.Index
    except Exception:
        section_number = None

    page_setup = getattr(section, "PageSetup", None)
    left_margin_cm = _points_to_cm(_safe_com_attr(page_setup, "LeftMargin")) if page_setup is not None else None
    right_margin_cm = _points_to_cm(_safe_com_attr(page_setup, "RightMargin")) if page_setup is not None else None
    return section_number, left_margin_cm, right_margin_cm


def _format_optional_pt(value: float | None) -> str:
    if value is None:
        return "None"
    return f"{value:g}"


def _is_meaningful_word_font_size(value: object) -> bool:
    try:
        size = float(value)
    except (TypeError, ValueError):
        return False
    return 0 < size < 200 and abs(size) not in {9999999.0, 999999.0}


def _word_paragraph_font_sizes(paragraph) -> tuple[float | None, float | None]:
    try:
        text_range = paragraph.Range.Duplicate
    except Exception:
        text_range = None

    if text_range is not None:
        try:
            if text_range.End > text_range.Start:
                text_range.End = text_range.End - 1
        except Exception:
            pass

    range_font_size = None
    if text_range is not None:
        try:
            value = text_range.Font.Size
            if _is_meaningful_word_font_size(value):
                range_font_size = float(value)
        except Exception:
            pass

    if range_font_size is not None:
        return range_font_size, range_font_size

    weighted_sizes: dict[float, int] = {}
    order: list[float] = []
    try:
        chars = text_range.Characters if text_range is not None else paragraph.Range.Characters
        count = int(chars.Count)
    except Exception:
        return None, None

    for index in range(1, count + 1):
        try:
            char = chars(index)
        except Exception:
            try:
                char = chars.Item(index)
            except Exception:
                continue
        try:
            size_value = char.Font.Size
        except Exception:
            try:
                size_value = char.Range.Font.Size
            except Exception:
                continue
        if not _is_meaningful_word_font_size(size_value):
            continue
        size = float(size_value)
        try:
            text = str(char.Text)
            weight = len(text.strip()) or 1
        except Exception:
            weight = 1
        if size not in weighted_sizes:
            weighted_sizes[size] = 0
            order.append(size)
        weighted_sizes[size] += weight

    if not weighted_sizes:
        return None, None

    dominant_size = max(order, key=lambda size: (weighted_sizes[size], -order.index(size)))
    return range_font_size, dominant_size


def _verify_and_fix_body_indents_with_word_com_in_process(
    output_docx: Path,
    body_indent_records: list[dict[str, object]],
) -> list[str]:
    if not body_indent_records:
        return ["WORD_COM_BODY_INDENT_FIX_SKIPPED reason=no_records"]

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        return ["WORD_COM_BODY_INDENT_FIX_SKIPPED reason=win32com_unavailable"]

    word = None
    doc = None
    logs: list[str] = []
    try:
        word = create_word_application_for_com_fix(logs)
        word.Visible = False
        doc = word.Documents.Open(
            str(output_docx.resolve()),
            ReadOnly=False,
            AddToRecentFiles=False,
            Visible=False,
        )

        paragraphs = [doc.Paragraphs(i).Range.Text for i in range(1, doc.Paragraphs.Count + 1)]
        for record in body_indent_records:
            match_index = find_word_paragraph_index_for_record(paragraphs, record)
            preview = str(record.get("text_preview") or "")
            kind = str(record.get("kind") or "body")
            expected_left_cm = float(record.get("expected_left_cm") or 0.0)
            expected_left_points = float(record.get("expected_left_points") or 0.0)
            expected_firstline_cm = float(record.get("expected_firstline_cm") or 0.0)
            expected_firstline_points = float(record.get("expected_firstline_points") or 0.0)
            apply_only_if_word_font_size_is_14 = bool(record.get("apply_only_if_word_font_size_is_14"))

            if match_index is None:
                logs.append(
                    "WORD_COM_BODY_INDENT_FIX: "
                    f"paragraph_index={record.get('paragraph_index')}; "
                    f"text={preview!r}; expected_left_cm={expected_left_cm:.2f}; status=not_found"
                )
                continue

            paragraph = doc.Paragraphs(match_index)
            pf = paragraph.Format

            before_left_pt = _safe_com_attr(pf, "LeftIndent")
            before_left_cm = _points_to_cm(before_left_pt)
            before_first_line_pt = _safe_com_attr(pf, "FirstLineIndent")
            before_first_line_cm = _points_to_cm(before_first_line_pt)
            before_char_left = _safe_com_attr(pf, "CharacterUnitLeftIndent")
            before_char_first = _safe_com_attr(pf, "CharacterUnitFirstLineIndent")
            before_char_right = _safe_com_attr(pf, "CharacterUnitRightIndent")
            before_tab_count = _safe_tab_stops_count(pf)
            style_name_local = _safe_paragraph_style_name(paragraph)
            section_number, section_left_margin_cm, section_right_margin_cm = _safe_section_diagnostics(paragraph)
            word_range_font_size, word_dominant_font_size = (
                _word_paragraph_font_sizes(paragraph)
                if apply_only_if_word_font_size_is_14
                else (None, None)
            )
            word_font_check_pass = (
                not apply_only_if_word_font_size_is_14
                or (
                    word_dominant_font_size is not None
                    and abs(word_dominant_font_size - 14.0) <= 0.01
                )
            )

            if not word_font_check_pass:
                logs.append(
                    "WORD_COM_INDENT_VERIFY: "
                    f"paragraph_index={record.get('paragraph_index')}; "
                    f"matched_paragraph_index={match_index}; "
                    f"text_preview={preview!r}; "
                    f"kind={kind}; "
                    f"level={record.get('level')}; "
                    f"xml_font_size={record.get('xml_font_size')}; "
                    f"xml_font_size_source={record.get('xml_font_size_source')}; "
                    f"word_range_font_size={_format_optional_pt(word_range_font_size)}; "
                    f"word_dominant_font_size={_format_optional_pt(word_dominant_font_size)}; "
                    f"word_font_check_pass=False; "
                    f"decision=skipped_word_font_not_14; "
                    f"status=skipped_word_font_not_14"
                )
                logs.append(
                    "WORD_COM_BODY_INDENT_FIX: "
                    f"paragraph_index={record.get('paragraph_index')}; "
                    f"matched_paragraph_index={match_index}; "
                    f"status=skipped_word_font_not_14"
                )
                continue

            try:
                pf.CharacterUnitLeftIndent = 0
            except Exception:
                pass
            try:
                pf.CharacterUnitFirstLineIndent = 0
            except Exception:
                pass
            try:
                pf.CharacterUnitRightIndent = 0
            except Exception:
                pass
            tabs_cleared = _safe_clear_tab_stops(pf)

            pf.LeftIndent = expected_left_points
            pf.FirstLineIndent = expected_firstline_points

            actual_left_pt = _safe_com_attr(pf, "LeftIndent")
            actual_left_cm = _points_to_cm(actual_left_pt)
            actual_first_line_pt = _safe_com_attr(pf, "FirstLineIndent")
            actual_char_left = _safe_com_attr(pf, "CharacterUnitLeftIndent")
            actual_char_first = _safe_com_attr(pf, "CharacterUnitFirstLineIndent")
            actual_char_right = _safe_com_attr(pf, "CharacterUnitRightIndent")
            actual_tab_count = _safe_tab_stops_count(pf)
            first_line_indent_cm = _points_to_cm(actual_first_line_pt)
            absolute_text_start_cm = (
                section_left_margin_cm + actual_left_cm
                if section_left_margin_cm is not None and actual_left_cm is not None
                else None
            )
            status = (
                "ok"
                if actual_left_cm is not None and abs(actual_left_cm - expected_left_cm) <= 0.02
                and first_line_indent_cm is not None
                and abs(first_line_indent_cm - expected_firstline_cm) <= 0.02
                else "mismatch"
            )

            logs.append(
                "WORD_COM_INDENT_VERIFY: "
                f"paragraph_index={record.get('paragraph_index')}; "
                f"matched_paragraph_index={match_index}; "
                f"text_preview={preview!r}; "
                f"kind={kind}; "
                f"level={record.get('level')}; "
                f"expected_number_start_cm={record.get('expected_number_start_cm')}; "
                f"expected_hanging_cm={record.get('expected_hanging_cm')}; "
                f"expected_heading_left_cm={record.get('expected_heading_left_cm')}; "
                f"expected_body_left_cm={record.get('expected_body_left_cm')}; "
                f"expected_first_line_twips={record.get('expected_first_line_twips')}; "
                f"xml_font_size={record.get('xml_font_size')}; "
                f"xml_font_size_source={record.get('xml_font_size_source')}; "
                f"word_range_font_size={_format_optional_pt(word_range_font_size)}; "
                f"word_dominant_font_size={_format_optional_pt(word_dominant_font_size)}; "
                f"word_font_check_pass={word_font_check_pass}; "
                f"decision=apply_body_indent; "
                f"xml_written_left_cm={record.get('xml_written_left_cm')}; "
                f"xml_written_hanging_cm={record.get('xml_written_hanging_cm')}; "
                f"word_opened_left_cm={_format_optional_cm(before_left_cm)}; "
                f"word_opened_firstline_cm={_format_optional_cm(before_first_line_cm)}; "
                f"final_left_cm={_format_optional_cm(actual_left_cm)}; "
                f"final_firstline_cm={_format_optional_cm(first_line_indent_cm)}; "
                f"word_com_LeftIndent_cm={_format_optional_cm(actual_left_cm)}; "
                f"word_com_FirstLineIndent_cm={_format_optional_cm(first_line_indent_cm)}; "
                f"word_com_CharacterUnitLeftIndent={actual_char_left}; "
                f"word_com_CharacterUnitFirstLineIndent={actual_char_first}; "
                f"word_com_CharacterUnitRightIndent={actual_char_right}; "
                f"word_com_TabStops_Count={actual_tab_count}; "
                f"word_com_Style_NameLocal={style_name_local}; "
                f"word_com_Section_Number={section_number}; "
                f"section_index={section_number}; "
                f"section_left_margin_cm={_format_optional_cm(section_left_margin_cm)}; "
                f"section_right_margin_cm={_format_optional_cm(section_right_margin_cm)}; "
                f"paragraph_left_indent_cm={_format_optional_cm(actual_left_cm)}; "
                f"absolute_text_start_cm={_format_optional_cm(absolute_text_start_cm)}; "
                f"second_fix=yes; "
                f"status={status}"
            )
            logs.append(
                "WORD_COM_BODY_INDENT_FIX: "
                f"paragraph_index={record.get('paragraph_index')}; "
                f"matched_paragraph_index={match_index}; "
                f"text={preview!r}; "
                f"expected_left_cm={expected_left_cm:.2f}; "
                f"before_left_cm={_format_optional_cm(before_left_cm)}; "
                f"after_left_cm={_format_optional_cm(actual_left_cm)}; "
                f"before_char_left={before_char_left}; "
                f"after_char_left={actual_char_left}; "
                f"before_char_first={before_char_first}; "
                f"after_char_first={actual_char_first}; "
                f"before_char_right={before_char_right}; "
                f"after_char_right={actual_char_right}; "
                f"before_tab_count={before_tab_count}; "
                f"after_tab_count={actual_tab_count}; "
                f"tabs_cleared={tabs_cleared}; "
                f"first_line_indent_cm={_format_optional_cm(first_line_indent_cm)}; "
                f"status={status}"
            )

        doc.Save()
    except Exception as exc:
        retry_logs = [
            log
            for log in logs
            if log.startswith("WORD_COM_RETRY_")
            or log.startswith("WORD_COM_GEN_PY_")
            or log.startswith("WORD_COM_DYNAMIC_DISPATCH")
        ]
        retry_reason = f"; retries={' | '.join(retry_logs)}" if retry_logs else ""
        logs.append(f"WORD_COM_BODY_INDENT_FIX_SKIPPED reason={type(exc).__name__}:{exc}{retry_reason}")
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

    return logs


def _parse_powershell_log_output(output: str) -> list[str]:
    text = output.strip()
    if not text:
        return []

    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        return []

    if isinstance(parsed, str):
        return [parsed]
    if not isinstance(parsed, list):
        return []

    return [str(item) for item in parsed]


def _parse_plain_log_output(output: str) -> list[str]:
    return [line.strip() for line in output.splitlines() if line.strip()]


def _collect_word_com_powershell_logs(stdout: str, result_path: Path) -> list[str]:
    logs = _parse_powershell_log_output(stdout)
    if logs:
        return logs

    logs = _parse_plain_log_output(stdout)
    if logs:
        return logs

    if result_path.exists():
        try:
            return _parse_plain_log_output(result_path.read_text(encoding="utf-8"))
        except Exception:
            return []

    return []


def _parse_word_com_approved_records(script_logs: list[str]) -> list[dict[str, object]]:
    approved_records: list[dict[str, object]] = []
    prefix = "WORD_COM_APPROVED_RECORD_JSON "
    for log in script_logs:
        if not log.startswith(prefix):
            continue
        try:
            payload = json.loads(log[len(prefix) :])
        except json.JSONDecodeError:
            continue
        if isinstance(payload, dict):
            approved_records.append(payload)
    return approved_records


def _get_or_add_paragraph_properties(p):
    p_pr = p.find("w:pPr", NS)
    if p_pr is None:
        p_pr = etree.Element(qn("pPr"))
        p.insert(0, p_pr)
    return p_pr


def _get_or_add_indent(p_pr):
    ind = p_pr.find("w:ind", NS)
    if ind is None:
        ind = etree.Element(qn("ind"))
        p_pr.append(ind)
    return ind


def _apply_body_indent_record_to_paragraph(p, record: dict[str, object]) -> str:
    expected_left_twips = record.get("expected_left_twips")
    if expected_left_twips is None or str(expected_left_twips).strip() == "":
        return "skipped_missing_expected_left_twips"

    match_prefix = str(record.get("text_match_prefix") or "")
    if match_prefix and not paragraph_text_matches_preview(paragraph_text(p), match_prefix):
        return "skipped_text_mismatch"

    p_pr = _get_or_add_paragraph_properties(p)
    tabs = p_pr.find("w:tabs", NS)
    if tabs is not None:
        p_pr.remove(tabs)
    num_pr = p_pr.find("w:numPr", NS)
    if num_pr is not None:
        p_pr.remove(num_pr)

    ind = _get_or_add_indent(p_pr)
    for attr in (
        "left",
        "start",
        "right",
        "end",
        "firstLine",
        "hanging",
        "leftChars",
        "startChars",
        "rightChars",
        "endChars",
        "firstLineChars",
        "hangingChars",
    ):
        ind.attrib.pop(qn(attr), None)

    ind.set(qn("left"), str(expected_left_twips))
    ind.set(qn("start"), str(expected_left_twips))
    expected_first_line_twips = record.get("expected_first_line_twips")
    if expected_first_line_twips is not None and str(expected_first_line_twips).strip():
        ind.set(qn("firstLine"), str(expected_first_line_twips))
    else:
        ind.set(qn("firstLine"), "0")
    ind.set(qn("hanging"), "0")
    ind.set(qn("leftChars"), "0")
    ind.set(qn("startChars"), "0")
    ind.set(qn("firstLineChars"), "0")
    ind.set(qn("hangingChars"), "0")
    return "applied"


def apply_word_com_approved_body_indents_to_docx_xml(
    output_docx: Path,
    approved_records: list[dict[str, object]],
) -> list[str]:
    logs = [
        f"WORD_COM_XML_APPLY_STARTED approved_records={len(approved_records)}",
    ]
    if not approved_records:
        logs.append("WORD_COM_XML_APPLY_SKIPPED reason=no_approved_records")
        return logs

    applied = 0
    skipped = 0
    errors = 0
    temp_docx = output_docx.with_suffix(output_docx.suffix + ".word_com_xml.tmp")

    try:
        with ZipFile(output_docx, "r") as zin, ZipFile(temp_docx, "w", ZIP_DEFLATED) as zout:
            document_root = etree.fromstring(zin.read("word/document.xml"))
            paragraphs = document_root.xpath(".//w:p", namespaces=NS)

            for record in approved_records:
                try:
                    paragraph_index = int(record.get("paragraph_index") or 0)
                    if paragraph_index < 1 or paragraph_index > len(paragraphs):
                        status = "skipped_index_out_of_range"
                        skipped += 1
                    else:
                        status = _apply_body_indent_record_to_paragraph(paragraphs[paragraph_index - 1], record)
                        if status == "applied":
                            applied += 1
                        else:
                            skipped += 1
                    logs.append(
                        f"WORD_COM_XML_APPLY_RECORD paragraph_index={record.get('paragraph_index')} status={status}"
                    )
                except Exception as exc:
                    errors += 1
                    logs.append(
                        "WORD_COM_XML_APPLY_RECORD "
                        f"paragraph_index={record.get('paragraph_index')} "
                        f"status=error type={type(exc).__name__} message={exc}"
                    )

            document_xml = etree.tostring(
                document_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, document_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

        shutil.move(_long_path_compatible_str(temp_docx), _long_path_compatible_str(output_docx))
    except Exception as exc:
        errors += 1
        logs.append(f"WORD_COM_XML_APPLY_FAILED type={type(exc).__name__} message={exc}")
    finally:
        try:
            temp_docx.unlink()
        except FileNotFoundError:
            pass
        except Exception:
            pass

    logs.append(f"WORD_COM_XML_APPLY_DONE applied={applied} skipped={skipped} errors={errors}")
    return logs


def _word_com_temp_root() -> Path:
    root = Path(tempfile.gettempdir()) / WORD_COM_TEMP_DIR_NAME
    root.mkdir(parents=True, exist_ok=True)
    return root


def _word_com_temp_paths() -> tuple[Path, Path, Path, Path]:
    root = _word_com_temp_root()
    token = format(time.time_ns() % 0xFFFFFF, "06x")
    return (
        root / f"wfix_{token}.ps1",
        root / f"wfix_{token}.json",
        root / f"wfix_{token}.docx",
        root / f"wfix_{token}.result.txt",
    )


def _command_length(parts: list[str]) -> int:
    return sum(len(part) for part in parts) + max(len(parts) - 1, 0)


def _long_path_compatible_str(path: Path) -> str:
    resolved = str(path.resolve())
    if resolved.startswith("\\\\?\\"):
        return resolved
    if len(resolved) >= 240 and Path(resolved).drive:
        return "\\\\?\\" + resolved
    return resolved


def _build_word_com_body_indent_powershell_script_file() -> str:
    return f"""param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [Parameter(Mandatory = $true)][string]$RecordsPath,
    [Parameter(Mandatory = $true)][string]$ResultPath
)

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
[Console]::InputEncoding = $utf8NoBom

$ErrorActionPreference = 'Stop'
$pointPerCm = {POINTS_PER_CM}
$word = $null
$doc = $null
$processed = 0
$ok = 0
$mismatch = 0
$notFound = 0
$approved = 0
$skippedNot14 = 0
$errors = 0

function Add-Log([string]$msg) {{
    Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value $msg
}}

function Test-CodexStop {{
    $stopPath = $env:CODEX_STOP_PATH
    if ([string]::IsNullOrWhiteSpace($stopPath)) {{ return $false }}
    return [System.IO.File]::Exists($stopPath)
}}

function Normalize-WordParagraphText([string]$text) {{
    if ($null -eq $text) {{ return '' }}
    $cleaned = $text.Replace("`r", " ").Replace([string][char]7, " ").Replace([string][char]11, " ")
    return ([regex]::Replace($cleaned, '\\s+', ' ')).Trim()
}}

function Paragraph-TextMatchesPreview([string]$actual, [string]$preview) {{
    $normalizedActual = Normalize-WordParagraphText $actual
    $normalizedPreview = Normalize-WordParagraphText $preview
    if ([string]::IsNullOrWhiteSpace($normalizedPreview)) {{ return $false }}
    if ($normalizedPreview.EndsWith('...')) {{
        $normalizedPreview = $normalizedPreview.Substring(0, $normalizedPreview.Length - 3).TrimEnd()
    }}
    if ([string]::IsNullOrWhiteSpace($normalizedPreview)) {{ return $false }}
    return $normalizedActual.StartsWith($normalizedPreview)
}}

function Format-OptionalCm($value) {{
    if ($null -eq $value) {{ return 'None' }}
    return ('{{0:F2}}' -f [double]$value)
}}

function Get-SafeProperty($obj, [string]$name) {{
    try {{
        if ($null -eq $obj) {{ return $null }}
        return $obj.$name
    }} catch {{
        return $null
    }}
}}

function Get-ParagraphDiagnostic($paragraph) {{
    $pf = $null
    try {{ $pf = $paragraph.Format }} catch {{}}
    $leftPt = Get-SafeProperty $pf 'LeftIndent'
    $firstLinePt = Get-SafeProperty $pf 'FirstLineIndent'
    $leftCm = $null
    $firstLineCm = $null
    if ($null -ne $leftPt) {{ try {{ $leftCm = [double]$leftPt / $pointPerCm }} catch {{}} }}
    if ($null -ne $firstLinePt) {{ try {{ $firstLineCm = [double]$firstLinePt / $pointPerCm }} catch {{}} }}
    $charLeft = Get-SafeProperty $pf 'CharacterUnitLeftIndent'
    $charFirst = Get-SafeProperty $pf 'CharacterUnitFirstLineIndent'
    $charRight = Get-SafeProperty $pf 'CharacterUnitRightIndent'
    $tabCount = $null
    try {{ $tabCount = $pf.TabStops.Count }} catch {{}}
    $styleName = $null
    try {{ $styleName = $paragraph.Style.NameLocal }} catch {{}}
    $sectionNumber = $null
    $sectionLeftMarginCm = $null
    $sectionRightMarginCm = $null
    try {{
        $section = $paragraph.Range.Sections.Item(1)
        try {{ $sectionNumber = $section.Index }} catch {{}}
        try {{ $sectionLeftMarginCm = [double]$section.PageSetup.LeftMargin / $pointPerCm }} catch {{}}
        try {{ $sectionRightMarginCm = [double]$section.PageSetup.RightMargin / $pointPerCm }} catch {{}}
    }} catch {{}}
    $absoluteTextStartCm = $null
    if ($null -ne $sectionLeftMarginCm -and $null -ne $leftCm) {{
        $absoluteTextStartCm = [double]$sectionLeftMarginCm + [double]$leftCm
    }}
    return [ordered]@{{
        left_cm = Format-OptionalCm $leftCm
        firstline_cm = Format-OptionalCm $firstLineCm
        char_left = $charLeft
        char_first = $charFirst
        char_right = $charRight
        tab_count = $tabCount
        style_name = $styleName
        section_number = $sectionNumber
        section_left_margin_cm = Format-OptionalCm $sectionLeftMarginCm
        section_right_margin_cm = Format-OptionalCm $sectionRightMarginCm
        absolute_text_start_cm = Format-OptionalCm $absoluteTextStartCm
    }}
}}

function Get-RecordDouble($record, [string]$name, [double]$defaultValue) {{
    try {{
        $property = $record.PSObject.Properties[$name]
        if ($null -eq $property) {{ return $defaultValue }}
        $value = $property.Value
        if ($null -eq $value) {{ return $defaultValue }}
        return [double]$value
    }} catch {{
        return $defaultValue
    }}
}}

function Test-MeaningfulFontSize($value) {{
    try {{
        if ($null -eq $value) {{ return $false }}
        $size = [double]$value
        if ($size -le 0 -or $size -ge 200) {{ return $false }}
        if ([math]::Abs($size) -eq 9999999 -or [math]::Abs($size) -eq 999999) {{ return $false }}
        return $true
    }} catch {{
        return $false
    }}
}}

function Get-ParagraphFontSizes($paragraph) {{
    $rangeSize = $null
    $dominantSize = $null
    try {{
        $textRange = $paragraph.Range.Duplicate
        try {{
            if ($textRange.End -gt $textRange.Start) {{ $textRange.End = $textRange.End - 1 }}
        }} catch {{}}
        try {{
            $candidate = $textRange.Font.Size
            if (Test-MeaningfulFontSize $candidate) {{
                $rangeSize = [double]$candidate
                $dominantSize = $rangeSize
                return @($rangeSize, $dominantSize)
            }}
        }} catch {{}}

        $weights = @{{}}
        $order = New-Object System.Collections.Generic.List[double]
        $chars = $textRange.Characters
        for ($ci = 1; $ci -le $chars.Count; $ci++) {{
            try {{
                $ch = $chars.Item($ci)
                $candidate = $ch.Font.Size
                if (-not (Test-MeaningfulFontSize $candidate)) {{
                    try {{ $candidate = $ch.Range.Font.Size }} catch {{}}
                }}
                if (-not (Test-MeaningfulFontSize $candidate)) {{ continue }}
                $size = [double]$candidate
                $key = ('{{0:G}}' -f $size)
                if (-not $weights.ContainsKey($key)) {{
                    $weights[$key] = 0
                    $order.Add($size) | Out-Null
                }}
                $weight = 1
                try {{
                    $text = [string]$ch.Text
                    if (-not [string]::IsNullOrWhiteSpace($text)) {{ $weight = $text.Trim().Length }}
                }} catch {{}}
                $weights[$key] = [int]$weights[$key] + $weight
            }} catch {{}}
        }}
        $bestWeight = -1
        foreach ($size in $order) {{
            $key = ('{{0:G}}' -f $size)
            $weight = [int]$weights[$key]
            if ($weight -gt $bestWeight) {{
                $bestWeight = $weight
                $dominantSize = $size
            }}
        }}
    }} catch {{}}
    return @($rangeSize, $dominantSize)
}}

function Find-MatchingParagraphInWindow($doc, [int]$start, [int]$end, [string]$matchPrefix, [int]$recordIndex, [string]$source) {{
    Add-Log(("WORD_COM_LOCAL_SCAN_BEGIN record={{0}} source={{1}} start={{2}} end={{3}}" -f $recordIndex, $source, $start, $end)) | Out-Null
    for ($j = $start; $j -le $end; $j++) {{
        if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}

        if (($j - $start) % 100 -eq 0) {{
            Add-Log(("WORD_COM_LOCAL_SCAN_PROGRESS record={{0}} source={{1}} current={{2}} start={{3}} end={{4}}" -f $recordIndex, $source, $j, $start, $end)) | Out-Null
        }}

        $candidateText = [string]$doc.Paragraphs.Item($j).Range.Text
        if (Paragraph-TextMatchesPreview $candidateText $matchPrefix) {{
            return [int]$j
        }}
    }}
    return $null
}}

try {{
    Set-Content -LiteralPath $ResultPath -Encoding UTF8 -Value ''
    Add-Log 'WORD_COM_PS_STARTED'
    Add-Log 'WORD_COM_FONT_CHECK_ONLY_STARTED'
    Add-Log ('DOCX_PATH=' + $DocxPath)
    Add-Log ('RECORDS_PATH=' + $RecordsPath)

    if (-not (Test-Path -LiteralPath $DocxPath)) {{
        Add-Log ('WORD_COM_PS_ERROR missing_docx=' + $DocxPath)
        exit 2
    }}

    if (-not (Test-Path -LiteralPath $RecordsPath)) {{
        Add-Log ('WORD_COM_PS_ERROR missing_records=' + $RecordsPath)
        exit 3
    }}

    $recordsRaw = Get-Content -LiteralPath $RecordsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($recordsRaw -is [System.Array]) {{
        $records = @($recordsRaw)
    }} elseif ($null -eq $recordsRaw) {{
        $records = @()
    }} else {{
        $records = @($recordsRaw)
    }}
    Add-Log ('WORD_COM_PS_RECORDS_LOADED count=' + $records.Count)

    $word = New-Object -ComObject Word.Application
    Add-Log 'WORD_COM_PS_WORD_CREATED'
    $word.Visible = $false
    $doc = $word.Documents.Open($DocxPath, $false, $false, $false)
    Add-Log 'WORD_COM_PS_DOC_OPENED'
    Add-Log 'WORD_COM_PS_BEFORE_LOOP'
    $paragraphCount = $doc.Paragraphs.Count
    Add-Log ('WORD_COM_PS_PARAGRAPHS_COUNT count=' + $paragraphCount)
    Add-Log ('WORD_COM_PS_RECORD_LOOP_BEGIN count=' + $records.Count)
    $paragraphIndexOffset = $null
    $lastMatchedWordIndex = $null

    foreach ($record in $records) {{
        try {{
            if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}

            $processed += 1
            $targetIndex = 0
            try {{ $targetIndex = [int]$record.paragraph_index }} catch {{}}
            $preview = [string]$record.text_preview
            $matchPrefix = [string]$record.text_match_prefix
            if ([string]::IsNullOrWhiteSpace($matchPrefix)) {{
                $matchPrefix = $preview
            }}
            $kind = [string]$record.kind
            if ([string]::IsNullOrWhiteSpace($kind)) {{ $kind = 'body' }}
            $expectedLeftCm = Get-RecordDouble $record 'expected_left_cm' 0
            $expectedLeftPoints = Get-RecordDouble $record 'expected_left_points' 0
            $expectedFirstLineCm = Get-RecordDouble $record 'expected_firstline_cm' 0
            $expectedFirstLinePoints = Get-RecordDouble $record 'expected_firstline_points' 0
            $applyOnlyIfWordFontSizeIs14 = $false
            try {{ $applyOnlyIfWordFontSizeIs14 = [bool]$record.apply_only_if_word_font_size_is_14 }} catch {{}}
            Add-Log(("WORD_COM_RECORD_BEGIN i={{0}} paragraph_index={{1}} expected_left_cm={{2}} text={{3}}" -f $processed, $record.paragraph_index, $record.expected_left_cm, $preview))
            Add-Log(("WORD_COM_RECORD_MATCH_PREFIX i={{0}} length={{1}} preview_has_ellipsis={{2}}" -f $processed, $matchPrefix.Length, $preview.EndsWith('...')))

            $matchIndex = $null
            $matchSource = $null
            if ($targetIndex -ge 1 -and $targetIndex -le $paragraphCount) {{
                Add-Log(("WORD_COM_RECORD_TRY_DIRECT i={{0}} word_index={{1}}" -f $processed, $targetIndex))
                $candidateText = [string]$doc.Paragraphs.Item($targetIndex).Range.Text
                if (Paragraph-TextMatchesPreview $candidateText $matchPrefix) {{
                    $matchIndex = $targetIndex
                    $matchSource = 'direct'
                    Add-Log(("WORD_COM_RECORD_DIRECT_MATCHED i={{0}} word_index={{1}}" -f $processed, $matchIndex))
                }}
            }}

            if ($null -eq $matchIndex -and $null -ne $paragraphIndexOffset) {{
                $offsetIndex = $targetIndex + [int]$paragraphIndexOffset
                if ($offsetIndex -ge 1 -and $offsetIndex -le $paragraphCount) {{
                    Add-Log(("WORD_COM_RECORD_TRY_OFFSET i={{0}} word_index={{1}} offset={{2}}" -f $processed, $offsetIndex, $paragraphIndexOffset))
                    $candidateText = [string]$doc.Paragraphs.Item($offsetIndex).Range.Text
                    if (Paragraph-TextMatchesPreview $candidateText $matchPrefix) {{
                        $matchIndex = $offsetIndex
                        $matchSource = 'offset'
                        Add-Log(("WORD_COM_RECORD_OFFSET_MATCHED i={{0}} word_index={{1}} offset={{2}}" -f $processed, $matchIndex, $paragraphIndexOffset))
                    }}
                }}
            }}

            if ($null -eq $matchIndex) {{
                if ($null -ne $paragraphIndexOffset) {{
                    $offsetIndex = $targetIndex + [int]$paragraphIndexOffset
                    $start = [Math]::Max(1, $offsetIndex - 250)
                    $end = [Math]::Min($paragraphCount, $offsetIndex + 250)
                    $localMatch = Find-MatchingParagraphInWindow $doc $start $end $matchPrefix $processed 'offset_window'
                    if ($null -ne $localMatch) {{
                        $matchIndex = [int]$localMatch
                        $matchSource = 'local_offset_window'
                        Add-Log(("WORD_COM_RECORD_LOCAL_MATCHED i={{0}} word_index={{1}} source=offset_window" -f $processed, $matchIndex))
                    }}
                }}
            }}

            if ($null -eq $matchIndex) {{
                $start = [Math]::Max(1, $targetIndex - 250)
                $end = [Math]::Min($paragraphCount, $targetIndex + 250)
                $localMatch = Find-MatchingParagraphInWindow $doc $start $end $matchPrefix $processed 'target_window'
                if ($null -ne $localMatch) {{
                    $matchIndex = [int]$localMatch
                    $matchSource = 'local_target_window'
                    Add-Log(("WORD_COM_RECORD_LOCAL_MATCHED i={{0}} word_index={{1}} source=target_window" -f $processed, $matchIndex))
                }}
            }}

            if ($null -eq $matchIndex -and $null -ne $lastMatchedWordIndex) {{
                $start = [Math]::Max(1, [int]$lastMatchedWordIndex - 100)
                $end = [Math]::Min($paragraphCount, [int]$lastMatchedWordIndex + 500)
                $localMatch = Find-MatchingParagraphInWindow $doc $start $end $matchPrefix $processed 'last_match_window'
                if ($null -ne $localMatch) {{
                    $matchIndex = [int]$localMatch
                    $matchSource = 'local_last_match_window'
                    Add-Log(("WORD_COM_RECORD_LOCAL_MATCHED i={{0}} word_index={{1}} source=last_match_window" -f $processed, $matchIndex))
                }}
            }}

            if ($null -eq $matchIndex) {{
                Add-Log(("WORD_COM_FULL_FALLBACK_SCAN_BEGIN record={{0}} total={{1}}" -f $processed, $paragraphCount))
                for ($j = 1; $j -le $paragraphCount; $j++) {{
                    if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
                    if ($j % 200 -eq 0) {{
                        Add-Log(("WORD_COM_FALLBACK_SCAN_PROGRESS record={{0}} current={{1}} total={{2}}" -f $processed, $j, $paragraphCount))
                    }}

                    $candidateText = [string]$doc.Paragraphs.Item($j).Range.Text
                    if (Paragraph-TextMatchesPreview $candidateText $matchPrefix) {{
                        $matchIndex = $j
                        $matchSource = 'fallback'
                        Add-Log(("WORD_COM_RECORD_FALLBACK_MATCHED i={{0}} word_index={{1}}" -f $processed, $matchIndex))
                        break
                    }}
                }}
            }}

            if ($null -eq $matchIndex) {{
                $notFound += 1
                Add-Log(("WORD_COM_BODY_INDENT_FIX: i={{0}} paragraph_index={{1}} expected_left_cm={{2}} before_left_cm=None after_left_cm=None status=not_found" -f $processed, $record.paragraph_index, $record.expected_left_cm))
                continue
            }}

            if ($matchSource -ne 'direct' -and $matchSource -ne 'offset') {{
                $newOffset = [int]$matchIndex - $targetIndex
                if ($null -eq $paragraphIndexOffset) {{
                    $paragraphIndexOffset = $newOffset
                    Add-Log(("WORD_COM_PARAGRAPH_INDEX_OFFSET_LEARNED offset={{0}} from_record={{1}} xml_index={{2}} word_index={{3}}" -f $paragraphIndexOffset, $processed, $targetIndex, $matchIndex))
                }} elseif ([int]$paragraphIndexOffset -ne [int]$newOffset) {{
                    Add-Log(("WORD_COM_PARAGRAPH_INDEX_OFFSET_CHANGED old={{0}} new={{1}} record={{2}} xml_index={{3}} word_index={{4}}" -f $paragraphIndexOffset, $newOffset, $processed, $targetIndex, $matchIndex))
                    $paragraphIndexOffset = $newOffset
                }}
            }}
            $lastMatchedWordIndex = [int]$matchIndex
            Add-Log(("WORD_COM_RECORD_MATCHED i={{0}} word_index={{1}} source={{2}}" -f $processed, $matchIndex, $matchSource))
            Add-Log(("WORD_COM_RECORD_BEFORE_GET_PARAGRAPH i={{0}} word_index={{1}}" -f $processed, $matchIndex))
            $paragraph = $doc.Paragraphs.Item([int]$matchIndex)
            $diag = Get-ParagraphDiagnostic $paragraph
            Add-Log(("WORD_COM_RECORD_AFTER_GET_PARAGRAPH i={{0}}" -f $processed))
            $wordRangeFontSize = $null
            $wordDominantFontSize = $null
            $wordFontCheckPass = $true
            if ($applyOnlyIfWordFontSizeIs14) {{
                Add-Log(("WORD_COM_RECORD_BEFORE_FONT_CHECK i={{0}}" -f $processed))
                $fontSizes = Get-ParagraphFontSizes $paragraph
                if ($fontSizes.Count -ge 1) {{ $wordRangeFontSize = $fontSizes[0] }}
                if ($fontSizes.Count -ge 2) {{ $wordDominantFontSize = $fontSizes[1] }}
                Add-Log(("WORD_COM_RECORD_AFTER_FONT_CHECK i={{0}} range_font_size={{1}} dominant_font_size={{2}}" -f $processed, $wordRangeFontSize, $wordDominantFontSize))
                $wordFontCheckPass = ($null -ne $wordDominantFontSize -and [math]::Abs([double]$wordDominantFontSize - 14.0) -le 0.01)
            }}

            if (-not $wordFontCheckPass) {{
                $skippedNot14 += 1
                Add-Log(("WORD_COM_INDENT_VERIFY: paragraph_index={{0}}; matched_paragraph_index={{1}}; text_preview={{2}}; kind={{3}}; level={{4}}; xml_font_size={{5}}; xml_font_size_source={{6}}; word_range_font_size={{7}}; word_dominant_font_size={{8}}; word_font_check_pass=False; decision=skipped_word_font_not_14; status=skipped_word_font_not_14" -f $record.paragraph_index, $matchIndex, $preview, $kind, $record.level, $record.xml_font_size, $record.xml_font_size_source, $wordRangeFontSize, $wordDominantFontSize))
                Add-Log(("WORD_COM_BODY_INDENT_FIX: i={{0}} paragraph_index={{1}} matched_paragraph_index={{2}} status=skipped_word_font_not_14" -f $processed, $record.paragraph_index, $matchIndex))
                continue
            }}

            $status = 'approved'
            $approved += 1
            $ok += 1

            Add-Log(("WORD_COM_FONT_CHECK_APPROVED: record_index={{0}}; paragraph_index={{1}}; matched_paragraph_index={{2}}; word_dominant_font_size={{3}}" -f $processed, $record.paragraph_index, $matchIndex, $wordDominantFontSize))
            $approvedRecord = [ordered]@{{
                record_index = $processed
                paragraph_index = $record.paragraph_index
                matched_paragraph_index = $matchIndex
                word_dominant_font_size = $wordDominantFontSize
                expected_left_twips = $record.expected_left_twips
                expected_first_line_twips = $record.expected_first_line_twips
                text_match_prefix = $matchPrefix
            }}
            Add-Log(("WORD_COM_APPROVED_RECORD_JSON " + ($approvedRecord | ConvertTo-Json -Compress)))
            Add-Log(("WORD_COM_INDENT_VERIFY: paragraph_index={{0}}; matched_paragraph_index={{1}}; text_preview={{2}}; kind={{3}}; level={{4}}; expected_number_start_cm={{5}}; expected_hanging_cm={{6}}; expected_heading_left_cm={{7}}; expected_body_left_cm={{8}}; expected_first_line_twips={{9}}; xml_font_size={{10}}; xml_font_size_source={{11}}; word_range_font_size={{12}}; word_dominant_font_size={{13}}; word_font_check_pass={{14}}; decision=approved_for_xml_body_indent; word_opened_left_cm={{15}}; word_opened_firstline_cm={{16}}; final_left_cm=not_read; final_firstline_cm=not_read; word_com_LeftIndent_cm={{17}}; word_com_FirstLineIndent_cm={{18}}; word_com_CharacterUnitLeftIndent={{19}}; word_com_CharacterUnitFirstLineIndent={{20}}; word_com_CharacterUnitRightIndent={{21}}; word_com_TabStops_Count={{22}}; word_com_Style_NameLocal={{23}}; word_com_Section_Number={{24}}; section_index={{25}}; section_left_margin_cm={{26}}; section_right_margin_cm={{27}}; paragraph_left_indent_cm={{28}}; absolute_text_start_cm={{29}}; second_fix=python_xml; status={{30}}" -f $record.paragraph_index, $matchIndex, $preview, $kind, $record.level, $record.expected_number_start_cm, $record.expected_hanging_cm, $record.expected_heading_left_cm, $record.expected_body_left_cm, $record.expected_first_line_twips, $record.xml_font_size, $record.xml_font_size_source, $wordRangeFontSize, $wordDominantFontSize, $wordFontCheckPass, $diag.left_cm, $diag.firstline_cm, $diag.left_cm, $diag.firstline_cm, $diag.char_left, $diag.char_first, $diag.char_right, $diag.tab_count, $diag.style_name, $diag.section_number, $diag.section_number, $diag.section_left_margin_cm, $diag.section_right_margin_cm, $diag.left_cm, $diag.absolute_text_start_cm, $status))
            Add-Log(("WORD_COM_BODY_INDENT_FIX: i={{0}} paragraph_index={{1}} expected_left_cm={{2}} before_left_cm=not_read after_left_cm=not_read status={{3}}" -f $processed, $record.paragraph_index, $record.expected_left_cm, $status))
        }} catch {{
            $errors += 1
            Add-Log(("WORD_COM_RECORD_EXCEPTION i={{0}} paragraph_index={{1}} type={{2}} message={{3}}" -f $processed, $record.paragraph_index, $_.Exception.GetType().FullName, $_.Exception.Message))
            Add-Log(("WORD_COM_RECORD_STACK i={{0}} stack={{1}}" -f $processed, $_.ScriptStackTrace))
            continue
        }}
    }}

    Add-Log(("WORD_COM_BODY_INDENT_FIX_SUMMARY processed={{0}} ok={{1}} mismatch={{2}} not_found={{3}} errors={{4}}" -f $processed, $ok, $mismatch, $notFound, $errors))
    Add-Log(("WORD_COM_FONT_CHECK_SUMMARY processed={{0}} approved={{1}} skipped_not_14={{2}} not_found={{3}} errors={{4}}" -f $processed, $approved, $skippedNot14, $notFound, $errors))
    Add-Log 'WORD_COM_PS_DONE'
    exit 0
}} catch {{
    try {{
        Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value ("WORD_COM_PS_EXCEPTION type=" + $_.Exception.GetType().FullName + " message=" + $_.Exception.Message)
        Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value ("WORD_COM_PS_STACK " + $_.ScriptStackTrace)
    }} catch {{}}
    Write-Output ("WORD_COM_PS_EXCEPTION type=" + $_.Exception.GetType().FullName + " message=" + $_.Exception.Message)
    exit 1
}} finally {{
    try {{
        Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value 'WORD_COM_PS_FINALLY_BEGIN'
    }} catch {{}}

    if ($doc -ne $null) {{
        try {{
            $doc.Close($false) | Out-Null
            Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value 'WORD_COM_PS_DOC_CLOSED'
        }} catch {{
            Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value ("WORD_COM_PS_DOC_CLOSE_FAILED " + $_.Exception.Message)
        }}
    }}

    if ($word -ne $null) {{
        try {{
            $word.Quit() | Out-Null
            Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value 'WORD_COM_PS_WORD_QUIT'
        }} catch {{
            Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value ("WORD_COM_PS_WORD_QUIT_FAILED " + $_.Exception.Message)
        }}
    }}
}}
"""


def verify_and_fix_body_indents_with_word_com(
    output_docx: Path,
    body_indent_records: list[dict[str, object]],
    stop: StopController | None = None,
) -> list[str]:
    if not body_indent_records:
        return ["WORD_COM_BODY_INDENT_FIX_SKIPPED reason=no_records"]

    script_path, records_path, work_docx_path, result_path = _word_com_temp_paths()
    command_parts = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path),
        "-DocxPath",
        str(work_docx_path),
        "-RecordsPath",
        str(records_path),
        "-ResultPath",
        str(result_path),
    ]
    logs = [
        "WORD_COM_BODY_INDENT_FIX_STARTED",
        f"body_indent_records_count={len(body_indent_records)}",
        f"WORD_COM_POWERSHELL_SCRIPT_PATH={script_path}",
        f"WORD_COM_RECORDS_JSON_PATH={records_path}",
        f"WORD_COM_DOCX_WORK_PATH={work_docx_path}",
        f"WORD_COM_RESULT_LOG_PATH={result_path}",
        f"command_length={_command_length(command_parts)}",
        f"records_count={len(body_indent_records)}",
    ]

    try:
        script_path.write_text(_build_word_com_body_indent_powershell_script_file(), encoding="utf-8")
        records_path.write_text(
            json.dumps(body_indent_records, ensure_ascii=False),
            encoding="utf-8",
        )
        shutil.copy2(_long_path_compatible_str(output_docx), _long_path_compatible_str(work_docx_path))
        result_path.write_text("", encoding="utf-8")
        logs.append(f"WORD_COM_POWERSHELL_COMMAND={' '.join(command_parts)}")
        logs.append(f"WORD_COM_POWERSHELL_SCRIPT_EXISTS={script_path.exists()}")
        logs.append(f"WORD_COM_RECORDS_JSON_EXISTS={records_path.exists()}")
        logs.append(f"WORD_COM_DOCX_WORK_EXISTS={work_docx_path.exists()}")

        completed = run_powershell_file(
            script_path,
            arguments=[
                "-DocxPath",
                str(work_docx_path),
                "-RecordsPath",
                str(records_path),
                "-ResultPath",
                str(result_path),
            ],
            stop=stop,
            timeout=WORD_COM_TIMEOUT_SECONDS,
        )

        logs.append(f"WORD_COM_POWERSHELL_RETURN_CODE={completed.returncode}")
        logs.append(f"WORD_COM_POWERSHELL_STDOUT_LEN={len(completed.stdout)}")
        logs.append(f"WORD_COM_POWERSHELL_STDERR_LEN={len(completed.stderr)}")
        logs.append(f"WORD_COM_POWERSHELL_STDERR={completed.stderr.strip()}")
        logs.append(f"WORD_COM_RESULT_LOG_EXISTS={result_path.exists()}")
        result_size = result_path.stat().st_size if result_path.exists() else 0
        logs.append(f"WORD_COM_RESULT_LOG_SIZE={result_size}")

        script_logs = _collect_word_com_powershell_logs(completed.stdout, result_path)
        if script_logs:
            logs.extend(script_logs)
            has_script_failure = any(
                log.startswith("WORD_COM_PS_EXCEPTION")
                or log.startswith("WORD_COM_PS_ERROR")
                or log.startswith("WORD_COM_BODY_INDENT_FIX_SKIPPED")
                for log in script_logs
            )
            has_script_done = any(log.startswith("WORD_COM_PS_DONE") for log in script_logs)
            if completed.returncode != 0 or not has_script_done:
                logs.append("WORD_COM_BODY_INDENT_FIX_FAILED_AFTER_PARTIAL_LOGS")
                has_script_failure = True

            if not has_script_failure and completed.returncode == 0:
                approved_records = _parse_word_com_approved_records(script_logs)
                logs.append(f"WORD_COM_FONT_CHECK_APPROVED_COUNT={len(approved_records)}")
                logs.extend(apply_word_com_approved_body_indents_to_docx_xml(output_docx, approved_records))
            elif not any(log.startswith("WORD_COM_BODY_INDENT_FIX_SKIPPED") for log in logs):
                has_script_exception = any(
                    log.startswith("WORD_COM_PS_EXCEPTION") or log.startswith("WORD_COM_PS_ERROR")
                    for log in script_logs
                )
                if not has_script_done and not has_script_exception and not completed.stderr.strip():
                    last_partial_log = script_logs[-1] if script_logs else "None"
                    logs.append("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_interrupted_or_timeout")
                    logs.append(f"timeout_seconds={WORD_COM_TIMEOUT_SECONDS}")
                    logs.append(f"last_partial_log={last_partial_log}")
                else:
                    logs.append("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_script_failed")
            return logs

        stderr = completed.stderr.strip()
        reason = stderr if stderr else "empty_output"
        logs.append(f"WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_no_logs:{reason}")
        return logs
    except ProcessStopped:
        raise
    except Exception as exc:
        logs.append(f"WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_runner_failed:{type(exc).__name__}:{exc}")
        logs.append(f"exception_repr={exc!r}")
        return logs
    finally:
        for temp_path in (script_path, records_path, work_docx_path, result_path):
            try:
                temp_path.unlink()
            except FileNotFoundError:
                pass
            except Exception:
                pass


def _filter_word_com_body_indent_records(
    body_indent_records: list[dict[str, object]],
) -> list[dict[str, object]]:
    filtered: list[dict[str, object]] = []
    for record in body_indent_records:
        if not bool(record.get("apply_only_if_word_font_size_is_14")):
            continue
        try:
            xml_font_size = float(record.get("xml_font_size"))
        except (TypeError, ValueError):
            xml_font_size = None
        if xml_font_size is not None and xml_font_size <= 11.0:
            continue
        filtered.append(record)
    return filtered


def filter_word_com_body_indent_records(
    body_indent_records: list[dict[str, object]],
) -> list[dict[str, object]]:
    return _filter_word_com_body_indent_records(body_indent_records)
