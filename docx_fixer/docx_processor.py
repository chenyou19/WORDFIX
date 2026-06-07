from __future__ import annotations

import json
import re
import shutil
import tempfile
import time
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from .constants import NS, TEMPLATE_OUTLINE_INDENTS
from .exceptions import ProcessStopped
from .models import ProcessOptions, ProcessSummary
from .numbering import (
    apply_numbering_outline_format,
    apply_styles_outline_format_to_root,
    build_numbering_format_lookup,
    build_numbering_level_lookup,
    build_style_numbering_lookup,
    has_auto_numbering,
    paragraph_style_id,
)
from .outline import (
    detect_manual_numbering_prefix,
    fix_outline_paragraphs,
    force_all_paragraphs_to_body_outline_level,
    get_auto_number_identity,
    remove_all_outline_levels_from_any_root,
)
from .path_utils import is_same_file_path
from .process_runner import run_powershell_file, run_powershell_script
from .stop_controller import StopController
from .style_resolver import build_style_font_size_lookup
from .table_format import process_table, table_cell_count, table_column_count
from .xml_utils import paragraph_text, qn, remove_character_indent_attrs_from_root

POINTS_PER_CM = 28.3464567
WORD_COM_TIMEOUT_SECONDS = 120
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


def should_process_part(name: str) -> bool:
    if name == "word/document.xml":
        return True
    if name.startswith("word/header") and name.endswith(".xml"):
        return True
    if name.startswith("word/footer") and name.endswith(".xml"):
        return True
    if name in {"word/footnotes.xml", "word/endnotes.xml"}:
        return True
    return False


def should_remove_outline_part(name: str) -> bool:
    if should_process_part(name):
        return True
    if name in {"word/styles.xml", "word/numbering.xml"}:
        return True
    return False


def should_force_body_outline_part(name: str) -> bool:
    return should_process_part(name)


def should_sanitize_indent_unit_part(name: str) -> bool:
    if name == "word/document.xml":
        return True
    if name == "word/styles.xml":
        return True
    if name == "word/numbering.xml":
        return True
    if name.startswith("word/header") and name.endswith(".xml"):
        return True
    if name.startswith("word/footer") and name.endswith(".xml"):
        return True
    if name in {"word/footnotes.xml", "word/endnotes.xml"}:
        return True
    return False


def should_fix_paragraph_part(name: str) -> bool:
    """
    只對 `word/document.xml` 套用段落縮排與大綱修正。

    header/footer/footnotes/endnotes 不執行本文段落階層主邏輯，
    避免誤改非本文區域。

    表格、顏色與其他 XML 處理仍由各自函式判斷是否套用。
    """
    return name == "word/document.xml"


def _normalize_table_log_text(text: str, limit: int = 100) -> str:
    normalized = " ".join((text or "").replace("\t", " ").replace("\r", " ").replace("\n", " ").split())
    if not normalized:
        return "(empty)"
    if len(normalized) <= limit:
        return normalized
    return normalized[: limit - 3].rstrip() + "..."


def find_previous_paragraph_text_for_table(root, tbl) -> str:
    del root  # The table already belongs to the current XML part tree.
    paragraphs = tbl.xpath("preceding::w:p[not(ancestor::w:tbl)]", namespaces=NS)
    for p in reversed(paragraphs):
        text = _normalize_table_log_text(paragraph_text(p))
        if text != "(empty)":
            return text
    return "(empty)"


_TRADITIONAL_LEGAL_CHAPTER_NUMBERS = {
    1: "壹",
    2: "貳",
    3: "參",
    4: "肆",
    5: "伍",
    6: "陸",
    7: "柒",
    8: "捌",
    9: "玖",
    10: "拾",
}

_TRADITIONAL_COUNTING_CHAPTER_NUMBERS = {
    1: "一",
    2: "二",
    3: "三",
    4: "四",
    5: "五",
    6: "六",
    7: "七",
    8: "八",
    9: "九",
    10: "十",
}


def _effective_paragraph_numbering_identity(
    p,
    style_numbering_lookup,
) -> tuple[str | None, int | None]:
    num_id, ilvl = get_auto_number_identity(p)
    if num_id is not None and ilvl is not None:
        return num_id, ilvl

    style_id = paragraph_style_id(p)
    if style_id and style_numbering_lookup:
        return style_numbering_lookup.get(style_id, (None, None))

    return num_id, ilvl


def _outline_level_from_identity(
    num_id,
    ilvl,
    numbering_level_lookup,
) -> int | None:
    if num_id is None or ilvl is None:
        return None

    level = numbering_level_lookup.get((num_id, ilvl))
    if level is not None:
        return level
    if 0 <= ilvl <= 8:
        return ilvl
    return None


def _chapter_number_token_from_format(num_fmt: str | None, ordinal: int) -> str | None:
    if ordinal <= 0:
        return None

    fmt = (num_fmt or "").strip()
    if fmt in {"ideographLegalTraditional", "chineseLegalSimplified"}:
        return _TRADITIONAL_LEGAL_CHAPTER_NUMBERS.get(ordinal)
    if fmt in {"taiwaneseCountingThousand", "ideographTraditional", "chineseCounting"}:
        return _TRADITIONAL_COUNTING_CHAPTER_NUMBERS.get(ordinal)
    return None


def _count_same_stream_first_level_headings_before_paragraph(
    p,
    *,
    num_id,
    ilvl,
    numbering_level_lookup,
    style_numbering_lookup,
) -> int:
    count = 0
    paragraphs = p.xpath("preceding::w:p[not(ancestor::w:tbl)]", namespaces=NS)
    for candidate in [*paragraphs, p]:
        candidate_num_id, candidate_ilvl = _effective_paragraph_numbering_identity(
            candidate,
            style_numbering_lookup,
        )
        if (candidate_num_id, candidate_ilvl) != (num_id, ilvl):
            continue

        candidate_level = _outline_level_from_identity(
            candidate_num_id,
            candidate_ilvl,
            numbering_level_lookup,
        )
        if candidate_level == 0:
            count += 1

    return count


def _first_level_heading_prefix_for_paragraph(
    p,
    *,
    numbering_level_lookup,
    numbering_format_lookup,
    style_numbering_lookup,
) -> str | None:
    text = paragraph_text(p).strip()
    manual = detect_manual_numbering_prefix(text)
    if manual is not None:
        level, prefix = manual
        if level == 0:
            return prefix
        return None

    num_id, ilvl = _effective_paragraph_numbering_identity(p, style_numbering_lookup)
    level = _outline_level_from_identity(num_id, ilvl, numbering_level_lookup)
    if level != 0 or num_id is None or ilvl is None:
        return None

    level_format = numbering_format_lookup.get((num_id, ilvl), {})
    ordinal = _count_same_stream_first_level_headings_before_paragraph(
        p,
        num_id=num_id,
        ilvl=ilvl,
        numbering_level_lookup=numbering_level_lookup,
        style_numbering_lookup=style_numbering_lookup,
    )
    token = _chapter_number_token_from_format(level_format.get("numFmt"), ordinal)
    lvl_text = level_format.get("lvlText")
    if token is None or lvl_text is None or "%1" not in lvl_text:
        return None

    return lvl_text.replace("%1", token)


def is_table_under_chapter_three(
    tbl,
    numbering_level_lookup,
    numbering_format_lookup,
    style_numbering_lookup,
) -> bool:
    heading = find_table_first_level_heading(
        tbl,
        numbering_level_lookup,
        numbering_format_lookup,
        style_numbering_lookup,
    )
    return heading == "參、"


def find_table_first_level_heading(
    tbl,
    numbering_level_lookup,
    numbering_format_lookup,
    style_numbering_lookup,
) -> str | None:
    paragraphs = tbl.xpath("preceding::w:p[not(ancestor::w:tbl)]", namespaces=NS)
    for p in reversed(paragraphs):
        prefix = _first_level_heading_prefix_for_paragraph(
            p,
            numbering_level_lookup=numbering_level_lookup,
            numbering_format_lookup=numbering_format_lookup,
            style_numbering_lookup=style_numbering_lookup,
        )
        if prefix is not None:
            return prefix
    return None


def build_table_log_record(
    *,
    part_name: str,
    table_index: int,
    global_table_index: int,
    table_name: str,
    first_level_heading: str,
    cell_count: int,
    column_count: int,
    table_type: str,
    action: str,
    reason: str,
    special_layout_used: bool,
    layout_fixed: bool,
    color_fixed: bool,
    changed_to_gray: int,
    cleared_colors: int,
    shading_debug: list[str] | None = None,
) -> dict[str, object]:
    return {
        "part_name": part_name,
        "table_index": table_index,
        "global_table_index": global_table_index,
        "table_name": table_name,
        "first_level_heading": first_level_heading,
        "cell_count": cell_count,
        "column_count": column_count,
        "table_type": table_type,
        "action": action,
        "reason": reason,
        "special_layout_used": special_layout_used,
        "layout_fixed": layout_fixed,
        "color_fixed": color_fixed,
        "changed_to_gray": changed_to_gray,
        "cleared_colors": cleared_colors,
        "shading_debug": list(shading_debug or []),
    }


def _parse_twips_attr(element, attr_name: str) -> int | None:
    if element is None:
        return None
    value = element.get(qn(attr_name))
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _paragraph_outline_level_for_table_anchor(
    p,
    numbering_level_lookup,
    style_numbering_lookup,
) -> int | None:
    num_id, ilvl = _effective_paragraph_numbering_identity(p, style_numbering_lookup)
    level = _outline_level_from_identity(num_id, ilvl, numbering_level_lookup)
    if level is not None:
        return level

    manual = detect_manual_numbering_prefix(paragraph_text(p).strip())
    if manual is not None:
        return manual[0]

    return None


def _paragraph_text_start_twips(
    p,
    numbering_level_lookup,
    style_numbering_lookup,
) -> int | None:
    outline_level = _paragraph_outline_level_for_table_anchor(
        p,
        numbering_level_lookup,
        style_numbering_lookup,
    )
    if outline_level is not None:
        spec = TEMPLATE_OUTLINE_INDENTS.get(outline_level)
        if spec is not None:
            try:
                return int(spec.get("body_left", spec["left"]))
            except (KeyError, TypeError, ValueError):
                return None

    ind = p.find("./w:pPr/w:ind", NS)
    if ind is None:
        return None

    base = _parse_twips_attr(ind, "start")
    if base is None:
        base = _parse_twips_attr(ind, "left")
    if base is None:
        base = 0

    first_line = _parse_twips_attr(ind, "firstLine")
    hanging = _parse_twips_attr(ind, "hanging")

    if first_line is not None and first_line > 0:
        return base + first_line
    if hanging is not None:
        return base
    if first_line is not None and first_line < 0:
        return base
    return base


def _find_previous_effective_paragraph(tbl, style_numbering_lookup):
    paragraphs = tbl.xpath("preceding::w:p[not(ancestor::w:tbl)]", namespaces=NS)
    for p in reversed(paragraphs):
        text = paragraph_text(p).strip()
        if text:
            return p

        if has_auto_numbering(p):
            return p

        style_id = paragraph_style_id(p)
        if style_id and style_numbering_lookup and style_id in style_numbering_lookup:
            return p

    return None


def _find_table_section_properties(tbl):
    preceding_sect_pr = tbl.xpath("preceding::w:sectPr", namespaces=NS)
    if preceding_sect_pr:
        return preceding_sect_pr[-1]

    body_sect_pr = tbl.xpath("ancestor::w:body/w:sectPr", namespaces=NS)
    if body_sect_pr:
        return body_sect_pr[0]

    return None


def _page_text_width_twips(sect_pr) -> int | None:
    if sect_pr is None:
        return None

    pg_sz = sect_pr.find("w:pgSz", NS)
    pg_mar = sect_pr.find("w:pgMar", NS)
    if pg_sz is None or pg_mar is None:
        return None

    page_width = _parse_twips_attr(pg_sz, "w")
    left_margin = _parse_twips_attr(pg_mar, "left")
    right_margin = _parse_twips_attr(pg_mar, "right")
    if page_width is None or left_margin is None or right_margin is None:
        return None

    available_width = page_width - left_margin - right_margin
    if available_width <= 0:
        return None
    return available_width


def _resolve_special_table_geometry(
    tbl,
    numbering_level_lookup,
    style_numbering_lookup,
) -> tuple[int, int] | None:
    anchor_paragraph = _find_previous_effective_paragraph(tbl, style_numbering_lookup)
    if anchor_paragraph is None:
        return None

    left_indent_twips = _paragraph_text_start_twips(
        anchor_paragraph,
        numbering_level_lookup,
        style_numbering_lookup,
    )
    if left_indent_twips is None or left_indent_twips < 0:
        return None

    text_width_twips = _page_text_width_twips(_find_table_section_properties(tbl))
    if text_width_twips is None:
        return None

    width_twips = text_width_twips - left_indent_twips
    if width_twips <= 0:
        return None

    return left_indent_twips, width_twips


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
    if not normalized_preview:
        return False
    return normalized_actual.startswith(normalized_preview)


def find_word_paragraph_index_for_record(paragraph_texts: list[str], record: dict[str, object]) -> int | None:
    target_index = int(record.get("paragraph_index") or 0)
    preview = str(record.get("text_preview") or "")

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

            pf.LeftIndent = expected_left_points
            pf.FirstLineIndent = expected_firstline_points

            actual_left_pt = _safe_com_attr(pf, "LeftIndent")
            actual_left_cm = _points_to_cm(actual_left_pt)
            actual_first_line_pt = _safe_com_attr(pf, "FirstLineIndent")
            actual_char_left = _safe_com_attr(pf, "CharacterUnitLeftIndent")
            actual_char_first = _safe_com_attr(pf, "CharacterUnitFirstLineIndent")
            first_line_indent_cm = _points_to_cm(actual_first_line_pt)
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

$ErrorActionPreference = 'Stop'
$pointPerCm = {POINTS_PER_CM}
$word = $null
$doc = $null
$processed = 0
$ok = 0
$mismatch = 0
$notFound = 0
$errors = 0

function Add-Log([string]$msg) {{
    Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value $msg
    Write-Output $msg
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
    return $normalizedActual.StartsWith($normalizedPreview)
}}

function Format-OptionalCm($value) {{
    if ($null -eq $value) {{ return 'None' }}
    return ('{{0:F2}}' -f [double]$value)
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

try {{
    Set-Content -LiteralPath $ResultPath -Encoding UTF8 -Value ''
    Add-Log 'WORD_COM_PS_STARTED'
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
    Add-Log ('WORD_COM_PS_PARAGRAPHS_COUNT count=' + $doc.Paragraphs.Count)

    $paragraphs = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {{
        if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
        $paragraphs.Add([string]$doc.Paragraphs.Item($i).Range.Text)
    }}

    foreach ($record in $records) {{
        try {{
            if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}

            $processed += 1
            $targetIndex = 0
            try {{ $targetIndex = [int]$record.paragraph_index }} catch {{}}
            $preview = [string]$record.text_preview
            $kind = [string]$record.kind
            if ([string]::IsNullOrWhiteSpace($kind)) {{ $kind = 'body' }}
            $expectedLeftCm = Get-RecordDouble $record 'expected_left_cm' 0
            $expectedLeftPoints = Get-RecordDouble $record 'expected_left_points' 0
            $expectedFirstLineCm = Get-RecordDouble $record 'expected_firstline_cm' 0
            $expectedFirstLinePoints = Get-RecordDouble $record 'expected_firstline_points' 0
            $applyOnlyIfWordFontSizeIs14 = $false
            try {{ $applyOnlyIfWordFontSizeIs14 = [bool]$record.apply_only_if_word_font_size_is_14 }} catch {{}}
            Add-Log(("WORD_COM_RECORD_BEGIN i={{0}} paragraph_index={{1}} expected_left_cm={{2}} text={{3}}" -f $processed, $record.paragraph_index, $record.expected_left_cm, $preview))

            $matchIndex = $null
            if ($targetIndex -ge 1 -and $targetIndex -le $paragraphs.Count) {{
                if (Paragraph-TextMatchesPreview $paragraphs[$targetIndex - 1] $preview) {{
                    $matchIndex = $targetIndex
                }}
            }}

            if ($null -eq $matchIndex) {{
                for ($j = 0; $j -lt $paragraphs.Count; $j++) {{
                    if (Paragraph-TextMatchesPreview $paragraphs[$j] $preview) {{
                        $matchIndex = $j + 1
                        break
                    }}
                }}
            }}

            if ($null -eq $matchIndex) {{
                $notFound += 1
                Add-Log(("WORD_COM_BODY_INDENT_FIX: i={{0}} paragraph_index={{1}} expected_left_cm={{2}} before_left_cm=None after_left_cm=None status=not_found" -f $processed, $record.paragraph_index, $record.expected_left_cm))
                continue
            }}

            Add-Log(("WORD_COM_RECORD_MATCHED i={{0}} word_index={{1}}" -f $processed, $matchIndex))
            $paragraph = $doc.Paragraphs.Item([int]$matchIndex)
            $pf = $paragraph.Format
            $beforeLeftPt = $null
            $beforeFirstLinePt = $null
            $actualLeftPt = $null
            $actualFirstLinePt = $null
            try {{ $beforeLeftPt = [double]$pf.LeftIndent }} catch {{}}
            try {{ $beforeFirstLinePt = [double]$pf.FirstLineIndent }} catch {{}}
            $wordRangeFontSize = $null
            $wordDominantFontSize = $null
            $wordFontCheckPass = $true
            if ($applyOnlyIfWordFontSizeIs14) {{
                $fontSizes = Get-ParagraphFontSizes $paragraph
                if ($fontSizes.Count -ge 1) {{ $wordRangeFontSize = $fontSizes[0] }}
                if ($fontSizes.Count -ge 2) {{ $wordDominantFontSize = $fontSizes[1] }}
                $wordFontCheckPass = ($null -ne $wordDominantFontSize -and [math]::Abs([double]$wordDominantFontSize - 14.0) -le 0.01)
            }}

            if (-not $wordFontCheckPass) {{
                Add-Log(("WORD_COM_INDENT_VERIFY: paragraph_index={{0}}; matched_paragraph_index={{1}}; text_preview={{2}}; kind={{3}}; level={{4}}; xml_font_size={{5}}; xml_font_size_source={{6}}; word_range_font_size={{7}}; word_dominant_font_size={{8}}; word_font_check_pass=False; decision=skipped_word_font_not_14; status=skipped_word_font_not_14" -f $record.paragraph_index, $matchIndex, $preview, $kind, $record.level, $record.xml_font_size, $record.xml_font_size_source, $wordRangeFontSize, $wordDominantFontSize))
                Add-Log(("WORD_COM_BODY_INDENT_FIX: i={{0}} paragraph_index={{1}} matched_paragraph_index={{2}} status=skipped_word_font_not_14" -f $processed, $record.paragraph_index, $matchIndex))
                continue
            }}

            try {{ $pf.CharacterUnitLeftIndent = 0 }} catch {{
                Add-Log(("WORD_COM_RECORD_CHAR_LEFT_CLEAR_FAILED i={{0}} reason={{1}}" -f $processed, $_.Exception.Message))
            }}
            try {{ $pf.CharacterUnitFirstLineIndent = 0 }} catch {{
                Add-Log(("WORD_COM_RECORD_CHAR_FIRST_CLEAR_FAILED i={{0}} reason={{1}}" -f $processed, $_.Exception.Message))
            }}

            $pf.LeftIndent = $expectedLeftPoints
            $pf.FirstLineIndent = $expectedFirstLinePoints

            try {{ $actualLeftPt = [double]$pf.LeftIndent }} catch {{}}
            try {{ $actualFirstLinePt = [double]$pf.FirstLineIndent }} catch {{}}
            $beforeLeftCm = if ($null -eq $beforeLeftPt) {{ $null }} else {{ [double]$beforeLeftPt / $pointPerCm }}
            $beforeFirstLineCm = if ($null -eq $beforeFirstLinePt) {{ $null }} else {{ [double]$beforeFirstLinePt / $pointPerCm }}
            $afterLeftCm = if ($null -eq $actualLeftPt) {{ $null }} else {{ [double]$actualLeftPt / $pointPerCm }}
            $afterFirstLineCm = if ($null -eq $actualFirstLinePt) {{ $null }} else {{ [double]$actualFirstLinePt / $pointPerCm }}

            if ($null -ne $afterLeftCm -and [math]::Abs($afterLeftCm - $expectedLeftCm) -le 0.02 -and $null -ne $afterFirstLineCm -and [math]::Abs($afterFirstLineCm - $expectedFirstLineCm) -le 0.02) {{
                $status = 'ok'
                $ok += 1
            }} else {{
                $status = 'mismatch'
                $mismatch += 1
            }}

            Add-Log(("WORD_COM_INDENT_VERIFY: paragraph_index={{0}}; matched_paragraph_index={{1}}; text_preview={{2}}; kind={{3}}; level={{4}}; expected_number_start_cm={{5}}; expected_hanging_cm={{6}}; expected_heading_left_cm={{7}}; expected_body_left_cm={{8}}; expected_first_line_twips={{9}}; xml_font_size={{10}}; xml_font_size_source={{11}}; word_range_font_size={{12}}; word_dominant_font_size={{13}}; word_font_check_pass={{14}}; decision=apply_body_indent; xml_written_left_cm={{15}}; xml_written_hanging_cm={{16}}; word_opened_left_cm={{17}}; word_opened_firstline_cm={{18}}; final_left_cm={{19}}; final_firstline_cm={{20}}; second_fix=yes; status={{21}}" -f $record.paragraph_index, $matchIndex, $preview, $kind, $record.level, $record.expected_number_start_cm, $record.expected_hanging_cm, $record.expected_heading_left_cm, $record.expected_body_left_cm, $record.expected_first_line_twips, $record.xml_font_size, $record.xml_font_size_source, $wordRangeFontSize, $wordDominantFontSize, $wordFontCheckPass, $record.xml_written_left_cm, $record.xml_written_hanging_cm, (Format-OptionalCm $beforeLeftCm), (Format-OptionalCm $beforeFirstLineCm), (Format-OptionalCm $afterLeftCm), (Format-OptionalCm $afterFirstLineCm), $status))
            Add-Log(("WORD_COM_BODY_INDENT_FIX: i={{0}} paragraph_index={{1}} expected_left_cm={{2}} before_left_cm={{3}} after_left_cm={{4}} status={{5}}" -f $processed, $record.paragraph_index, $record.expected_left_cm, (Format-OptionalCm $beforeLeftCm), (Format-OptionalCm $afterLeftCm), $status))
        }} catch {{
            $errors += 1
            Add-Log(("WORD_COM_RECORD_EXCEPTION i={{0}} paragraph_index={{1}} type={{2}} message={{3}}" -f $processed, $record.paragraph_index, $_.Exception.GetType().FullName, $_.Exception.Message))
            Add-Log(("WORD_COM_RECORD_STACK i={{0}} stack={{1}}" -f $processed, $_.ScriptStackTrace))
            continue
        }}
    }}

    Add-Log 'WORD_COM_PS_BEFORE_SAVE'
    $doc.Save()
    Add-Log 'WORD_COM_PS_DOC_SAVED'
    Add-Log(("WORD_COM_BODY_INDENT_FIX_SUMMARY processed={{0}} ok={{1}} mismatch={{2}} not_found={{3}} errors={{4}}" -f $processed, $ok, $mismatch, $notFound, $errors))
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
            if completed.returncode == 1:
                logs.append("WORD_COM_BODY_INDENT_FIX_FAILED_AFTER_PARTIAL_LOGS")
                has_script_failure = True

            if not has_script_failure and completed.returncode == 0:
                try:
                    shutil.copy2(_long_path_compatible_str(work_docx_path), _long_path_compatible_str(output_docx))
                except Exception as exc:
                    logs.append(f"WORD_COM_BODY_INDENT_FIX_SKIPPED reason=copy_back_failed:{type(exc).__name__}:{exc}")
                    logs.append(f"exception_repr={exc!r}")
            elif not any(log.startswith("WORD_COM_BODY_INDENT_FIX_SKIPPED") for log in logs):
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


def fix_docx_fast(
    input_docx: str | Path,
    output_docx: str | Path,
    options: ProcessOptions,
    stop: StopController | None = None,
    progress_callback=None,
) -> ProcessSummary:
    input_docx = Path(input_docx)
    output_docx = Path(output_docx)

    if is_same_file_path(input_docx, output_docx):
        raise ValueError("Input and output paths must be different")

    if not input_docx.exists():
        raise FileNotFoundError(f"Input file does not exist: {input_docx}")

    if input_docx.suffix.lower() != ".docx":
        raise ValueError("Input file must be a .docx file")

    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    summary = ProcessSummary()
    global_table_index = 0

    with ZipFile(input_docx, "r") as zin, ZipFile(output_docx, "w", ZIP_DEFLATED) as zout:
        numbering_xml = zin.read("word/numbering.xml") if "word/numbering.xml" in zin.namelist() else None
        styles_xml = zin.read("word/styles.xml") if "word/styles.xml" in zin.namelist() else None
        formatted_numbering_xml = (
            apply_numbering_outline_format(numbering_xml, change_logs=summary.numbering_xml_logs)
            if options.fix_paragraph
            else numbering_xml
        )
        numbering_level_lookup = build_numbering_level_lookup(numbering_xml)
        numbering_format_lookup = build_numbering_format_lookup(formatted_numbering_xml)
        style_numbering_lookup = build_style_numbering_lookup(styles_xml)
        style_font_size_lookup = build_style_font_size_lookup(styles_xml)

        items = zin.infolist()
        total_items = max(len(items), 1)

        for item_index, item in enumerate(items):
            if stop:
                stop.check()

            if progress_callback:
                progress_callback(
                    percent=(item_index / total_items) * 100,
                    message=f"reading {item.filename}",
                )

            data = zin.read(item.filename)
            root = None

            if options.remove_all_outline_levels and should_remove_outline_part(item.filename):
                if progress_callback:
                    progress_callback(
                        percent=((item_index + 0.25) / total_items) * 100,
                        message=f"{item.filename}: removing outline levels",
                    )
                root = etree.fromstring(data, parser)
                if should_force_body_outline_part(item.filename):
                    force_all_paragraphs_to_body_outline_level(
                        root,
                        stop=stop,
                        summary=summary,
                    )
                else:
                    remove_all_outline_levels_from_any_root(
                        root,
                        stop=stop,
                        summary=summary,
                    )
                data = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

            # Normalize numbering definitions before document paragraph formatting.
            if item.filename == "word/numbering.xml" and options.fix_paragraph:
                if progress_callback:
                    progress_callback(
                        percent=((item_index + 0.5) / total_items) * 100,
                        message="word/numbering.xml: normalizing numbering indents",
                    )
                data = formatted_numbering_xml or data
                if options.remove_all_outline_levels:
                    root = etree.fromstring(data, parser)
                    remove_all_outline_levels_from_any_root(
                        root,
                        stop=stop,
                    )
                    data = etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )

            if item.filename == "word/styles.xml" and options.fix_paragraph:
                if root is None:
                    root = etree.fromstring(data, parser)
                apply_styles_outline_format_to_root(
                    root,
                    numbering_level_lookup=numbering_level_lookup,
                    style_numbering_lookup=style_numbering_lookup,
                    change_logs=summary.numbering_xml_logs,
                )
                data = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

            if should_process_part(item.filename):
                if root is None:
                    root = etree.fromstring(data, parser)

                if (
                    (
                        options.fix_paragraph
                        or options.indent_preface_paragraphs
                        or options.outline_preface_paragraphs
                    )
                    and should_fix_paragraph_part(item.filename)
                ):
                    if progress_callback:
                        message = "processing preface paragraphs"
                        if options.fix_paragraph:
                            message = "processing outline paragraphs"
                        progress_callback(
                            percent=((item_index + 0.95) / total_items) * 100,
                            message=f"{item.filename}: {message}",
                        )

                    changed_paragraphs = fix_outline_paragraphs(
                        root,
                        include_tables=False,
                        stop=stop,
                        numbering_level_lookup=numbering_level_lookup,
                        numbering_format_lookup=numbering_format_lookup,
                        style_numbering_lookup=style_numbering_lookup,
                        style_font_size_lookup=style_font_size_lookup,
                        change_logs=summary.paragraph_logs,
                        part_name=item.filename,
                        summary=summary,
                        fix_numbered_paragraphs=options.fix_paragraph,
                        indent_preface_paragraphs=options.indent_preface_paragraphs,
                        outline_preface_paragraphs=options.outline_preface_paragraphs,
                        enable_level1_level2_body_first_line_indent=options.enable_level1_level2_body_first_line_indent,
                        word_com_check_body_font_when_xml_not_14=options.word_com_check_body_font_when_xml_not_14,
                    )
                    summary.paragraphs += changed_paragraphs

                if options.fix_table_layout or options.fix_color:
                    tables = root.xpath(".//w:tbl", namespaces=NS)
                    table_count = len(tables)

                    for table_index, tbl in enumerate(tables, start=1):
                        if stop:
                            stop.check()

                        global_table_index += 1
                        cell_count = table_cell_count(tbl)
                        column_count = table_column_count(tbl)
                        table_name = find_previous_paragraph_text_for_table(root, tbl)
                        first_level_heading = "(none)"
                        if item.filename == "word/document.xml":
                            first_level_heading = (
                                find_table_first_level_heading(
                                    tbl,
                                    numbering_level_lookup,
                                    numbering_format_lookup,
                                    style_numbering_lookup,
                                )
                                or "(none)"
                            )

                        if item.filename == "word/document.xml" and table_index == 1:
                            summary.skipped_first_page_tables += 1
                            summary.table_log_records.append(
                                build_table_log_record(
                                    part_name=item.filename,
                                    table_index=table_index,
                                    global_table_index=global_table_index,
                                    table_name=table_name,
                                    first_level_heading=first_level_heading,
                                    cell_count=cell_count,
                                    column_count=column_count,
                                    table_type="skipped_first_table",
                                    action="skipped",
                                    reason="first table in word/document.xml",
                                    special_layout_used=False,
                                    layout_fixed=False,
                                    color_fixed=False,
                                    changed_to_gray=0,
                                    cleared_colors=0,
                                )
                            )
                            continue

                        if cell_count <= 4:
                            summary.skipped_small_tables += 1
                            summary.table_log_records.append(
                                build_table_log_record(
                                    part_name=item.filename,
                                    table_index=table_index,
                                    global_table_index=global_table_index,
                                    table_name=table_name,
                                    first_level_heading=first_level_heading,
                                    cell_count=cell_count,
                                    column_count=column_count,
                                    table_type="skipped_small_table",
                                    action="skipped",
                                    reason="cell_count <= 4",
                                    special_layout_used=False,
                                    layout_fixed=False,
                                    color_fixed=False,
                                    changed_to_gray=0,
                                    cleared_colors=0,
                                )
                            )
                            continue

                        skip_special_layout_under_chapter_three = (
                            options.skip_special_table_layout_under_chapter_three
                            and item.filename == "word/document.xml"
                            and is_table_under_chapter_three(
                                tbl,
                                numbering_level_lookup,
                                numbering_format_lookup,
                                style_numbering_lookup,
                            )
                        )
                        special_layout = (
                            options.fix_table_layout
                            and column_count <= 4
                            and not skip_special_layout_under_chapter_three
                        )
                        special_table_geometry = None
                        if special_layout:
                            special_table_geometry = _resolve_special_table_geometry(
                                tbl,
                                numbering_level_lookup,
                                style_numbering_lookup,
                            )
                        changed_to_gray, cleared_colors, shading_debug = process_table(
                            tbl,
                            options,
                            stop=stop,
                            special_layout=special_layout,
                            special_table_geometry=special_table_geometry,
                        )
                        layout_fixed = bool(options.fix_table_layout)
                        color_fixed = bool(options.fix_color)
                        if options.fix_table_layout:
                            if special_layout:
                                table_type = "special_table"
                                action = (
                                    "apply_special_table_format_and_color"
                                    if options.fix_color
                                    else "apply_special_table_format"
                                )
                                reason = "column_count <= 4"
                            else:
                                table_type = "normal_table"
                                action = (
                                    "apply_normal_table_format_and_color"
                                    if options.fix_color
                                    else "apply_normal_table_format"
                                )
                                reason = (
                                    "skipped special layout under chapter 參"
                                    if skip_special_layout_under_chapter_three and column_count <= 4
                                    else "column_count > 4"
                                )
                        elif options.fix_color:
                            table_type = "color_only_table"
                            action = "apply_color_only"
                            reason = "fix_table_layout disabled but fix_color enabled"
                        else:
                            table_type = "skipped"
                            action = "skipped"
                            reason = "no table actions enabled"
                        summary.changed_to_gray += changed_to_gray
                        summary.cleared_colors += cleared_colors
                        if special_layout:
                            summary.special_autofit_right_tables += 1
                        else:
                            summary.normal_processed_tables += 1
                        summary.table_log_records.append(
                            build_table_log_record(
                                part_name=item.filename,
                                table_index=table_index,
                                global_table_index=global_table_index,
                                table_name=table_name,
                                first_level_heading=first_level_heading,
                                cell_count=cell_count,
                                column_count=column_count,
                                table_type=table_type,
                                action=action,
                                reason=reason,
                                special_layout_used=special_layout,
                                layout_fixed=layout_fixed,
                                color_fixed=color_fixed,
                                changed_to_gray=changed_to_gray,
                                cleared_colors=cleared_colors,
                                shading_debug=shading_debug,
                            )
                        )

                        if progress_callback and table_count:
                            inner_fraction = table_index / table_count
                            percent = ((item_index + inner_fraction) / total_items) * 100
                            progress_callback(
                                percent=percent,
                                message=f"{item.filename}: table {table_index}/{table_count}",
                            )

                    summary.tables += table_count

                data = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

            if should_sanitize_indent_unit_part(item.filename):
                if root is None:
                    root = etree.fromstring(data, parser)
                removed_char_indent_attrs = remove_character_indent_attrs_from_root(root)
                if removed_char_indent_attrs:
                    summary.character_indent_attrs_removed += removed_char_indent_attrs
                data = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

            zout.writestr(item, data)

    if options.normalize_with_word_com:
        summary.word_com_body_indent_logs.extend(
            verify_and_fix_body_indents_with_word_com(output_docx, summary.body_indent_records, stop=stop)
        )
    else:
        summary.word_com_body_indent_logs.append("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=disabled")

    if progress_callback:
        progress_callback(percent=100, message="done")

    return summary


