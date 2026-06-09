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
from .indent_settings import twips_to_cm
from .models import ProcessOptions, ProcessSummary
from .numbering import (
    apply_numbering_outline_format,
    apply_styles_outline_format_to_root,
    build_numbering_format_lookup,
    build_numbering_level_lookup,
    build_style_numbering_lookup,
    has_auto_numbering,
    is_toc_style_definition,
    paragraph_style_id,
    style_name_value,
)
from .outline import (
    collect_all_toc_paragraph_ids,
    detect_manual_numbering_prefix,
    fix_outline_paragraphs,
    force_all_paragraphs_to_body_outline_level,
    get_auto_number_identity,
    remove_all_outline_levels_from_any_root,
    should_skip_style_numbering,
)
from .path_utils import is_same_file_path
from .process_runner import run_powershell_file, run_powershell_script
from .stop_controller import StopController
from .style_resolver import build_style_font_size_lookup
from .table_format import process_table, table_cell_count, table_column_count
from .xml_utils import paragraph_text, qn, remove_character_indent_attrs, remove_character_indent_attrs_from_root

POINTS_PER_CM = 28.3464567
WORD_COM_TIMEOUT_SECONDS = 600
WORD_COM_TEMP_DIR_NAME = "wfix"
CHAPTER_THREE_SKIP_TITLE = "價格形成之主要因素分析"
CHAPTER_THREE_SKIP_VISIBLE_PREFIXES = (
    "參、價格形成之主要因素分析",
)


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
    Apply paragraph indentation and outline fixes only to `word/document.xml`.

    Headers, footers, footnotes, and endnotes do not run the main body
    paragraph hierarchy logic, which avoids changing non-body regions.

    Table, color, and other XML handlers still decide their own applicability.
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


def _compact_heading_text(text: str) -> str:
    return "".join((text or "").split())


def is_chapter_three_start_marker(
    p,
    text: str,
    *,
    numbering_level_lookup=None,
    numbering_format_lookup=None,
    style_numbering_lookup=None,
) -> bool:
    del numbering_format_lookup  # Kept for parity with chapter prefix helpers.
    compact = _compact_heading_text(text)
    visible_prefixes = tuple(
        _compact_heading_text(prefix)
        for prefix in CHAPTER_THREE_SKIP_VISIBLE_PREFIXES
    )
    if compact.startswith(visible_prefixes):
        return True

    if not compact.startswith(_compact_heading_text(CHAPTER_THREE_SKIP_TITLE)):
        return False

    level = None
    if has_auto_numbering(p):
        num_id, ilvl = _effective_paragraph_numbering_identity(p, style_numbering_lookup)
        level = _outline_level_from_identity(num_id, ilvl, numbering_level_lookup)

    if level is None:
        num_id, ilvl = _effective_paragraph_numbering_identity(p, style_numbering_lookup)
        level = _outline_level_from_identity(num_id, ilvl, numbering_level_lookup)

    if level is None:
        manual = detect_manual_numbering_prefix(text)
        if manual is not None:
            level = manual[0]

    return level == 0


def is_table_under_chapter_three(
    tbl,
    numbering_level_lookup,
    numbering_format_lookup,
    style_numbering_lookup,
) -> bool:
    paragraphs = tbl.xpath("preceding::w:p[not(ancestor::w:tbl)]", namespaces=NS)
    for p in reversed(paragraphs):
        text = paragraph_text(p)
        prefix = _first_level_heading_prefix_for_paragraph(
            p,
            numbering_level_lookup=numbering_level_lookup,
            numbering_format_lookup=numbering_format_lookup,
            style_numbering_lookup=style_numbering_lookup,
        )
        if prefix is not None:
            return is_chapter_three_start_marker(
                p,
                text,
                numbering_level_lookup=numbering_level_lookup,
                numbering_format_lookup=numbering_format_lookup,
                style_numbering_lookup=style_numbering_lookup,
            )
    return False


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


def collect_chapter_three_paragraph_ids(
    root,
    *,
    numbering_level_lookup,
    numbering_format_lookup,
    style_numbering_lookup,
    toc_paragraph_ids=None,
) -> set[int]:
    """Collect paragraphs from the target chapter 參 title until the next first-level heading."""
    skip_ids: set[int] = set()
    toc_ids = toc_paragraph_ids or set()
    in_chapter_three = False

    for p in root.xpath(".//w:p", namespaces=NS):
        paragraph_id = id(p)
        if paragraph_id in toc_ids:
            continue

        text = paragraph_text(p)
        is_first_level_heading = False
        prefix = _first_level_heading_prefix_for_paragraph(
            p,
            numbering_level_lookup=numbering_level_lookup,
            numbering_format_lookup=numbering_format_lookup,
            style_numbering_lookup=style_numbering_lookup,
        )
        if prefix is not None:
            is_first_level_heading = True
        else:
            num_id, ilvl = _effective_paragraph_numbering_identity(p, style_numbering_lookup)
            level = _outline_level_from_identity(num_id, ilvl, numbering_level_lookup)

            if level is None:
                manual = detect_manual_numbering_prefix(text.strip())
                if manual is not None:
                    level = manual[0]

            is_first_level_heading = level == 0

        if is_chapter_three_start_marker(
            p,
            text,
            numbering_level_lookup=numbering_level_lookup,
            numbering_format_lookup=numbering_format_lookup,
            style_numbering_lookup=style_numbering_lookup,
        ):
            in_chapter_three = True
        elif in_chapter_three and is_first_level_heading:
            in_chapter_three = False

        if in_chapter_three:
            skip_ids.add(paragraph_id)

    return skip_ids


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


def _twips_to_log_cm(value: object) -> float | None:
    if value is None:
        return None
    try:
        return round(twips_to_cm(int(value)), 2)
    except (TypeError, ValueError):
        return None


def _manual_suffix_details(text: str, prefix: str) -> dict[str, object]:
    stripped = text.lstrip()
    prefix_start = len(text) - len(stripped)
    separator_start = prefix_start + len(prefix)
    separator_end = separator_start
    while separator_end < len(text) and text[separator_end] in {" ", "\t", "\u3000"}:
        separator_end += 1

    raw_separator = text[separator_start:separator_end]
    if raw_separator == "":
        suffix = "nothing"
    elif raw_separator[0] == "\t":
        suffix = "tab"
    elif raw_separator[0] in {" ", "\u3000"}:
        suffix = "space"
    else:
        suffix = "other"

    return {
        "suffix": suffix,
        "space_count": raw_separator.count(" ") + raw_separator.count("\u3000"),
        "tab_count": raw_separator.count("\t"),
        "raw_separator_repr": repr(raw_separator),
    }


def _auto_suffix_details(
    num_id: str | None,
    ilvl: int | None,
    numbering_format_lookup,
) -> dict[str, object]:
    level_format = numbering_format_lookup.get((num_id, ilvl), {}) if num_id is not None and ilvl is not None else {}
    raw_suffix = level_format.get("suff")
    if raw_suffix is None:
        suffix = "missing"
    elif raw_suffix in {"nothing", "tab", "space"}:
        suffix = raw_suffix
    else:
        suffix = "other"
    effective_suffix = "tab" if suffix == "missing" else suffix

    tab_pos = level_format.get("tab_pos")
    left = level_format.get("left")
    hanging = level_format.get("hanging")
    number_start = level_format.get("number_start")
    return {
        "suffix": suffix,
        "raw_suffix": suffix,
        "effective_suffix": effective_suffix,
        "numId": num_id,
        "ilvl": ilvl,
        "numFmt": level_format.get("numFmt"),
        "lvlText": level_format.get("lvlText"),
        "has_tab_stop": tab_pos is not None,
        "tab_pos_twips": tab_pos,
        "tab_pos_cm": _twips_to_log_cm(tab_pos),
        "left_twips": left,
        "hanging_twips": hanging,
        "number_start_twips": number_start,
        "left_cm": _twips_to_log_cm(left),
        "hanging_cm": _twips_to_log_cm(hanging),
        "number_start_cm": _twips_to_log_cm(number_start),
    }


def collect_heading_suffix_records_from_docx(docx_path: str | Path) -> list[dict[str, object]]:
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    records: list[dict[str, object]] = []
    with ZipFile(docx_path, "r") as zin:
        names = set(zin.namelist())
        numbering_xml = zin.read("word/numbering.xml") if "word/numbering.xml" in names else None
        styles_xml = zin.read("word/styles.xml") if "word/styles.xml" in names else None
        numbering_level_lookup = build_numbering_level_lookup(numbering_xml)
        numbering_format_lookup = build_numbering_format_lookup(numbering_xml)
        style_numbering_lookup = build_style_numbering_lookup(styles_xml)

        for part_name in sorted(name for name in names if should_process_part(name)):
            try:
                root = etree.fromstring(zin.read(part_name), parser)
            except Exception:
                continue

            for paragraph_index, p in enumerate(root.xpath(".//w:p", namespaces=NS), start=1):
                if p.xpath("ancestor::w:tbl", namespaces=NS):
                    continue

                text = paragraph_text(p)
                if not text or not text.strip():
                    continue

                num_id = None
                ilvl = None
                level = None
                source = None
                number_token = None
                details: dict[str, object] | None = None

                if has_auto_numbering(p):
                    num_id, ilvl = get_auto_number_identity(p)
                    level = _outline_level_from_identity(num_id, ilvl, numbering_level_lookup)
                    if level is not None:
                        source = "auto_numbering_xml"
                        details = _auto_suffix_details(num_id, ilvl, numbering_format_lookup)

                if level is None and not should_skip_style_numbering(text):
                    num_id, ilvl = _effective_paragraph_numbering_identity(p, style_numbering_lookup)
                    level = _outline_level_from_identity(num_id, ilvl, numbering_level_lookup)
                    if level is not None:
                        source = "auto_numbering_xml"
                        details = _auto_suffix_details(num_id, ilvl, numbering_format_lookup)

                if level is None:
                    manual = detect_manual_numbering_prefix(text)
                    if manual is not None:
                        level, number_token = manual
                        source = "manual_text"
                        details = _manual_suffix_details(text, number_token)

                if level is None or source is None or details is None:
                    continue
                if level < 0 or level > 8:
                    continue

                if number_token is None:
                    number_token = (details.get("lvlText") if details else None) or "(auto)"

                records.append(
                    {
                        "part_name": part_name,
                        "paragraph_index": paragraph_index,
                        "source": source,
                        "outline_level": level,
                        "heading_text": text,
                        "number_token": number_token,
                        **details,
                    }
                )

    return records


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
    expected_first_line_twips = record.get("expected_first_line_twips")
    if expected_first_line_twips is not None and str(expected_first_line_twips).strip():
        ind.set(qn("firstLine"), str(expected_first_line_twips))
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
            Add-Log(("WORD_COM_INDENT_VERIFY: paragraph_index={{0}}; matched_paragraph_index={{1}}; text_preview={{2}}; kind={{3}}; level={{4}}; expected_number_start_cm={{5}}; expected_hanging_cm={{6}}; expected_heading_left_cm={{7}}; expected_body_left_cm={{8}}; expected_first_line_twips={{9}}; xml_font_size={{10}}; xml_font_size_source={{11}}; word_range_font_size={{12}}; word_dominant_font_size={{13}}; word_font_check_pass={{14}}; decision=approved_for_xml_body_indent; word_opened_left_cm=not_read; word_opened_firstline_cm=not_read; final_left_cm=not_read; final_firstline_cm=not_read; second_fix=python_xml; status={{15}}" -f $record.paragraph_index, $matchIndex, $preview, $kind, $record.level, $record.expected_number_start_cm, $record.expected_hanging_cm, $record.expected_heading_left_cm, $record.expected_body_left_cm, $record.expected_first_line_twips, $record.xml_font_size, $record.xml_font_size_source, $wordRangeFontSize, $wordDominantFontSize, $wordFontCheckPass, $status))
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


def collect_toc_numbering_exclusions(
    document_root,
    toc_paragraph_ids: set[int],
    style_numbering_lookup: dict[str, tuple[str, int]],
    numbering_xml: bytes | None,
    paragraphs=None,
) -> tuple[set[tuple[str, int]], set[str], set[str]]:
    pairs: set[tuple[str, int]] = set()
    num_ids: set[str] = set()
    abstract_ids: set[str] = set()
    num_to_abstract_id: dict[str, str] = {}

    if numbering_xml:
        try:
            numbering_root = etree.fromstring(numbering_xml)
            for num in numbering_root.xpath("./w:num", namespaces=NS):
                num_id = num.get(qn("numId"))
                abstract_el = num.find("w:abstractNumId", NS)
                abstract_id = abstract_el.get(qn("val")) if abstract_el is not None else None
                if num_id is not None and abstract_id is not None:
                    num_to_abstract_id[num_id] = abstract_id
        except Exception:
            pass

    paragraphs = paragraphs if paragraphs is not None else document_root.xpath(".//w:p", namespaces=NS)
    for p in paragraphs:
        if id(p) not in toc_paragraph_ids:
            continue

        num_id = None
        ilvl = None
        if has_auto_numbering(p):
            num_id, ilvl = get_auto_number_identity(p)
        if num_id is None:
            style_id = paragraph_style_id(p)
            if style_id:
                num_id, ilvl = style_numbering_lookup.get(style_id, (None, None))

        if num_id is None:
            continue
        if ilvl is None:
            ilvl = 0

        num_ids.add(str(num_id))
        pairs.add((str(num_id), int(ilvl)))
        abstract_id = num_to_abstract_id.get(str(num_id))
        if abstract_id is not None:
            abstract_ids.add(abstract_id)

    return pairs, num_ids, abstract_ids


def remove_character_indent_attrs_from_styles_root_excluding_toc(
    root,
    change_logs: list[str] | None = None,
) -> int:
    removed = 0
    for style in root.xpath("./w:style[@w:type='paragraph']", namespaces=NS):
        style_id = style.get(qn("styleId")) or ""
        style_name = style_name_value(style)
        if is_toc_style_definition(style_id, style_name):
            continue
        for ind in style.xpath(".//w:ind", namespaces=NS):
            removed += remove_character_indent_attrs(ind)
    return removed


def remove_character_indent_attrs_from_numbering_root_excluding_toc(
    root,
    excluded_numbering_pairs: set[tuple[str, int]],
    excluded_num_ids: set[str],
    excluded_abstract_ids: set[str],
) -> int:
    removed = 0
    num_to_abstract_id: dict[str, str] = {}
    for num in root.xpath("./w:num", namespaces=NS):
        num_id = num.get(qn("numId"))
        abstract_el = num.find("w:abstractNumId", NS)
        abstract_id = abstract_el.get(qn("val")) if abstract_el is not None else None
        if num_id is not None and abstract_id is not None:
            num_to_abstract_id[num_id] = abstract_id

    def should_skip(num_id: str | None, ilvl: int | None, abstract_id: str | None) -> bool:
        if abstract_id is not None and abstract_id in excluded_abstract_ids:
            return True
        if num_id is not None and num_id in excluded_num_ids:
            return True
        if num_id is not None and ilvl is not None and (num_id, ilvl) in excluded_numbering_pairs:
            return True
        return False

    for lvl in root.xpath("./w:abstractNum/w:lvl", namespaces=NS):
        abstract_num = lvl.getparent()
        abstract_id = abstract_num.get(qn("abstractNumId")) if abstract_num is not None else None
        try:
            ilvl = int(lvl.get(qn("ilvl")))
        except Exception:
            ilvl = None
        if should_skip(None, ilvl, abstract_id):
            continue
        for ind in lvl.xpath(".//w:ind", namespaces=NS):
            removed += remove_character_indent_attrs(ind)

    for lvl in root.xpath("./w:num/w:lvlOverride/w:lvl", namespaces=NS):
        override = lvl.getparent()
        num = override.getparent() if override is not None else None
        num_id = num.get(qn("numId")) if num is not None else None
        abstract_id = num_to_abstract_id.get(num_id or "")
        try:
            ilvl = int(override.get(qn("ilvl"))) if override is not None else None
        except Exception:
            ilvl = None
        if should_skip(num_id, ilvl, abstract_id):
            continue
        for ind in lvl.xpath(".//w:ind", namespaces=NS):
            removed += remove_character_indent_attrs(ind)

    return removed


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
    try:
        summary.heading_suffix_before_records = collect_heading_suffix_records_from_docx(input_docx)
    except Exception as exc:
        summary.heading_suffix_before_records = [
            {
                "part_name": "(scan_error)",
                "paragraph_index": 0,
                "source": "error",
                "outline_level": None,
                "heading_text": f"BEFORE_FIX scan failed: {exc!r}",
                "number_token": None,
                "suffix": "other",
            }
        ]

    with ZipFile(input_docx, "r") as zin, ZipFile(output_docx, "w", ZIP_DEFLATED) as zout:
        numbering_xml = zin.read("word/numbering.xml") if "word/numbering.xml" in zin.namelist() else None
        styles_xml = zin.read("word/styles.xml") if "word/styles.xml" in zin.namelist() else None
        numbering_level_lookup = build_numbering_level_lookup(numbering_xml)
        style_numbering_lookup = build_style_numbering_lookup(styles_xml)
        style_font_size_lookup = build_style_font_size_lookup(styles_xml)
        document_root_for_toc = None
        document_toc_paragraph_ids: set[int] = set()
        document_chapter_three_paragraph_ids: set[int] = set()
        toc_numbering_pairs: set[tuple[str, int]] = set()
        toc_num_ids: set[str] = set()
        toc_abstract_ids: set[str] = set()
        chapter_three_numbering_pairs: set[tuple[str, int]] = set()
        chapter_three_num_ids: set[str] = set()
        chapter_three_abstract_ids: set[str] = set()
        chapter_three_style_ids: set[str] = set()
        original_numbering_format_lookup = build_numbering_format_lookup(numbering_xml)
        if "word/document.xml" in zin.namelist():
            try:
                document_root_for_toc = etree.fromstring(zin.read("word/document.xml"), parser)
                document_paragraphs_for_toc = document_root_for_toc.xpath(".//w:p", namespaces=NS)
                document_toc_paragraph_ids = collect_all_toc_paragraph_ids(
                    document_root_for_toc,
                    numbering_level_lookup=numbering_level_lookup,
                    style_numbering_lookup=style_numbering_lookup,
                    paragraphs=document_paragraphs_for_toc,
                )
                toc_numbering_pairs, toc_num_ids, toc_abstract_ids = collect_toc_numbering_exclusions(
                    document_root_for_toc,
                    document_toc_paragraph_ids,
                    style_numbering_lookup,
                    numbering_xml,
                    paragraphs=document_paragraphs_for_toc,
                )
                if options.skip_all_under_chapter_three:
                    document_chapter_three_paragraph_ids = collect_chapter_three_paragraph_ids(
                        document_root_for_toc,
                        numbering_level_lookup=numbering_level_lookup,
                        numbering_format_lookup=original_numbering_format_lookup,
                        style_numbering_lookup=style_numbering_lookup,
                        toc_paragraph_ids=document_toc_paragraph_ids,
                    )
                    summary.numbering_xml_logs.append(
                        f"CHAPTER_THREE_SKIP_IDS collected={len(document_chapter_three_paragraph_ids)}"
                    )
                    (
                        chapter_three_numbering_pairs,
                        chapter_three_num_ids,
                        chapter_three_abstract_ids,
                    ) = collect_toc_numbering_exclusions(
                        document_root_for_toc,
                        document_chapter_three_paragraph_ids,
                        style_numbering_lookup,
                        numbering_xml,
                        paragraphs=document_paragraphs_for_toc,
                    )
                    chapter_three_style_ids = {
                        style_id
                        for p in document_paragraphs_for_toc
                        if id(p) in document_chapter_three_paragraph_ids
                        for style_id in [paragraph_style_id(p)]
                        if style_id
                    }
            except Exception:
                document_root_for_toc = None
                document_toc_paragraph_ids = set()
                document_chapter_three_paragraph_ids = set()
                chapter_three_style_ids = set()
        excluded_numbering_pairs = set(toc_numbering_pairs)
        excluded_numbering_pairs.update(chapter_three_numbering_pairs)
        excluded_num_ids = set(toc_num_ids)
        excluded_num_ids.update(chapter_three_num_ids)
        excluded_abstract_ids = set(toc_abstract_ids)
        excluded_abstract_ids.update(chapter_three_abstract_ids)
        formatted_numbering_xml = (
            apply_numbering_outline_format(
                numbering_xml,
                change_logs=summary.numbering_xml_logs,
                excluded_numbering_pairs=excluded_numbering_pairs,
                excluded_num_ids=excluded_num_ids,
                excluded_abstract_ids=excluded_abstract_ids,
            )
            if options.fix_paragraph
            else numbering_xml
        )
        numbering_format_lookup = build_numbering_format_lookup(formatted_numbering_xml)

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
            if item.filename == "word/document.xml" and document_root_for_toc is not None:
                root = document_root_for_toc
            chapter_three_exclude_paragraph_ids = (
                document_chapter_three_paragraph_ids
                if item.filename == "word/document.xml" and document_chapter_three_paragraph_ids
                else None
            )

            if options.remove_all_outline_levels and should_remove_outline_part(item.filename):
                if progress_callback:
                    progress_callback(
                        percent=((item_index + 0.25) / total_items) * 100,
                        message=f"{item.filename}: removing outline levels",
                    )
                if root is None:
                    root = etree.fromstring(data, parser)
                if should_force_body_outline_part(item.filename):
                    force_all_paragraphs_to_body_outline_level(
                        root,
                        stop=stop,
                        summary=summary,
                        exclude_paragraph_ids=chapter_three_exclude_paragraph_ids,
                    )
                elif (
                    item.filename == "word/numbering.xml"
                    and options.skip_all_under_chapter_three
                    and excluded_abstract_ids
                ):
                    summary.numbering_xml_logs.append(
                        "WARNING: chapter 參 uses numbering definitions excluded from "
                        "remove_all_outline_levels; numbering.xml outline removal skipped"
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
                    if options.skip_all_under_chapter_three and excluded_abstract_ids:
                        summary.numbering_xml_logs.append(
                            "WARNING: chapter 參 uses numbering definitions excluded from "
                            "remove_all_outline_levels; numbering.xml normalization may still affect shared definitions"
                        )
                    else:
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
                    excluded_style_ids=chapter_three_style_ids,
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
                        skip_paragraph_ids=chapter_three_exclude_paragraph_ids,
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

                        skip_all_under_chapter_three = (
                            options.skip_all_under_chapter_three
                            and item.filename == "word/document.xml"
                            and chapter_three_exclude_paragraph_ids
                            and any(
                                id(p) in chapter_three_exclude_paragraph_ids
                                for p in tbl.xpath(".//w:p", namespaces=NS)
                            )
                        )
                        if skip_all_under_chapter_three:
                            summary.table_log_records.append(
                                build_table_log_record(
                                    part_name=item.filename,
                                    table_index=table_index,
                                    global_table_index=global_table_index,
                                    table_name=table_name,
                                    first_level_heading=first_level_heading,
                                    cell_count=cell_count,
                                    column_count=column_count,
                                    table_type="skipped_chapter_three_table",
                                    action="skipped",
                                    reason=(
                                        "under chapter 參、價格形成之主要因素分析; "
                                        "all table layout and color fixes skipped"
                                    ),
                                    special_layout_used=False,
                                    layout_fixed=False,
                                    color_fixed=False,
                                    changed_to_gray=0,
                                    cleared_colors=0,
                                    shading_debug=[],
                                )
                            )
                            continue

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

                        special_layout = (
                            options.fix_table_layout
                            and column_count <= 4
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
                                reason = "column_count > 4"
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
                exclude_paragraph_ids = None
                if item.filename == "word/document.xml":
                    exclude_set: set[int] = set()
                    exclude_set.update(document_toc_paragraph_ids)
                    exclude_set.update(document_chapter_three_paragraph_ids)
                    exclude_paragraph_ids = exclude_set or None
                if item.filename == "word/styles.xml":
                    removed_char_indent_attrs = remove_character_indent_attrs_from_styles_root_excluding_toc(
                        root,
                        change_logs=summary.numbering_xml_logs,
                    )
                elif item.filename == "word/numbering.xml":
                    removed_char_indent_attrs = remove_character_indent_attrs_from_numbering_root_excluding_toc(
                        root,
                        toc_numbering_pairs,
                        toc_num_ids,
                        toc_abstract_ids,
                    )
                else:
                    removed_char_indent_attrs = remove_character_indent_attrs_from_root(
                        root,
                        exclude_paragraph_ids=exclude_paragraph_ids,
                        change_logs=summary.numbering_xml_logs,
                        part_name=item.filename,
                    )
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
        word_com_records = _filter_word_com_body_indent_records(summary.body_indent_records)
        summary.word_com_body_indent_logs.append(
            "WORD_COM_BODY_INDENT_RECORD_FILTER "
            f"total_records={len(summary.body_indent_records)} "
            f"word_com_records={len(word_com_records)} "
            f"skipped_records={len(summary.body_indent_records) - len(word_com_records)} "
            "criteria=apply_only_if_word_font_size_is_14_and_xml_font_size_gt_11"
        )
        if word_com_records:
            summary.word_com_body_indent_logs.extend(
                verify_and_fix_body_indents_with_word_com(output_docx, word_com_records, stop=stop)
            )
        else:
            summary.word_com_body_indent_logs.append(
                "WORD_COM_BODY_INDENT_FIX_SKIPPED reason=no_font_check_records"
            )
    else:
        summary.word_com_body_indent_logs.append("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=disabled")

    try:
        summary.heading_suffix_after_records = collect_heading_suffix_records_from_docx(output_docx)
    except Exception as exc:
        summary.heading_suffix_after_records = [
            {
                "part_name": "(scan_error)",
                "paragraph_index": 0,
                "source": "error",
                "outline_level": None,
                "heading_text": f"AFTER_FIX scan failed: {exc!r}",
                "number_token": None,
                "suffix": "other",
            }
        ]

    if progress_callback:
        progress_callback(percent=100, message="done")

    return summary


