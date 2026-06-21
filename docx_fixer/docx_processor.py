from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from .constants import NS
from .exceptions import ProcessStopped
from .indent_settings import twips_to_cm
from .indent_sanitizer import (
    remove_character_indent_attrs_from_numbering_root_excluding_protected,
    remove_character_indent_attrs_from_styles_root_excluding_protected,
)
from .models import ProcessOptions, ProcessSummary
from .note_alignment import force_note_paragraph_left_alignment_in_docx
from .note_debug_log import write_note_debug_log_for_docx
from .numbering_cleanup import force_clean_numbering_suffix_tabs_in_docx
from .numbering import (
    apply_numbering_outline_format,
    apply_styles_outline_format_to_root,
    build_numbering_format_lookup,
    build_numbering_level_lookup,
    build_style_numbering_lookup,
    has_auto_numbering,
)
from .outline import (
    detect_manual_numbering_prefix,
    fix_outline_paragraphs,
    force_all_paragraphs_to_body_outline_level,
    get_auto_number_identity,
    remove_all_outline_levels_from_any_root,
    restore_outline_levels_for_protected_paragraphs,
    should_skip_style_numbering,
)
from .path_utils import is_same_file_path
from .protected_region import (
    _effective_paragraph_numbering_identity,
    ProtectedRegionContext,
    _outline_level_from_identity,
)
from .stop_controller import StopController
from .style_resolver import build_style_font_size_lookup
from .table_fallback import fallback_normal_table_autofit_in_docx
from .table_footer_postprocess import apply_table_footer_source_format_in_docx
from .table_pipeline import process_tables_in_part
from .table_word_com import WORD_COM_AUTOFIT_SEQUENCE, apply_table_autofit_with_word_com
from . import word_com_indent
from .xml_utils import paragraph_text, qn, remove_character_indent_attrs_from_root






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





def _twips_to_log_cm(value: object) -> float | None:
    if value is None:
        return None
    try:
        return round(twips_to_cm(int(value)), 2)
    except (TypeError, ValueError):
        return None


def _tabs_summary_from_pPr(pPr) -> str:
    if pPr is None:
        return "none"
    tabs = pPr.findall("./w:tabs/w:tab", NS)
    if not tabs:
        return "none"
    values = []
    for tab in tabs:
        pos = tab.get(qn("pos")) or "unknown"
        val = tab.get(qn("val")) or "unknown"
        values.append(f"{val}@{pos}")
    return ",".join(values)


def _style_tabs_lookup(styles_xml: bytes | None) -> dict[str, str]:
    if not styles_xml:
        return {}
    try:
        root = etree.fromstring(styles_xml)
    except Exception:
        return {}
    lookup: dict[str, str] = {}
    for style in root.xpath("./w:style[@w:type='paragraph']", namespaces=NS):
        style_id = style.get(qn("styleId"))
        if not style_id:
            continue
        lookup[style_id] = _tabs_summary_from_pPr(style.find("./w:pPr", NS))
    return lookup


def _numbering_context_details(
    p,
    *,
    numbering_format_lookup,
    style_numbering_lookup,
    style_tabs_lookup,
) -> dict[str, object]:
    pPr = p.find("./w:pPr", NS)
    paragraph_num_pr = pPr.find("w:numPr", NS) if pPr is not None else None
    paragraph_num_id, paragraph_ilvl = get_auto_number_identity(p)
    style_id = None
    style_el = p.find("./w:pPr/w:pStyle", NS)
    if style_el is not None:
        style_id = style_el.get(qn("val"))

    style_num_id = None
    style_ilvl = None
    if style_id and style_numbering_lookup:
        style_num_id, style_ilvl = style_numbering_lookup.get(style_id, (None, None))

    effective_num_id = paragraph_num_id if paragraph_num_id is not None else style_num_id
    effective_ilvl = paragraph_ilvl if paragraph_num_id is not None else style_ilvl
    level_format = (
        numbering_format_lookup.get((effective_num_id, effective_ilvl), {})
        if effective_num_id is not None and effective_ilvl is not None
        else {}
    )
    style_num_pr = (
        f"{style_num_id}:{style_ilvl}"
        if style_num_id is not None and style_ilvl is not None
        else "none"
    )
    return {
        "paragraph_has_numPr": paragraph_num_pr is not None,
        "paragraph_tabs": _tabs_summary_from_pPr(pPr),
        "numId": effective_num_id,
        "ilvl": effective_ilvl,
        "numbering_suff": level_format.get("suff"),
        "numbering_tab_pos": level_format.get("tab_pos"),
        "style_numPr": style_num_pr,
        "style_tabs": style_tabs_lookup.get(style_id or "", "none"),
    }


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
    lvl_text = level_format.get("lvlText")
    return {
        "suffix": suffix,
        "raw_suffix": suffix,
        "effective_suffix": effective_suffix,
        "numId": num_id,
        "ilvl": ilvl,
        "numbering_suff": suffix,
        "numbering_tab_pos": tab_pos,
        "numFmt": level_format.get("numFmt"),
        "lvlText": lvl_text,
        "lvlText_has_trailing_space": isinstance(lvl_text, str) and lvl_text.endswith((" ", "\t", "\u3000")),
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
        style_tabs_lookup = _style_tabs_lookup(styles_xml)

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
                context_details = _numbering_context_details(
                    p,
                    numbering_format_lookup=numbering_format_lookup,
                    style_numbering_lookup=style_numbering_lookup,
                    style_tabs_lookup=style_tabs_lookup,
                )

                manual = detect_manual_numbering_prefix(text)
                if manual is not None:
                    level, number_token = manual
                    source = "manual_text"
                    details = _manual_suffix_details(text, number_token)

                if level is None and has_auto_numbering(p):
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
                        **context_details,
                        **details,
                    }
                )

    return records


def _mark_word_com_table_autofit_applied(
    summary: ProcessSummary,
    applied_global_table_indices: set[int],
) -> None:
    if not applied_global_table_indices:
        return

    for record in summary.table_log_records:
        try:
            global_table_index = int(record.get("global_table_index", 0))
        except (TypeError, ValueError):
            continue
        if global_table_index not in applied_global_table_indices:
            continue
        record["word_com_autofit_applied"] = True
        record["word_com_autofit_sequence"] = WORD_COM_AUTOFIT_SEQUENCE
        record["word_com_autofit_status"] = "word_com"


def _mark_word_com_table_autofit_fallback(
    summary: ProcessSummary,
    fallback_applied_indices: set[int],
    failed_indices: set[int],
) -> None:
    if not failed_indices:
        return

    for record in summary.table_log_records:
        try:
            global_table_index = int(record.get("global_table_index", 0))
        except (TypeError, ValueError):
            continue
        if global_table_index not in failed_indices:
            continue
        if global_table_index in fallback_applied_indices:
            record["word_com_autofit_fallback_applied"] = True
            record["word_com_autofit_status"] = "xml_fallback"
        else:
            record["word_com_autofit_fallback_applied"] = False
            record["word_com_autofit_status"] = "failed"


def _mark_table_footer_source_format_applied(
    summary: ProcessSummary,
    footer_results: dict[int, dict[str, object]],
) -> None:
    if not footer_results:
        return

    for record in summary.table_log_records:
        try:
            global_table_index = int(record.get("global_table_index", 0))
        except (TypeError, ValueError):
            continue
        result = footer_results.get(global_table_index)
        if result is None:
            continue
        record["table_footer_note_source_format_applied"] = True
        record["table_footer_note_source_format_skipped_reason"] = "none"
        record["outer_double_border_applied_by_footer_source_format"] = bool(
            result.get("outer_double_border_applied", False)
        )
        record["first_row_single_cell_border_adjusted"] = bool(
            result.get("first_row_single_cell_border_adjusted", False)
        )
        record["footer_note_cells_adjusted"] = int(
            result.get("footer_note_cells_adjusted", 0)
        )
        record["footer_note_cell_matches"] = list(
            result.get("footer_note_cell_matches", [])
        )
        record["footer_note_cell_debug"] = list(result.get("footer_note_cell_debug", []))
        record["footer_rows_processed"] = int(result.get("footer_rows_processed", 0))
        record["footer_row_matches"] = list(result.get("footer_row_matches", []))

    summary.table_footer_source_format_tables = len(footer_results)


def _format_id_level_pairs(pairs: set[tuple[str, int]]) -> str:
    if not pairs:
        return "none"
    return ",".join(f"{ident}:{level}" for ident, level in sorted(pairs))


def _build_num_to_abstract_id_map(numbering_xml: bytes | None) -> dict[str, str]:
    mapping: dict[str, str] = {}
    if not numbering_xml:
        return mapping
    try:
        root = etree.fromstring(numbering_xml)
    except Exception:
        return mapping
    for num in root.xpath("./w:num", namespaces=NS):
        num_id = num.get(qn("numId"))
        abstract_el = num.find("w:abstractNumId", NS)
        abstract_id = abstract_el.get(qn("val")) if abstract_el is not None else None
        if num_id is not None and abstract_id is not None:
            mapping[num_id] = abstract_id
    return mapping


def _note_debug_log_path(output_docx: Path) -> Path:
    return output_docx.with_name(f"{output_docx.stem}_note_debug_log.txt")


def _write_note_debug_log_safely(
    output_docx: Path,
    summary: ProcessSummary,
    stage: str,
    *,
    append: bool,
) -> None:
    try:
        log_path = _note_debug_log_path(output_docx)
        write_note_debug_log_for_docx(output_docx, log_path, stage, append=append)
        summary.paragraph_logs.append(
            f"NOTE_DEBUG_LOG_WRITTEN stage={stage} path={log_path}"
        )
    except Exception as exc:
        summary.paragraph_logs.append(
            f"NOTE_DEBUG_LOG_FAILED stage={stage} reason={type(exc).__name__}:{exc}"
        )








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
        protected_context = ProtectedRegionContext()
        original_numbering_format_lookup = build_numbering_format_lookup(numbering_xml)
        if "word/document.xml" in zin.namelist():
            try:
                document_root_for_toc = etree.fromstring(zin.read("word/document.xml"), parser)
                protect_chapter_three = (
                    options.skip_chapter_three_table_layout
                    or options.skip_chapter_three_table_color
                    or options.skip_chapter_three_indents
                    or options.skip_chapter_three_numbering_suffix_cleanup
                )
                collect_section_three_note_region = bool(
                    options.move_table_notes_below and options.skip_chapter_three_table_notes
                )
                protected_context = ProtectedRegionContext.from_document(
                    document_root_for_toc,
                    protect_chapter_three=protect_chapter_three,
                    numbering_level_lookup=numbering_level_lookup,
                    numbering_format_lookup=original_numbering_format_lookup,
                    style_numbering_lookup=style_numbering_lookup,
                    numbering_xml=numbering_xml,
                    summary=summary,
                    use_generic_section_three=options.skip_chapter_three_adjustments,
                    collect_section_three_note_region=collect_section_three_note_region,
                )
            except Exception:
                document_root_for_toc = None
                protected_context = ProtectedRegionContext()
        indent_excluded_numbering_pairs = set(protected_context.toc_numbering_pairs)
        indent_excluded_num_ids = set(protected_context.toc_num_ids)
        indent_excluded_abstract_ids = set(protected_context.toc_abstract_ids)
        indent_excluded_style_ids: set[str] = set()

        # 「參、不要清理編號後綴 tab/space」(skip_chapter_three_numbering_suffix_cleanup):
        # when enabled, the numbering suffix/tab/lvlText cleanup passes
        # (apply_numbering_outline_format + the final force_clean) must also
        # exclude the chapter 參 numbering definitions, on top of the TOC
        # exclusions. This only governs numbering.xml suffix cleanup; it does
        # not touch table layout/color, paragraph indents, or the numbering
        # char-indent sanitizer (which keeps using the TOC-only set above).
        skip_chapter_three_numbering_suffix_cleanup = bool(
            getattr(options, "skip_chapter_three_numbering_suffix_cleanup", True)
        )
        chapter_three_numbering_pairs = set(protected_context.chapter_three_numbering_pairs)
        # Express the 參 pairs as precise (abstractNumId, ilvl) protection for the
        # log and abstract-level matching — never as whole abstractNumIds, which
        # 壹/貳/參/肆 commonly share.
        num_to_abstract_id_map = _build_num_to_abstract_id_map(numbering_xml)
        chapter_three_abstract_levels = {
            (num_to_abstract_id_map[num_id], ilvl)
            for (num_id, ilvl) in chapter_three_numbering_pairs
            if num_id in num_to_abstract_id_map
        }
        if skip_chapter_three_numbering_suffix_cleanup:
            # Chapter 參 contributes ONLY its actual (numId, ilvl) pairs — never
            # whole numIds or abstractNumIds, which are shared with the other
            # chapters. TOC keeps its original abstractId / numId / pair
            # exclusions so the TOC behaviour is unchanged.
            numbering_suffix_excluded_numbering_pairs = (
                set(protected_context.toc_numbering_pairs) | chapter_three_numbering_pairs
            )
            numbering_suffix_excluded_num_ids = set(protected_context.toc_num_ids)
            numbering_suffix_excluded_abstract_ids = set(protected_context.toc_abstract_ids)
            # Final cleanup keeps the original full body-heading re-include set;
            # chapter 參 pairs are passed separately as a hard protection that
            # wins over that re-include (force_clean_numbering_suffix_tabs).
            final_included_numbering_pairs = set(protected_context.body_heading_numbering_pairs)
            final_included_num_ids = set(protected_context.body_heading_num_ids)
            final_included_abstract_ids = set(protected_context.body_heading_abstract_ids)
            final_protected_numbering_pairs = set(chapter_three_numbering_pairs)
            summary.numbering_xml_logs.append(
                "CHAPTER_THREE_NUMBERING_SUFFIX_CLEANUP_SKIP enabled=true "
                f"protected_pairs={_format_id_level_pairs(chapter_three_numbering_pairs)} "
                f"protected_abstract_levels={_format_id_level_pairs(chapter_three_abstract_levels)} "
                "protected_abstractIds_not_used_for_chapter_three=true"
            )
        else:
            numbering_suffix_excluded_numbering_pairs = set(protected_context.toc_numbering_pairs)
            numbering_suffix_excluded_num_ids = set(protected_context.toc_num_ids)
            numbering_suffix_excluded_abstract_ids = set(protected_context.toc_abstract_ids)
            final_included_numbering_pairs = set(protected_context.body_heading_numbering_pairs)
            final_included_num_ids = set(protected_context.body_heading_num_ids)
            final_included_abstract_ids = set(protected_context.body_heading_abstract_ids)
            final_protected_numbering_pairs = set()
            summary.numbering_xml_logs.append(
                "CHAPTER_THREE_NUMBERING_SUFFIX_CLEANUP_SKIP enabled=false"
            )

        final_suffix_excluded_numbering_pairs = numbering_suffix_excluded_numbering_pairs
        final_suffix_excluded_num_ids = numbering_suffix_excluded_num_ids
        final_suffix_excluded_abstract_ids = numbering_suffix_excluded_abstract_ids
        formatted_numbering_xml = (
            apply_numbering_outline_format(
                numbering_xml,
                change_logs=summary.numbering_xml_logs,
                excluded_numbering_pairs=numbering_suffix_excluded_numbering_pairs,
                excluded_num_ids=numbering_suffix_excluded_num_ids,
                excluded_abstract_ids=numbering_suffix_excluded_abstract_ids,
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
            toc_paragraph_ids = protected_context.toc_paragraph_ids_for_part(item.filename)
            # remove_all_outline_levels intentionally also applies to TOC
            # paragraphs. Other cleanup steps still use toc_paragraph_ids to
            # avoid changing TOC indents, tabs, fields, styles, and numbering.
            outline_body_level_exclude_paragraph_ids = None
            chapter_three_paragraph_ids = protected_context.chapter_three_paragraph_ids_for_part(
                item.filename
            )
            indent_skip_paragraph_ids = (
                chapter_three_paragraph_ids
                if options.skip_chapter_three_indents
                else None
            )
            sanitize_exclude_paragraph_ids = toc_paragraph_ids
            if options.skip_chapter_three_indents and chapter_three_paragraph_ids:
                sanitize_ids = set(sanitize_exclude_paragraph_ids or set())
                sanitize_ids.update(chapter_three_paragraph_ids)
                sanitize_exclude_paragraph_ids = sanitize_ids

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
                        exclude_paragraph_ids=outline_body_level_exclude_paragraph_ids,
                    )
                    restore_outline_levels_for_protected_paragraphs(
                        root,
                        indent_skip_paragraph_ids,
                        stop=stop,
                        numbering_level_lookup=numbering_level_lookup,
                        style_numbering_lookup=style_numbering_lookup,
                        change_logs=summary.paragraph_logs,
                        part_name=item.filename,
                    )
                elif (
                    item.filename == "word/numbering.xml"
                    and options.skip_chapter_three_indents
                    and protected_context.chapter_three_abstract_ids
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
                    if options.skip_chapter_three_indents and protected_context.chapter_three_abstract_ids:
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
                    excluded_style_ids=indent_excluded_style_ids,
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
                        normalize_body_style_to_none=options.normalize_body_style_to_none,
                        skip_paragraph_ids=indent_skip_paragraph_ids,
                    )
                    summary.paragraphs += changed_paragraphs

                if options.fix_table_layout or options.fix_color or options.move_table_notes_below:
                    global_table_index = process_tables_in_part(
                        root=root,
                        part_name=item.filename,
                        options=options,
                        stop=stop,
                        summary=summary,
                        global_table_index=global_table_index,
                        numbering_level_lookup=numbering_level_lookup,
                        numbering_format_lookup=numbering_format_lookup,
                        style_numbering_lookup=style_numbering_lookup,
                        protected_context=protected_context,
                        progress_callback=progress_callback,
                        item_index=item_index,
                        total_items=total_items,
                    )

                data = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

            if should_sanitize_indent_unit_part(item.filename):
                if root is None:
                    root = etree.fromstring(data, parser)
                if item.filename == "word/styles.xml":
                    removed_char_indent_attrs = remove_character_indent_attrs_from_styles_root_excluding_protected(
                        root,
                        excluded_style_ids=indent_excluded_style_ids,
                        change_logs=summary.numbering_xml_logs,
                    )
                elif item.filename == "word/numbering.xml":
                    removed_char_indent_attrs = remove_character_indent_attrs_from_numbering_root_excluding_protected(
                        root,
                        indent_excluded_numbering_pairs,
                        indent_excluded_num_ids,
                        indent_excluded_abstract_ids,
                    )
                else:
                    removed_char_indent_attrs = remove_character_indent_attrs_from_root(
                        root,
                        exclude_paragraph_ids=sanitize_exclude_paragraph_ids,
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

    if options.write_note_debug_log:
        _write_note_debug_log_safely(
            output_docx,
            summary,
            "after_xml_pipeline_before_word_com",
            append=False,
        )

    if options.normalize_with_word_com:
        table_autofit_records = list(summary.word_com_table_autofit_records)
        if table_autofit_records:
            requested_indices = {
                int(record.get("global_table_index", 0)) for record in table_autofit_records
            }
            try:
                table_autofit_logs, applied_indices, failed_indices = (
                    apply_table_autofit_with_word_com(
                        output_docx,
                        table_autofit_records,
                        stop=stop,
                    )
                )
            except ProcessStopped:
                raise
            except Exception as exc:
                table_autofit_logs = [
                    "WORD_COM_TABLE_AUTOFIT_SKIPPED "
                    f"reason=runner_failed:{type(exc).__name__}:{exc}"
                ]
                applied_indices = set()
                failed_indices = set(requested_indices)
            summary.word_com_table_autofit_logs.extend(table_autofit_logs)
            _mark_word_com_table_autofit_applied(summary, applied_indices)
            if failed_indices:
                failed_records = [
                    record
                    for record in table_autofit_records
                    if int(record.get("global_table_index", 0)) in failed_indices
                ]
                fallback_logs: list[str] = []
                fallback_applied_indices = fallback_normal_table_autofit_in_docx(
                    output_docx,
                    failed_records,
                    fallback_logs,
                    stop=stop,
                )
                summary.word_com_table_autofit_logs.extend(fallback_logs)
                _mark_word_com_table_autofit_fallback(
                    summary,
                    fallback_applied_indices,
                    failed_indices,
                )
        else:
            summary.word_com_table_autofit_logs.append(
                "WORD_COM_TABLE_AUTOFIT_SKIPPED reason=no_records"
            )
    else:
        summary.word_com_table_autofit_logs.append("WORD_COM_TABLE_AUTOFIT_SKIPPED reason=disabled")

    if options.normalize_with_word_com:
        word_com_records = word_com_indent.filter_word_com_body_indent_records(summary.body_indent_records)
        summary.word_com_body_indent_logs.append(
            "WORD_COM_BODY_INDENT_RECORD_FILTER "
            f"total_records={len(summary.body_indent_records)} "
            f"word_com_records={len(word_com_records)} "
            f"skipped_records={len(summary.body_indent_records) - len(word_com_records)} "
            "criteria=apply_only_if_word_font_size_is_14_and_xml_font_size_gt_11"
        )
        if word_com_records:
            summary.word_com_body_indent_logs.extend(
                word_com_indent.verify_and_fix_body_indents_with_word_com(
                    output_docx,
                    word_com_records,
                    stop=stop,
                )
            )
        else:
            summary.word_com_body_indent_logs.append(
                "WORD_COM_BODY_INDENT_FIX_SKIPPED reason=no_font_check_records"
            )
    else:
        summary.word_com_body_indent_logs.append("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=disabled")

    # 「表格最後一列說明格式化」 is the final table format step: it runs after Word
    # COM AutoFit and the XML fallback (both of which can reset table fonts /
    # alignment) so the last-row footer cells keep 10 pt + left/right alignment
    # + the nil/double borders. Only the tables recorded during the XML pipeline
    # are re-located and formatted, so unrelated tables are never touched.
    if options.enable_table_footer_source_format and summary.table_footer_source_format_records:
        if progress_callback:
            progress_callback(percent=99, message="reapplying table footer formatting")
        footer_results = apply_table_footer_source_format_in_docx(
            output_docx,
            summary.table_footer_source_format_records,
            logs=summary.table_footer_source_format_logs,
            stop=stop,
        )
        _mark_table_footer_source_format_applied(summary, footer_results)
    elif options.enable_table_footer_source_format:
        summary.table_footer_source_format_logs.append(
            "FOOTER_SOURCE_FORMAT_REAPPLY_SKIPPED reason=no_eligible_tables"
        )
    else:
        summary.table_footer_source_format_logs.append(
            "FOOTER_SOURCE_FORMAT_REAPPLY_SKIPPED reason=disabled"
        )

    if progress_callback:
        progress_callback(percent=99, message="word/numbering.xml: final suffix cleanup")
    force_clean_numbering_suffix_tabs_in_docx(
        output_docx,
        logs=summary.numbering_xml_logs,
        excluded_numbering_pairs=final_suffix_excluded_numbering_pairs,
        excluded_num_ids=final_suffix_excluded_num_ids,
        excluded_abstract_ids=final_suffix_excluded_abstract_ids,
        included_numbering_pairs=final_included_numbering_pairs,
        included_num_ids=final_included_num_ids,
        included_abstract_ids=final_included_abstract_ids,
        protected_numbering_pairs=final_protected_numbering_pairs,
    )
    if options.force_note_paragraph_left_alignment:
        force_note_paragraph_left_alignment_in_docx(
            output_docx,
            logs=summary.paragraph_logs,
        )
    else:
        summary.paragraph_logs.append(
            "FINAL_NOTE_ALIGNMENT_FIX_SKIPPED reason=disabled"
        )
    if options.write_note_debug_log:
        _write_note_debug_log_safely(
            output_docx,
            summary,
            "after_final_output",
            append=True,
        )
    else:
        summary.paragraph_logs.append("NOTE_DEBUG_LOG_SKIPPED reason=disabled")

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


