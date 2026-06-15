from __future__ import annotations

from dataclasses import replace

from .constants import NS, TEMPLATE_OUTLINE_INDENTS
from .numbering import has_auto_numbering, paragraph_style_id
from .outline import detect_manual_numbering_prefix
from .protected_region import (
    _effective_paragraph_numbering_identity,
    _outline_level_from_identity,
    find_table_first_level_heading,
)
from .table_format import (
    apply_double_black_table_borders,
    clear_matching_special_colors,
    process_table,
    table_cell_count,
    table_column_count,
    table_has_special_skip_color,
)
from .table_notes import move_table_note_cells_below
from .xml_utils import paragraph_text, qn


def _note_log_fields(enabled: bool, result) -> dict[str, object]:
    if result is None:
        return {
            "move_table_notes_below_enabled": enabled,
            "note_cells_moved": False,
            "moved_note_count": 0,
            "deleted_note_cells": 0,
            "deleted_note_rows": 0,
            "inserted_note_paragraphs": 0,
            "moved_notes": [],
            "note_move_warnings": [],
        }
    return {
        "move_table_notes_below_enabled": enabled,
        "note_cells_moved": result.note_cells_moved,
        "moved_note_count": result.moved_note_count,
        "deleted_note_cells": result.deleted_note_cells,
        "deleted_note_rows": result.deleted_note_rows,
        "inserted_note_paragraphs": result.inserted_note_paragraphs,
        "moved_notes": [
            {
                "note_text": note.note_text,
                "delete_action": note.delete_action,
                "row_index": note.row_index,
                "cell_index": note.cell_index,
            }
            for note in result.moved_notes
        ],
        "note_move_warnings": list(result.warnings),
    }


def _section_three_fields(
    *,
    enabled: bool,
    source: str,
    in_protected: bool,
    skipped_by_protection: bool,
) -> dict[str, object]:
    return {
        "skip_section_three_adjustments_enabled": enabled,
        "section_three_detection_source": source,
        "in_section_three_protected": in_protected,
        "skipped_by_section_three_protection": skipped_by_protection,
    }

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


def is_nested_table(tbl) -> bool:
    return bool(tbl.xpath("ancestor::w:tbl", namespaces=NS))


def contains_nested_table(tbl) -> bool:
    return bool(tbl.xpath(".//w:tc//w:tbl", namespaces=NS))





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
    chapter_three_table_layout_skipped: bool = False,
    chapter_three_table_color_skipped: bool = False,
    word_com_autofit_applied: bool = False,
    word_com_autofit_sequence: str = "none",
    word_com_autofit_fallback_applied: bool = False,
    word_com_autofit_status: str = "not_needed",
    shading_debug: list[str] | None = None,
    special_left_indent_twips: int | None = None,
    special_width_twips: int | None = None,
    special_text_width_twips: int | None = None,
    special_color_skip_matched: bool = False,
    special_color_skip_colors: list[str] | None = None,
    special_color_cleared_count: int = 0,
    table_keep_colors: list[str] | None = None,
    table_gray_colors: list[str] | None = None,
    table_gray_target: str = "D9D9D9",
    move_table_notes_below_enabled: bool = False,
    note_cells_moved: bool = False,
    moved_note_count: int = 0,
    deleted_note_cells: int = 0,
    deleted_note_rows: int = 0,
    inserted_note_paragraphs: int = 0,
    moved_notes: list[dict[str, object]] | None = None,
    note_move_warnings: list[str] | None = None,
    double_border_applied: bool = False,
    skip_section_three_adjustments_enabled: bool = False,
    in_section_three_protected: bool = False,
    section_three_detection_source: str = "none",
    skipped_by_section_three_protection: bool = False,
) -> dict[str, object]:
    special_right_edge_twips: int | None = None
    special_overflow_twips: int | None = None
    if special_left_indent_twips is not None and special_width_twips is not None:
        special_right_edge_twips = special_left_indent_twips + special_width_twips
        if special_text_width_twips is not None:
            special_overflow_twips = max(0, special_right_edge_twips - special_text_width_twips)
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
        "chapter_three_table_layout_skipped": chapter_three_table_layout_skipped,
        "chapter_three_table_color_skipped": chapter_three_table_color_skipped,
        "word_com_autofit_applied": word_com_autofit_applied,
        "word_com_autofit_sequence": word_com_autofit_sequence,
        "word_com_autofit_fallback_applied": word_com_autofit_fallback_applied,
        "word_com_autofit_status": word_com_autofit_status,
        "special_left_indent_twips": special_left_indent_twips,
        "special_width_twips": special_width_twips,
        "special_text_width_twips": special_text_width_twips,
        "special_right_edge_twips": special_right_edge_twips,
        "special_overflow_twips": special_overflow_twips,
        "special_color_skip_matched": special_color_skip_matched,
        "special_color_skip_colors": list(special_color_skip_colors or []),
        "special_color_cleared_count": special_color_cleared_count,
        "table_keep_colors": list(table_keep_colors or []),
        "table_gray_colors": list(table_gray_colors or []),
        "table_gray_target": table_gray_target,
        "move_table_notes_below_enabled": move_table_notes_below_enabled,
        "note_cells_moved": note_cells_moved,
        "moved_note_count": moved_note_count,
        "deleted_note_cells": deleted_note_cells,
        "deleted_note_rows": deleted_note_rows,
        "inserted_note_paragraphs": inserted_note_paragraphs,
        "moved_notes": list(moved_notes or []),
        "note_move_warnings": list(note_move_warnings or []),
        "double_border_applied": double_border_applied,
        "skip_section_three_adjustments_enabled": skip_section_three_adjustments_enabled,
        "in_section_three_protected": in_section_three_protected,
        "section_three_detection_source": section_three_detection_source,
        "skipped_by_section_three_protection": skipped_by_section_three_protection,
        "shading_debug": list(shading_debug or []),
    }


def _chapter_three_table_skip_reason(*, layout_skipped: bool, color_skipped: bool) -> str:
    if layout_skipped and color_skipped:
        return "chapter three protected table; layout and color skipped"
    if layout_skipped:
        return "chapter three protected table; layout skipped; color allowed"
    if color_skipped:
        return "chapter three protected table; layout allowed; color skipped"
    return "chapter three protected table"

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
    # OOXML places a section's sectPr at the end of that section, so the
    # sectPr governing this table is the first one after the table. Using a
    # preceding sectPr would pick up the previous section's page geometry in
    # multi-section documents.
    following_sect_pr = tbl.xpath("following::w:sectPr", namespaces=NS)
    if following_sect_pr:
        return following_sect_pr[0]

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
) -> tuple[int, int, int] | None:
    anchor_paragraph = _find_previous_effective_paragraph(tbl, style_numbering_lookup)
    if anchor_paragraph is None:
        return None

    left_indent_twips = _paragraph_text_start_twips(
        anchor_paragraph,
        numbering_level_lookup,
        style_numbering_lookup,
    )
    if left_indent_twips is None:
        return None

    text_width_twips = _page_text_width_twips(_find_table_section_properties(tbl))
    if text_width_twips is None:
        return None

    left_indent_twips = max(left_indent_twips, 0)
    if left_indent_twips >= text_width_twips:
        return None

    width_twips = text_width_twips - left_indent_twips
    if width_twips <= 0:
        return None

    return left_indent_twips, width_twips, text_width_twips


def process_tables_in_part(
    *,
    root,
    part_name: str,
    options,
    stop,
    summary,
    global_table_index: int,
    numbering_level_lookup,
    numbering_format_lookup,
    style_numbering_lookup,
    protected_context,
    progress_callback=None,
    item_index: int = 0,
    total_items: int = 1,
) -> int:
    tables = root.xpath(".//w:tbl", namespaces=NS)
    table_count = len(tables)

    color_settings_fields = {
        "table_keep_colors": list(getattr(options, "table_keep_colors", ()) or ()),
        "table_gray_colors": list(getattr(options, "table_gray_colors", ()) or ()),
        "table_gray_target": str(getattr(options, "table_gray_target", "D9D9D9") or "D9D9D9"),
    }
    special_color_skip_colors = tuple(getattr(options, "special_color_skip_colors", ()) or ())
    skip_special_color_tables = bool(getattr(options, "skip_special_color_tables", False))
    clear_special_colors_after_skip = bool(
        getattr(options, "clear_special_colors_after_skip", False)
    )
    move_notes_enabled = bool(getattr(options, "move_table_notes_below", False))
    section_three_enabled = bool(getattr(options, "skip_chapter_three_adjustments", False))
    section_three_source = getattr(protected_context, "section_three_detection_source", "none")

    def section_fields(*, in_protected: bool, skipped_by_protection: bool) -> dict[str, object]:
        return _section_three_fields(
            enabled=section_three_enabled,
            source=section_three_source,
            in_protected=in_protected,
            skipped_by_protection=skipped_by_protection,
        )

    for table_index, tbl in enumerate(tables, start=1):
        if stop:
            stop.check()

        global_table_index += 1
        cell_count = table_cell_count(tbl)
        column_count = table_column_count(tbl)
        table_name = find_previous_paragraph_text_for_table(root, tbl)
        first_level_heading = "(none)"
        if part_name == "word/document.xml":
            first_level_heading = (
                find_table_first_level_heading(
                    tbl,
                    numbering_level_lookup,
                    numbering_format_lookup,
                    style_numbering_lookup,
                )
                or "(none)"
            )

        if part_name == "word/document.xml" and table_index == 1:
            summary.table_log_records.append(
                build_table_log_record(
                    part_name=part_name,
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
                    **_note_log_fields(move_notes_enabled, None),
                    **section_fields(in_protected=False, skipped_by_protection=False),
                    **color_settings_fields,
                )
            )
            summary.skipped_first_page_tables += 1
            continue

        if getattr(options, "skip_nested_tables", True) and (
            is_nested_table(tbl) or contains_nested_table(tbl)
        ):
            summary.skipped_nested_tables += 1
            summary.table_log_records.append(
                build_table_log_record(
                    part_name=part_name,
                    table_index=table_index,
                    global_table_index=global_table_index,
                    table_name=table_name,
                    first_level_heading=first_level_heading,
                    cell_count=cell_count,
                    column_count=column_count,
                    table_type="skipped_nested_table",
                    action="skipped",
                    reason="nested table protected; table contains or is inside another table",
                    special_layout_used=False,
                    layout_fixed=False,
                    color_fixed=False,
                    changed_to_gray=0,
                    cleared_colors=0,
                    shading_debug=[],
                    **_note_log_fields(move_notes_enabled, None),
                    **section_fields(in_protected=False, skipped_by_protection=False),
                    **color_settings_fields,
                )
            )
            continue

        is_chapter_three_table = protected_context.is_table_protected(tbl, part_name)
        chapter_three_table_layout_skipped = bool(
            is_chapter_three_table and getattr(options, "skip_chapter_three_table_layout", False)
        )
        chapter_three_table_color_skipped = bool(
            is_chapter_three_table and getattr(options, "skip_chapter_three_table_color", False)
        )
        effective_fix_table_layout = bool(options.fix_table_layout)
        effective_fix_color = bool(options.fix_color)
        if chapter_three_table_layout_skipped:
            effective_fix_table_layout = False
        if chapter_three_table_color_skipped:
            effective_fix_color = False
        effective_options = replace(
            options,
            fix_table_layout=effective_fix_table_layout,
            fix_color=effective_fix_color,
        )

        if is_chapter_three_table and not effective_fix_table_layout and not effective_fix_color:
            if section_three_enabled:
                summary.section_three_protected_tables += 1
            summary.table_log_records.append(
                build_table_log_record(
                    part_name=part_name,
                    table_index=table_index,
                    global_table_index=global_table_index,
                    table_name=table_name,
                    first_level_heading=first_level_heading,
                    cell_count=cell_count,
                    column_count=column_count,
                    table_type="skipped_chapter_three_table",
                    action="skipped",
                    reason=_chapter_three_table_skip_reason(
                        layout_skipped=chapter_three_table_layout_skipped,
                        color_skipped=chapter_three_table_color_skipped,
                    ),
                    special_layout_used=False,
                    layout_fixed=False,
                    color_fixed=False,
                    changed_to_gray=0,
                    cleared_colors=0,
                    chapter_three_table_layout_skipped=chapter_three_table_layout_skipped,
                    chapter_three_table_color_skipped=chapter_three_table_color_skipped,
                    shading_debug=[],
                    **_note_log_fields(move_notes_enabled, None),
                    **section_fields(
                        in_protected=True,
                        skipped_by_protection=section_three_enabled,
                    ),
                    **color_settings_fields,
                )
            )
            continue

        # Step 1: move in-table note cells below the table before any layout,
        # font, color, or border work. Protected 參、 tables never reach here
        # because they were fully skipped above.
        note_result = None
        if move_notes_enabled and not is_chapter_three_table:
            note_result = move_table_note_cells_below(tbl)
            if note_result.note_cells_moved:
                summary.note_cells_moved_tables += 1
                summary.moved_note_count += note_result.moved_note_count
                summary.deleted_note_cells += note_result.deleted_note_cells
                summary.deleted_note_rows += note_result.deleted_note_rows
                summary.inserted_note_paragraphs += note_result.inserted_note_paragraphs
            # Step 2: recompute counts/type after deleting cells or rows.
            cell_count = table_cell_count(tbl)
            column_count = table_column_count(tbl)

        note_fields = _note_log_fields(move_notes_enabled, note_result)

        if skip_special_color_tables and special_color_skip_colors:
            special_color_matched, matched_skip_colors = table_has_special_skip_color(
                tbl,
                special_color_skip_colors,
            )
            if special_color_matched:
                special_color_cleared_count = 0
                if clear_special_colors_after_skip:
                    special_color_cleared_count = clear_matching_special_colors(
                        tbl,
                        special_color_skip_colors,
                    )
                summary.special_color_skipped_tables += 1
                summary.table_log_records.append(
                    build_table_log_record(
                        part_name=part_name,
                        table_index=table_index,
                        global_table_index=global_table_index,
                        table_name=table_name,
                        first_level_heading=first_level_heading,
                        cell_count=cell_count,
                        column_count=column_count,
                        table_type="special_color_skipped_table",
                        action="skipped_special_color_table",
                        reason="matched special color skip list",
                        special_layout_used=False,
                        layout_fixed=False,
                        color_fixed=special_color_cleared_count > 0,
                        changed_to_gray=0,
                        cleared_colors=0,
                        special_color_skip_matched=True,
                        special_color_skip_colors=matched_skip_colors,
                        special_color_cleared_count=special_color_cleared_count,
                        shading_debug=[],
                        **note_fields,
                        **section_fields(in_protected=False, skipped_by_protection=False),
                        **color_settings_fields,
                    )
                )
                continue

        if cell_count <= 4:
            summary.skipped_small_tables += 1
            summary.table_log_records.append(
                build_table_log_record(
                    part_name=part_name,
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
                    **note_fields,
                    **section_fields(in_protected=False, skipped_by_protection=False),
                    **color_settings_fields,
                )
            )
            continue

        special_layout = effective_options.fix_table_layout and column_count <= 4
        special_table_geometry = None
        special_text_width_twips = None
        if special_layout:
            resolved_geometry = _resolve_special_table_geometry(
                tbl,
                numbering_level_lookup,
                style_numbering_lookup,
            )
            if resolved_geometry is not None:
                geometry_left, geometry_width, special_text_width_twips = resolved_geometry
                special_table_geometry = (geometry_left, geometry_width)
        changed_to_gray, cleared_colors, shading_debug = process_table(
            tbl,
            effective_options,
            stop=stop,
            special_layout=special_layout,
            special_table_geometry=special_table_geometry,
        )
        # Step 6: re-apply the black double-line border to normal and special
        # tables, after note cells/rows were removed. 參、 protected tables are
        # excluded so they keep their original frame.
        double_border_applied = False
        if not is_chapter_three_table:
            apply_double_black_table_borders(tbl)
            double_border_applied = True
            summary.double_border_tables += 1
        layout_fixed = bool(effective_options.fix_table_layout)
        color_fixed = bool(effective_options.fix_color)
        if effective_options.fix_table_layout:
            if special_layout:
                table_type = "special_table"
                action = (
                    "apply_special_table_format_and_color"
                    if effective_options.fix_color
                    else "apply_special_table_format"
                )
                reason = "column_count <= 4"
            else:
                table_type = "normal_table"
                action = (
                    "apply_normal_table_format_and_color"
                    if effective_options.fix_color
                    else "apply_normal_table_format"
                )
                reason = "column_count > 4"
        elif effective_options.fix_color:
            table_type = "color_only_table"
            action = "apply_color_only"
            reason = "fix_table_layout disabled but fix_color enabled"
        else:
            table_type = "skipped"
            action = "skipped"
            reason = "no table actions enabled"
        if is_chapter_three_table and (
            chapter_three_table_layout_skipped or chapter_three_table_color_skipped
        ):
            reason = _chapter_three_table_skip_reason(
                layout_skipped=chapter_three_table_layout_skipped,
                color_skipped=chapter_three_table_color_skipped,
            )
        if table_type == "normal_table":
            summary.word_com_table_autofit_records.append(
                {
                    "part_name": part_name,
                    "table_index": table_index,
                    "global_table_index": global_table_index,
                    "table_name": table_name,
                    "first_level_heading": first_level_heading,
                    "cell_count": cell_count,
                    "column_count": column_count,
                }
            )
        summary.changed_to_gray += changed_to_gray
        summary.cleared_colors += cleared_colors
        if special_layout:
            summary.special_autofit_right_tables += 1
        else:
            summary.normal_processed_tables += 1
        special_left_indent_twips = None
        special_width_twips = None
        if special_table_geometry is not None:
            special_left_indent_twips, special_width_twips = special_table_geometry
        summary.table_log_records.append(
            build_table_log_record(
                part_name=part_name,
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
                chapter_three_table_layout_skipped=chapter_three_table_layout_skipped,
                chapter_three_table_color_skipped=chapter_three_table_color_skipped,
                special_left_indent_twips=special_left_indent_twips,
                special_width_twips=special_width_twips,
                special_text_width_twips=special_text_width_twips,
                shading_debug=shading_debug,
                double_border_applied=double_border_applied,
                **note_fields,
                **section_fields(
                    in_protected=is_chapter_three_table,
                    skipped_by_protection=False,
                ),
                **color_settings_fields,
            )
        )

        if progress_callback and table_count:
            inner_fraction = table_index / table_count
            percent = ((item_index + inner_fraction) / total_items) * 100
            progress_callback(
                percent=percent,
                message=f"{part_name}: table {table_index}/{table_count}",
            )

    summary.tables += table_count
    return global_table_index








