from __future__ import annotations

from pathlib import Path

from .indent_settings import current_indent_settings, format_cm
from .models import ProcessSummary


def get_process_log_path(output_docx: str | Path) -> Path:
    output_path = Path(output_docx)
    return output_path.with_name(f"{output_path.stem}_log.txt")


def get_table_log_path(output_docx: str | Path) -> Path:
    output_path = Path(output_docx)
    return output_path.with_name(f"{output_path.stem}_table_log.txt")


def get_heading_suffix_log_path(output_docx: str | Path) -> Path:
    output_path = Path(output_docx)
    return output_path.with_name(f"{output_path.stem}_heading_suffix_log.txt")


def _suffix_record_key(record: dict[str, object]) -> tuple[str, int]:
    return str(record.get("part_name", "")), int(record.get("paragraph_index", 0))


def _suffix_count(records: list[dict[str, object]], suffix: str) -> int:
    return sum(1 for record in records if record.get("suffix") == suffix)


def _tab_stop_count(records: list[dict[str, object]]) -> int:
    return sum(1 for record in records if record.get("has_tab_stop") is True)


def _effective_suffix_count(records: list[dict[str, object]], suffix: str) -> int:
    return sum(1 for record in records if record.get("effective_suffix", record.get("suffix")) == suffix)


def _lvl_text_trailing_space_count(records: list[dict[str, object]]) -> int:
    return sum(1 for record in records if record.get("lvlText_has_trailing_space") is True)


def _change_type(before_record: dict[str, object], after_record: dict[str, object]) -> str:
    before_suffix = before_record.get("suffix")
    after_suffix = after_record.get("suffix")
    before_effective = before_record.get("effective_suffix", before_suffix)
    after_effective = after_record.get("effective_suffix", after_suffix)
    if before_suffix == "missing" and before_effective != before_suffix:
        return f"missing_effective_{before_effective}_to_{after_effective}"
    if before_effective != after_effective:
        return f"{before_effective}_to_{after_effective}"
    if before_suffix != after_suffix:
        return f"{before_suffix}_to_{after_suffix}"
    return f"{before_suffix}_to_{after_suffix}"


def _format_optional_cm(value: object) -> str:
    if value is None:
        return "None"
    return f"{float(value):.2f}"


def _format_suffix_record_value(record: dict[str, object], key: str) -> str:
    value = record.get(key)
    if value is None:
        return "None"
    if isinstance(value, bool):
        return _bool_text(value)
    return str(value)


def format_heading_suffix_log_lines(summary: ProcessSummary) -> list[str]:
    before = list(summary.heading_suffix_before_records)
    after = list(summary.heading_suffix_after_records)
    before_by_key = {_suffix_record_key(record): record for record in before}
    after_by_key = {_suffix_record_key(record): record for record in after}
    keys = sorted(set(before_by_key) | set(after_by_key), key=lambda item: (item[0], item[1]))

    before_manual = sum(1 for record in before if record.get("source") == "manual_text")
    before_auto = sum(1 for record in before if record.get("source") == "auto_numbering_xml")
    after_manual = sum(1 for record in after if record.get("source") == "manual_text")
    after_auto = sum(1 for record in after if record.get("source") == "auto_numbering_xml")

    changed_to_nothing = 0
    still_tab = 0
    still_space = 0
    still_missing = 0
    still_other = 0
    after_raw_suffix_missing_count = _suffix_count(after, "missing")
    after_effective_suffix_tab_count = _effective_suffix_count(after, "tab")
    after_suffix_space_count = _suffix_count(after, "space")
    after_tab_stop_remaining_count = _tab_stop_count(after)
    after_lvl_text_trailing_space_count = _lvl_text_trailing_space_count(after)
    for key in keys:
        before_record = before_by_key.get(key, {})
        after_record = after_by_key.get(key, {})
        before_suffix = before_record.get("suffix")
        after_suffix = after_record.get("suffix")
        if before_suffix != "nothing" and after_suffix == "nothing":
            changed_to_nothing += 1
        if after_suffix == "tab":
            still_tab += 1
        elif after_suffix == "space":
            still_space += 1
        elif after_suffix == "missing":
            still_missing += 1
        elif after_suffix == "other":
            still_other += 1

    lines = [
        "HEADING_SUFFIX_LOG",
        "",
        "===== SUMMARY BEFORE_FIX =====",
        f"total_headings: {len(before)}",
        f"manual_headings: {before_manual}",
        f"auto_numbering_headings: {before_auto}",
        f"suffix_nothing: {_suffix_count(before, 'nothing')}",
        f"suffix_tab: {_suffix_count(before, 'tab')}",
        f"suffix_space: {_suffix_count(before, 'space')}",
        f"suffix_missing: {_suffix_count(before, 'missing')}",
        f"suffix_other: {_suffix_count(before, 'other')}",
        f"tab_stop_remaining: {_tab_stop_count(before)}",
        "",
        "===== SUMMARY AFTER_FIX =====",
        f"total_headings: {len(after)}",
        f"manual_headings: {after_manual}",
        f"auto_numbering_headings: {after_auto}",
        f"suffix_nothing: {_suffix_count(after, 'nothing')}",
        f"suffix_tab: {_suffix_count(after, 'tab')}",
        f"suffix_space: {_suffix_count(after, 'space')}",
        f"suffix_missing: {_suffix_count(after, 'missing')}",
        f"suffix_other: {_suffix_count(after, 'other')}",
        f"tab_stop_remaining: {_tab_stop_count(after)}",
        f"after_raw_suffix_missing_count: {after_raw_suffix_missing_count}",
        f"after_effective_suffix_tab_count: {after_effective_suffix_tab_count}",
        f"after_suffix_space_count: {after_suffix_space_count}",
        f"after_tab_stop_remaining_count: {after_tab_stop_remaining_count}",
        f"after_lvlText_trailing_space_count: {after_lvl_text_trailing_space_count}",
        "",
        "===== SUMMARY CHANGE =====",
        f"changed_to_nothing: {changed_to_nothing}",
        f"still_tab: {still_tab}",
        f"still_space: {still_space}",
        f"still_missing: {still_missing}",
        f"still_other: {still_other}",
        "",
    ]
    if any(
        count
        for count in (
            after_raw_suffix_missing_count,
            after_effective_suffix_tab_count,
            _suffix_count(after, "tab"),
            after_suffix_space_count,
            after_tab_stop_remaining_count,
            after_lvl_text_trailing_space_count,
        )
    ):
        lines.extend(
            [
                "WARNING: AFTER_FIX numbering suffix/tab cleanup still has remaining issues.",
                f"WARNING raw_suffix_after_missing={after_raw_suffix_missing_count}",
                f"WARNING effective_suffix_after_tab={after_effective_suffix_tab_count}",
                f"WARNING suffix_after_tab={_suffix_count(after, 'tab')}",
                f"WARNING suffix_after_space={after_suffix_space_count}",
                f"WARNING has_tab_stop_after_true={after_tab_stop_remaining_count}",
                f"WARNING lvlText_after_has_trailing_space={after_lvl_text_trailing_space_count}",
                "",
            ]
        )

    if not keys:
        lines.append("No heading suffix records.")
        return lines

    for index, key in enumerate(keys, start=1):
        before_record = before_by_key.get(key, {})
        after_record = after_by_key.get(key, {})
        record = after_record or before_record
        matched = bool(before_record and after_record)
        if matched:
            match_status = "matched"
        elif before_record:
            match_status = "before_only"
        else:
            match_status = "after_only"

        source = record.get("source", "unknown")
        changed = matched and (
            before_record.get("suffix") != after_record.get("suffix")
            or before_record.get("effective_suffix", before_record.get("suffix"))
            != after_record.get("effective_suffix", after_record.get("suffix"))
            or before_record.get("has_tab_stop") != after_record.get("has_tab_stop")
        )
        lines.extend(
            [
                f"===== Heading {index} =====",
                f"part_name: {key[0]}",
                f"paragraph_index: {key[1]}",
                f"matched: {_bool_text(matched)}",
                f"match_status: {match_status}",
                f"source_before: {_format_suffix_record_value(before_record, 'source')}",
                f"source_after: {_format_suffix_record_value(after_record, 'source')}",
                f"outline_level_before: {_format_suffix_record_value(before_record, 'outline_level')}",
                f"outline_level_after: {_format_suffix_record_value(after_record, 'outline_level')}",
                f"heading_text_before: {_format_suffix_record_value(before_record, 'heading_text')}",
                f"heading_text_after: {_format_suffix_record_value(after_record, 'heading_text')}",
                f"number_token_before: {_format_suffix_record_value(before_record, 'number_token')}",
                f"number_token_after: {_format_suffix_record_value(after_record, 'number_token')}",
                f"suffix_before: {_format_suffix_record_value(before_record, 'suffix')}",
                f"suffix_after: {_format_suffix_record_value(after_record, 'suffix')}",
                f"paragraph_has_numPr_before: {_format_suffix_record_value(before_record, 'paragraph_has_numPr')}",
                f"paragraph_has_numPr_after: {_format_suffix_record_value(after_record, 'paragraph_has_numPr')}",
                f"paragraph_tabs_before: {_format_suffix_record_value(before_record, 'paragraph_tabs')}",
                f"paragraph_tabs_after: {_format_suffix_record_value(after_record, 'paragraph_tabs')}",
                f"numId_before: {_format_suffix_record_value(before_record, 'numId')}",
                f"numId_after: {_format_suffix_record_value(after_record, 'numId')}",
                f"ilvl_before: {_format_suffix_record_value(before_record, 'ilvl')}",
                f"ilvl_after: {_format_suffix_record_value(after_record, 'ilvl')}",
                f"numbering_suff_before: {_format_suffix_record_value(before_record, 'numbering_suff')}",
                f"numbering_suff_after: {_format_suffix_record_value(after_record, 'numbering_suff')}",
                f"numbering_tab_pos_before: {_format_suffix_record_value(before_record, 'numbering_tab_pos')}",
                f"numbering_tab_pos_after: {_format_suffix_record_value(after_record, 'numbering_tab_pos')}",
                f"style_numPr_before: {_format_suffix_record_value(before_record, 'style_numPr')}",
                f"style_numPr_after: {_format_suffix_record_value(after_record, 'style_numPr')}",
                f"style_tabs_before: {_format_suffix_record_value(before_record, 'style_tabs')}",
                f"style_tabs_after: {_format_suffix_record_value(after_record, 'style_tabs')}",
            ]
        )

        if source == "auto_numbering_xml":
            for key_name in ("raw_suffix", "effective_suffix"):
                lines.append(
                    f"{key_name}_before: {_format_suffix_record_value(before_record, key_name)}"
                )
                lines.append(
                    f"{key_name}_after: {_format_suffix_record_value(after_record, key_name)}"
                )
            for key_name in ("numId", "ilvl", "numFmt", "lvlText", "lvlText_has_trailing_space"):
                lines.append(
                    f"{key_name}_before: {_format_suffix_record_value(before_record, key_name)}"
                )
                lines.append(
                    f"{key_name}_after: {_format_suffix_record_value(after_record, key_name)}"
                )
            for key_name in ("has_tab_stop", "tab_pos_twips", "left_twips", "hanging_twips", "number_start_twips"):
                lines.append(
                    f"{key_name}_before: {_format_suffix_record_value(before_record, key_name)}"
                )
                lines.append(
                    f"{key_name}_after: {_format_suffix_record_value(after_record, key_name)}"
                )
            for key_name in ("tab_pos_cm", "left_cm", "hanging_cm", "number_start_cm"):
                lines.append(f"{key_name}_before: {_format_optional_cm(before_record.get(key_name))}")
                lines.append(f"{key_name}_after: {_format_optional_cm(after_record.get(key_name))}")
        else:
            for key_name in ("space_count", "tab_count", "raw_separator_repr"):
                lines.append(
                    f"{key_name}_before: {_format_suffix_record_value(before_record, key_name)}"
                )
                lines.append(
                    f"{key_name}_after: {_format_suffix_record_value(after_record, key_name)}"
                )

        lines.extend(
            [
                f"changed: {_bool_text(changed)}",
                f"change_type: {_change_type(before_record, after_record) if matched else match_status}",
                "",
            ]
        )

    return lines[:-1] if lines and lines[-1] == "" else lines


def format_numbering_indent_log_lines(summary: ProcessSummary) -> list[str]:
    lines = ["編號縮排量測紀錄："]
    if not summary.numbering_measurements:
        lines.append("沒有編號縮排量測資料。")
        return lines

    records = sorted(
        summary.numbering_measurements.values(),
        key=lambda item: (
            str(item.get("section", "")),
            int(item.get("level", 0)),
            str(item.get("prefix", "")),
        ),
    )
    current_section = None
    for record in records:
        section = str(record["section"])
        if section != current_section:
            lines.append(f"{section}：")
            current_section = section

        lines.append(
            f"第 {int(record['level']) + 1} 階 {record['prefix']}："
            f"文字起點 {format_cm(float(record['text_start_cm']))} cm，"
            f"編號起點 {format_cm(float(record['number_start_cm']))} cm，"
            f"編號寬度 {format_cm(float(record['number_size_cm']))} cm，"
            f"字型 {record['font_name']} {format_cm(float(record['font_size_pt']))} pt，"
            f"樣本數 {record['count']}"
        )

    return lines


def format_numbering_debug_log_lines(summary: ProcessSummary) -> list[str]:
    lines = ["Numbering XML debug:"]
    if summary.numbering_xml_logs:
        lines.extend(summary.numbering_xml_logs)
    if not summary.numbering_debug_logs:
        if not summary.numbering_xml_logs:
            lines.append("No numbering XML debug records.")
        return lines
    lines.extend(summary.numbering_debug_logs)
    return lines


def format_body_indent_debug_log_lines(summary: ProcessSummary) -> list[str]:
    lines = ["Body indent debug:"]
    if not summary.body_indent_debug_logs:
        lines.append("No body indent debug records.")
        return lines
    lines.extend(summary.body_indent_debug_logs)
    return lines


def format_indent_settings_log_lines() -> list[str]:
    lines = [
        "Indent settings snapshot:",
        "Note: local indent_defaults.json overrides constants.py built-in indent defaults. Delete it or use the GUI restore button to apply the latest built-ins.",
    ]
    settings = current_indent_settings()
    for row in settings["body"]:
        level = int(row["level"])
        left_cm = float(row["number_start_cm"]) + float(row["hanging_cm"])
        body_left_twips = round(float(row["body_left_cm"]) * 20 * 28.3464567)
        lines.append(
            f"level={level}; "
            f"number_start_cm={format_cm(float(row['number_start_cm']))}; "
            f"hanging_cm={format_cm(float(row['hanging_cm']))}; "
            f"left_cm={format_cm(left_cm)}; "
            f"body_left_cm={format_cm(float(row['body_left_cm']))}; "
            f"body_left_twips={body_left_twips}"
        )
    return lines


def format_word_com_body_indent_log_lines(summary: ProcessSummary) -> list[str]:
    lines = ["Word COM body indent fix:"]
    if not summary.word_com_body_indent_logs:
        lines.append("No Word COM body indent logs.")
        return lines
    lines.extend(summary.word_com_body_indent_logs)
    return lines


def format_word_com_table_autofit_log_lines(summary: ProcessSummary) -> list[str]:
    lines = ["Word COM table AutoFit:"]
    if not summary.word_com_table_autofit_logs:
        lines.append("No Word COM table AutoFit logs.")
        return lines
    lines.extend(summary.word_com_table_autofit_logs)
    return lines


def format_table_footer_source_format_log_lines(summary: ProcessSummary) -> list[str]:
    lines = ["Table footer (最後一列說明) formatting (final post-process):"]
    if not summary.table_footer_source_format_logs:
        lines.append("No table footer formatting logs.")
        return lines
    lines.extend(summary.table_footer_source_format_logs)
    return lines


def _bool_text(value: object) -> str:
    return "true" if bool(value) else "false"


def _optional_int_text(value: object) -> str:
    if value is None:
        return "none"
    return str(value)


def _color_list_text(value: object) -> str:
    if not value:
        return "none"
    return ",".join(str(item) for item in value)


def _moved_notes_text(value: object) -> str:
    if not value:
        return "none"
    parts = []
    for note in value:
        if not isinstance(note, dict):
            continue
        parts.append(
            f"note_text={note.get('note_text', '')};"
            f"delete_action={note.get('delete_action', '')};"
            f"row_index={note.get('row_index', '')};"
            f"cell_index={note.get('cell_index', '')}"
        )
    return " | ".join(parts) if parts else "none"


def format_table_log_lines(summary: ProcessSummary) -> list[str]:
    lines = ["表格處理紀錄："]
    if not summary.table_log_records:
        lines.append("沒有表格紀錄。")
        return lines

    for index, record in enumerate(summary.table_log_records, start=1):
        lines.extend(
            [
                f"===== Table {index} =====",
                f"part_name: {record['part_name']}",
                f"table_index: {record['table_index']}",
                f"global_table_index: {record['global_table_index']}",
                f"table_name: {record['table_name']}",
                f"first_level_heading: {record.get('first_level_heading', '(none)')}",
                f"cell_count: {record['cell_count']}",
                f"column_count: {record['column_count']}",
                f"table_type: {record['table_type']}",
                f"action: {record['action']}",
                f"reason: {record['reason']}",
                f"special_layout_used: {_bool_text(record['special_layout_used'])}",
                f"layout_fixed: {_bool_text(record['layout_fixed'])}",
                f"color_fixed: {_bool_text(record['color_fixed'])}",
                f"chapter_three_table_layout_skipped: {_bool_text(record.get('chapter_three_table_layout_skipped', False))}",
                f"chapter_three_table_color_skipped: {_bool_text(record.get('chapter_three_table_color_skipped', False))}",
                f"word_com_autofit_applied: {_bool_text(record.get('word_com_autofit_applied', False))}",
                f"word_com_autofit_sequence: {record.get('word_com_autofit_sequence', 'none')}",
                f"word_com_autofit_fallback_applied: {_bool_text(record.get('word_com_autofit_fallback_applied', False))}",
                f"word_com_autofit_status: {record.get('word_com_autofit_status', 'not_needed')}",
                f"special_left_indent_twips: {_optional_int_text(record.get('special_left_indent_twips'))}",
                f"special_width_twips: {_optional_int_text(record.get('special_width_twips'))}",
                f"special_text_width_twips: {_optional_int_text(record.get('special_text_width_twips'))}",
                f"special_right_edge_twips: {_optional_int_text(record.get('special_right_edge_twips'))}",
                f"special_overflow_twips: {_optional_int_text(record.get('special_overflow_twips'))}",
                f"special_color_skip_matched: {_bool_text(record.get('special_color_skip_matched', False))}",
                f"special_color_skip_colors: {_color_list_text(record.get('special_color_skip_colors'))}",
                f"special_color_cleared_count: {record.get('special_color_cleared_count', 0)}",
                f"table_keep_colors: {_color_list_text(record.get('table_keep_colors'))}",
                f"table_gray_colors: {_color_list_text(record.get('table_gray_colors'))}",
                f"table_gray_target: {record.get('table_gray_target', 'D9D9D9')}",
                f"double_border_enabled: {_bool_text(record.get('double_border_enabled', False))}",
                f"double_border_applied: {_bool_text(record.get('double_border_applied', False))}",
                f"table_footer_note_source_format_enabled: {_bool_text(record.get('table_footer_note_source_format_enabled', False))}",
                f"table_footer_note_source_format_should_apply: {_bool_text(record.get('table_footer_note_source_format_should_apply', False))}",
                f"table_footer_note_source_format_applied: {_bool_text(record.get('table_footer_note_source_format_applied', False))}",
                f"outer_double_border_applied_by_footer_source_format: {_bool_text(record.get('outer_double_border_applied_by_footer_source_format', False))}",
                f"first_row_single_cell_border_adjusted: {_bool_text(record.get('first_row_single_cell_border_adjusted', False))}",
                f"footer_rows_processed: {record.get('footer_rows_processed', 0)}",
                f"footer_row_matches: {' | '.join(record.get('footer_row_matches', [])) if record.get('footer_row_matches') else 'none'}",
                f"footer_note_cells_adjusted: {record.get('footer_note_cells_adjusted', 0)}",
                f"footer_note_cell_matches: {_color_list_text(record.get('footer_note_cell_matches'))}",
                f"footer_note_cell_debug: {' | '.join(record.get('footer_note_cell_debug', [])) if record.get('footer_note_cell_debug') else 'none'}",
                f"table_footer_note_source_format_skipped_reason: {record.get('table_footer_note_source_format_skipped_reason', 'none')}",
                f"table_note_move_gui_hidden: {_bool_text(record.get('table_note_move_gui_hidden', False))}",
                f"table_note_move_forced_false: {_bool_text(record.get('table_note_move_forced_false', False))}",
                f"skip_chapter_three_table_note_move_forced_false: {_bool_text(record.get('skip_chapter_three_table_note_move_forced_false', False))}",
                f"skip_section_three_adjustments_enabled: {_bool_text(record.get('skip_section_three_adjustments_enabled', False))}",
                f"in_section_three_protected: {_bool_text(record.get('in_section_three_protected', False))}",
                f"section_three_detection_source: {record.get('section_three_detection_source', 'none')}",
                f"skipped_by_section_three_protection: {_bool_text(record.get('skipped_by_section_three_protection', False))}",
                f"move_table_notes_below_enabled: {_bool_text(record.get('move_table_notes_below_enabled', False))}",
                f"skip_chapter_three_table_notes_enabled: {_bool_text(record.get('skip_chapter_three_table_notes_enabled', False))}",
                f"table_notes_skipped_by_chapter_three: {_bool_text(record.get('table_notes_skipped_by_chapter_three', False))}",
                f"note_cells_moved: {_bool_text(record.get('note_cells_moved', False))}",
                f"moved_note_count: {record.get('moved_note_count', 0)}",
                f"deleted_note_cells: {record.get('deleted_note_cells', 0)}",
                f"deleted_note_rows: {record.get('deleted_note_rows', 0)}",
                f"inserted_note_paragraphs: {record.get('inserted_note_paragraphs', 0)}",
                f"moved_notes: {_moved_notes_text(record.get('moved_notes'))}",
                f"note_move_warnings: {' | '.join(record.get('note_move_warnings', [])) if record.get('note_move_warnings') else 'none'}",
                f"changed_to_gray: {record['changed_to_gray']}",
                f"cleared_colors: {record['cleared_colors']}",
                f"shading_debug: {' | '.join(record.get('shading_debug', [])) if record.get('shading_debug') else 'none'}",
                "",
            ]
        )

    return lines[:-1] if lines and lines[-1] == "" else lines


def write_table_log_file(output_docx: str | Path, summary: ProcessSummary) -> Path:
    log_path = get_table_log_path(output_docx)
    log_path.write_text("\n".join(format_table_log_lines(summary)) + "\n", encoding="utf-8")
    return log_path


def write_heading_suffix_log_file(output_docx: str | Path, summary: ProcessSummary) -> Path:
    log_path = get_heading_suffix_log_path(output_docx)
    log_path.write_text("\n".join(format_heading_suffix_log_lines(summary)) + "\n", encoding="utf-8")
    return log_path


def write_process_log(output_docx: str | Path, summary: ProcessSummary) -> Path:
    log_path = get_process_log_path(output_docx)
    lines = [
        "Word DOCX 快速整理工具處理紀錄",
        f"輸出檔案：{Path(output_docx)}",
        "",
        "表格摘要：",
        f"表格總數：{summary.tables}",
        f"跳過第一張表格數：{summary.skipped_first_page_tables}",
        f"因格子數小於等於 4 而跳過的表格數：{summary.skipped_small_tables}",
        f"因表格中有表格而跳過的表格數：{summary.skipped_nested_tables}",
        f"因特殊顏色而跳過的表格數：{summary.special_color_skipped_tables}",
        f"因「參、不要調整」而保護跳過的表格數：{summary.section_three_protected_tables}",
        f"套用黑色雙線外框的表格數：{summary.double_border_tables}",
        f"套用最後一列說明格式化的表格數：{summary.table_footer_source_format_tables}",
        f"搬移表格內註記的表格數：{summary.note_cells_moved_tables}",
        f"因「參、不要表格註記搬移」而未搬移的表格數：{summary.note_move_skipped_by_chapter_three_tables}",
        f"搬移註記筆數：{summary.moved_note_count}",
        f"刪除註記儲存格數：{summary.deleted_note_cells}",
        f"刪除註記整列數：{summary.deleted_note_rows}",
        f"表格下方新增註記段落數：{summary.inserted_note_paragraphs}",
        f"跨頁表格數：{summary.cross_page_tables}",
        f"跨頁已解決的表格數：{summary.cross_page_resolved_tables}",
        f"跨頁未解決的表格數：{summary.cross_page_still_split_tables}",
        f"調整儲存格 padding 的表格數：{summary.adjusted_cell_padding_tables}",
        f"調整表格段落間距的表格數：{summary.adjusted_table_spacing_tables}",
        f"改成自動列高的表格數：{summary.auto_height_tables}",
        f"移到下一頁後解決跨頁的表格數：{summary.moved_next_page_resolved_tables}",
        f"無法避免跨頁的表格數：{summary.cannot_avoid_cross_page_tables}",
        f"跨頁處理失敗的表格數：{summary.failed_cross_page_tables}",
        f"套用特殊版面表格數：{summary.special_autofit_right_tables}",
        f"一般表格處理數：{summary.normal_processed_tables}",
        f"Word COM AutoFit 成功表格數：{summary.word_com_table_autofit_applied_count}",
        f"Word COM 失敗改用 XML fallback 修復表格數：{summary.word_com_table_autofit_fallback_count}",
        f"Word COM 與 XML fallback 都失敗表格數：{summary.word_com_table_autofit_failed_count}",
        "",
        "段落摘要：",
        f"段落總數：{summary.total_paragraphs}",
        f"跳過目錄段落數：{summary.skipped_toc_paragraphs}",
        f"跳過表格段落數：{summary.skipped_table_paragraphs}",
        f"移除大綱層級的段落數：{summary.removed_all_outline_paragraphs}",
        f"套用前言縮排的段落數：{summary.indented_preface_paragraphs}",
        f"套用前言大綱的段落數：{summary.outlined_preface_paragraphs}",
        f"移除字元縮排屬性數：{summary.character_indent_attrs_removed}",
        *[
            f"第 {level} 階段落數：{count}"
            for level, count in enumerate(summary.paragraph_level_counts, start=1)
        ],
        f"未分類段落數：{summary.unknown_paragraphs}",
        "",
        *format_indent_settings_log_lines(),
        "",
        *format_numbering_indent_log_lines(summary),
        "",
        *format_numbering_debug_log_lines(summary),
        "",
        *format_body_indent_debug_log_lines(summary),
        "",
        *format_word_com_table_autofit_log_lines(summary),
        "",
        *format_table_footer_source_format_log_lines(summary),
        "",
        *format_word_com_body_indent_log_lines(summary),
        "",
        "段落變更紀錄：",
    ]

    if summary.paragraph_logs:
        lines.extend(summary.paragraph_logs)
    else:
        lines.append("沒有段落變更紀錄。")

    log_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return log_path
