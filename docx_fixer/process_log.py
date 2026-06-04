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


def _bool_text(value: object) -> str:
    return "true" if bool(value) else "false"


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
                f"cell_count: {record['cell_count']}",
                f"column_count: {record['column_count']}",
                f"table_type: {record['table_type']}",
                f"action: {record['action']}",
                f"reason: {record['reason']}",
                f"special_layout_used: {_bool_text(record['special_layout_used'])}",
                f"layout_fixed: {_bool_text(record['layout_fixed'])}",
                f"color_fixed: {_bool_text(record['color_fixed'])}",
                f"changed_to_gray: {record['changed_to_gray']}",
                f"cleared_colors: {record['cleared_colors']}",
                "",
            ]
        )

    return lines[:-1] if lines and lines[-1] == "" else lines


def write_table_log_file(output_docx: str | Path, summary: ProcessSummary) -> Path:
    log_path = get_table_log_path(output_docx)
    log_path.write_text("\n".join(format_table_log_lines(summary)) + "\n", encoding="utf-8")
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
