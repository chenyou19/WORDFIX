from __future__ import annotations

from pathlib import Path

from .indent_settings import format_cm
from .models import ProcessSummary


def get_process_log_path(output_docx: str | Path) -> Path:
    output_path = Path(output_docx)
    return output_path.with_name(f"{output_path.stem}_log.txt")


def format_numbering_indent_log_lines(summary: ProcessSummary) -> list[str]:
    lines = ["實際編號格式量測紀錄："]
    if not summary.numbering_measurements:
        lines.append("沒有量測到可見的手動編號格式。")
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
            f"文字起點 {format_cm(float(record['text_start_cm']))} cm；"
            f"編號起點 {format_cm(float(record['number_start_cm']))} cm；"
            f"編號大小 {format_cm(float(record['number_size_cm']))} cm；"
            f"字型 {record['font_name']} {format_cm(float(record['font_size_pt']))} pt；"
            f"量測次數 {record['count']}"
        )

    return lines


def write_process_log(output_docx: str | Path, summary: ProcessSummary) -> Path:
    log_path = get_process_log_path(output_docx)
    lines = [
        "Word DOCX 快速整理工具處理紀錄",
        f"輸出檔案：{Path(output_docx)}",
        "",
        "表格處理摘要",
        f"總表格數：{summary.tables}",
        f"跳過第 1 頁的表格數：{summary.skipped_first_page_tables}",
        f"因格子數小於等於 4 而跳過的表格數：{summary.skipped_small_tables}",
        f"偵測到跨頁的表格數：{summary.cross_page_tables}",
        f"成功調整後不再跨頁的表格數：{summary.cross_page_resolved_tables}",
        f"調整後仍跨頁的表格數：{summary.cross_page_still_split_tables}",
        f"儲存格內距被調整的表格數：{summary.adjusted_cell_padding_tables}",
        f"行距或段距被調整的表格數：{summary.adjusted_table_spacing_tables}",
        f"列高改為自動高度的表格數：{summary.auto_height_tables}",
        f"移到下一頁後成功不跨頁的表格數：{summary.moved_next_page_resolved_tables}",
        f"不縮小字體下無法完全避免跨頁的表格數：{summary.cannot_avoid_cross_page_tables}",
        f"處理失敗但已略過的表格數：{summary.failed_cross_page_tables}",
        f"套用「內容大小＋靠右對齊」的表格數：{summary.special_autofit_right_tables}",
        f"其他正常處理的表格數：{summary.normal_processed_tables}",
        "",
        "編號段落處理摘要",
        f"總段落數：{summary.total_paragraphs}",
        f"跳過目錄段落數：{summary.skipped_toc_paragraphs}",
        f"跳過表格段落數：{summary.skipped_table_paragraphs}",
        f"移除全文件既有大綱階層的段落數：{summary.removed_all_outline_paragraphs}",
        f"套用壹、序言前縮排的段落數：{summary.indented_preface_paragraphs}",
        f"套用壹、序言前大綱階層的段落數：{summary.outlined_preface_paragraphs}",
        f"移除字元縮排屬性數：{summary.character_indent_attrs_removed}",
        *[
            f"成功套用第 {level} 階數量：{count}"
            for level, count in enumerate(summary.paragraph_level_counts, start=1)
        ],
        f"無法判斷而跳過的段落數：{summary.unknown_paragraphs}",
        "",
        *format_numbering_indent_log_lines(summary),
        "",
        "段落大綱階層修改紀錄：",
    ]

    if summary.paragraph_logs:
        lines.extend(summary.paragraph_logs)
    else:
        lines.append("沒有段落大綱階層修改紀錄。")

    log_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return log_path
