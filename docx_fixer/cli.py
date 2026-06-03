from __future__ import annotations

import argparse

from .constants import DEFAULT_GRAY
from .docx_processor import fix_docx_fast
from .indent_settings import load_saved_indent_settings
from .models import ProcessOptions
from .process_log import write_process_log
from .stop_controller import StopController

def run_cli(args) -> int:
    load_saved_indent_settings()

    options = ProcessOptions(
        fix_table_layout=args.table,
        fix_color=args.color,
        fix_paragraph=args.paragraph,
        include_tables_in_paragraph=args.paragraph_in_tables,
        remove_preface_outline=args.remove_preface_outline,
        remove_all_outline_levels=args.remove_all_outline,
    )

    if not (
        options.fix_table_layout
        or options.fix_color
        or options.fix_paragraph
        or options.remove_preface_outline
        or options.remove_all_outline_levels
    ):
        options = ProcessOptions(
            fix_table_layout=True,
            fix_color=True,
            fix_paragraph=True,
            include_tables_in_paragraph=args.paragraph_in_tables,
            remove_preface_outline=True,
            remove_all_outline_levels=False,
        )

    stop = StopController()

    def print_progress(percent: float, message: str) -> None:
        print(f"[{percent:6.2f}%] {message}")

    summary = fix_docx_fast(
        input_docx=args.input_docx,
        output_docx=args.output_docx,
        options=options,
        stop=stop,
        progress_callback=print_progress if not args.quiet else None,
    )

    print("完成！")
    print(f"已處理表格數：{summary.tables}")
    print(f"跳過第 1 頁的表格數：{summary.skipped_first_page_tables}")
    print(f"因格子數小於等於 4 而跳過的表格數：{summary.skipped_small_tables}")
    print(f"偵測到跨頁的表格數：{summary.cross_page_tables}")
    print(f"成功調整後不再跨頁的表格數：{summary.cross_page_resolved_tables}")
    print(f"調整後仍跨頁的表格數：{summary.cross_page_still_split_tables}")
    print(f"儲存格內距被調整的表格數：{summary.adjusted_cell_padding_tables}")
    print(f"行距或段距被調整的表格數：{summary.adjusted_table_spacing_tables}")
    print(f"列高改為自動高度的表格數：{summary.auto_height_tables}")
    print(f"移到下一頁後成功不跨頁的表格數：{summary.moved_next_page_resolved_tables}")
    print(f"不縮小字體下無法完全避免跨頁的表格數：{summary.cannot_avoid_cross_page_tables}")
    print(f"處理失敗但已略過的表格數：{summary.failed_cross_page_tables}")
    print(f"套用「內容大小＋靠右對齊」的表格數：{summary.special_autofit_right_tables}")
    print(f"其他正常處理的表格數：{summary.normal_processed_tables}")
    print(f"顏色調整總數：{summary.changed_colors}")
    print(f"指定色碼改成 {DEFAULT_GRAY} 的儲存格數：{summary.changed_to_gray}")
    print(f"其他顏色改成無色彩的儲存格數：{summary.cleared_colors}")
    print(f"已套用階層縮排與大綱階層的段落數：{summary.paragraphs}")
    print(f"總段落數：{summary.total_paragraphs}")
    print(f"跳過目錄段落數：{summary.skipped_toc_paragraphs}")
    print(f"跳過表格段落數：{summary.skipped_table_paragraphs}")
    print(f"移除全文件既有大綱階層的段落數：{summary.removed_all_outline_paragraphs}")
    print(f"移除壹、序言前大綱階層的段落數：{summary.removed_preface_outline_paragraphs}")
    for level, count in enumerate(summary.paragraph_level_counts, start=1):
        print(f"成功套用第 {level} 階數量：{count}")
    print(f"無法判斷而跳過的段落數：{summary.unknown_paragraphs}")
    print(f"輸出檔案路徑：{args.output_docx}")
    log_path = write_process_log(args.output_docx, summary)
    print(f"處理紀錄檔：{log_path}")
    return 0


def parse_args(argv: list[str]):
    parser = argparse.ArgumentParser(description="Word DOCX 快速整理工具。未帶參數時會開啟視窗。")
    parser.add_argument("input_docx", nargs="?", help="輸入 .docx 檔案路徑")
    parser.add_argument("output_docx", nargs="?", help="輸出 .docx 檔案路徑")
    parser.add_argument("--table", action="store_true", help="啟用表格基本格式整理")
    parser.add_argument("--color", action="store_true", help="啟用顏色整理：BFBFBF／A6A6A6／808080 改 D9D9D9，F2F2F2 保持，其他顏色改無色彩")
    parser.add_argument("--paragraph", action="store_true", help="啟用文件編號階層縮排，並依範本.docx的縮排標準套用")
    parser.add_argument("--remove-all-outline", action="store_true", help="去除文件中所有既有 Word 大綱階層")
    parser.add_argument("--remove-preface-outline", action="store_true", help="移除第一次壹、序言前既有的 Word 大綱階層")
    parser.add_argument("--paragraph-in-tables", action="store_true", help="文件編號階層縮排也處理表格內段落")
    parser.add_argument("--quiet", action="store_true", help="命令列模式不輸出進度")
    return parser.parse_args(argv)
