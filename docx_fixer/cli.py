from __future__ import annotations

import argparse

from .constants import DEFAULT_GRAY
from .docx_processor import fix_docx_fast
from .indent_settings import load_saved_indent_settings
from .models import ProcessOptions
from .process_log import write_process_log, write_table_log_file
from .stop_controller import StopController


def run_cli(args) -> int:
    load_saved_indent_settings()

    options = ProcessOptions(
        fix_table_layout=args.table,
        fix_color=args.color,
        fix_paragraph=args.paragraph,
        include_tables_in_paragraph=args.paragraph_in_tables,
        remove_all_outline_levels=args.remove_all_outline,
        indent_preface_paragraphs=args.indent_preface,
        outline_preface_paragraphs=args.outline_preface,
    )

    if not (
        options.fix_table_layout
        or options.fix_color
        or options.fix_paragraph
        or options.remove_all_outline_levels
        or options.indent_preface_paragraphs
        or options.outline_preface_paragraphs
    ):
        options = ProcessOptions(
            fix_table_layout=True,
            fix_color=True,
            fix_paragraph=True,
            include_tables_in_paragraph=args.paragraph_in_tables,
            remove_all_outline_levels=False,
            indent_preface_paragraphs=False,
            outline_preface_paragraphs=False,
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

    print("完成")
    print(f"表格總數：{summary.tables}")
    print(f"跳過第一張表格數：{summary.skipped_first_page_tables}")
    print(f"因格子數小於等於 4 而跳過的表格數：{summary.skipped_small_tables}")
    print(f"跨頁表格數：{summary.cross_page_tables}")
    print(f"跨頁已解決的表格數：{summary.cross_page_resolved_tables}")
    print(f"跨頁未解決的表格數：{summary.cross_page_still_split_tables}")
    print(f"調整儲存格 padding 的表格數：{summary.adjusted_cell_padding_tables}")
    print(f"調整表格段落間距的表格數：{summary.adjusted_table_spacing_tables}")
    print(f"改成自動列高的表格數：{summary.auto_height_tables}")
    print(f"移到下一頁後解決跨頁的表格數：{summary.moved_next_page_resolved_tables}")
    print(f"無法避免跨頁的表格數：{summary.cannot_avoid_cross_page_tables}")
    print(f"跨頁處理失敗的表格數：{summary.failed_cross_page_tables}")
    print(f"套用特殊版面表格數：{summary.special_autofit_right_tables}")
    print(f"一般表格處理數：{summary.normal_processed_tables}")
    print(f"顏色調整總數：{summary.changed_colors}")
    print(f"改成 {DEFAULT_GRAY} 的儲存格數：{summary.changed_to_gray}")
    print(f"清除底色的儲存格數：{summary.cleared_colors}")
    print(f"被修改的大綱段落數：{summary.paragraphs}")
    print(f"段落總數：{summary.total_paragraphs}")
    print(f"跳過目錄段落數：{summary.skipped_toc_paragraphs}")
    print(f"跳過表格段落數：{summary.skipped_table_paragraphs}")
    print(f"移除大綱層級的段落數：{summary.removed_all_outline_paragraphs}")
    print(f"套用前言縮排的段落數：{summary.indented_preface_paragraphs}")
    print(f"套用前言大綱的段落數：{summary.outlined_preface_paragraphs}")
    for level, count in enumerate(summary.paragraph_level_counts, start=1):
        print(f"第 {level} 階段落數：{count}")
    print(f"未分類段落數：{summary.unknown_paragraphs}")
    print(f"輸出檔案：{args.output_docx}")
    log_path = write_process_log(args.output_docx, summary)
    table_log_path = write_table_log_file(args.output_docx, summary)
    print(f"處理紀錄：{log_path}")
    print(f"表格紀錄：{table_log_path}")
    return 0


def parse_args(argv: list[str]):
    parser = argparse.ArgumentParser(description="Word DOCX 格式修正工具")
    parser.add_argument("input_docx", nargs="?", help="輸入 .docx 檔案路徑")
    parser.add_argument("output_docx", nargs="?", help="輸出 .docx 檔案路徑")
    parser.add_argument("--table", action="store_true", help="修正表格版面")
    parser.add_argument("--color", action="store_true", help="修正表格底色")
    parser.add_argument("--paragraph", action="store_true", help="修正段落大綱格式")
    parser.add_argument("--remove-all-outline", action="store_true", help="移除所有段落大綱層級")
    parser.add_argument("--indent-preface", action="store_true", help="套用前言段落縮排")
    parser.add_argument("--outline-preface", action="store_true", help="套用前言段落大綱層級")
    parser.add_argument("--paragraph-in-tables", action="store_true", help="也處理表格中的段落")
    parser.add_argument("--quiet", action="store_true", help="不顯示進度訊息")
    return parser.parse_args(argv)
