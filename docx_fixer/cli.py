from __future__ import annotations

import argparse
import sys

from .constants import DEFAULT_GRAY
from .docx_processor import fix_docx_fast
from .indent_settings import load_saved_indent_settings
from .models import ProcessOptions
from .process_log import write_heading_suffix_log_file, write_process_log, write_table_log_file
from .stop_controller import StopController
from .table_color_settings import (
    current_table_color_settings,
    normalize_hex_color,
    parse_color_list_text,
)


def _chapter_three_options_from_args(args) -> tuple[bool, bool, bool]:
    skip_table_layout = args.skip_chapter_three_table_layout
    skip_table_color = args.skip_chapter_three_table_color
    skip_indents = args.skip_chapter_three_indents

    if args.skip_chapter_three_tables is not None:
        skip_table_layout = args.skip_chapter_three_tables
        skip_table_color = args.skip_chapter_three_tables

    if args.skip_all_under_chapter_three is not None:
        skip_table_layout = args.skip_all_under_chapter_three
        skip_table_color = args.skip_all_under_chapter_three
        skip_indents = args.skip_all_under_chapter_three

    if args.skip_special_layout_under_chapter_three:
        skip_table_layout = True
        skip_table_color = True
        skip_indents = True

    return skip_table_layout, skip_table_color, skip_indents


def _table_color_options_from_args(args) -> dict[str, object]:
    saved = current_table_color_settings()
    keep_colors = (
        parse_color_list_text(args.table_keep_colors)
        if args.table_keep_colors is not None
        else list(saved["keep_colors"])
    )
    gray_colors = (
        parse_color_list_text(args.table_gray_colors)
        if args.table_gray_colors is not None
        else list(saved["gray_colors"])
    )
    gray_target = (
        normalize_hex_color(args.table_gray_target)
        if args.table_gray_target is not None
        else str(saved["gray_target"])
    )
    special_skip_colors = (
        parse_color_list_text(args.special_color_skip_colors)
        if args.special_color_skip_colors is not None
        else list(saved["special_color_skip_colors"])
    )
    return {
        "table_keep_colors": tuple(keep_colors),
        "table_gray_colors": tuple(gray_colors),
        "table_gray_target": gray_target,
        "skip_special_color_tables": args.skip_special_color_tables,
        "special_color_skip_colors": tuple(special_skip_colors),
        "clear_special_colors_after_skip": args.clear_special_colors_after_skip,
    }


def _build_process_options(args, *, enable_default_actions: bool = False) -> ProcessOptions:
    (
        skip_chapter_three_table_layout,
        skip_chapter_three_table_color,
        skip_chapter_three_indents,
    ) = _chapter_three_options_from_args(args)
    return ProcessOptions(
        fix_table_layout=True if enable_default_actions else args.table,
        fix_color=True if enable_default_actions else args.color,
        fix_paragraph=True if enable_default_actions else args.paragraph,
        remove_all_outline_levels=False if enable_default_actions else args.remove_all_outline,
        indent_preface_paragraphs=False if enable_default_actions else args.indent_preface,
        outline_preface_paragraphs=False if enable_default_actions else args.outline_preface,
        enable_level1_level2_body_first_line_indent=args.level1_level2_body_first_line_indent,
        word_com_check_body_font_when_xml_not_14=args.word_com_check_body_font,
        normalize_body_style_to_none=args.normalize_body_style_to_none,
        skip_chapter_three_table_layout=skip_chapter_three_table_layout,
        skip_chapter_three_table_color=skip_chapter_three_table_color,
        skip_chapter_three_indents=skip_chapter_three_indents,
        skip_chapter_three_numbering_suffix_cleanup=args.skip_chapter_three_numbering_suffix_cleanup,
        skip_chapter_three_adjustments=args.skip_chapter_three_adjustments,
        move_table_notes_below=args.move_table_notes_below,
        skip_chapter_three_table_notes=args.skip_chapter_three_table_notes,
        force_note_paragraph_left_alignment=args.force_note_paragraph_left_alignment,
        enable_double_black_table_borders=args.enable_double_black_table_borders,
        enable_table_footer_source_format=args.enable_table_footer_source_format,
        write_note_debug_log=args.write_note_debug_log,
        skip_nested_tables=args.skip_nested_tables,
        skip_log_output=args.skip_log_output,
        **_table_color_options_from_args(args),
    )


def _configure_stdio_utf8() -> None:
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream is None:
            continue
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            reconfigure(encoding="utf-8", errors="replace")


def run_cli(args) -> int:
    _configure_stdio_utf8()
    load_saved_indent_settings()

    options = _build_process_options(args)

    if not (
        options.fix_table_layout
        or options.fix_color
        or options.fix_paragraph
        or options.remove_all_outline_levels
        or options.indent_preface_paragraphs
        or options.outline_preface_paragraphs
        or options.move_table_notes_below
    ):
        options = _build_process_options(args, enable_default_actions=True)

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

    print("Done")
    print(f"tables={summary.tables}")
    print(f"skipped_first_page_tables={summary.skipped_first_page_tables}")
    print(f"skipped_small_tables={summary.skipped_small_tables}")
    print(f"skipped_nested_tables={summary.skipped_nested_tables}")
    print(f"special_color_skipped_tables={summary.special_color_skipped_tables}")
    print(f"cross_page_tables={summary.cross_page_tables}")
    print(f"cross_page_resolved_tables={summary.cross_page_resolved_tables}")
    print(f"cross_page_still_split_tables={summary.cross_page_still_split_tables}")
    print(f"adjusted_cell_padding_tables={summary.adjusted_cell_padding_tables}")
    print(f"adjusted_table_spacing_tables={summary.adjusted_table_spacing_tables}")
    print(f"auto_height_tables={summary.auto_height_tables}")
    print(f"moved_next_page_resolved_tables={summary.moved_next_page_resolved_tables}")
    print(f"cannot_avoid_cross_page_tables={summary.cannot_avoid_cross_page_tables}")
    print(f"failed_cross_page_tables={summary.failed_cross_page_tables}")
    print(f"special_autofit_right_tables={summary.special_autofit_right_tables}")
    print(f"normal_processed_tables={summary.normal_processed_tables}")
    print(f"word_com_table_autofit_applied_count={summary.word_com_table_autofit_applied_count}")
    print(f"word_com_table_autofit_fallback_count={summary.word_com_table_autofit_fallback_count}")
    print(f"word_com_table_autofit_failed_count={summary.word_com_table_autofit_failed_count}")
    if summary.word_com_table_autofit_failed_count:
        print(
            "WARNING: word_com_table_autofit_failed_count="
            f"{summary.word_com_table_autofit_failed_count} "
            "(Word COM AutoFit failed and XML fallback could not repair these tables)"
        )
    print(f"changed_colors={summary.changed_colors}")
    print(f"changed_to_gray_{DEFAULT_GRAY}={summary.changed_to_gray}")
    print(f"cleared_colors={summary.cleared_colors}")
    print(f"paragraphs_changed={summary.paragraphs}")
    print(f"total_paragraphs={summary.total_paragraphs}")
    print(f"skipped_toc_paragraphs={summary.skipped_toc_paragraphs}")
    print(f"skipped_table_paragraphs={summary.skipped_table_paragraphs}")
    print(f"removed_all_outline_paragraphs={summary.removed_all_outline_paragraphs}")
    print(f"indented_preface_paragraphs={summary.indented_preface_paragraphs}")
    print(f"outlined_preface_paragraphs={summary.outlined_preface_paragraphs}")
    for level, count in enumerate(summary.paragraph_level_counts, start=1):
        print(f"level_{level}_paragraphs={count}")
    print(f"unknown_paragraphs={summary.unknown_paragraphs}")
    print(f"output_docx={args.output_docx}")
    log_path = None
    table_log_path = None
    heading_suffix_log_path = None
    if not options.skip_log_output:
        log_path = write_process_log(args.output_docx, summary)
        table_log_path = write_table_log_file(args.output_docx, summary)
        heading_suffix_log_path = write_heading_suffix_log_file(args.output_docx, summary)
    print(f"process_log={log_path}")
    print(f"table_log={table_log_path}")
    print(f"heading_suffix_log={heading_suffix_log_path}")
    return 0


def parse_args(argv: list[str]):
    parser = argparse.ArgumentParser(description="Fix Word DOCX formatting")
    parser.add_argument("input_docx", nargs="?", help="Input .docx path")
    parser.add_argument("output_docx", nargs="?", help="Output .docx path")
    parser.add_argument("--table", action="store_true", help="Fix table layout")
    parser.add_argument("--color", action="store_true", help="Fix table shading colors")
    parser.add_argument("--paragraph", action="store_true", help="Fix paragraph outline levels and indents")
    parser.add_argument("--remove-all-outline", action="store_true", help="Force all paragraph outline levels to body text")
    parser.add_argument("--indent-preface", action="store_true", help="Indent preface paragraphs before the main body marker")
    parser.add_argument("--outline-preface", action="store_true", help="Apply outline levels to preface paragraphs before the main body marker")
    parser.add_argument(
        "--level1-level2-body-first-line-indent",
        "--level2-body-first-line-indent",
        action="store_true",
        dest="level1_level2_body_first_line_indent",
        help="Apply 560 twips first-line indent to plain body text under level 1 and level 2 headings",
    )
    parser.add_argument("--word-com-check-body-font", action="store_true", help="XML body font is not 14pt: ask Word COM to verify before applying body indent")
    parser.add_argument(
        "--normalize-body-style-to-none",
        action="store_true",
        dest="normalize_body_style_to_none",
        default=False,
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--no-normalize-body-style-to-none",
        "--no-normalize-body-style-to-default-text",
        action="store_false",
        dest="normalize_body_style_to_none",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--skip-special-layout-under-chapter-three",
        action="store_true",
        help="Deprecated alias: skip both table and indent fixes under chapter 參、價格形成之主要因素分析",
    )
    parser.add_argument(
        "--skip-chapter-three-tables",
        action="store_true",
        default=None,
        dest="skip_chapter_three_tables",
        help="Deprecated alias: do not modify table layout or colors under chapter 參、價格形成之主要因素分析",
    )
    parser.add_argument(
        "--no-skip-chapter-three-tables",
        action="store_false",
        dest="skip_chapter_three_tables",
        help="Deprecated alias: allow table layout and color changes under chapter 參、價格形成之主要因素分析",
    )
    parser.add_argument(
        "--skip-chapter-three-table-layout",
        action="store_true",
        default=True,
        dest="skip_chapter_three_table_layout",
        help="Do not modify table layout under chapter 參、價格形成之主要因素分析",
    )
    parser.add_argument(
        "--no-skip-chapter-three-table-layout",
        action="store_false",
        dest="skip_chapter_three_table_layout",
        help="Allow table layout changes under chapter 參、價格形成之主要因素分析",
    )
    parser.add_argument(
        "--skip-chapter-three-table-color",
        action="store_true",
        default=True,
        dest="skip_chapter_three_table_color",
        help="Do not modify table shading colors under chapter 參、價格形成之主要因素分析",
    )
    parser.add_argument(
        "--no-skip-chapter-three-table-color",
        action="store_false",
        dest="skip_chapter_three_table_color",
        help="Allow table shading color changes under chapter 參、價格形成之主要因素分析",
    )
    parser.add_argument(
        "--skip-chapter-three-indents",
        action="store_true",
        default=True,
        dest="skip_chapter_three_indents",
        help="Do not modify paragraph indents under chapter 參、價格形成之主要因素分析",
    )
    parser.add_argument(
        "--no-skip-chapter-three-indents",
        action="store_false",
        dest="skip_chapter_three_indents",
        help="Allow paragraph indent changes under chapter 參、價格形成之主要因素分析",
    )
    parser.add_argument(
        "--skip-chapter-three-numbering-suffix-cleanup",
        action="store_true",
        default=True,
        dest="skip_chapter_three_numbering_suffix_cleanup",
        help="參、不要清理編號後綴 tab/space: do not clean numbering.xml suffix/tab/lvlText trailing whitespace for chapter 參 numbering definitions",
    )
    parser.add_argument(
        "--no-skip-chapter-three-numbering-suffix-cleanup",
        action="store_false",
        dest="skip_chapter_three_numbering_suffix_cleanup",
        help="Allow numbering suffix/tab cleanup for chapter 參 numbering definitions",
    )
    parser.add_argument(
        "--skip-chapter-three-adjustments",
        "--protect-section-three",
        action="store_true",
        default=False,
        dest="skip_chapter_three_adjustments",
        help="參、不要調整: protect the whole body chapter 參、 from every adjustment (layout, color, font, borders, note moving, indents, Word COM)",
    )
    parser.add_argument(
        "--no-skip-chapter-three-adjustments",
        "--no-protect-section-three",
        action="store_false",
        dest="skip_chapter_three_adjustments",
        help="Allow adjustments inside chapter 參、",
    )
    parser.add_argument(
        "--move-table-notes-below",
        action="store_true",
        default=False,
        dest="move_table_notes_below",
        help="Move note cells (註：/註1：/註一、 ...) out of each table into paragraphs below it",
    )
    parser.add_argument(
        "--no-move-table-notes-below",
        action="store_false",
        dest="move_table_notes_below",
        help="Keep note cells inside tables",
    )
    parser.add_argument(
        "--skip-chapter-three-table-notes",
        action="store_true",
        default=True,
        dest="skip_chapter_three_table_notes",
        help="參、不要表格註記搬移: do not move note cells for tables inside the body chapter 參、 (only affects note moving)",
    )
    parser.add_argument(
        "--no-skip-chapter-three-table-notes",
        action="store_false",
        dest="skip_chapter_three_table_notes",
        help="Allow moving note cells even for tables inside the body chapter 參、",
    )
    parser.add_argument(
        "--force-note-paragraph-left-alignment",
        action="store_true",
        default=False,
        dest="force_note_paragraph_left_alignment",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--no-force-note-paragraph-left-alignment",
        action="store_false",
        dest="force_note_paragraph_left_alignment",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--enable-double-black-table-borders",
        action="store_true",
        default=False,
        dest="enable_double_black_table_borders",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--no-enable-double-black-table-borders",
        action="store_false",
        dest="enable_double_black_table_borders",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--write-note-debug-log",
        action="store_true",
        default=False,
        dest="write_note_debug_log",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--no-write-note-debug-log",
        action="store_false",
        dest="write_note_debug_log",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--enable-table-footer-source-format",
        "--table-footer-source-format",
        action="store_true",
        default=False,
        dest="enable_table_footer_source_format",
        help="表格最後一列說明格式化: format the table outer frame, the single-cell title row, and the last-row 基期：/資料來源：/註記 cells",
    )
    parser.add_argument(
        "--no-enable-table-footer-source-format",
        "--no-table-footer-source-format",
        action="store_false",
        dest="enable_table_footer_source_format",
        help="Do not apply the last-row 基期/資料來源/註記 footer table format",
    )
    parser.add_argument(
        "--skip-nested-tables",
        action="store_true",
        default=True,
        dest="skip_nested_tables",
        help="Do not modify tables that contain another table or are inside another table",
    )
    parser.add_argument(
        "--no-skip-nested-tables",
        action="store_false",
        dest="skip_nested_tables",
        help="Allow changes to nested tables and tables that contain nested tables",
    )
    parser.add_argument(
        "--skip-all-under-chapter-three",
        action="store_true",
        default=None,
        dest="skip_all_under_chapter_three",
        help="Deprecated alias: skip both table and indent fixes under chapter 參、價格形成之主要因素分析",
    )
    parser.add_argument(
        "--no-skip-all-under-chapter-three",
        action="store_false",
        dest="skip_all_under_chapter_three",
        help="Deprecated alias: allow both table and indent fixes under chapter 參、價格形成之主要因素分析",
    )
    parser.add_argument(
        "--skip-special-color-tables",
        action="store_true",
        default=False,
        dest="skip_special_color_tables",
        help="Skip the whole table when any cell shading matches the special color skip list",
    )
    parser.add_argument(
        "--no-skip-special-color-tables",
        action="store_false",
        dest="skip_special_color_tables",
        help="Do not skip tables based on the special color skip list",
    )
    parser.add_argument(
        "--clear-special-colors-after-skip",
        action="store_true",
        default=False,
        dest="clear_special_colors_after_skip",
        help="After skipping a special color table, clear only the cells matching the skip list to no color",
    )
    parser.add_argument(
        "--no-clear-special-colors-after-skip",
        action="store_false",
        dest="clear_special_colors_after_skip",
        help="Keep the matched special colors when skipping a special color table",
    )
    parser.add_argument(
        "--table-keep-colors",
        default=None,
        dest="table_keep_colors",
        help="Comma separated HEX shading colors to keep unchanged (default: saved table color settings)",
    )
    parser.add_argument(
        "--table-gray-colors",
        default=None,
        dest="table_gray_colors",
        help="Comma separated HEX shading colors converted to the gray target (default: saved table color settings)",
    )
    parser.add_argument(
        "--table-gray-target",
        default=None,
        dest="table_gray_target",
        help="Target gray HEX used when converting shading colors (default: saved table color settings)",
    )
    parser.add_argument(
        "--special-color-skip-colors",
        default=None,
        dest="special_color_skip_colors",
        help="Comma separated HEX colors; a table is skipped entirely when any cell matches (default: saved table color settings)",
    )
    parser.add_argument(
        "--no-log",
        "--skip-log-output",
        action="store_true",
        default=False,
        dest="skip_log_output",
        help="Do not write process, table, or heading suffix log files",
    )
    parser.add_argument("--quiet", action="store_true", help="Do not print progress")
    return parser.parse_args(argv)


