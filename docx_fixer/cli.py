from __future__ import annotations

import argparse
import sys

from .constants import DEFAULT_GRAY
from .docx_processor import fix_docx_fast
from .indent_settings import load_saved_indent_settings
from .models import ProcessOptions
from .process_log import write_heading_suffix_log_file, write_process_log, write_table_log_file
from .stop_controller import StopController


def _chapter_three_options_from_args(args) -> tuple[bool, bool]:
    skip_tables = args.skip_chapter_three_tables
    skip_indents = args.skip_chapter_three_indents

    if args.skip_all_under_chapter_three is not None:
        skip_tables = args.skip_all_under_chapter_three
        skip_indents = args.skip_all_under_chapter_three

    if args.skip_special_layout_under_chapter_three:
        skip_tables = True
        skip_indents = True

    return skip_tables, skip_indents


def _build_process_options(args, *, enable_default_actions: bool = False) -> ProcessOptions:
    skip_chapter_three_tables, skip_chapter_three_indents = _chapter_three_options_from_args(args)
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
        skip_chapter_three_tables=skip_chapter_three_tables,
        skip_chapter_three_indents=skip_chapter_three_indents,
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
        default=True,
        dest="skip_chapter_three_tables",
        help="Do not modify table layout or colors under chapter 參、價格形成之主要因素分析",
    )
    parser.add_argument(
        "--no-skip-chapter-three-tables",
        action="store_false",
        dest="skip_chapter_three_tables",
        help="Allow table layout and color changes under chapter 參、價格形成之主要因素分析",
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
    parser.add_argument("--quiet", action="store_true", help="Do not print progress")
    return parser.parse_args(argv)


