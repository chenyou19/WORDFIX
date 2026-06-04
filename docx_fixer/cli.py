from __future__ import annotations

import argparse
import sys

from .constants import DEFAULT_GRAY
from .docx_processor import fix_docx_fast
from .indent_settings import load_saved_indent_settings
from .models import ProcessOptions
from .process_log import write_process_log, write_table_log_file
from .stop_controller import StopController


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

    options = ProcessOptions(
        fix_table_layout=args.table,
        fix_color=args.color,
        fix_paragraph=args.paragraph,
        remove_all_outline_levels=args.remove_all_outline,
        indent_preface_paragraphs=args.indent_preface,
        outline_preface_paragraphs=args.outline_preface,
        enable_level2_body_first_line_indent=args.level2_body_first_line_indent,
        word_com_check_body_font_when_xml_not_14=args.word_com_check_body_font,
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
            remove_all_outline_levels=False,
            indent_preface_paragraphs=False,
            outline_preface_paragraphs=False,
            enable_level2_body_first_line_indent=args.level2_body_first_line_indent,
            word_com_check_body_font_when_xml_not_14=args.word_com_check_body_font,
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
    print(f"process_log={log_path}")
    print(f"table_log={table_log_path}")
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
    parser.add_argument("--level2-body-first-line-indent", action="store_true", help="Apply 560 twips first-line indent to plain body text under level 2 headings")
    parser.add_argument("--word-com-check-body-font", action="store_true", help="XML body font is not 14pt: ask Word COM to verify before applying body indent")
    parser.add_argument("--quiet", action="store_true", help="Do not print progress")
    return parser.parse_args(argv)
