from __future__ import annotations

import unittest
from contextlib import redirect_stderr, redirect_stdout
from dataclasses import fields
from io import StringIO
from pathlib import Path

from docx_fixer.cli import _build_process_options, parse_args
from docx_fixer.models import ProcessOptions


class CliOptionTests(unittest.TestCase):
    def test_remove_preface_outline_argument_is_no_longer_supported(self):
        with redirect_stderr(StringIO()):
            with self.assertRaises(SystemExit):
                parse_args(["input.docx", "output.docx", "--remove-preface-outline"])

    def test_new_preface_arguments_are_supported(self):
        args = parse_args(["input.docx", "output.docx", "--indent-preface", "--outline-preface"])

        self.assertTrue(args.indent_preface)
        self.assertTrue(args.outline_preface)

    def test_process_options_default_skip_log_output_is_true(self):
        options = ProcessOptions(True, True, True)

        self.assertTrue(options.skip_log_output)
        self.assertTrue(options.skip_nested_tables)

    def test_cli_preserves_log_output_by_default_and_supports_no_log_aliases(self):
        args = parse_args(["input.docx", "output.docx"])
        options = _build_process_options(args)

        self.assertFalse(args.skip_log_output)
        self.assertFalse(options.skip_log_output)

        no_log_args = parse_args(["input.docx", "output.docx", "--no-log"])
        no_log_options = _build_process_options(no_log_args)
        self.assertTrue(no_log_args.skip_log_output)
        self.assertTrue(no_log_options.skip_log_output)

        alias_args = parse_args(["input.docx", "output.docx", "--skip-log-output"])
        alias_options = _build_process_options(alias_args)
        self.assertTrue(alias_args.skip_log_output)
        self.assertTrue(alias_options.skip_log_output)

    def test_process_options_do_not_keep_old_special_table_flag(self):
        field_names = {field.name for field in fields(ProcessOptions)}

        self.assertNotIn("skip_special_table_layout_under_chapter_three", field_names)

    def test_new_body_indent_arguments_are_supported(self):
        args = parse_args([
            "input.docx",
            "output.docx",
            "--level1-level2-body-first-line-indent",
            "--word-com-check-body-font",
            "--skip-special-layout-under-chapter-three",
            "--skip-chapter-three-tables",
            "--skip-chapter-three-table-layout",
            "--skip-chapter-three-table-color",
            "--skip-chapter-three-indents",
        ])

        self.assertTrue(args.level1_level2_body_first_line_indent)
        self.assertTrue(args.word_com_check_body_font)
        self.assertTrue(args.skip_special_layout_under_chapter_three)
        self.assertTrue(args.skip_chapter_three_tables)
        self.assertTrue(args.skip_chapter_three_table_layout)
        self.assertTrue(args.skip_chapter_three_table_color)
        self.assertTrue(args.skip_chapter_three_indents)

    def test_body_style_normalization_hidden_argument_can_enable_internal_option(self):
        args = parse_args([
            "input.docx",
            "output.docx",
            "--normalize-body-style-to-none",
        ])
        options = _build_process_options(args)

        self.assertTrue(args.normalize_body_style_to_none)
        self.assertTrue(options.normalize_body_style_to_none)

    def test_body_style_normalization_is_disabled_by_default_and_old_disable_alias_is_supported(self):
        default_args = parse_args(["input.docx", "output.docx"])
        default_options = _build_process_options(default_args)
        self.assertFalse(default_args.normalize_body_style_to_none)
        self.assertFalse(default_options.normalize_body_style_to_none)

        args = parse_args([
            "input.docx",
            "output.docx",
            "--no-normalize-body-style-to-default-text",
        ])
        options = _build_process_options(args)

        self.assertFalse(args.normalize_body_style_to_none)
        self.assertFalse(options.normalize_body_style_to_none)

    def test_body_style_normalization_argument_is_hidden_from_help(self):
        stdout = StringIO()
        with redirect_stdout(stdout):
            with self.assertRaises(SystemExit):
                parse_args(["--help"])

        self.assertNotIn("normalize-body-style-to-none", stdout.getvalue())
        self.assertNotIn("normalize-body-style-to-default-text", stdout.getvalue())

    def test_chapter_three_table_and_indent_skips_are_enabled_by_default(self):
        args = parse_args(["input.docx", "output.docx"])
        options = _build_process_options(args)

        self.assertIsNone(args.skip_chapter_three_tables)
        self.assertTrue(args.skip_chapter_three_table_layout)
        self.assertTrue(args.skip_chapter_three_table_color)
        self.assertTrue(args.skip_chapter_three_indents)
        self.assertTrue(options.skip_chapter_three_table_layout)
        self.assertTrue(options.skip_chapter_three_table_color)
        self.assertTrue(options.skip_chapter_three_indents)

    def test_nested_table_protection_is_enabled_by_default_and_can_be_disabled(self):
        args = parse_args(["input.docx", "output.docx"])
        options = _build_process_options(args)

        self.assertTrue(args.skip_nested_tables)
        self.assertTrue(options.skip_nested_tables)

        disabled_args = parse_args(["input.docx", "output.docx", "--no-skip-nested-tables"])
        disabled_options = _build_process_options(disabled_args)
        self.assertFalse(disabled_args.skip_nested_tables)
        self.assertFalse(disabled_options.skip_nested_tables)

    def test_new_chapter_three_table_options_can_be_disabled_independently(self):
        args = parse_args([
            "input.docx",
            "output.docx",
            "--no-skip-chapter-three-table-layout",
        ])
        options = _build_process_options(args)

        self.assertFalse(options.skip_chapter_three_table_layout)
        self.assertTrue(options.skip_chapter_three_table_color)
        self.assertTrue(options.skip_chapter_three_indents)

    def test_no_skip_all_under_chapter_three_disables_both_legacy_protections(self):
        args = parse_args(["input.docx", "output.docx", "--no-skip-all-under-chapter-three"])
        options = _build_process_options(args)

        self.assertFalse(args.skip_all_under_chapter_three)
        self.assertFalse(options.skip_chapter_three_table_layout)
        self.assertFalse(options.skip_chapter_three_table_color)
        self.assertFalse(options.skip_chapter_three_indents)

    def test_legacy_chapter_three_tables_alias_maps_to_layout_and_color(self):
        args = parse_args(["input.docx", "output.docx", "--no-skip-chapter-three-tables"])
        options = _build_process_options(args)

        self.assertFalse(args.skip_chapter_three_tables)
        self.assertFalse(options.skip_chapter_three_table_layout)
        self.assertFalse(options.skip_chapter_three_table_color)

    def test_old_special_layout_alias_maps_to_both_chapter_three_options(self):
        args = parse_args([
            "input.docx",
            "output.docx",
            "--no-skip-all-under-chapter-three",
            "--skip-special-layout-under-chapter-three",
        ])
        options = _build_process_options(args)

        self.assertTrue(args.skip_special_layout_under_chapter_three)
        self.assertTrue(options.skip_chapter_three_table_layout)
        self.assertTrue(options.skip_chapter_three_table_color)
        self.assertTrue(options.skip_chapter_three_indents)

    def test_old_level_two_body_indent_argument_is_kept_as_alias(self):
        args = parse_args(["input.docx", "output.docx", "--level2-body-first-line-indent"])

        self.assertTrue(args.level1_level2_body_first_line_indent)

    def test_paragraph_in_tables_argument_is_no_longer_supported(self):
        with redirect_stderr(StringIO()):
            with self.assertRaises(SystemExit):
                parse_args(["input.docx", "output.docx", "--paragraph-in-tables"])

    def test_readme_no_longer_mentions_removed_preface_argument(self):
        readme = Path("README.md").read_text(encoding="utf-8")

        self.assertNotIn("remove-preface-outline", readme)


if __name__ == "__main__":
    unittest.main()

