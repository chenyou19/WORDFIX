from __future__ import annotations

import unittest
from contextlib import redirect_stderr
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
            "--skip-all-under-chapter-three",
        ])

        self.assertTrue(args.level1_level2_body_first_line_indent)
        self.assertTrue(args.word_com_check_body_font)
        self.assertTrue(args.skip_special_layout_under_chapter_three)
        self.assertTrue(args.skip_all_under_chapter_three)

    def test_chapter_three_protection_is_enabled_by_default(self):
        args = parse_args(["input.docx", "output.docx"])
        options = _build_process_options(args)

        self.assertTrue(args.skip_all_under_chapter_three)
        self.assertTrue(options.skip_all_under_chapter_three)

    def test_no_skip_all_under_chapter_three_disables_protection(self):
        args = parse_args(["input.docx", "output.docx", "--no-skip-all-under-chapter-three"])
        options = _build_process_options(args)

        self.assertFalse(args.skip_all_under_chapter_three)
        self.assertFalse(options.skip_all_under_chapter_three)

    def test_old_special_layout_alias_maps_to_chapter_three_protection(self):
        args = parse_args([
            "input.docx",
            "output.docx",
            "--no-skip-all-under-chapter-three",
            "--skip-special-layout-under-chapter-three",
        ])
        options = _build_process_options(args)

        self.assertTrue(args.skip_special_layout_under_chapter_three)
        self.assertTrue(options.skip_all_under_chapter_three)

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

