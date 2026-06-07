from __future__ import annotations

import unittest
from contextlib import redirect_stderr
from io import StringIO
from pathlib import Path

from docx_fixer.cli import parse_args


class CliOptionTests(unittest.TestCase):
    def test_remove_preface_outline_argument_is_no_longer_supported(self):
        with redirect_stderr(StringIO()):
            with self.assertRaises(SystemExit):
                parse_args(["input.docx", "output.docx", "--remove-preface-outline"])

    def test_new_preface_arguments_are_supported(self):
        args = parse_args(["input.docx", "output.docx", "--indent-preface", "--outline-preface"])

        self.assertTrue(args.indent_preface)
        self.assertTrue(args.outline_preface)

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

