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

    def test_readme_no_longer_mentions_removed_preface_argument(self):
        readme = Path("README.md").read_text(encoding="utf-8")

        self.assertNotIn("remove-preface-outline", readme)


if __name__ == "__main__":
    unittest.main()
