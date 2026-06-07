from __future__ import annotations

import unittest
from pathlib import Path

from docx_fixer.outline import format_font_size_for_log


class TextSanityTests(unittest.TestCase):
    def test_readme_has_no_common_mojibake_fragments(self):
        readme = Path("README.md").read_text(encoding="utf-8")

        for fragment in ("嚗", "銝", "蝮", "撅", "閮", "?", "?"):
            self.assertNotIn(fragment, readme)

    def test_format_font_size_for_log_uses_human_readable_fallback(self):
        self.assertEqual(format_font_size_for_log(None), "未知字號")


if __name__ == "__main__":
    unittest.main()
