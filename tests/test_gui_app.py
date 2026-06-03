from __future__ import annotations

import unittest

from docx_fixer.gui_app import DEFAULT_WINDOW_GEOMETRY, MIN_WINDOW_SIZE


class GuiAppTests(unittest.TestCase):
    def test_default_and_min_window_sizes_keep_action_buttons_visible(self):
        width, height = [int(part) for part in DEFAULT_WINDOW_GEOMETRY.split("x")]
        min_width, min_height = MIN_WINDOW_SIZE

        self.assertGreaterEqual(width, 900)
        self.assertGreaterEqual(height, 720)
        self.assertGreaterEqual(min_width, 820)
        self.assertGreaterEqual(min_height, 640)
        self.assertLessEqual(min_width, width)
        self.assertLessEqual(min_height, height)


if __name__ == "__main__":
    unittest.main()
