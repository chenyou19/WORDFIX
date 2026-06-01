from __future__ import annotations

import unittest

from docx_fixer.exceptions import ProcessStopped
from docx_fixer.stop_controller import StopController


class StopControllerTests(unittest.TestCase):
    def test_stop_invokes_registered_cleanup_callbacks(self):
        stop = StopController()
        called: list[str] = []

        def cleanup() -> None:
            called.append("cleanup")

        stop.register_stop_callback(cleanup)
        stop.stop()

        self.assertEqual(called, ["cleanup"])
        with self.assertRaises(ProcessStopped):
            stop.check()

    def test_unregister_callback_prevents_cleanup_call(self):
        stop = StopController()
        called: list[str] = []

        def cleanup() -> None:
            called.append("cleanup")

        stop.register_stop_callback(cleanup)
        stop.unregister_stop_callback(cleanup)
        stop.stop()

        self.assertEqual(called, [])


if __name__ == "__main__":
    unittest.main()
