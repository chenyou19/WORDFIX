from __future__ import annotations

import subprocess
import unittest
from unittest.mock import patch

from docx_fixer import process_runner


class _FakeProcess:
    """Minimal stand-in for subprocess.Popen that finishes immediately."""

    def __init__(self):
        self.args = ["powershell"]
        self.returncode = 0

    def poll(self):
        return 0

    def communicate(self, timeout=None):
        return ("stdout-data", "stderr-data")


class _FakeStartupInfo:
    def __init__(self):
        self.dwFlags = 0
        self.wShowWindow = None


def _patch_simulated_windows():
    """Patch the Windows-only subprocess constants so the hidden-window branch
    can be exercised on Linux/macOS test runners too."""
    return (
        patch.object(process_runner.os, "name", "nt"),
        patch.object(process_runner.subprocess, "CREATE_NO_WINDOW", 0x08000000, create=True),
        patch.object(process_runner.subprocess, "STARTUPINFO", _FakeStartupInfo, create=True),
        patch.object(process_runner.subprocess, "STARTF_USESHOWWINDOW", 0x00000001, create=True),
        patch.object(process_runner.subprocess, "SW_HIDE", 0, create=True),
    )


class HiddenWindowKwargsTests(unittest.TestCase):
    def test_non_windows_returns_empty_mapping(self):
        with patch.object(process_runner.os, "name", "posix"):
            self.assertEqual(process_runner._hidden_window_popen_kwargs(), {})

    def test_windows_returns_creationflags_and_startupinfo(self):
        patches = _patch_simulated_windows()
        with patches[0], patches[1], patches[2], patches[3], patches[4]:
            kwargs = process_runner._hidden_window_popen_kwargs()

        self.assertEqual(kwargs["creationflags"], 0x08000000)
        startupinfo = kwargs["startupinfo"]
        self.assertIsInstance(startupinfo, _FakeStartupInfo)
        self.assertTrue(startupinfo.dwFlags & 0x00000001)
        self.assertEqual(startupinfo.wShowWindow, 0)


class RunPowershellProcessTests(unittest.TestCase):
    # These tests patch _hidden_window_popen_kwargs directly instead of
    # os.name. Forcing os.name to a different platform would break pathlib
    # (Path cannot instantiate PosixPath on Windows), and patching the helper
    # keeps the forwarding assertions independent of the host platform.
    def test_passes_hidden_window_args_to_popen_on_windows(self):
        startupinfo = _FakeStartupInfo()
        startupinfo.dwFlags |= 0x00000001
        startupinfo.wShowWindow = 0
        hidden_kwargs = {"creationflags": 0x08000000, "startupinfo": startupinfo}

        with patch.object(
            process_runner, "_hidden_window_popen_kwargs", return_value=dict(hidden_kwargs)
        ), patch.object(
            process_runner.subprocess, "Popen", return_value=_FakeProcess()
        ) as popen:
            process_runner._run_powershell_process(
                ["powershell", "-NoProfile", "-Command", "echo hi"],
                stop_token="test_token",
                timeout=5,
            )

        kwargs = popen.call_args.kwargs
        self.assertEqual(kwargs["creationflags"], 0x08000000)
        self.assertIs(kwargs["startupinfo"], startupinfo)

    def test_does_not_pass_window_args_on_non_windows(self):
        with patch.object(
            process_runner, "_hidden_window_popen_kwargs", return_value={}
        ), patch.object(
            process_runner.subprocess, "Popen", return_value=_FakeProcess()
        ) as popen:
            process_runner._run_powershell_process(
                ["powershell", "-NoProfile", "-Command", "echo hi"],
                stop_token="test_token",
                timeout=5,
            )

        kwargs = popen.call_args.kwargs
        self.assertNotIn("creationflags", kwargs)
        self.assertNotIn("startupinfo", kwargs)

    def test_preserves_existing_popen_behavior(self):
        with patch.object(
            process_runner, "_hidden_window_popen_kwargs", return_value={}
        ), patch.object(
            process_runner.subprocess, "Popen", return_value=_FakeProcess()
        ) as popen:
            result = process_runner._run_powershell_process(
                ["powershell", "-NoProfile", "-Command", "echo hi"],
                env={"EXTRA_VAR": "value"},
                stop_token="test_token",
                timeout=5,
            )

        kwargs = popen.call_args.kwargs
        self.assertEqual(kwargs["stdout"], subprocess.PIPE)
        self.assertEqual(kwargs["stderr"], subprocess.PIPE)
        self.assertTrue(kwargs["text"])
        self.assertEqual(kwargs["encoding"], "utf-8")
        self.assertEqual(kwargs["errors"], "replace")
        self.assertEqual(kwargs["env"]["PYTHONUTF8"], "1")
        self.assertEqual(kwargs["env"]["PYTHONIOENCODING"], "utf-8")
        self.assertEqual(kwargs["env"]["EXTRA_VAR"], "value")
        self.assertIn("CODEX_STOP_PATH", kwargs["env"])
        self.assertEqual(result.returncode, 0)
        self.assertEqual(result.stdout, "stdout-data")
        self.assertEqual(result.stderr, "stderr-data")


class PowershellCommandTests(unittest.TestCase):
    def _capture_command(self, runner, *args, **kwargs):
        captured: dict[str, object] = {}

        def fake_run(command, **run_kwargs):
            captured["command"] = command
            captured["kwargs"] = run_kwargs
            return subprocess.CompletedProcess(command, 0, "", "")

        with patch.object(process_runner, "_run_powershell_process", side_effect=fake_run):
            runner(*args, **kwargs)
        return captured

    def _assert_hidden_window_style(self, command):
        self.assertEqual(command[0], "powershell")
        self.assertIn("-WindowStyle", command)
        index = command.index("-WindowStyle")
        self.assertEqual(command[index + 1], "Hidden")

    def test_run_powershell_script_includes_hidden_window_style(self):
        captured = self._capture_command(
            process_runner.run_powershell_script, "Write-Output 'hi'"
        )
        self._assert_hidden_window_style(captured["command"])
        self.assertIn("-Command", captured["command"])

    def test_run_powershell_file_includes_hidden_window_style(self):
        captured = self._capture_command(
            process_runner.run_powershell_file,
            "C:/tmp/script.ps1",
            arguments=["-DocxPath", "C:/tmp/doc.docx"],
        )
        command = captured["command"]
        self._assert_hidden_window_style(command)
        self.assertIn("-File", command)
        self.assertIn("-DocxPath", command)


if __name__ == "__main__":
    unittest.main()
