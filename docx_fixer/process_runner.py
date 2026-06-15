from __future__ import annotations

import os
import subprocess
import tempfile
import time
from pathlib import Path

from .stop_controller import StopController


def _ps_single_quoted(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def _hidden_window_popen_kwargs() -> dict[str, object]:
    """Build Popen kwargs that hide the PowerShell console window on Windows.

    On Windows the launched ``powershell.exe`` would otherwise flash a black
    console window when the GUI/exe is double-clicked. ``CREATE_NO_WINDOW``
    plus a ``STARTUPINFO`` with ``STARTF_USESHOWWINDOW`` / ``SW_HIDE`` keeps it
    hidden without affecting stdout/stderr capture.

    The relevant ``subprocess`` constants only exist on Windows, so on other
    platforms (and during Linux/macOS test runs) an empty mapping is returned
    and ``Popen`` behaves exactly as before.
    """
    if os.name != "nt":
        return {}

    kwargs: dict[str, object] = {}

    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    if creation_flags:
        kwargs["creationflags"] = creation_flags

    startupinfo_cls = getattr(subprocess, "STARTUPINFO", None)
    if startupinfo_cls is not None:
        startupinfo = startupinfo_cls()
        startupinfo.dwFlags |= getattr(subprocess, "STARTF_USESHOWWINDOW", 0)
        startupinfo.wShowWindow = getattr(subprocess, "SW_HIDE", 0)
        kwargs["startupinfo"] = startupinfo

    return kwargs


def _run_powershell_process(
    command: list[str],
    *,
    env: dict[str, str] | None = None,
    stop: StopController | None = None,
    timeout: float = 180,
    stop_grace_seconds: float = 8,
    stop_token: str,
) -> subprocess.CompletedProcess:
    stop_path = Path(tempfile.gettempdir()) / f"{stop_token}_{time.time_ns()}.flag"

    def signal_stop() -> None:
        try:
            stop_path.write_text("stop", encoding="utf-8")
        except Exception:
            pass

    process_env = os.environ.copy()
    if env:
        process_env.update(env)
    process_env["CODEX_STOP_PATH"] = str(stop_path)
    process_env["PYTHONIOENCODING"] = "utf-8"
    process_env["PYTHONUTF8"] = "1"

    if stop is not None:
        stop.register_stop_callback(signal_stop)

    process = subprocess.Popen(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=process_env,
        **_hidden_window_popen_kwargs(),
    )

    start_time = time.monotonic()
    stop_started: float | None = None

    try:
        while process.poll() is None:
            if stop is not None and stop.is_stopped():
                signal_stop()
                if stop_started is None:
                    stop_started = time.monotonic()
                elif time.monotonic() - stop_started > stop_grace_seconds:
                    process.terminate()

            if timeout and time.monotonic() - start_time > timeout:
                process.terminate()
                break

            time.sleep(0.1)

        try:
            stdout, stderr = process.communicate(timeout=5)
        except subprocess.TimeoutExpired:
            process.kill()
            stdout, stderr = process.communicate()

        if stop is not None:
            stop.check()

        return subprocess.CompletedProcess(
            process.args,
            process.returncode,
            stdout,
            stderr,
        )
    finally:
        if stop is not None:
            stop.unregister_stop_callback(signal_stop)
        try:
            stop_path.unlink()
        except FileNotFoundError:
            pass
        except Exception:
            pass


def run_powershell_script(
    script: str,
    *,
    stop: StopController | None = None,
    timeout: float = 180,
    stop_grace_seconds: float = 8,
) -> subprocess.CompletedProcess:
    wrapped_script = f"""
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
[Console]::InputEncoding = $utf8NoBom
$CodexStopPath = $env:CODEX_STOP_PATH
function Test-CodexStop {{
    return [System.IO.File]::Exists($CodexStopPath)
}}
{script}
"""
    return _run_powershell_process(
        [
            "powershell",
            "-NoProfile",
            "-WindowStyle",
            "Hidden",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            wrapped_script,
        ],
        stop=stop,
        timeout=timeout,
        stop_grace_seconds=stop_grace_seconds,
        stop_token="docx_fixer_stop_script",
    )


def run_powershell_file(
    script_path: str | Path,
    *,
    arguments: list[str] | None = None,
    stop: StopController | None = None,
    timeout: float = 180,
    stop_grace_seconds: float = 8,
) -> subprocess.CompletedProcess:
    return _run_powershell_process(
        [
            "powershell",
            "-NoProfile",
            "-WindowStyle",
            "Hidden",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_path),
            *(arguments or []),
        ],
        stop=stop,
        timeout=timeout,
        stop_grace_seconds=stop_grace_seconds,
        stop_token=f"docx_fixer_stop_{Path(script_path).stem}",
    )
