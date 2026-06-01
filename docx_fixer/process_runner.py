from __future__ import annotations

import subprocess
import tempfile
import time
from pathlib import Path

from .stop_controller import StopController


def _ps_single_quoted(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def run_powershell_script(
    script: str,
    *,
    stop: StopController | None = None,
    timeout: float = 180,
    stop_grace_seconds: float = 8,
) -> subprocess.CompletedProcess:
    stop_path = Path(tempfile.gettempdir()) / f"docx_fixer_stop_{id(script)}_{time.time_ns()}.flag"
    stop_literal = _ps_single_quoted(str(stop_path))
    wrapped_script = f"""
$CodexStopPath = {stop_literal}
function Test-CodexStop {{
    return [System.IO.File]::Exists($CodexStopPath)
}}
{script}
"""

    def signal_stop() -> None:
        try:
            stop_path.write_text("stop", encoding="utf-8")
        except Exception:
            pass

    if stop is not None:
        stop.register_stop_callback(signal_stop)

    process = subprocess.Popen(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", wrapped_script],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
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
