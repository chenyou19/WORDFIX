from __future__ import annotations

import json
import tempfile
import time
from pathlib import Path

from .exceptions import ProcessStopped
from .process_runner import run_powershell_file
from .stop_controller import StopController

WORD_COM_TABLE_AUTOFIT_TIMEOUT_SECONDS = 600
WORD_COM_AUTOFIT_SEQUENCE = "content_then_window"


def _table_autofit_temp_paths() -> tuple[Path, Path]:
    root = Path(tempfile.gettempdir()) / "wfix"
    root.mkdir(parents=True, exist_ok=True)
    token = f"table_autofit_{time.time_ns()}"
    return root / f"{token}.ps1", root / f"{token}_records.json"


def _build_table_autofit_powershell_script() -> str:
    return r"""
param(
    [Parameter(Mandatory=$true)][string]$DocxPath,
    [Parameter(Mandatory=$true)][string]$RecordsPath
)

$ErrorActionPreference = 'Stop'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
[Console]::InputEncoding = $utf8NoBom
$CodexStopPath = $env:CODEX_STOP_PATH
function Test-CodexStop {
    return [System.IO.File]::Exists($CodexStopPath)
}

$wdAutoFitContent = 1
$wdAutoFitWindow = 2
$word = $null
$doc = $null
$applied = 0
$errors = 0
$notFound = 0

try {
    $recordsJson = Get-Content -LiteralPath $RecordsPath -Raw -Encoding UTF8
    $records = @($recordsJson | ConvertFrom-Json)
    Write-Output ("WORD_COM_TABLE_AUTOFIT_STARTED records_count={0}" -f $records.Count)

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($DocxPath, $false, $false, $false)
    $tableCount = [int]$doc.Tables.Count
    Write-Output ("WORD_COM_TABLE_AUTOFIT_DOC_TABLES count={0}" -f $tableCount)

    foreach ($record in $records) {
        if (Test-CodexStop) { throw 'STOPPED_BY_USER' }
        $globalIndex = [int]$record.global_table_index
        try {
            if ($globalIndex -lt 1 -or $globalIndex -gt $tableCount) {
                $notFound += 1
                Write-Output ("WORD_COM_TABLE_AUTOFIT_NOT_FOUND global_table_index={0} doc_table_count={1}" -f $globalIndex, $tableCount)
                continue
            }

            $table = $doc.Tables.Item($globalIndex)
            $table.AutoFitBehavior($wdAutoFitContent)
            if (Test-CodexStop) { throw 'STOPPED_BY_USER' }
            $table.AutoFitBehavior($wdAutoFitWindow)
            $applied += 1
            Write-Output ("WORD_COM_TABLE_AUTOFIT_APPLIED global_table_index={0} sequence=content_then_window" -f $globalIndex)
        } catch {
            $errors += 1
            Write-Output ("WORD_COM_TABLE_AUTOFIT_ERROR global_table_index={0} type={1} message={2}" -f $globalIndex, $_.Exception.GetType().FullName, $_.Exception.Message)
            continue
        }
    }

    $doc.Save()
    Write-Output ("WORD_COM_TABLE_AUTOFIT_SUMMARY applied={0} not_found={1} errors={2}" -f $applied, $notFound, $errors)
    Write-Output 'WORD_COM_TABLE_AUTOFIT_DONE'
    exit 0
} catch {
    Write-Output ("WORD_COM_TABLE_AUTOFIT_EXCEPTION type={0} message={1}" -f $_.Exception.GetType().FullName, $_.Exception.Message)
    exit 1
} finally {
    if ($doc -ne $null) {
        try { $doc.Close($false) | Out-Null } catch {}
    }
    if ($word -ne $null) {
        try { $word.Quit() | Out-Null } catch {}
    }
}
"""


def _parse_applied_indices(logs: list[str]) -> set[int]:
    applied: set[int] = set()
    prefix = "WORD_COM_TABLE_AUTOFIT_APPLIED global_table_index="
    for log in logs:
        if not log.startswith(prefix):
            continue
        value = log[len(prefix):].split(" ", 1)[0]
        try:
            applied.add(int(value))
        except ValueError:
            continue
    return applied


def apply_table_autofit_with_word_com(
    output_docx: str | Path,
    records: list[dict[str, object]],
    stop: StopController | None = None,
) -> tuple[list[str], set[int]]:
    if not records:
        return ["WORD_COM_TABLE_AUTOFIT_SKIPPED reason=no_records"], set()

    script_path, records_path = _table_autofit_temp_paths()
    logs = [
        "WORD_COM_TABLE_AUTOFIT_REQUESTED",
        f"records_count={len(records)}",
        f"sequence={WORD_COM_AUTOFIT_SEQUENCE}",
        f"WORD_COM_TABLE_AUTOFIT_SCRIPT_PATH={script_path}",
        f"WORD_COM_TABLE_AUTOFIT_RECORDS_PATH={records_path}",
    ]

    try:
        script_path.write_text(_build_table_autofit_powershell_script(), encoding="utf-8")
        records_path.write_text(json.dumps(records, ensure_ascii=False), encoding="utf-8")
        completed = run_powershell_file(
            script_path,
            arguments=[
                "-DocxPath",
                str(Path(output_docx).resolve()),
                "-RecordsPath",
                str(records_path),
            ],
            stop=stop,
            timeout=WORD_COM_TABLE_AUTOFIT_TIMEOUT_SECONDS,
        )

        logs.append(f"WORD_COM_TABLE_AUTOFIT_RETURN_CODE={completed.returncode}")
        if completed.stderr.strip():
            logs.append(f"WORD_COM_TABLE_AUTOFIT_STDERR={completed.stderr.strip()}")
        script_logs = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
        logs.extend(script_logs)

        applied_indices = _parse_applied_indices(script_logs)
        if completed.returncode != 0 and not any(
            log.startswith("WORD_COM_TABLE_AUTOFIT_EXCEPTION") for log in script_logs
        ):
            logs.append("WORD_COM_TABLE_AUTOFIT_FAILED reason=powershell_nonzero_exit")
        return logs, applied_indices
    except ProcessStopped:
        raise
    except Exception as exc:
        logs.append(f"WORD_COM_TABLE_AUTOFIT_SKIPPED reason=runner_failed:{type(exc).__name__}:{exc}")
        return logs, set()
    finally:
        for temp_path in (script_path, records_path):
            try:
                temp_path.unlink()
            except FileNotFoundError:
                pass
            except Exception:
                pass
