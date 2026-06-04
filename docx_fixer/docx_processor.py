from __future__ import annotations

import json
import re
import shutil
import tempfile
import time
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from .constants import NS
from .exceptions import ProcessStopped
from .models import ProcessOptions, ProcessSummary
from .numbering import (
    apply_numbering_outline_format,
    apply_styles_outline_format_to_root,
    build_numbering_format_lookup,
    build_numbering_level_lookup,
    build_style_numbering_lookup,
)
from .outline import (
    fix_outline_paragraphs,
    force_all_paragraphs_to_body_outline_level,
    remove_all_outline_levels_from_any_root,
)
from .path_utils import is_same_file_path
from .process_runner import run_powershell_file, run_powershell_script
from .stop_controller import StopController
from .style_resolver import build_style_font_size_lookup
from .table_format import process_table, table_cell_count, table_column_count
from .xml_utils import qn, remove_character_indent_attrs_from_root

POINTS_PER_CM = 28.3464567
WORD_COM_TIMEOUT_SECONDS = 120
WORD_COM_TEMP_DIR_NAME = "wfix"


def create_word_application_for_com_fix(logs: list[str]):
    try:
        import win32com.client  # type: ignore[import-not-found]

        return win32com.client.DispatchEx("Word.Application")
    except AttributeError as exc:
        if "CLSIDToClassMap" not in str(exc):
            raise

        logs.append(f"WORD_COM_RETRY_CLEAR_GEN_PY reason={exc!r}")

        try:
            import shutil
            import win32com.client.gencache as gencache  # type: ignore[import-not-found]

            gen_path = gencache.GetGeneratePath()
            logs.append(f"WORD_COM_GEN_PY_PATH path={gen_path}")
            shutil.rmtree(gen_path, ignore_errors=True)
            try:
                gencache.Rebuild()
            except Exception as rebuild_exc:
                logs.append(f"WORD_COM_GEN_PY_REBUILD_FAILED reason={rebuild_exc!r}")
        except Exception as cleanup_exc:
            logs.append(f"WORD_COM_GEN_PY_CLEANUP_FAILED reason={cleanup_exc!r}")

        try:
            import win32com.client  # type: ignore[import-not-found]

            word = win32com.client.DispatchEx("Word.Application")
            logs.append("WORD_COM_RETRY_AFTER_CLEAR_GEN_PY status=ok")
            return word
        except Exception as retry_exc:
            logs.append(f"WORD_COM_RETRY_AFTER_CLEAR_GEN_PY status=failed reason={retry_exc!r}")

        try:
            from win32com.client import dynamic  # type: ignore[import-not-found]

            word = dynamic.Dispatch("Word.Application")
            logs.append("WORD_COM_DYNAMIC_DISPATCH status=ok")
            return word
        except Exception as dynamic_exc:
            logs.append(f"WORD_COM_DYNAMIC_DISPATCH status=failed reason={dynamic_exc!r}")
            raise


def should_process_part(name: str) -> bool:
    if name == "word/document.xml":
        return True
    if name.startswith("word/header") and name.endswith(".xml"):
        return True
    if name.startswith("word/footer") and name.endswith(".xml"):
        return True
    if name in {"word/footnotes.xml", "word/endnotes.xml"}:
        return True
    return False


def should_remove_outline_part(name: str) -> bool:
    if should_process_part(name):
        return True
    if name in {"word/styles.xml", "word/numbering.xml"}:
        return True
    return False


def should_force_body_outline_part(name: str) -> bool:
    return should_process_part(name)


def should_sanitize_indent_unit_part(name: str) -> bool:
    if name == "word/document.xml":
        return True
    if name == "word/styles.xml":
        return True
    if name == "word/numbering.xml":
        return True
    if name.startswith("word/header") and name.endswith(".xml"):
        return True
    if name.startswith("word/footer") and name.endswith(".xml"):
        return True
    if name in {"word/footnotes.xml", "word/endnotes.xml"}:
        return True
    return False


def should_fix_paragraph_part(name: str) -> bool:
    """
    文件階層縮排只處理本文。

    頁首頁尾常有頁碼；頁碼文字可能只是「1」「2」這種數字，
    若拿同一套文件編號規則判斷，會被誤認為第 3 階編號，
    導致置中頁碼被套用 left/hanging 縮排而偏掉。
    """
    return name == "word/document.xml"


def get_word_table_start_pages(input_docx: Path, stop: StopController | None = None) -> list[int | None]:
    wd_collapse_start = 1
    wd_active_end_page_number = 3

    word = None
    doc = None
    pages: list[int | None] = []
    com_failed = False

    try:
        word = create_word_application_for_com_fix([])
        word.Visible = False
        doc = word.Documents.Open(
            str(input_docx.resolve()),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        doc.Repaginate()

        for table in doc.Tables:
            if stop:
                stop.check()

            try:
                table_range = table.Range.Duplicate
                table_range.Collapse(wd_collapse_start)
                pages.append(int(table_range.Information(wd_active_end_page_number)))
            except Exception:
                pages.append(None)
    except Exception:
        com_failed = True
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

    if com_failed:
        return get_word_table_start_pages_with_powershell(input_docx, stop=stop)

    return pages


def get_word_table_start_pages_with_powershell(
    input_docx: Path,
    stop: StopController | None = None,
) -> list[int | None]:
    path_literal = "'" + str(input_docx.resolve()).replace("'", "''") + "'"
    script = f"""
$ErrorActionPreference = 'Stop'
$path = {path_literal}
$word = $null
$doc = $null
$pages = New-Object System.Collections.Generic.List[object]
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($path, $false, $true, $false)
    $doc.Repaginate()
    foreach ($table in $doc.Tables) {{
        if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
        $range = $table.Range.Duplicate
        $range.Collapse(1)
        $pages.Add([int]$range.Information(3))
    }}
    $pages | ConvertTo-Json -Compress
}} catch {{
    "[]"
}} finally {{
    if ($doc -ne $null) {{ $doc.Close($false) | Out-Null }}
    if ($word -ne $null) {{ $word.Quit() | Out-Null }}
}}
"""

    try:
        completed = run_powershell_script(script, stop=stop, timeout=120)
    except ProcessStopped:
        raise
    except Exception:
        return []

    output = completed.stdout.strip()
    if not output:
        return []

    try:
        parsed = json.loads(output)
    except json.JSONDecodeError:
        return []

    if isinstance(parsed, int):
        return [parsed]
    if not isinstance(parsed, list):
        return []

    pages: list[int | None] = []
    for page in parsed:
        try:
            pages.append(int(page))
        except (TypeError, ValueError):
            pages.append(None)
    return pages


def get_rendered_table_start_pages(root) -> list[int]:
    pages: list[int] = []
    current_page = 1

    for element in root.iter():
        if element.tag == qn("tbl"):
            pages.append(current_page)
        elif element.tag == qn("lastRenderedPageBreak"):
            current_page += 1

    return pages


def normalize_word_paragraph_text(text: str) -> str:
    cleaned = (text or "").replace("\r", " ").replace("\x07", " ").replace("\x0b", " ")
    return re.sub(r"\s+", " ", cleaned).strip()


def paragraph_text_matches_preview(actual_text: str, preview: str) -> bool:
    normalized_actual = normalize_word_paragraph_text(actual_text)
    normalized_preview = normalize_word_paragraph_text(preview)
    if not normalized_preview:
        return False
    return normalized_actual.startswith(normalized_preview)


def find_word_paragraph_index_for_record(paragraph_texts: list[str], record: dict[str, object]) -> int | None:
    target_index = int(record.get("paragraph_index") or 0)
    preview = str(record.get("text_preview") or "")

    if 1 <= target_index <= len(paragraph_texts):
        if paragraph_text_matches_preview(paragraph_texts[target_index - 1], preview):
            return target_index

    for index, text in enumerate(paragraph_texts, start=1):
        if paragraph_text_matches_preview(text, preview):
            return index

    return None


def _safe_com_attr(obj, name: str):
    try:
        return getattr(obj, name)
    except Exception:
        return None


def _points_to_cm(value) -> float | None:
    try:
        return float(value) / POINTS_PER_CM
    except (TypeError, ValueError):
        return None


def _format_optional_cm(value: float | None) -> str:
    if value is None:
        return "None"
    return f"{value:.2f}"


def _verify_and_fix_body_indents_with_word_com_in_process(
    output_docx: Path,
    body_indent_records: list[dict[str, object]],
) -> list[str]:
    if not body_indent_records:
        return ["WORD_COM_BODY_INDENT_FIX_SKIPPED reason=no_records"]

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        return ["WORD_COM_BODY_INDENT_FIX_SKIPPED reason=win32com_unavailable"]

    word = None
    doc = None
    logs: list[str] = []
    try:
        word = create_word_application_for_com_fix(logs)
        word.Visible = False
        doc = word.Documents.Open(
            str(output_docx.resolve()),
            ReadOnly=False,
            AddToRecentFiles=False,
            Visible=False,
        )

        paragraphs = [doc.Paragraphs(i).Range.Text for i in range(1, doc.Paragraphs.Count + 1)]
        for record in body_indent_records:
            match_index = find_word_paragraph_index_for_record(paragraphs, record)
            preview = str(record.get("text_preview") or "")
            kind = str(record.get("kind") or "body")
            expected_left_cm = float(record.get("expected_left_cm") or 0.0)
            expected_left_points = float(record.get("expected_left_points") or 0.0)
            expected_firstline_cm = float(record.get("expected_firstline_cm") or 0.0)
            expected_firstline_points = float(record.get("expected_firstline_points") or 0.0)

            if match_index is None:
                logs.append(
                    "WORD_COM_BODY_INDENT_FIX: "
                    f"paragraph_index={record.get('paragraph_index')}; "
                    f"text={preview!r}; expected_left_cm={expected_left_cm:.2f}; status=not_found"
                )
                continue

            paragraph = doc.Paragraphs(match_index)
            pf = paragraph.Format

            before_left_pt = _safe_com_attr(pf, "LeftIndent")
            before_left_cm = _points_to_cm(before_left_pt)
            before_first_line_pt = _safe_com_attr(pf, "FirstLineIndent")
            before_first_line_cm = _points_to_cm(before_first_line_pt)
            before_char_left = _safe_com_attr(pf, "CharacterUnitLeftIndent")
            before_char_first = _safe_com_attr(pf, "CharacterUnitFirstLineIndent")

            try:
                pf.CharacterUnitLeftIndent = 0
            except Exception:
                pass
            try:
                pf.CharacterUnitFirstLineIndent = 0
            except Exception:
                pass

            pf.LeftIndent = expected_left_points
            pf.FirstLineIndent = expected_firstline_points

            actual_left_pt = _safe_com_attr(pf, "LeftIndent")
            actual_left_cm = _points_to_cm(actual_left_pt)
            actual_first_line_pt = _safe_com_attr(pf, "FirstLineIndent")
            actual_char_left = _safe_com_attr(pf, "CharacterUnitLeftIndent")
            actual_char_first = _safe_com_attr(pf, "CharacterUnitFirstLineIndent")
            first_line_indent_cm = _points_to_cm(actual_first_line_pt)
            status = (
                "ok"
                if actual_left_cm is not None and abs(actual_left_cm - expected_left_cm) <= 0.02
                and first_line_indent_cm is not None
                and abs(first_line_indent_cm - expected_firstline_cm) <= 0.02
                else "mismatch"
            )

            logs.append(
                "WORD_COM_INDENT_VERIFY: "
                f"paragraph_index={record.get('paragraph_index')}; "
                f"matched_paragraph_index={match_index}; "
                f"text_preview={preview!r}; "
                f"kind={kind}; "
                f"level={record.get('level')}; "
                f"expected_number_start_cm={record.get('expected_number_start_cm')}; "
                f"expected_hanging_cm={record.get('expected_hanging_cm')}; "
                f"expected_heading_left_cm={record.get('expected_heading_left_cm')}; "
                f"expected_body_left_cm={record.get('expected_body_left_cm')}; "
                f"xml_written_left_cm={record.get('xml_written_left_cm')}; "
                f"xml_written_hanging_cm={record.get('xml_written_hanging_cm')}; "
                f"word_opened_left_cm={_format_optional_cm(before_left_cm)}; "
                f"word_opened_firstline_cm={_format_optional_cm(before_first_line_cm)}; "
                f"final_left_cm={_format_optional_cm(actual_left_cm)}; "
                f"final_firstline_cm={_format_optional_cm(first_line_indent_cm)}; "
                f"second_fix=yes; "
                f"status={status}"
            )
            logs.append(
                "WORD_COM_BODY_INDENT_FIX: "
                f"paragraph_index={record.get('paragraph_index')}; "
                f"matched_paragraph_index={match_index}; "
                f"text={preview!r}; "
                f"expected_left_cm={expected_left_cm:.2f}; "
                f"before_left_cm={_format_optional_cm(before_left_cm)}; "
                f"after_left_cm={_format_optional_cm(actual_left_cm)}; "
                f"before_char_left={before_char_left}; "
                f"after_char_left={actual_char_left}; "
                f"before_char_first={before_char_first}; "
                f"after_char_first={actual_char_first}; "
                f"first_line_indent_cm={_format_optional_cm(first_line_indent_cm)}; "
                f"status={status}"
            )

        doc.Save()
    except Exception as exc:
        retry_logs = [
            log
            for log in logs
            if log.startswith("WORD_COM_RETRY_")
            or log.startswith("WORD_COM_GEN_PY_")
            or log.startswith("WORD_COM_DYNAMIC_DISPATCH")
        ]
        retry_reason = f"; retries={' | '.join(retry_logs)}" if retry_logs else ""
        logs.append(f"WORD_COM_BODY_INDENT_FIX_SKIPPED reason={type(exc).__name__}:{exc}{retry_reason}")
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

    return logs


def _parse_powershell_log_output(output: str) -> list[str]:
    text = output.strip()
    if not text:
        return []

    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        return []

    if isinstance(parsed, str):
        return [parsed]
    if not isinstance(parsed, list):
        return []

    return [str(item) for item in parsed]


def _parse_plain_log_output(output: str) -> list[str]:
    return [line.strip() for line in output.splitlines() if line.strip()]


def _collect_word_com_powershell_logs(stdout: str, result_path: Path) -> list[str]:
    logs = _parse_powershell_log_output(stdout)
    if logs:
        return logs

    logs = _parse_plain_log_output(stdout)
    if logs:
        return logs

    if result_path.exists():
        try:
            return _parse_plain_log_output(result_path.read_text(encoding="utf-8"))
        except Exception:
            return []

    return []


def _word_com_temp_root() -> Path:
    root = Path(tempfile.gettempdir()) / WORD_COM_TEMP_DIR_NAME
    root.mkdir(parents=True, exist_ok=True)
    return root


def _word_com_temp_paths() -> tuple[Path, Path, Path, Path]:
    root = _word_com_temp_root()
    token = format(time.time_ns() % 0xFFFFFF, "06x")
    return (
        root / f"wfix_{token}.ps1",
        root / f"wfix_{token}.json",
        root / f"wfix_{token}.docx",
        root / f"wfix_{token}.result.txt",
    )


def _command_length(parts: list[str]) -> int:
    return sum(len(part) for part in parts) + max(len(parts) - 1, 0)


def _long_path_compatible_str(path: Path) -> str:
    resolved = str(path.resolve())
    if resolved.startswith("\\\\?\\"):
        return resolved
    if len(resolved) >= 240 and Path(resolved).drive:
        return "\\\\?\\" + resolved
    return resolved


def _build_word_com_body_indent_powershell_script_file() -> str:
    return f"""param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [Parameter(Mandatory = $true)][string]$RecordsPath,
    [Parameter(Mandatory = $true)][string]$ResultPath
)

$ErrorActionPreference = 'Stop'
$pointPerCm = {POINTS_PER_CM}
$word = $null
$doc = $null
$processed = 0
$ok = 0
$mismatch = 0
$notFound = 0
$errors = 0

function Add-Log([string]$msg) {{
    Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value $msg
    Write-Output $msg
}}

function Test-CodexStop {{
    $stopPath = $env:CODEX_STOP_PATH
    if ([string]::IsNullOrWhiteSpace($stopPath)) {{ return $false }}
    return [System.IO.File]::Exists($stopPath)
}}

function Normalize-WordParagraphText([string]$text) {{
    if ($null -eq $text) {{ return '' }}
    $cleaned = $text.Replace("`r", " ").Replace([string][char]7, " ").Replace([string][char]11, " ")
    return ([regex]::Replace($cleaned, '\\s+', ' ')).Trim()
}}

function Paragraph-TextMatchesPreview([string]$actual, [string]$preview) {{
    $normalizedActual = Normalize-WordParagraphText $actual
    $normalizedPreview = Normalize-WordParagraphText $preview
    if ([string]::IsNullOrWhiteSpace($normalizedPreview)) {{ return $false }}
    return $normalizedActual.StartsWith($normalizedPreview)
}}

function Format-OptionalCm($value) {{
    if ($null -eq $value) {{ return 'None' }}
    return ('{{0:F2}}' -f [double]$value)
}}

function Get-RecordDouble($record, [string]$name, [double]$defaultValue) {{
    try {{
        $property = $record.PSObject.Properties[$name]
        if ($null -eq $property) {{ return $defaultValue }}
        $value = $property.Value
        if ($null -eq $value) {{ return $defaultValue }}
        return [double]$value
    }} catch {{
        return $defaultValue
    }}
}}

try {{
    Set-Content -LiteralPath $ResultPath -Encoding UTF8 -Value ''
    Add-Log 'WORD_COM_PS_STARTED'
    Add-Log ('DOCX_PATH=' + $DocxPath)
    Add-Log ('RECORDS_PATH=' + $RecordsPath)

    if (-not (Test-Path -LiteralPath $DocxPath)) {{
        Add-Log ('WORD_COM_PS_ERROR missing_docx=' + $DocxPath)
        exit 2
    }}

    if (-not (Test-Path -LiteralPath $RecordsPath)) {{
        Add-Log ('WORD_COM_PS_ERROR missing_records=' + $RecordsPath)
        exit 3
    }}

    $recordsRaw = Get-Content -LiteralPath $RecordsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($recordsRaw -is [System.Array]) {{
        $records = @($recordsRaw)
    }} elseif ($null -eq $recordsRaw) {{
        $records = @()
    }} else {{
        $records = @($recordsRaw)
    }}
    Add-Log ('WORD_COM_PS_RECORDS_LOADED count=' + $records.Count)

    $word = New-Object -ComObject Word.Application
    Add-Log 'WORD_COM_PS_WORD_CREATED'
    $word.Visible = $false
    $doc = $word.Documents.Open($DocxPath, $false, $false, $false)
    Add-Log 'WORD_COM_PS_DOC_OPENED'
    Add-Log 'WORD_COM_PS_BEFORE_LOOP'
    Add-Log ('WORD_COM_PS_PARAGRAPHS_COUNT count=' + $doc.Paragraphs.Count)

    $paragraphs = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {{
        if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
        $paragraphs.Add([string]$doc.Paragraphs.Item($i).Range.Text)
    }}

    foreach ($record in $records) {{
        try {{
            if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}

            $processed += 1
            $targetIndex = 0
            try {{ $targetIndex = [int]$record.paragraph_index }} catch {{}}
            $preview = [string]$record.text_preview
            $kind = [string]$record.kind
            if ([string]::IsNullOrWhiteSpace($kind)) {{ $kind = 'body' }}
            $expectedLeftCm = Get-RecordDouble $record 'expected_left_cm' 0
            $expectedLeftPoints = Get-RecordDouble $record 'expected_left_points' 0
            $expectedFirstLineCm = Get-RecordDouble $record 'expected_firstline_cm' 0
            $expectedFirstLinePoints = Get-RecordDouble $record 'expected_firstline_points' 0
            Add-Log(("WORD_COM_RECORD_BEGIN i={{0}} paragraph_index={{1}} expected_left_cm={{2}} text={{3}}" -f $processed, $record.paragraph_index, $record.expected_left_cm, $preview))

            $matchIndex = $null
            if ($targetIndex -ge 1 -and $targetIndex -le $paragraphs.Count) {{
                if (Paragraph-TextMatchesPreview $paragraphs[$targetIndex - 1] $preview) {{
                    $matchIndex = $targetIndex
                }}
            }}

            if ($null -eq $matchIndex) {{
                for ($j = 0; $j -lt $paragraphs.Count; $j++) {{
                    if (Paragraph-TextMatchesPreview $paragraphs[$j] $preview) {{
                        $matchIndex = $j + 1
                        break
                    }}
                }}
            }}

            if ($null -eq $matchIndex) {{
                $notFound += 1
                Add-Log(("WORD_COM_BODY_INDENT_FIX: i={{0}} paragraph_index={{1}} expected_left_cm={{2}} before_left_cm=None after_left_cm=None status=not_found" -f $processed, $record.paragraph_index, $record.expected_left_cm))
                continue
            }}

            Add-Log(("WORD_COM_RECORD_MATCHED i={{0}} word_index={{1}}" -f $processed, $matchIndex))
            $paragraph = $doc.Paragraphs.Item([int]$matchIndex)
            $pf = $paragraph.Format
            $beforeLeftPt = $null
            $beforeFirstLinePt = $null
            $actualLeftPt = $null
            $actualFirstLinePt = $null
            try {{ $beforeLeftPt = [double]$pf.LeftIndent }} catch {{}}
            try {{ $beforeFirstLinePt = [double]$pf.FirstLineIndent }} catch {{}}

            try {{ $pf.CharacterUnitLeftIndent = 0 }} catch {{
                Add-Log(("WORD_COM_RECORD_CHAR_LEFT_CLEAR_FAILED i={{0}} reason={{1}}" -f $processed, $_.Exception.Message))
            }}
            try {{ $pf.CharacterUnitFirstLineIndent = 0 }} catch {{
                Add-Log(("WORD_COM_RECORD_CHAR_FIRST_CLEAR_FAILED i={{0}} reason={{1}}" -f $processed, $_.Exception.Message))
            }}

            $pf.LeftIndent = $expectedLeftPoints
            $pf.FirstLineIndent = $expectedFirstLinePoints

            try {{ $actualLeftPt = [double]$pf.LeftIndent }} catch {{}}
            try {{ $actualFirstLinePt = [double]$pf.FirstLineIndent }} catch {{}}
            $beforeLeftCm = if ($null -eq $beforeLeftPt) {{ $null }} else {{ [double]$beforeLeftPt / $pointPerCm }}
            $beforeFirstLineCm = if ($null -eq $beforeFirstLinePt) {{ $null }} else {{ [double]$beforeFirstLinePt / $pointPerCm }}
            $afterLeftCm = if ($null -eq $actualLeftPt) {{ $null }} else {{ [double]$actualLeftPt / $pointPerCm }}
            $afterFirstLineCm = if ($null -eq $actualFirstLinePt) {{ $null }} else {{ [double]$actualFirstLinePt / $pointPerCm }}

            if ($null -ne $afterLeftCm -and [math]::Abs($afterLeftCm - $expectedLeftCm) -le 0.02 -and $null -ne $afterFirstLineCm -and [math]::Abs($afterFirstLineCm - $expectedFirstLineCm) -le 0.02) {{
                $status = 'ok'
                $ok += 1
            }} else {{
                $status = 'mismatch'
                $mismatch += 1
            }}

            Add-Log(("WORD_COM_INDENT_VERIFY: paragraph_index={{0}}; matched_paragraph_index={{1}}; text_preview={{2}}; kind={{3}}; level={{4}}; expected_number_start_cm={{5}}; expected_hanging_cm={{6}}; expected_heading_left_cm={{7}}; expected_body_left_cm={{8}}; xml_written_left_cm={{9}}; xml_written_hanging_cm={{10}}; word_opened_left_cm={{11}}; word_opened_firstline_cm={{12}}; final_left_cm={{13}}; final_firstline_cm={{14}}; second_fix=yes; status={{15}}" -f $record.paragraph_index, $matchIndex, $preview, $kind, $record.level, $record.expected_number_start_cm, $record.expected_hanging_cm, $record.expected_heading_left_cm, $record.expected_body_left_cm, $record.xml_written_left_cm, $record.xml_written_hanging_cm, (Format-OptionalCm $beforeLeftCm), (Format-OptionalCm $beforeFirstLineCm), (Format-OptionalCm $afterLeftCm), (Format-OptionalCm $afterFirstLineCm), $status))
            Add-Log(("WORD_COM_BODY_INDENT_FIX: i={{0}} paragraph_index={{1}} expected_left_cm={{2}} before_left_cm={{3}} after_left_cm={{4}} status={{5}}" -f $processed, $record.paragraph_index, $record.expected_left_cm, (Format-OptionalCm $beforeLeftCm), (Format-OptionalCm $afterLeftCm), $status))
        }} catch {{
            $errors += 1
            Add-Log(("WORD_COM_RECORD_EXCEPTION i={{0}} paragraph_index={{1}} type={{2}} message={{3}}" -f $processed, $record.paragraph_index, $_.Exception.GetType().FullName, $_.Exception.Message))
            Add-Log(("WORD_COM_RECORD_STACK i={{0}} stack={{1}}" -f $processed, $_.ScriptStackTrace))
            continue
        }}
    }}

    Add-Log 'WORD_COM_PS_BEFORE_SAVE'
    $doc.Save()
    Add-Log 'WORD_COM_PS_DOC_SAVED'
    Add-Log(("WORD_COM_BODY_INDENT_FIX_SUMMARY processed={{0}} ok={{1}} mismatch={{2}} not_found={{3}} errors={{4}}" -f $processed, $ok, $mismatch, $notFound, $errors))
    Add-Log 'WORD_COM_PS_DONE'
    exit 0
}} catch {{
    try {{
        Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value ("WORD_COM_PS_EXCEPTION type=" + $_.Exception.GetType().FullName + " message=" + $_.Exception.Message)
        Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value ("WORD_COM_PS_STACK " + $_.ScriptStackTrace)
    }} catch {{}}
    Write-Output ("WORD_COM_PS_EXCEPTION type=" + $_.Exception.GetType().FullName + " message=" + $_.Exception.Message)
    exit 1
}} finally {{
    try {{
        Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value 'WORD_COM_PS_FINALLY_BEGIN'
    }} catch {{}}

    if ($doc -ne $null) {{
        try {{
            $doc.Close($false) | Out-Null
            Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value 'WORD_COM_PS_DOC_CLOSED'
        }} catch {{
            Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value ("WORD_COM_PS_DOC_CLOSE_FAILED " + $_.Exception.Message)
        }}
    }}

    if ($word -ne $null) {{
        try {{
            $word.Quit() | Out-Null
            Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value 'WORD_COM_PS_WORD_QUIT'
        }} catch {{
            Add-Content -LiteralPath $ResultPath -Encoding UTF8 -Value ("WORD_COM_PS_WORD_QUIT_FAILED " + $_.Exception.Message)
        }}
    }}
}}
"""


def verify_and_fix_body_indents_with_word_com(
    output_docx: Path,
    body_indent_records: list[dict[str, object]],
    stop: StopController | None = None,
) -> list[str]:
    if not body_indent_records:
        return ["WORD_COM_BODY_INDENT_FIX_SKIPPED reason=no_records"]

    script_path, records_path, work_docx_path, result_path = _word_com_temp_paths()
    command_parts = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path),
        "-DocxPath",
        str(work_docx_path),
        "-RecordsPath",
        str(records_path),
        "-ResultPath",
        str(result_path),
    ]
    logs = [
        "WORD_COM_BODY_INDENT_FIX_STARTED",
        f"body_indent_records_count={len(body_indent_records)}",
        f"WORD_COM_POWERSHELL_SCRIPT_PATH={script_path}",
        f"WORD_COM_RECORDS_JSON_PATH={records_path}",
        f"WORD_COM_DOCX_WORK_PATH={work_docx_path}",
        f"WORD_COM_RESULT_LOG_PATH={result_path}",
        f"command_length={_command_length(command_parts)}",
        f"records_count={len(body_indent_records)}",
    ]

    try:
        script_path.write_text(_build_word_com_body_indent_powershell_script_file(), encoding="utf-8")
        records_path.write_text(
            json.dumps(body_indent_records, ensure_ascii=False),
            encoding="utf-8",
        )
        shutil.copy2(_long_path_compatible_str(output_docx), _long_path_compatible_str(work_docx_path))
        result_path.write_text("", encoding="utf-8")
        logs.append(f"WORD_COM_POWERSHELL_COMMAND={' '.join(command_parts)}")
        logs.append(f"WORD_COM_POWERSHELL_SCRIPT_EXISTS={script_path.exists()}")
        logs.append(f"WORD_COM_RECORDS_JSON_EXISTS={records_path.exists()}")
        logs.append(f"WORD_COM_DOCX_WORK_EXISTS={work_docx_path.exists()}")

        completed = run_powershell_file(
            script_path,
            arguments=[
                "-DocxPath",
                str(work_docx_path),
                "-RecordsPath",
                str(records_path),
                "-ResultPath",
                str(result_path),
            ],
            stop=stop,
            timeout=WORD_COM_TIMEOUT_SECONDS,
        )

        logs.append(f"WORD_COM_POWERSHELL_RETURN_CODE={completed.returncode}")
        logs.append(f"WORD_COM_POWERSHELL_STDOUT_LEN={len(completed.stdout)}")
        logs.append(f"WORD_COM_POWERSHELL_STDERR_LEN={len(completed.stderr)}")
        logs.append(f"WORD_COM_POWERSHELL_STDERR={completed.stderr.strip()}")
        logs.append(f"WORD_COM_RESULT_LOG_EXISTS={result_path.exists()}")
        result_size = result_path.stat().st_size if result_path.exists() else 0
        logs.append(f"WORD_COM_RESULT_LOG_SIZE={result_size}")

        script_logs = _collect_word_com_powershell_logs(completed.stdout, result_path)
        if script_logs:
            logs.extend(script_logs)
            has_script_failure = any(
                log.startswith("WORD_COM_PS_EXCEPTION")
                or log.startswith("WORD_COM_PS_ERROR")
                or log.startswith("WORD_COM_BODY_INDENT_FIX_SKIPPED")
                for log in script_logs
            )
            if completed.returncode == 1:
                logs.append("WORD_COM_BODY_INDENT_FIX_FAILED_AFTER_PARTIAL_LOGS")
                has_script_failure = True

            if not has_script_failure and completed.returncode == 0:
                try:
                    shutil.copy2(_long_path_compatible_str(work_docx_path), _long_path_compatible_str(output_docx))
                except Exception as exc:
                    logs.append(f"WORD_COM_BODY_INDENT_FIX_SKIPPED reason=copy_back_failed:{type(exc).__name__}:{exc}")
                    logs.append(f"exception_repr={exc!r}")
            elif not any(log.startswith("WORD_COM_BODY_INDENT_FIX_SKIPPED") for log in logs):
                logs.append("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_script_failed")
            return logs

        stderr = completed.stderr.strip()
        reason = stderr if stderr else "empty_output"
        logs.append(f"WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_no_logs:{reason}")
        return logs
    except ProcessStopped:
        raise
    except Exception as exc:
        logs.append(f"WORD_COM_BODY_INDENT_FIX_SKIPPED reason=powershell_runner_failed:{type(exc).__name__}:{exc}")
        logs.append(f"exception_repr={exc!r}")
        return logs
    finally:
        for temp_path in (script_path, records_path, work_docx_path, result_path):
            try:
                temp_path.unlink()
            except FileNotFoundError:
                pass
            except Exception:
                pass


def fix_docx_fast(
    input_docx: str | Path,
    output_docx: str | Path,
    options: ProcessOptions,
    stop: StopController | None = None,
    progress_callback=None,
) -> ProcessSummary:
    input_docx = Path(input_docx)
    output_docx = Path(output_docx)

    if is_same_file_path(input_docx, output_docx):
        raise ValueError("輸出檔案不可與原始檔案相同，避免覆蓋原檔。請選擇另一個檔名或資料夾。")

    if not input_docx.exists():
        raise FileNotFoundError(f"找不到輸入檔案：{input_docx}")

    if input_docx.suffix.lower() != ".docx":
        raise ValueError("目前只支援 .docx 檔案，不支援 .doc。")

    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    summary = ProcessSummary()
    document_table_pages: list[int | None] = []
    if options.fix_table_layout or options.fix_color:
        document_table_pages = get_word_table_start_pages(input_docx, stop=stop)

    with ZipFile(input_docx, "r") as zin, ZipFile(output_docx, "w", ZIP_DEFLATED) as zout:
        numbering_xml = zin.read("word/numbering.xml") if "word/numbering.xml" in zin.namelist() else None
        styles_xml = zin.read("word/styles.xml") if "word/styles.xml" in zin.namelist() else None
        formatted_numbering_xml = (
            apply_numbering_outline_format(numbering_xml, change_logs=summary.numbering_xml_logs)
            if options.fix_paragraph
            else numbering_xml
        )
        numbering_level_lookup = build_numbering_level_lookup(numbering_xml)
        numbering_format_lookup = build_numbering_format_lookup(formatted_numbering_xml)
        style_numbering_lookup = build_style_numbering_lookup(styles_xml)
        style_font_size_lookup = build_style_font_size_lookup(styles_xml)

        items = zin.infolist()
        total_items = max(len(items), 1)

        for item_index, item in enumerate(items):
            if stop:
                stop.check()

            if progress_callback:
                progress_callback(
                    percent=(item_index / total_items) * 100,
                    message=f"讀取：{item.filename}",
                )

            data = zin.read(item.filename)
            root = None

            if options.remove_all_outline_levels and should_remove_outline_part(item.filename):
                if progress_callback:
                    progress_callback(
                        percent=((item_index + 0.25) / total_items) * 100,
                        message=f"{item.filename}：去除所有大綱階層",
                    )
                root = etree.fromstring(data, parser)
                if should_force_body_outline_part(item.filename):
                    force_all_paragraphs_to_body_outline_level(
                        root,
                        stop=stop,
                        summary=summary,
                    )
                else:
                    remove_all_outline_levels_from_any_root(
                        root,
                        stop=stop,
                        summary=summary,
                    )
                data = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

            # 自動編號的縮排與「編號後方 tab/space」主要記在 numbering.xml；
            # 若只改 document.xml 的段落 pPr，Word 仍可能用舊 tab stop 造成留白。
            if item.filename == "word/numbering.xml" and options.fix_paragraph:
                if progress_callback:
                    progress_callback(
                        percent=((item_index + 0.5) / total_items) * 100,
                        message="word/numbering.xml：修正自動編號縮排與後方留白",
                    )
                data = formatted_numbering_xml or data
                if options.remove_all_outline_levels:
                    root = etree.fromstring(data, parser)
                    remove_all_outline_levels_from_any_root(
                        root,
                        stop=stop,
                    )
                    data = etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )

            if item.filename == "word/styles.xml" and options.fix_paragraph:
                if root is None:
                    root = etree.fromstring(data, parser)
                apply_styles_outline_format_to_root(
                    root,
                    numbering_level_lookup=numbering_level_lookup,
                    style_numbering_lookup=style_numbering_lookup,
                    change_logs=summary.numbering_xml_logs,
                )
                data = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

            if should_process_part(item.filename):
                if root is None:
                    root = etree.fromstring(data, parser)
                if item.filename == "word/document.xml" and not document_table_pages:
                    document_table_pages = get_rendered_table_start_pages(root)

                if (
                    (
                        options.fix_paragraph
                        or options.indent_preface_paragraphs
                        or options.outline_preface_paragraphs
                    )
                    and should_fix_paragraph_part(item.filename)
                ):
                    if progress_callback:
                        message = "處理壹、序言前段落"
                        if options.fix_paragraph:
                            message = "處理文件編號段落與大綱階層（跳過目錄）"
                        progress_callback(
                            percent=((item_index + 0.95) / total_items) * 100,
                            message=f"{item.filename}：{message}",
                        )

                    changed_paragraphs = fix_outline_paragraphs(
                        root,
                        include_tables=options.include_tables_in_paragraph,
                        stop=stop,
                        numbering_level_lookup=numbering_level_lookup,
                        numbering_format_lookup=numbering_format_lookup,
                        style_numbering_lookup=style_numbering_lookup,
                        style_font_size_lookup=style_font_size_lookup,
                        change_logs=summary.paragraph_logs,
                        part_name=item.filename,
                        summary=summary,
                        fix_numbered_paragraphs=options.fix_paragraph,
                        indent_preface_paragraphs=options.indent_preface_paragraphs,
                        outline_preface_paragraphs=options.outline_preface_paragraphs,
                    )
                    summary.paragraphs += changed_paragraphs

                if options.fix_table_layout or options.fix_color:
                    tables = root.xpath(".//w:tbl", namespaces=NS)
                    table_count = len(tables)
                    if item.filename == "word/document.xml" and len(document_table_pages) != table_count:
                        rendered_table_pages = get_rendered_table_start_pages(root)
                        if len(rendered_table_pages) == table_count:
                            document_table_pages = rendered_table_pages

                    for table_index, tbl in enumerate(tables, start=1):
                        if stop:
                            stop.check()

                        table_page = None
                        if item.filename == "word/document.xml" and table_index <= len(document_table_pages):
                            table_page = document_table_pages[table_index - 1]
                        elif item.filename == "word/document.xml":
                            table_page = 1

                        if table_page == 1:
                            summary.skipped_first_page_tables += 1
                            continue

                        cell_count = table_cell_count(tbl)
                        if cell_count <= 4:
                            summary.skipped_small_tables += 1
                            continue

                        special_layout = options.fix_table_layout and table_column_count(tbl) < 4
                        changed_to_gray, cleared_colors = process_table(
                            tbl,
                            options,
                            stop=stop,
                            special_layout=special_layout,
                        )
                        summary.changed_to_gray += changed_to_gray
                        summary.cleared_colors += cleared_colors
                        if special_layout:
                            summary.special_autofit_right_tables += 1
                        else:
                            summary.normal_processed_tables += 1

                        if progress_callback and table_count:
                            inner_fraction = table_index / table_count
                            percent = ((item_index + inner_fraction) / total_items) * 100
                            progress_callback(
                                percent=percent,
                                message=f"{item.filename}：處理表格 {table_index}/{table_count}",
                            )

                    summary.tables += table_count

                data = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

            if should_sanitize_indent_unit_part(item.filename):
                if root is None:
                    root = etree.fromstring(data, parser)
                removed_char_indent_attrs = remove_character_indent_attrs_from_root(root)
                if removed_char_indent_attrs:
                    summary.character_indent_attrs_removed += removed_char_indent_attrs
                data = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

            zout.writestr(item, data)

    if options.normalize_with_word_com:
        summary.word_com_body_indent_logs.extend(
            verify_and_fix_body_indents_with_word_com(output_docx, summary.body_indent_records, stop=stop)
        )
    else:
        summary.word_com_body_indent_logs.append("WORD_COM_BODY_INDENT_FIX_SKIPPED reason=disabled")

    if progress_callback:
        progress_callback(percent=100, message="完成")

    return summary
