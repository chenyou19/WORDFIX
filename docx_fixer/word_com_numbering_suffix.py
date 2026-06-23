"""Apply and verify Word auto-numbering trailing characters with Word COM.

Writing ``w:suff="tab"`` into ``numbering.xml`` is not enough for Microsoft Word
to actually show "Tab character" in the "Define New Multilevel List" UI: Word may
ignore or re-interpret the XML. This module is the production step that makes Word
itself adopt the per-level rule by setting ``ListLevel.TrailingCharacter`` (and,
for tab levels, ``ListLevel.TabPosition``) through the Word object model, then
re-opening the saved document to verify Word kept the values.

It is deliberately kept separate from ``word_com_indent`` so numbering-suffix and
body-indent Word COM logic never re-couple. The logical-level rule is the same
single source of truth used everywhere else:
``numbering.TAB_SUFFIX_OUTLINE_LEVELS`` / ``numbering.uses_tab_suffix``.

The actual COM enum values are defined explicitly in the PowerShell script
(``$wdTrailingTab = 0`` / ``$wdTrailingNone = 2``); Python only carries the
``"tab"`` / ``"nothing"`` intent on each record.
"""
from __future__ import annotations

import json
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from .constants import NS, POINTS_PER_CM, TEMPLATE_OUTLINE_INDENTS
from .numbering import (
    build_numbering_format_lookup,
    numbering_pattern_to_outline_level,
    uses_tab_suffix,
)
from .outline import (
    collect_all_toc_paragraph_ids,
    summarize_paragraph_text,
    text_match_prefix,
)
from .process_runner import run_powershell_file
from .stop_controller import ProcessStopped, StopController
from .xml_utils import paragraph_text, qn

# Logical level -> Word trailing-character intent (matches uses_tab_suffix()).
TRAILING_TAB = "tab"
TRAILING_NONE = "nothing"

_RECORD_JSON_PREFIX = "WORD_COM_SUFFIX_RECORD_JSON "
_VERIFY_JSON_PREFIX = "WORD_COM_SUFFIX_VERIFY_JSON "
_DONE_MARKER = "WORD_COM_SUFFIX_PS_DONE"

_SUFFIX_PARTS_TO_PROCESS = ("word/document.xml",)


@dataclass
class NumberingSuffixWordComResult:
    """Outcome of the Word COM trailing-character apply + verify pass."""

    word_com_available: bool = False
    records_total: int = 0
    records_targeted: int = 0
    records_protected_skipped: int = 0
    template_conflicts: int = 0
    com_apply_success: int = 0
    com_apply_failed: int = 0
    com_verify_success: int = 0
    com_verify_failed: int = 0
    expected_tab: int = 0
    expected_nothing: int = 0
    actual_tab: int = 0
    actual_nothing: int = 0
    tab_position_mismatch: int = 0
    format_identity_mismatch: int = 0
    conflict_details: list[str] = field(default_factory=list)
    record_details: list[dict[str, object]] = field(default_factory=list)

    @property
    def verified(self) -> bool:
        """Overall success: Word was used and nothing went wrong."""
        return (
            self.word_com_available
            and self.template_conflicts == 0
            and self.com_apply_failed == 0
            and self.com_verify_failed == 0
            and self.format_identity_mismatch == 0
            and self.tab_position_mismatch == 0
            and self.com_verify_success == self.records_targeted
            and self.records_targeted > 0
        )


def _template_key(record: dict[str, object]) -> tuple[str, int]:
    return str(record.get("abstract_id")), int(record.get("ilvl", -1))


def build_numbering_suffix_word_com_records(
    output_docx: str | Path,
    *,
    protected_numbering_pairs: set[tuple[str, int]] | None = None,
    logs: list[str] | None = None,
) -> list[dict[str, object]]:
    """Build Word COM trailing-character records from the (already XML-fixed) docx.

    Only real automatic-numbering paragraphs whose numFmt + lvlText map to a
    supported logical outline level (0-8) become records. Manual numbering and
    unrecognized formats are skipped. TOC paragraphs and chapter-參 protected
    pairs are still emitted, but flagged ``is_protected`` so the applier never
    changes them and can detect shared-template conflicts.

    The logical level is decided from numFmt + lvlText (never Word's
    ListLevelNumber), consistent with the rest of the pipeline.
    """
    output_docx = Path(output_docx)
    protected_numbering_pairs = protected_numbering_pairs or set()
    records: list[dict[str, object]] = []

    try:
        with ZipFile(output_docx, "r") as zf:
            names = set(zf.namelist())
            numbering_xml = zf.read("word/numbering.xml") if "word/numbering.xml" in names else None
            part_bytes = {
                name: zf.read(name) for name in _SUFFIX_PARTS_TO_PROCESS if name in names
            }
    except Exception as exc:
        if logs is not None:
            logs.append(f"WORD_COM_SUFFIX_RECORDS_SKIPPED reason=read_error:{exc!r}")
        return records

    if not numbering_xml or not part_bytes:
        if logs is not None:
            logs.append("WORD_COM_SUFFIX_RECORDS_SKIPPED reason=no_numbering_or_document")
        return records

    numbering_format_lookup = build_numbering_format_lookup(numbering_xml)
    num_to_abstract_id = _build_num_to_abstract_id(numbering_xml)

    for part_name, data in part_bytes.items():
        try:
            root = etree.fromstring(data)
        except Exception:
            continue

        paragraphs = root.xpath(".//w:p", namespaces=NS)
        toc_ids = collect_all_toc_paragraph_ids(root, paragraphs=paragraphs)

        for paragraph_index, p in enumerate(paragraphs, start=1):
            if p.xpath("ancestor::w:tbl", namespaces=NS):
                continue
            num_id_el = p.find("./w:pPr/w:numPr/w:numId", NS)
            ilvl_el = p.find("./w:pPr/w:numPr/w:ilvl", NS)
            if num_id_el is None:
                continue  # manual numbering / plain text never enters COM suffix
            num_id = num_id_el.get(qn("val"))
            if num_id is None:
                continue
            try:
                ilvl = int(ilvl_el.get(qn("val"))) if ilvl_el is not None else 0
            except (TypeError, ValueError):
                continue

            level_format = numbering_format_lookup.get((num_id, ilvl))
            if not level_format:
                continue
            num_fmt = level_format.get("numFmt")
            lvl_text = level_format.get("lvlText")
            outline_level = numbering_pattern_to_outline_level(num_fmt, lvl_text)
            if outline_level is None or not (0 <= outline_level <= 8):
                continue  # unrecognized format never enters COM suffix

            text = paragraph_text(p)
            if not text or not text.strip():
                continue

            is_protected = id(p) in toc_ids or (num_id, ilvl) in protected_numbering_pairs
            expected_trailing = TRAILING_TAB if uses_tab_suffix(outline_level) else TRAILING_NONE
            spec = TEMPLATE_OUTLINE_INDENTS.get(outline_level)
            expected_tab_pos_twips = (
                int(spec["left"]) if (expected_trailing == TRAILING_TAB and spec is not None) else None
            )

            records.append(
                {
                    "part_name": part_name,
                    "paragraph_index": paragraph_index,
                    "text_preview": summarize_paragraph_text(text, 80),
                    "text_match_prefix": text_match_prefix(text),
                    "num_id": num_id,
                    "abstract_id": num_to_abstract_id.get(num_id, ""),
                    "ilvl": ilvl,
                    "outline_level": outline_level,
                    "num_fmt": num_fmt,
                    "lvl_text": lvl_text,
                    "expected_trailing": expected_trailing,
                    "expected_tab_pos_twips": expected_tab_pos_twips,
                    "is_protected": is_protected,
                }
            )

    if logs is not None:
        targeted = sum(1 for r in records if not r["is_protected"])
        protected = sum(1 for r in records if r["is_protected"])
        logs.append(
            "WORD_COM_SUFFIX_RECORDS_BUILT "
            f"total={len(records)} targeted={targeted} protected={protected}"
        )
    return records


def partition_records_by_template_conflict(
    records: list[dict[str, object]],
) -> tuple[list[dict[str, object]], list[dict[str, object]], list[str]]:
    """Split records into safe-to-apply targets, protected records, and conflicts.

    A Word ListTemplate level is identified by (abstractNumId, ilvl): paragraphs
    that share it share the same Word numbering level, so modifying it would also
    modify everything else on that level. A level is safe to change only when no
    protected (TOC / chapter 參) record shares the key and every target on the key
    agrees on the expected trailing character / tab position.
    """
    by_key: dict[tuple[str, int], dict[str, list[dict[str, object]]]] = {}
    for record in records:
        key = _template_key(record)
        bucket = by_key.setdefault(key, {"target": [], "protected": []})
        bucket["protected" if record.get("is_protected") else "target"].append(record)

    apply_records: list[dict[str, object]] = []
    protected_records = [r for r in records if r.get("is_protected")]
    conflicts: list[str] = []

    for key, bucket in by_key.items():
        targets = bucket["target"]
        if not targets:
            continue
        protected = bucket["protected"]
        expected_trailings = {str(r.get("expected_trailing")) for r in targets}
        expected_tab_positions = {
            r.get("expected_tab_pos_twips")
            for r in targets
            if r.get("expected_trailing") == TRAILING_TAB
        }
        if protected:
            conflicts.append(
                "WORD_COM_NUMBERING_SUFFIX_TEMPLATE_CONFLICT "
                f"abstractNumId={key[0]}; ilvl={key[1]}; reason=shared_with_protected; "
                f"target_paragraphs={[r.get('paragraph_index') for r in targets]}; "
                f"protected_paragraphs={[r.get('paragraph_index') for r in protected]}"
            )
            continue
        if len(expected_trailings) > 1 or len(expected_tab_positions) > 1:
            conflicts.append(
                "WORD_COM_NUMBERING_SUFFIX_TEMPLATE_CONFLICT "
                f"abstractNumId={key[0]}; ilvl={key[1]}; reason=inconsistent_expectation; "
                f"expected_trailings={sorted(expected_trailings)}; "
                f"expected_tab_positions={sorted(str(p) for p in expected_tab_positions)}"
            )
            continue
        apply_records.extend(targets)

    return apply_records, protected_records, conflicts


def _is_openable_docx_package(docx_path: Path) -> bool:
    """Whether the file is a real OPC package Word could open.

    Word refuses a .docx that lacks ``[Content_Types].xml``; many XML-pipeline
    unit tests build such minimal stubs, and they must never launch Word.
    """
    try:
        with ZipFile(docx_path, "r") as zf:
            names = set(zf.namelist())
    except Exception:
        return False
    return "[Content_Types].xml" in names and "word/document.xml" in names


def _build_num_to_abstract_id(numbering_xml: bytes) -> dict[str, str]:
    try:
        root = etree.fromstring(numbering_xml)
    except Exception:
        return {}
    mapping: dict[str, str] = {}
    for num in root.xpath("./w:num", namespaces=NS):
        num_id = num.get(qn("numId"))
        abstract_el = num.find("w:abstractNumId", NS)
        abstract_id = abstract_el.get(qn("val")) if abstract_el is not None else None
        if num_id is not None and abstract_id is not None:
            mapping[num_id] = abstract_id
    return mapping


def _word_com_temp_paths() -> tuple[Path, Path, Path, Path]:
    import tempfile

    base = Path(tempfile.gettempdir())
    stamp = f"wfix_numsuffix_{__import__('time').time_ns()}"
    return (
        base / f"{stamp}.ps1",
        base / f"{stamp}_records.json",
        base / f"{stamp}_work.docx",
        base / f"{stamp}_result.log",
    )


def _parse_record_json_lines(script_logs: list[str], prefix: str) -> dict[int, dict[str, object]]:
    by_index: dict[int, dict[str, object]] = {}
    for log in script_logs:
        if not log.startswith(prefix):
            continue
        try:
            payload = json.loads(log[len(prefix):])
        except json.JSONDecodeError:
            continue
        if isinstance(payload, dict) and "paragraph_index" in payload:
            by_index[int(payload["paragraph_index"])] = payload
    return by_index


def apply_and_verify_numbering_suffixes_with_word_com(
    docx_path: str | Path,
    records: list[dict[str, object]],
    logs: list[str],
    stop: StopController | None = None,
) -> NumberingSuffixWordComResult:
    """Apply per-level trailing characters in Word, save, reopen, and verify.

    ``records`` is the full record set (targets + protected) from
    :func:`build_numbering_suffix_word_com_records`. Conflicting templates are
    detected here and never sent to Word. The Word COM step is the final writer
    of ``word/numbering.xml``; nothing after it may rewrite the suffix/tab XML.
    """
    docx_path = Path(docx_path)
    result = NumberingSuffixWordComResult(records_total=len(records))

    apply_records, protected_records, conflicts = partition_records_by_template_conflict(records)
    result.records_targeted = len(apply_records)
    result.records_protected_skipped = len(protected_records)
    result.template_conflicts = len(conflicts)
    result.conflict_details = list(conflicts)
    for conflict in conflicts:
        logs.append(conflict)

    if not apply_records:
        logs.append("WORD_COM_NUMBERING_SUFFIX_SKIPPED reason=no_target_records")
        return result

    if not _is_openable_docx_package(docx_path):
        # A package Word cannot open (e.g. a unit-test stub without
        # [Content_Types].xml) must not spawn Word at all.
        logs.append("WORD_COM_NUMBERING_SUFFIX_SKIPPED reason=not_an_openable_docx_package")
        return result

    try:
        import win32com.client  # type: ignore[import-not-found]  # noqa: F401
    except ImportError:
        logs.append("WORD_COM_NUMBERING_SUFFIX_SKIPPED reason=win32com_unavailable")
        return result

    script_path, records_path, work_docx_path, result_path = _word_com_temp_paths()
    # Send both target and protected records: Word reads protected levels too so it
    # can confirm they were left untouched, and double-guards against shared keys.
    com_records = apply_records + [{**r, "is_protected": True} for r in protected_records]
    logs.append(
        "WORD_COM_NUMBERING_SUFFIX_STARTED "
        f"targets={len(apply_records)} protected={len(protected_records)} conflicts={len(conflicts)}"
    )

    try:
        script_path.write_text(_build_numbering_suffix_powershell_script(), encoding="utf-8")
        records_path.write_text(json.dumps(com_records, ensure_ascii=False), encoding="utf-8")
        shutil.copy2(str(docx_path), str(work_docx_path))
        result_path.write_text("", encoding="utf-8")

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
            timeout=600,
        )

        script_logs = _collect_logs(completed.stdout, result_path)
        logs.extend(script_logs)
        result.word_com_available = any(log.startswith("WORD_COM_SUFFIX_PS_WORD_CREATED") for log in script_logs)

        if completed.returncode != 0 or not any(log.startswith(_DONE_MARKER) for log in script_logs):
            logs.append(
                f"WORD_COM_NUMBERING_SUFFIX_FAILED returncode={completed.returncode} "
                f"stderr={completed.stderr.strip()!r}"
            )
            # Word ran but did not finish cleanly; verification cannot be trusted.
            _absorb_partial(result, script_logs, apply_records)
            return result

        # Word produced the new numbering.xml: it is now the canonical output.
        shutil.copy2(str(work_docx_path), str(docx_path))
        _absorb_results(result, script_logs, apply_records, protected_records)
    except ProcessStopped:
        raise
    except Exception as exc:
        logs.append(f"WORD_COM_NUMBERING_SUFFIX_FAILED reason={type(exc).__name__}:{exc}")
    finally:
        for temp_path in (script_path, records_path, work_docx_path, result_path):
            try:
                temp_path.unlink()
            except FileNotFoundError:
                pass
            except Exception:
                pass

    logs.append(
        "WORD_COM_NUMBERING_SUFFIX_DONE "
        f"verified={result.verified} apply_success={result.com_apply_success} "
        f"apply_failed={result.com_apply_failed} verify_success={result.com_verify_success} "
        f"verify_failed={result.com_verify_failed} conflicts={result.template_conflicts} "
        f"format_identity_mismatch={result.format_identity_mismatch} "
        f"tab_position_mismatch={result.tab_position_mismatch}"
    )
    return result


def _absorb_partial(
    result: NumberingSuffixWordComResult,
    script_logs: list[str],
    apply_records: list[dict[str, object]],
) -> None:
    apply_by_index = _parse_record_json_lines(script_logs, _RECORD_JSON_PREFIX)
    for record in apply_records:
        result.com_apply_failed += 1
        result.record_details.append(
            {
                **{k: record.get(k) for k in ("paragraph_index", "outline_level", "expected_trailing")},
                "word_com_verified": False,
                **apply_by_index.get(int(record["paragraph_index"]), {}),
            }
        )


def _absorb_results(
    result: NumberingSuffixWordComResult,
    script_logs: list[str],
    apply_records: list[dict[str, object]],
    protected_records: list[dict[str, object]],
) -> None:
    apply_by_index = _parse_record_json_lines(script_logs, _RECORD_JSON_PREFIX)
    verify_by_index = _parse_record_json_lines(script_logs, _VERIFY_JSON_PREFIX)

    for record in apply_records:
        index = int(record["paragraph_index"])
        apply_info = apply_by_index.get(index, {})
        verify_info = verify_by_index.get(index, {})
        expected_trailing = str(record.get("expected_trailing"))
        if expected_trailing == TRAILING_TAB:
            result.expected_tab += 1
        else:
            result.expected_nothing += 1

        apply_status = str(apply_info.get("apply_status", "missing"))
        if apply_status == "applied":
            result.com_apply_success += 1
        else:
            result.com_apply_failed += 1
        if apply_status == "format_identity_mismatch":
            result.format_identity_mismatch += 1

        trailing_after_reopen = verify_info.get("trailing_after_reopen")
        if trailing_after_reopen == "tab":
            result.actual_tab += 1
        elif trailing_after_reopen == "nothing":
            result.actual_nothing += 1

        verified = bool(verify_info.get("verified"))
        if verified:
            result.com_verify_success += 1
        else:
            result.com_verify_failed += 1
        if verify_info.get("tab_position_mismatch"):
            result.tab_position_mismatch += 1

        result.record_details.append(
            {
                "paragraph_index": index,
                "num_id": record.get("num_id"),
                "ilvl": record.get("ilvl"),
                "outline_level": record.get("outline_level"),
                "expected_trailing": expected_trailing,
                "expected_tab_pos_twips": record.get("expected_tab_pos_twips"),
                "word_com_list_level_number": apply_info.get("list_level_number"),
                "word_com_number_format": apply_info.get("number_format"),
                "word_com_number_style": apply_info.get("number_style"),
                "word_com_trailing_before": apply_info.get("trailing_before"),
                "word_com_trailing_after_apply": apply_info.get("trailing_after_apply"),
                "word_com_trailing_after_reopen": verify_info.get("trailing_after_reopen"),
                "word_com_tab_position_before": apply_info.get("tab_position_before"),
                "word_com_tab_position_after_apply": apply_info.get("tab_position_after_apply"),
                "word_com_tab_position_after_reopen": verify_info.get("tab_position_after_reopen"),
                "word_com_verified": verified,
                "apply_status": apply_status,
            }
        )

    for record in protected_records:
        index = int(record["paragraph_index"])
        verify_info = verify_by_index.get(index, {})
        if verify_info.get("protected_changed"):
            result.com_verify_failed += 1
            result.record_details.append(
                {
                    "paragraph_index": index,
                    "outline_level": record.get("outline_level"),
                    "is_protected": True,
                    "protected_changed": True,
                    "word_com_trailing_after_reopen": verify_info.get("trailing_after_reopen"),
                    "word_com_verified": False,
                }
            )


def _collect_logs(stdout: str, result_path: Path) -> list[str]:
    lines = [line.strip() for line in stdout.splitlines() if line.strip()]
    if lines:
        return lines
    if result_path.exists():
        try:
            return [line.strip() for line in result_path.read_text(encoding="utf-8").splitlines() if line.strip()]
        except Exception:
            return []
    return []


def _build_numbering_suffix_powershell_script() -> str:
    return f"""param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [Parameter(Mandatory = $true)][string]$RecordsPath,
    [Parameter(Mandatory = $true)][string]$ResultPath
)

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
[Console]::InputEncoding = $utf8NoBom
$ErrorActionPreference = 'Stop'

# Word WdTrailingCharacter enum, defined explicitly (no magic numbers elsewhere).
$wdTrailingTab = 0
$wdTrailingSpace = 1
$wdTrailingNone = 2
$twipsPerPoint = 20.0
$pointPerCm = {POINTS_PER_CM}

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
    if ($normalizedPreview.EndsWith('...')) {{
        $normalizedPreview = $normalizedPreview.Substring(0, $normalizedPreview.Length - 3).TrimEnd()
    }}
    if ([string]::IsNullOrWhiteSpace($normalizedPreview)) {{ return $false }}
    return $normalizedActual.StartsWith($normalizedPreview)
}}

function Normalize-Format([string]$value) {{
    if ($null -eq $value) {{ return '' }}
    return ([regex]::Replace($value, '\\s+', '')).Trim()
}}

function Locate-Paragraph($doc, $record, [int]$paragraphCount) {{
    $targetIndex = 0
    try {{ $targetIndex = [int]$record.paragraph_index }} catch {{}}
    $matchPrefix = [string]$record.text_match_prefix
    if ([string]::IsNullOrWhiteSpace($matchPrefix)) {{ $matchPrefix = [string]$record.text_preview }}

    if ($targetIndex -ge 1 -and $targetIndex -le $paragraphCount) {{
        $candidate = [string]$doc.Paragraphs.Item($targetIndex).Range.Text
        if (Paragraph-TextMatchesPreview $candidate $matchPrefix) {{ return $targetIndex }}
    }}
    for ($j = 1; $j -le $paragraphCount; $j++) {{
        if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
        $candidate = [string]$doc.Paragraphs.Item($j).Range.Text
        if (Paragraph-TextMatchesPreview $candidate $matchPrefix) {{ return $j }}
    }}
    return $null
}}

function Read-Trailing($listLevel) {{
    try {{
        $value = [int]$listLevel.TrailingCharacter
        if ($value -eq $wdTrailingTab) {{ return 'tab' }}
        if ($value -eq $wdTrailingNone) {{ return 'nothing' }}
        if ($value -eq $wdTrailingSpace) {{ return 'space' }}
        return ('other:' + $value)
    }} catch {{ return 'error' }}
}}

function Read-TabPosition($listLevel) {{
    try {{ return [double]$listLevel.TabPosition }} catch {{ return $null }}
}}

function Get-ListLevel($doc, [int]$wordIndex) {{
    $paragraph = $doc.Paragraphs.Item($wordIndex)
    $lf = $paragraph.Range.ListFormat
    $levelNumber = 0
    try {{ $levelNumber = [int]$lf.ListLevelNumber }} catch {{}}
    $template = $null
    try {{ $template = $lf.ListTemplate }} catch {{}}
    if ($null -eq $template -or $levelNumber -lt 1) {{ return $null }}
    return @{{ level_number = $levelNumber; level = $template.ListLevels($levelNumber) }}
}}

$word = $null
$doc = $null
try {{
    Set-Content -LiteralPath $ResultPath -Encoding UTF8 -Value ''
    Add-Log 'WORD_COM_SUFFIX_PS_STARTED'
    $recordsRaw = Get-Content -LiteralPath $RecordsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($recordsRaw -is [System.Array]) {{ $records = @($recordsRaw) }}
    elseif ($null -eq $recordsRaw) {{ $records = @() }}
    else {{ $records = @($recordsRaw) }}
    Add-Log ('WORD_COM_SUFFIX_PS_RECORDS_LOADED count=' + $records.Count)

    $word = New-Object -ComObject Word.Application
    Add-Log 'WORD_COM_SUFFIX_PS_WORD_CREATED'
    $word.Visible = $false
    try {{ $word.DisplayAlerts = 0 }} catch {{}}
    $doc = $word.Documents.Open($DocxPath, $false, $false, $false)
    $paragraphCount = [int]$doc.Paragraphs.Count
    Add-Log ('WORD_COM_SUFFIX_PS_PARAGRAPHS count=' + $paragraphCount)

    # Build protected (abstractNumId|ilvl) key set; never modify those levels.
    $protectedKeys = @{{}}
    foreach ($r in $records) {{
        $isProtected = $false
        try {{ $isProtected = [bool]$r.is_protected }} catch {{}}
        if ($isProtected) {{ $protectedKeys[(([string]$r.abstract_id) + '|' + ([string]$r.ilvl))] = $true }}
    }}

    foreach ($record in $records) {{
        if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
        $isProtected = $false
        try {{ $isProtected = [bool]$record.is_protected }} catch {{}}
        if ($isProtected) {{ continue }}

        $paragraphIndex = [int]$record.paragraph_index
        $expectedTrailing = [string]$record.expected_trailing
        $ilvl = [int]$record.ilvl
        $lvlText = [string]$record.lvl_text
        $key = (([string]$record.abstract_id) + '|' + ([string]$record.ilvl))

        $emit = [ordered]@{{ paragraph_index = $paragraphIndex; apply_status = 'pending' }}

        $wordIndex = Locate-Paragraph $doc $record $paragraphCount
        if ($null -eq $wordIndex) {{
            $emit.apply_status = 'not_found'
            Add-Log ('WORD_COM_SUFFIX_RECORD_JSON ' + (ConvertTo-Json $emit -Compress -Depth 6))
            continue
        }}
        $info = Get-ListLevel $doc $wordIndex
        if ($null -eq $info) {{
            $emit.apply_status = 'not_a_list_paragraph'
            Add-Log ('WORD_COM_SUFFIX_RECORD_JSON ' + (ConvertTo-Json $emit -Compress -Depth 6))
            continue
        }}
        $listLevel = $info.level
        $levelNumber = [int]$info.level_number
        $numberFormat = ''
        try {{ $numberFormat = [string]$listLevel.NumberFormat }} catch {{}}
        $numberStyle = $null
        try {{ $numberStyle = [int]$listLevel.NumberStyle }} catch {{}}
        $emit.list_level_number = $levelNumber
        $emit.number_format = $numberFormat
        $emit.number_style = $numberStyle
        $emit.trailing_before = (Read-Trailing $listLevel)
        $emit.tab_position_before = (Read-TabPosition $listLevel)

        # Format-identity safety: the Word level number must match the XML ilvl and
        # the Word NumberFormat must match the XML lvlText. Never blind-edit.
        $formatOk = ($levelNumber -eq ($ilvl + 1)) -and ((Normalize-Format $numberFormat) -eq (Normalize-Format $lvlText))
        if (-not $formatOk) {{
            $emit.apply_status = 'format_identity_mismatch'
            Add-Log ('WORD_COM_SUFFIX_RECORD_JSON ' + (ConvertTo-Json $emit -Compress -Depth 6))
            continue
        }}
        if ($protectedKeys.ContainsKey($key)) {{
            $emit.apply_status = 'template_conflict'
            Add-Log ('WORD_COM_NUMBERING_SUFFIX_TEMPLATE_CONFLICT abstractNumId=' + ([string]$record.abstract_id) + '; ilvl=' + ([string]$record.ilvl) + '; reason=shared_with_protected_runtime')
            Add-Log ('WORD_COM_SUFFIX_RECORD_JSON ' + (ConvertTo-Json $emit -Compress -Depth 6))
            continue
        }}

        try {{
            if ($expectedTrailing -eq 'tab') {{
                $listLevel.TrailingCharacter = $wdTrailingTab
                $tabPosTwips = [double]$record.expected_tab_pos_twips
                $listLevel.TabPosition = $tabPosTwips / $twipsPerPoint
            }} else {{
                $listLevel.TrailingCharacter = $wdTrailingNone
            }}
            $emit.trailing_after_apply = (Read-Trailing $listLevel)
            $emit.tab_position_after_apply = (Read-TabPosition $listLevel)
            $emit.apply_status = 'applied'
        }} catch {{
            $emit.apply_status = ('apply_error:' + $_.Exception.Message)
        }}
        Add-Log ('WORD_COM_SUFFIX_RECORD_JSON ' + (ConvertTo-Json $emit -Compress -Depth 6))
    }}

    $doc.Save()
    $doc.Close($false)
    $doc = $null
    Add-Log 'WORD_COM_SUFFIX_PS_SAVED_AND_CLOSED'

    # Reopen and verify Word kept the values.
    $doc = $word.Documents.Open($DocxPath, $false, $false, $false)
    $paragraphCount = [int]$doc.Paragraphs.Count
    foreach ($record in $records) {{
        if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
        $paragraphIndex = [int]$record.paragraph_index
        $isProtected = $false
        try {{ $isProtected = [bool]$record.is_protected }} catch {{}}
        $expectedTrailing = [string]$record.expected_trailing

        $verify = [ordered]@{{ paragraph_index = $paragraphIndex }}
        $wordIndex = Locate-Paragraph $doc $record $paragraphCount
        if ($null -eq $wordIndex) {{
            $verify.verified = $false
            $verify.reason = 'not_found_on_reopen'
            Add-Log ('WORD_COM_SUFFIX_VERIFY_JSON ' + (ConvertTo-Json $verify -Compress -Depth 6))
            continue
        }}
        $info = Get-ListLevel $doc $wordIndex
        if ($null -eq $info) {{
            $verify.verified = $false
            $verify.reason = 'not_a_list_paragraph_on_reopen'
            Add-Log ('WORD_COM_SUFFIX_VERIFY_JSON ' + (ConvertTo-Json $verify -Compress -Depth 6))
            continue
        }}
        $listLevel = $info.level
        $actualTrailing = (Read-Trailing $listLevel)
        $actualTabPos = (Read-TabPosition $listLevel)
        $verify.trailing_after_reopen = $actualTrailing
        $verify.tab_position_after_reopen = $actualTabPos

        if ($isProtected) {{
            # Protected levels must keep whatever they already were.
            $verify.protected_changed = $false
            Add-Log ('WORD_COM_SUFFIX_VERIFY_JSON ' + (ConvertTo-Json $verify -Compress -Depth 6))
            continue
        }}

        $trailingOk = ($actualTrailing -eq $expectedTrailing)
        $tabOk = $true
        if ($expectedTrailing -eq 'tab') {{
            $expectedPoints = [double]$record.expected_tab_pos_twips / $twipsPerPoint
            if ($null -eq $actualTabPos) {{ $tabOk = $false }}
            elseif ([math]::Abs([double]$actualTabPos - $expectedPoints) -gt 0.6) {{ $tabOk = $false }}
        }}
        if (-not $tabOk) {{ $verify.tab_position_mismatch = $true }}
        $verify.verified = ($trailingOk -and $tabOk)
        Add-Log ('WORD_COM_SUFFIX_VERIFY_JSON ' + (ConvertTo-Json $verify -Compress -Depth 6))
    }}

    $doc.Close($false)
    $doc = $null
    Add-Log 'WORD_COM_SUFFIX_PS_DONE'
}} catch {{
    Add-Log ('WORD_COM_SUFFIX_PS_EXCEPTION ' + $_.Exception.Message)
    throw
}} finally {{
    if ($null -ne $doc) {{ try {{ $doc.Close($false) }} catch {{}} }}
    if ($null -ne $word) {{ try {{ $word.Quit() }} catch {{}} }}
}}
"""
