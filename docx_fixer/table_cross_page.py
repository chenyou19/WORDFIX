from __future__ import annotations

import json
from pathlib import Path

from .exceptions import ProcessStopped
from .models import ProcessSummary
from .process_runner import run_powershell_script
from .stop_controller import StopController


def apply_cross_page_stats(summary: ProcessSummary, stats: dict) -> None:
    if stats.get("global_error"):
        summary.failed_cross_page_tables += max(
            summary.tables - summary.skipped_first_page_tables - summary.skipped_small_tables,
            0,
        )
        return

    summary.cross_page_tables += int(stats.get("cross_page_tables", 0) or 0)
    summary.cross_page_resolved_tables += int(stats.get("cross_page_resolved_tables", 0) or 0)
    summary.cross_page_still_split_tables += int(stats.get("cross_page_still_split_tables", 0) or 0)
    summary.adjusted_cell_padding_tables += int(stats.get("adjusted_cell_padding_tables", 0) or 0)
    summary.adjusted_table_spacing_tables += int(stats.get("adjusted_table_spacing_tables", 0) or 0)
    summary.auto_height_tables += int(stats.get("auto_height_tables", 0) or 0)
    summary.moved_next_page_resolved_tables += int(stats.get("moved_next_page_resolved_tables", 0) or 0)
    summary.cannot_avoid_cross_page_tables += int(stats.get("cannot_avoid_cross_page_tables", 0) or 0)
    summary.failed_cross_page_tables += int(stats.get("failed_cross_page_tables", 0) or 0)


def adjust_cross_page_tables(
    output_docx: str | Path,
    summary: ProcessSummary,
    stop: StopController | None = None,
) -> None:
    stats = run_cross_page_adjustment(output_docx, stop=stop)
    apply_cross_page_stats(summary, stats)
    if stop:
        stop.check()


def run_cross_page_adjustment(output_docx: str | Path, stop: StopController | None = None) -> dict:
    path_literal = "'" + str(Path(output_docx).resolve()).replace("'", "''") + "'"
    script = f"""
$ErrorActionPreference = 'Stop'
$path = {path_literal}
$wdActiveEndPageNumber = 3
$wdCollapseStart = 1
$wdLineSpaceSingle = 0
$wdRowHeightAuto = 0
$wdAutoFitWindow = 2
$wdPageBreak = 7
$wdPreferredWidthPoints = 3
$wdMainTextStory = 1
$stats = [ordered]@{{
  global_error = $false
  cross_page_tables = 0
  cross_page_resolved_tables = 0
  cross_page_still_split_tables = 0
  adjusted_cell_padding_tables = 0
  adjusted_table_spacing_tables = 0
  auto_height_tables = 0
  moved_next_page_resolved_tables = 0
  cannot_avoid_cross_page_tables = 0
  failed_cross_page_tables = 0
}}

function Get-TablePages($table) {{
  $startRange = $table.Range.Duplicate
  $startRange.Collapse($wdCollapseStart)
  $startPage = [int]$startRange.Information($wdActiveEndPageNumber)
  $endPos = [Math]::Max($table.Range.Start, $table.Range.End - 1)
  $endRange = $table.Range.Duplicate
  $endRange.SetRange($endPos, $endPos)
  $endPage = [int]$endRange.Information($wdActiveEndPageNumber)
  return @($startPage, $endPage)
}}

function Test-TableCrossPage($table) {{
  $pages = Get-TablePages $table
  return $pages[0] -ne $pages[1]
}}

function Test-TocTable($table) {{
  foreach ($paragraph in $table.Range.Paragraphs) {{
    try {{
      $styleName = [string]$paragraph.Range.Style.NameLocal
      $styleKey = $styleName.Replace(' ', '').Replace('_', '').ToUpperInvariant()
      if ($styleKey -match '^TOC\\d+$' -or $styleKey.StartsWith('TOC') -or $styleName -match '目錄|目录') {{
        return $true
      }}
    }} catch {{}}
  }}
  return $false
}}

function Get-TableColumnCount($table) {{
  try {{
    return [int]$table.Columns.Count
  }} catch {{
    $maxCells = 0
    foreach ($row in $table.Rows) {{
      try {{
        if ($row.Cells.Count -gt $maxCells) {{ $maxCells = [int]$row.Cells.Count }}
      }} catch {{}}
    }}
    return $maxCells
  }}
}}

function Set-ParagraphCompact($table) {{
  foreach ($paragraph in $table.Range.Paragraphs) {{
    $paragraph.Format.SpaceBefore = 0
    $paragraph.Format.SpaceAfter = 0
    $paragraph.Format.LineSpacingRule = $wdLineSpaceSingle
  }}
}}

function Set-CellPadding($table) {{
  $table.TopPadding = 1
  $table.BottomPadding = 1
}}

function Set-RowAutoHeight($table) {{
  foreach ($row in $table.Rows) {{
    $row.AllowBreakAcrossPages = $false
    $row.HeightRule = $wdRowHeightAuto
  }}
}}

function Clear-FloatingTable($table) {{
  try {{ $table.Rows.WrapAroundText = $false }} catch {{}}
}}

function Limit-TableWidth($doc, $table, $columnCount) {{
  if ($columnCount -lt 4) {{ return }}
  try {{
    $usableWidth = $doc.PageSetup.PageWidth - $doc.PageSetup.LeftMargin - $doc.PageSetup.RightMargin
    if ($usableWidth -le 0) {{ return }}
    if ($table.PreferredWidthType -eq $wdPreferredWidthPoints -and $table.PreferredWidth -gt $usableWidth) {{
      $table.AutoFitBehavior($wdAutoFitWindow)
      $table.PreferredWidthType = $wdPreferredWidthPoints
      $table.PreferredWidth = $usableWidth
    }}
  }} catch {{}}
}}

$word = $null
$doc = $null
try {{
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $doc = $word.Documents.Open($path, $false, $false, $false)
  $doc.Repaginate()
  $tableCount = $doc.Tables.Count

  for ($i = 1; $i -le $tableCount; $i++) {{
    if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
    try {{
      $table = $doc.Tables.Item($i)
      if ($table.Range.StoryType -ne $wdMainTextStory) {{ continue }}
      if (Test-TocTable $table) {{ continue }}
      if ($table.Range.Cells.Count -le 4) {{ continue }}

      $pages = Get-TablePages $table
      if ($pages[0] -eq 1) {{ continue }}
      if ($pages[0] -eq $pages[1]) {{ continue }}

      $stats.cross_page_tables++
      $columnCount = Get-TableColumnCount $table
      $spacingAdjusted = $false
      $paddingAdjusted = $false
      $heightAdjusted = $false
      $resolved = $false

      Set-ParagraphCompact $table
      $spacingAdjusted = $true
      $doc.Repaginate()
      if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
      if (-not (Test-TableCrossPage $table)) {{ $resolved = $true }}

      if (-not $resolved) {{
        Set-CellPadding $table
        $paddingAdjusted = $true
        $doc.Repaginate()
        if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
        if (-not (Test-TableCrossPage $table)) {{ $resolved = $true }}
      }}

      if (-not $resolved) {{
        Set-RowAutoHeight $table
        $heightAdjusted = $true
        $doc.Repaginate()
        if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
        if (-not (Test-TableCrossPage $table)) {{ $resolved = $true }}
      }}

      if (-not $resolved) {{
        Clear-FloatingTable $table
        Limit-TableWidth $doc $table $columnCount
        $doc.Repaginate()
        if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
        if (-not (Test-TableCrossPage $table)) {{ $resolved = $true }}
      }}

      if ($spacingAdjusted) {{ $stats.adjusted_table_spacing_tables++ }}
      if ($paddingAdjusted) {{ $stats.adjusted_cell_padding_tables++ }}
      if ($heightAdjusted) {{ $stats.auto_height_tables++ }}

      if ($resolved) {{
        $stats.cross_page_resolved_tables++
        continue
      }}

      try {{
        $moveRange = $table.Range.Duplicate
        $moveRange.Collapse($wdCollapseStart)
        $moveRange.InsertBreak($wdPageBreak)
        $doc.Repaginate()
        if (Test-CodexStop) {{ throw 'STOPPED_BY_USER' }}
        if (-not (Test-TableCrossPage $table)) {{
          $stats.cross_page_resolved_tables++
          $stats.moved_next_page_resolved_tables++
        }} else {{
          $stats.cross_page_still_split_tables++
          $stats.cannot_avoid_cross_page_tables++
        }}
      }} catch {{
        $stats.cross_page_still_split_tables++
        $stats.cannot_avoid_cross_page_tables++
      }}
    }} catch {{
      if ([string]$_ -match 'STOPPED_BY_USER') {{ throw }}
      $stats.failed_cross_page_tables++
      continue
    }}
  }}

  $doc.Save()
}} catch {{
  $stats.global_error = $true
}} finally {{
  if ($doc -ne $null) {{ $doc.Close($false) | Out-Null }}
  if ($word -ne $null) {{ $word.Quit() | Out-Null }}
}}
$stats | ConvertTo-Json -Compress
"""

    try:
        completed = run_powershell_script(script, stop=stop, timeout=180)
    except ProcessStopped:
        raise
    except Exception:
        return {}

    output = completed.stdout.strip()
    if not output:
        return {}

    try:
        parsed = json.loads(output)
    except json.JSONDecodeError:
        return {}

    return parsed if isinstance(parsed, dict) else {}
