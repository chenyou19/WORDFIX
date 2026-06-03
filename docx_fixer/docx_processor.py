from __future__ import annotations

import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from .constants import NS
from .exceptions import ProcessStopped
from .models import ProcessOptions, ProcessSummary
from .numbering import (
    apply_numbering_outline_format,
    build_numbering_level_lookup,
    build_style_numbering_lookup,
)
from .outline import (
    fix_outline_paragraphs,
    force_all_paragraphs_to_body_outline_level,
    remove_all_outline_levels_from_any_root,
)
from .path_utils import is_same_file_path
from .process_runner import run_powershell_script
from .stop_controller import StopController
from .table_format import process_table, table_cell_count, table_column_count
from .xml_utils import qn

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


def should_fix_paragraph_part(name: str) -> bool:
    """
    文件階層縮排只處理本文。

    頁首頁尾常有頁碼；頁碼文字可能只是「1」「2」這種數字，
    若拿同一套文件編號規則判斷，會被誤認為第 3 階編號，
    導致置中頁碼被套用 left/hanging 縮排而偏掉。
    """
    return name == "word/document.xml"


def get_word_table_start_pages(input_docx: Path, stop: StopController | None = None) -> list[int | None]:
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        return get_word_table_start_pages_with_powershell(input_docx, stop=stop)

    wd_collapse_start = 1
    wd_active_end_page_number = 3

    word = None
    doc = None
    pages: list[int | None] = []
    com_failed = False

    try:
        word = win32com.client.DispatchEx("Word.Application")
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
        numbering_level_lookup = build_numbering_level_lookup(numbering_xml)
        style_numbering_lookup = build_style_numbering_lookup(styles_xml)

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
                data = apply_numbering_outline_format(data) or data
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

            if should_process_part(item.filename):
                if root is None:
                    root = etree.fromstring(data, parser)
                if item.filename == "word/document.xml" and not document_table_pages:
                    document_table_pages = get_rendered_table_start_pages(root)

                if (
                    (
                        options.fix_paragraph
                        or (
                            options.remove_preface_outline
                            and not options.remove_all_outline_levels
                        )
                    )
                    and should_fix_paragraph_part(item.filename)
                ):
                    if progress_callback:
                        message = "移除壹、序言前的大綱階層"
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
                        style_numbering_lookup=style_numbering_lookup,
                        change_logs=summary.paragraph_logs,
                        part_name=item.filename,
                        summary=summary,
                        remove_preface_outline=options.remove_preface_outline,
                        fix_numbered_paragraphs=options.fix_paragraph,
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

            zout.writestr(item, data)

    if progress_callback:
        progress_callback(percent=100, message="完成")

    return summary
