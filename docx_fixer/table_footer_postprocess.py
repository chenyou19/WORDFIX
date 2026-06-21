from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from .constants import NS
from .exceptions import ProcessStopped
from .stop_controller import StopController
from .table_format import apply_table_footer_source_format


def apply_table_footer_source_format_in_docx(
    output_docx: str | Path,
    records: list[dict[str, object]],
    logs: list[str] | None = None,
    stop: StopController | None = None,
) -> dict[int, dict[str, object]]:
    """Re-apply 「表格最後一列說明格式化」 to the recorded tables in a saved docx.

    This runs AFTER Word COM AutoFit and the XML fallback so neither pass can
    clobber the footer formatting (the fallback re-applies ``apply_table_format``
    which resets every run to 11 pt and every paragraph to centered, and Word
    COM re-saves the document). It is the *final* table format step.

    Each table is located per part by its ``table_index`` within
    ``.//w:tbl`` — the same stable scheme the AutoFit fallback uses — so only
    the recorded tables are touched (no blind full-document rescan). Returns a
    mapping ``{global_table_index: result}`` for the tables that were formatted.
    """
    if logs is None:
        logs = []
    results: dict[int, dict[str, object]] = {}

    if not records:
        logs.append("FOOTER_SOURCE_FORMAT_REAPPLY_SKIPPED reason=no_records")
        return results

    logs.append(f"FOOTER_SOURCE_FORMAT_REAPPLY_STARTED records_count={len(records)}")

    records_by_part: dict[str, list[dict[str, object]]] = {}
    for record in records:
        part_name = str(record.get("part_name", "word/document.xml"))
        records_by_part.setdefault(part_name, []).append(record)

    try:
        output_path = Path(output_docx)
        with ZipFile(output_path) as zin:
            items = zin.infolist()
            data_by_name = {item.filename: zin.read(item.filename) for item in items}

        changed = False
        for part_name, part_records in records_by_part.items():
            if part_name not in data_by_name:
                for record in part_records:
                    logs.append(
                        "FOOTER_SOURCE_FORMAT_REAPPLY_PART_MISSING "
                        f"part_name={part_name} "
                        f"global_table_index={record.get('global_table_index')}"
                    )
                continue

            root = etree.fromstring(data_by_name[part_name])
            tables = root.xpath(".//w:tbl", namespaces=NS)
            part_changed = False

            for record in part_records:
                if stop:
                    stop.check()
                try:
                    table_index = int(record.get("table_index", 0))
                    global_table_index = int(record.get("global_table_index", 0))
                except (TypeError, ValueError):
                    continue

                if table_index < 1 or table_index > len(tables):
                    logs.append(
                        "FOOTER_SOURCE_FORMAT_REAPPLY_NOT_FOUND "
                        f"part_name={part_name} global_table_index={global_table_index} "
                        f"table_index={table_index} table_count={len(tables)}"
                    )
                    continue

                result = apply_table_footer_source_format(tables[table_index - 1], stop=stop)
                results[global_table_index] = result
                part_changed = True
                logs.append(
                    "FOOTER_SOURCE_FORMAT_REAPPLY_APPLIED "
                    f"part_name={part_name} global_table_index={global_table_index} "
                    f"table_index={table_index} "
                    f"footer_rows_processed={result['footer_rows_processed']} "
                    f"footer_row_matches={'|'.join(result['footer_row_matches']) or 'none'} "
                    f"footer_note_cells_adjusted={result['footer_note_cells_adjusted']}"
                )

            if part_changed:
                data_by_name[part_name] = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )
                changed = True

        if changed:
            buffer = BytesIO()
            with ZipFile(buffer, "w", ZIP_DEFLATED) as zout:
                for item in items:
                    zout.writestr(item, data_by_name[item.filename])
            output_path.write_bytes(buffer.getvalue())

        logs.append(f"FOOTER_SOURCE_FORMAT_REAPPLY_DONE applied={len(results)}")
        return results
    except ProcessStopped:
        raise
    except Exception as exc:
        logs.append(
            "FOOTER_SOURCE_FORMAT_REAPPLY_ERROR "
            f"type={type(exc).__name__} message={exc}"
        )
        return results
