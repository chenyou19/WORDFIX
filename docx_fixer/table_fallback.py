from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from .constants import NS
from .exceptions import ProcessStopped
from .stop_controller import StopController
from .table_format import apply_table_format


def fallback_normal_table_autofit_in_docx(
    output_docx: str | Path,
    failed_records: list[dict[str, object]],
    logs: list[str],
    stop: StopController | None = None,
) -> set[int]:
    """Re-apply the safe window-width XML format to normal tables that
    Word COM AutoFit failed to process, writing the result back into
    output_docx.

    Only records from word/document.xml are handled; tables are located by
    table_index within that part so header/footer indices can never be
    confused with Word COM's doc.Tables numbering. Returns the
    global_table_index set that was successfully repaired and persisted.
    """
    logs.append(
        f"WORD_COM_TABLE_AUTOFIT_FALLBACK_STARTED failed_records_count={len(failed_records)}"
    )

    document_records = [
        record
        for record in failed_records
        if record.get("part_name") == "word/document.xml"
    ]
    skipped_other_parts = len(failed_records) - len(document_records)
    if skipped_other_parts:
        logs.append(
            "WORD_COM_TABLE_AUTOFIT_FALLBACK_SKIPPED_NON_DOCUMENT_PART "
            f"count={skipped_other_parts}"
        )

    if not document_records:
        logs.append("WORD_COM_TABLE_AUTOFIT_FALLBACK_DONE applied=0")
        return set()

    applied: set[int] = set()
    try:
        output_path = Path(output_docx)
        with ZipFile(output_path) as zin:
            items = zin.infolist()
            data_by_name = {item.filename: zin.read(item.filename) for item in items}

        root = etree.fromstring(data_by_name["word/document.xml"])
        tables = root.xpath(".//w:tbl", namespaces=NS)

        for record in document_records:
            if stop:
                stop.check()
            try:
                table_index = int(record.get("table_index", 0))
                global_table_index = int(record.get("global_table_index", 0))
            except (TypeError, ValueError):
                continue
            if table_index < 1 or table_index > len(tables):
                logs.append(
                    "WORD_COM_TABLE_AUTOFIT_FALLBACK_NOT_FOUND "
                    f"global_table_index={global_table_index} "
                    f"table_index={table_index} document_table_count={len(tables)}"
                )
                continue
            apply_table_format(tables[table_index - 1], stop=stop)
            applied.add(global_table_index)
            logs.append(
                "WORD_COM_TABLE_AUTOFIT_FALLBACK_APPLIED "
                f"global_table_index={global_table_index} table_index={table_index}"
            )

        data_by_name["word/document.xml"] = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        buffer = BytesIO()
        with ZipFile(buffer, "w", ZIP_DEFLATED) as zout:
            for item in items:
                zout.writestr(item, data_by_name[item.filename])
        output_path.write_bytes(buffer.getvalue())

        logs.append(f"WORD_COM_TABLE_AUTOFIT_FALLBACK_DONE applied={len(applied)}")
        return applied
    except ProcessStopped:
        raise
    except Exception as exc:
        logs.append(
            "WORD_COM_TABLE_AUTOFIT_FALLBACK_ERROR "
            f"type={type(exc).__name__} message={exc}"
        )
        return set()
