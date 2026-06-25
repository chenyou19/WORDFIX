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
                    f"footer_row_count={result['footer_row_count']} "
                    f"footer_top_row_index={result['footer_top_row_index']} "
                    f"footer_cell_matches={'|'.join(result['footer_cell_matches']) or 'none'} "
                    f"footer_block_top_border_applied={result['footer_block_top_border_applied']} "
                    f"footer_internal_top_borders_cleared={result['footer_internal_top_borders_cleared']} "
                    f"table_bottom_border_mode={result['table_bottom_border_mode']} "
                    f"table_bottom_border_cell_count={result['table_bottom_border_cell_count']} "
                    f"table_bottom_border_xml_verified={result['table_bottom_border_xml_verified']} "
                    f"footer_terminal_bottom_none_applied={result['footer_terminal_bottom_none_applied']} "
                    f"footer_terminal_bottom_none_cell_count={result['footer_terminal_bottom_none_cell_count']} "
                    f"table_bottom_double_border_applied={result['table_bottom_double_border_applied']} "
                    f"table_bottom_double_border_cell_count={result['table_bottom_double_border_cell_count']} "
                    f"table_bottom_double_border_xml_verified={result['table_bottom_double_border_xml_verified']} "
                    f"table_top_border_mode={result['table_top_border_mode']} "
                    f"table_top_border_cell_count={result['table_top_border_cell_count']} "
                    f"table_top_border_xml_verified={result['table_top_border_xml_verified']} "
                    f"first_row_single_cell_title={result['first_row_single_cell_title']} "
                    f"first_row_single_cell_border_mode={result['first_row_single_cell_border_mode']} "
                    f"first_row_single_cell_border_xml_verified={result['first_row_single_cell_border_xml_verified']} "
                    f"data_rows_outer_left_double_applied={result['data_rows_outer_left_double_applied']} "
                    f"data_rows_outer_right_double_applied={result['data_rows_outer_right_double_applied']} "
                    f"data_rows_outer_left_target_count={result['data_rows_outer_left_target_count']} "
                    f"data_rows_outer_right_target_count={result['data_rows_outer_right_target_count']} "
                    f"data_rows_outer_left_vmerge_owner_target_count={result['data_rows_outer_left_vmerge_owner_target_count']} "
                    f"data_rows_outer_right_vmerge_owner_target_count={result['data_rows_outer_right_vmerge_owner_target_count']} "
                    f"footer_rows_outer_left_none_applied={result['footer_rows_outer_left_none_applied']} "
                    f"footer_rows_outer_right_none_applied={result['footer_rows_outer_right_none_applied']} "
                    f"footer_rows_outer_left_target_count={result['footer_rows_outer_left_target_count']} "
                    f"footer_rows_outer_right_target_count={result['footer_rows_outer_right_target_count']} "
                    f"outer_vertical_border_policy_xml_verified={result['outer_vertical_border_policy_xml_verified']} "
                    f"last_row_physical_cell_count={result['last_row_physical_cell_count']} "
                    f"last_row_grid_span_sum={result['last_row_grid_span_sum']} "
                    f"last_row_vmerge_states={result['last_row_vmerge_states']} "
                    f"last_row_bottom_edge_target_count={result['last_row_bottom_edge_target_count']} "
                    f"table_border_schema_order_valid={result['table_border_schema_order_valid']} "
                    f"tblPr_child_order={result['tblPr_child_order']} "
                    f"last_row_tcPr_child_orders={'|'.join(result['last_row_tcPr_child_orders']) or 'none'} "
                    "table_bottom_border_verify_detail="
                    f"{result['table_bottom_border_verify_detail']} "
                    "table_top_border_verify_detail="
                    f"{result['table_top_border_verify_detail']} "
                    "first_row_single_cell_border_verify_detail="
                    f"{result['first_row_single_cell_border_verify_detail']} "
                    "outer_vertical_border_policy_verify_detail="
                    f"{result['outer_vertical_border_policy_verify_detail']} "
                    "table_bottom_double_border_verify_detail="
                    f"{result['table_bottom_double_border_verify_detail']}"
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
