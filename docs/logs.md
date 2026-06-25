# Log 輸出說明

WORDFIX 目前有三種主要 log，由 `docx_fixer/process_log.py` 寫出。

## 輸出預設

- GUI 預設不輸出 log，因為 GUI 內建勾選「不要輸出 log 檔」。
- GUI 取消「不要輸出 log 檔」後，才會輸出 log。
- CLI 預設輸出 log。
- CLI 使用 `--no-log` 或 `--skip-log-output` 時不輸出 log。

## Log 檔案

假設輸出檔是 `output.docx`：

- `process_log`：`output_log.txt`
- `table_log`：`output_table_log.txt`
- `heading_suffix_log`：`output_heading_suffix_log.txt`

這三個是正式 log，由 `skip_log_output` 控制（GUI 預設不輸出、CLI 預設輸出）。

## 註記 debug log（開發用，預設關閉）

`*_note_debug_log.txt`（GUI 在暫存檔上處理時會變成 `*.__tmp___note_debug_log.txt`）是 `docx_processor._write_note_debug_log_safely()` 產生的**開發診斷 log**，由 `ProcessOptions.write_note_debug_log` 控制，**預設 `False`**：

- 一般 GUI／CLI 處理 docx **不會**再產生此檔；段落紀錄會出現 `NOTE_DEBUG_LOG_SKIPPED reason=disabled`。
- GUI 沒有此勾選項；`built_in_gui_defaults()` 為 `False`，且列入 `FORCED_FALSE_GUI_DEFAULTS`，舊設定檔即使存成 `True` 也會在載入／重存時被強制覆蓋為 `False`。
- CLI 提供隱藏開發旗標 `--write-note-debug-log`（help 以 `argparse.SUPPRESS` 隱藏），預設 `False`；只有明確指定時才輸出，輸出時段落紀錄會出現 `NOTE_DEBUG_LOG_WRITTEN`。
- 此開關與正式 log（process／table／heading suffix）完全獨立，關閉它不影響任何表格、註記、基期/資料來源或段落縮排處理。

## process_log

`process_log` 記錄整體摘要與段落處理細節，包含：

- 表格摘要，例如總表格數、跳過第一張表格數、跳過小表格數、跳過巢狀表格數。
- 段落摘要，例如總段落數、跳過 TOC 段落數、跳過表格段落數、移除大綱階層數。
- 縮排設定快照。
- 編號縮排量測紀錄。
- numbering XML debug。
- Body indent debug。
- Word COM table AutoFit log。
- 表格最後一列說明格式化 final post-process log（`FOOTER_SOURCE_FORMAT_REAPPLY_*`：在 Word COM／fallback 之後重新套用 footer 格式的紀錄）。
- Word COM body indent fix log。
- 段落變更紀錄。

## table_log

`table_log` 逐張表格記錄處理結果。欄位包含：

- `part_name`
- `table_index`
- `global_table_index`
- `table_name`
- `first_level_heading`
- `cell_count`
- `column_count`
- `table_type`
- `action`
- `reason`
- `special_layout_used`
- `layout_fixed`
- `color_fixed`
- `chapter_three_table_layout_skipped`
- `chapter_three_table_color_skipped`
- `word_com_autofit_applied`
- `word_com_autofit_sequence`
- `double_border_enabled`：隱藏選項 `enable_double_black_table_borders` 是否啟用（預設 `false`）。
- `double_border_applied`：該表是否實際套用黑色雙線外框（隱藏選項預設關閉時恆為 `false`）。
- `table_footer_note_source_format_enabled`：選項 `enable_table_footer_source_format` 是否啟用（核心 / CLI 預設 `false`；GUI 內建預設 `true`，以設定檔中的 `gui_defaults` 為準）。
- `table_footer_note_source_format_should_apply`：XML pipeline 判定該表需要 footer 格式化（已記錄、待最後 post-process 套用）。
- `table_footer_note_source_format_applied`：最後 post-process 是否實際套用「表格最後一列說明格式化」（在 Word COM AutoFit／fallback 之後執行）。
- `outer_double_border_applied_by_footer_source_format`：本功能是否套用外圍黑色雙線。
- `table_bottom_border_mode`：最後底邊決策。`data_double` 表示一般資料表底部套黑色雙線；`footer_none` 表示 footer 說明區底部清為無線；`not_applied` 表示未套用。
- `table_bottom_border_cell_count`：本次 mode 實際處理的底邊 cell 數。
- `table_bottom_border_xml_verified`：XML 是否符合該 mode 的預期。這只表示最終 DOCX XML 條件通過，不代表已做 Word 畫面或像素級驗證。
- `table_bottom_border_verify_detail`：XML 驗證摘要，例如 `table_bottom_border_mode=data_double;tbl_bottom=double/4/000000;last_row_tc_bottoms=...;table_border_schema_order_valid=true;tblPr_child_order=...;last_row_tcPr_child_orders=...;last_row_grid_span_sum=...;last_row_vmerge_states=...`。
- `table_bottom_double_border_applied`：舊相容欄位，只有 `data_double` mode 時才會是 `true`；`footer_none` mode 必須為 `false`。
- `table_bottom_double_border_cell_count`：舊相容欄位，`data_double` mode 的底邊 cell 數；`footer_none` 為 `0`。
- `table_bottom_double_border_xml_verified`：舊相容欄位，僅驗證 `data_double`；`footer_none` 為 `false`，不可把 footer 下方無線視為錯誤。
- `table_bottom_double_border_verify_detail`：舊相容欄位，保留 XML 摘要。
- `footer_terminal_bottom_none_applied`：footer mode 是否已清除 footer 最末列下方底線。
- `footer_terminal_bottom_none_cell_count`：footer 最末列實際清除 bottom 的實體 cell 數。
- `last_row_physical_cell_count`、`last_row_grid_span_sum`、`last_row_vmerge_states`、`last_row_bottom_edge_target_count`：最後列實體 cell、邏輯欄寬、垂直合併與底邊目標診斷。
- `table_border_schema_order_valid`、`tblPr_child_order`、`last_row_tcPr_child_orders`：`tblBorders`/`tcBorders` 是否位於合法 OOXML schema 位置，以及實際 child order 摘要。若 `tblBorders` 在 `tblLayout` 後或 `tcBorders` 在 `vAlign` 後，Word 可能保留 XML 但不照預期顯示。
- `first_row_single_cell_border_adjusted`：第一列單 cell 是否被本功能調整。
- `footer_row_count`：從表格底部往上連續處理的 footer 列數（中斷於第一個無命中的列）。
- `footer_cell_matches`：每列命中的類型（由上往下），例如 `note | base_period,source`。
- `footer_note_cells_adjusted`：所有 footer 列命中「基期：」「資料來源：」或註記（`^註(?:\d+)?[：:]`）並格式化的 cell 總數。
- `footer_note_cell_matches`：所有命中 cell 的類型清單（`note`／`base_period`／`source`）。
- `footer_note_cell_debug`：每個命中 cell 文字前 50 字與套用動作。
- `table_footer_note_source_format_skipped_reason`：未套用時的原因（`feature_disabled`、`layout not adjusted for this table` 或各跳過原因）。
- `table_note_move_gui_hidden`：表格註記搬移功能已從 GUI 隱藏（恆 `true`）。
- `table_note_move_forced_false`：`move_table_notes_below` 是否為關（GUI 一律強制 `true`）。
- `skip_chapter_three_table_note_move_forced_false`：`skip_chapter_three_table_notes` 是否為關（GUI 一律強制 `true`）。
- `move_table_notes_below_enabled`、`skip_chapter_three_table_notes_enabled`、`table_notes_skipped_by_chapter_three`、`note_cells_moved`、`moved_note_count`、`deleted_note_cells`、`deleted_note_rows`、`inserted_note_paragraphs`、`moved_notes`、`note_move_warnings`：表格內註記搬移結果（GUI 已隱藏並強制關閉，預設皆為未搬移）。
- `skip_section_three_adjustments_enabled`、`in_section_three_protected`、`section_three_detection_source`、`skipped_by_section_three_protection`：相容用的舊「參、不要調整」整章保護狀態（GUI 已移除）。
- `changed_to_gray`
- `cleared_colors`
- `shading_debug`

`table_type` 常見值包含：

- `skipped_first_table`
- `skipped_nested_table`
- `skipped_chapter_three_table`
- `skipped_small_table`
- `special_color_skipped_table`
- `special_table`
- `normal_table`
- `color_only_table`

`process_log` 的表格摘要另含：套用黑色雙線外框的表格數、套用最後一列說明格式化的表格數、搬移註記的表格數（GUI 強制關閉後為 0）、因「參、不要表格註記搬移」而未搬移的表格數、搬移註記筆數、刪除註記儲存格/整列數、表格下方新增註記段落數。

## 註記段落強制靠左（隱藏功能）

「註…」開頭段落強制靠左（`force_note_paragraph_left_alignment_in_docx`）預設**關閉**。預設情況下 `process_log` 的段落紀錄會看到 `FINAL_NOTE_ALIGNMENT_FIX_SKIPPED reason=disabled`，且不會出現 `FINAL_NOTE_ALIGNMENT_FIX` / `FINAL_NOTE_ALIGNMENT_SUMMARY`。只有以隱藏旗標 `--force-note-paragraph-left-alignment` 或 `ProcessOptions(force_note_paragraph_left_alignment=True)` 啟用時，才會執行並輸出 `FINAL_NOTE_ALIGNMENT_FIX` 紀錄。GUI 不提供此選項。

## heading_suffix_log

`heading_suffix_log` 比對修正前後標題編號 suffix 與 tab 狀態，包含：

- 修正前後標題總數。
- 手動編號與自動編號數量。
- `suffix=nothing`、`tab`、`space`、`missing`、`other` 統計。
- tab stop 是否仍存在。
- 每一個標題修正前後的 suffix、numId、ilvl、樣式與文字。

這個 log 主要用來追查 Word 編號後方 tab 或空白是否仍殘留。
