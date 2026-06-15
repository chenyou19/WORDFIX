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

## process_log

`process_log` 記錄整體摘要與段落處理細節，包含：

- 表格摘要，例如總表格數、跳過第一張表格數、跳過小表格數、跳過巢狀表格數。
- 段落摘要，例如總段落數、跳過 TOC 段落數、跳過表格段落數、移除大綱階層數。
- 縮排設定快照。
- 編號縮排量測紀錄。
- numbering XML debug。
- Body indent debug。
- Word COM table AutoFit log。
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
- `double_border_applied`：是否套用黑色雙線外框。
- `skip_section_three_adjustments_enabled`、`in_section_three_protected`、`section_three_detection_source`、`skipped_by_section_three_protection`：「參、不要調整」整章保護狀態。
- `move_table_notes_below_enabled`、`note_cells_moved`、`moved_note_count`、`deleted_note_cells`、`deleted_note_rows`、`inserted_note_paragraphs`、`moved_notes`、`note_move_warnings`：表格內註記搬移結果。
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

`process_log` 的表格摘要另含：因「參、不要調整」保護跳過的表格數、套用黑色雙線外框的表格數、搬移註記的表格數與筆數、刪除註記儲存格/整列數、表格下方新增註記段落數。

## heading_suffix_log

`heading_suffix_log` 比對修正前後標題編號 suffix 與 tab 狀態，包含：

- 修正前後標題總數。
- 手動編號與自動編號數量。
- `suffix=nothing`、`tab`、`space`、`missing`、`other` 統計。
- tab stop 是否仍存在。
- 每一個標題修正前後的 suffix、numId、ilvl、樣式與文字。

這個 log 主要用來追查 Word 編號後方 tab 或空白是否仍殘留。
