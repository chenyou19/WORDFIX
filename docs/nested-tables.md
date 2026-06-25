# 巢狀表格處理

「巢狀表格不調整版面（仍調整顏色）」對應 `ProcessOptions.skip_nested_tables`。GUI 內建預設與 CLI 預設都是 `True`，底層欄位名稱維持不變以相容既有設定檔。

## 判斷方式

`docx_fixer/table_pipeline.py` 會用兩種條件判斷巢狀表格：

- `is_nested_table(tbl)`：表格本身有 `ancestor::w:tbl`，代表位於另一張表格內。
- `contains_nested_table(tbl)`：表格內部有 `.//w:tc//w:tbl`，代表表格中含有另一張表格。

## 預設保護

當 `skip_nested_tables=True` 時，巢狀表格不套用表格版面調整：不 AutoFit、不改 `tblW` / `tblGrid` / `tcW` / `tblLayout` / `tblInd`，也不改表格內文字級、段落對齊、列高、框線、表格註記搬移、最後一列說明格式化或 Word COM table AutoFit。

若全域 `fix_color=False`，巢狀表格完整跳過，`table_type=skipped_nested_table`，且 `skipped_nested_tables` 會增加。

若全域 `fix_color=True`，巢狀表格進入 color-only 分支，只套用既有 `apply_table_color()` 規則。外表格只處理自己的直屬 `./w:tr/w:tc`，內表格會在自己的 table 迴圈中處理自己的直屬 cell，因此不會重複統計或重複 log。

## 保護優先順序

以下保護仍會阻止巢狀表格顏色修改：

1. `word/document.xml` 第一張表格完整跳過。
2. 「參、表格顏色不調整」保護。
3. 指定特殊顏色跳過整張表規則。

## GUI 與 CLI

GUI 顯示文字：

- 巢狀表格不調整版面（仍調整顏色）

CLI 參數：

- `--skip-nested-tables`：保護巢狀表格版面；若同時啟用 `--fix-color`，仍會套用表格顏色規則。
- `--no-skip-nested-tables`：允許完整處理巢狀表格與含巢狀表格的表格。

## Log

巢狀表格只調顏色時，`table_log` 會記錄：

- `table_type = nested_table_color_only`
- `action = apply_nested_table_color_only`
- `layout_fixed = false`
- `color_fixed = true`
- `changed_to_gray` / `cleared_colors` / `shading_debug`

完整跳過時仍使用：

- `table_type = skipped_nested_table`
- `action = skipped`
- `reason = nested table protected; table contains or is inside another table`

`process_log` 摘要會分開記錄 `skipped_nested_tables` 與 `nested_table_color_only_tables`。
