# 巢狀表格處理

「表格中有表格不調整」對應 `ProcessOptions.skip_nested_tables`。GUI 內建預設與 CLI 預設都是 `True`。

## 什麼情況算巢狀表格

`docx_fixer/table_pipeline.py` 會把兩種情況視為需要保護：

- 表格本身在另一個表格裡：`ancestor::w:tbl`。
- 表格內含另一張表格：`.//w:tc//w:tbl`。

只要符合其中一種，且 `skip_nested_tables=True`，這張表格就不做版面或底色處理。

## GUI 與 CLI

GUI 對應選項：

- 表格中有表格不調整

CLI 對應參數：

- `--skip-nested-tables`：跳過巢狀表格，這是預設行為。
- `--no-skip-nested-tables`：允許處理巢狀表格。

## Log

被跳過時，`table_log` 會記錄：

- `table_type = skipped_nested_table`
- `action = skipped`
- `reason = nested table protected; table contains or is inside another table`

同時 `process_log` 摘要裡的 `skipped_nested_tables` 計數會增加。
