# WORDFIX 文件索引

這個目錄收納 WORDFIX 的詳細使用與維護文件。專案入口請先看根目錄的 [README.md](../README.md)，需要查規則或參數時再進入下列專題頁。

## 文件列表

- [GUI 選項](gui-options.md)：說明 GUI 第一頁勾選項目、內建預設值與保存預設方案。
- [CLI 參數](cli-options.md)：說明命令列參數、預設行為、新版參數與相容舊名稱。
- [表格處理規則](table-rules.md)：說明表格跳過條件、一般表格、特殊表格、顏色與 Word COM AutoFit。
- [巢狀表格處理](nested-tables.md)：說明「表格中有表格不調整」的判斷、預設值與 log。
- [參章保護邏輯](chapter-three-protection.md)：說明「參、價格形成之主要因素分析」的偵測與分項保護。
- [段落、大綱與縮排規則](paragraph-indent-rules.md)：說明本文、序言、TOC、編號與內文縮排處理。
- [預設設定檔](default-settings.md)：說明 `indent_defaults.json` 的位置、內容與載入規則。
- [Log 輸出說明](logs.md)：說明 `process_log`、`table_log`、`heading_suffix_log` 與欄位。
- [EXE 與設定攜帶](exe-build-and-settings.md)：說明 GUI 執行、EXE 封裝與跨電腦攜帶設定。
- [開發者維護筆記](developer-notes.md)：列出主要程式檔案責任，方便後續維護。

## 建議閱讀順序

1. 使用 GUI：先看 [GUI 選項](gui-options.md)，再看 [預設設定檔](default-settings.md)。
2. 使用 CLI：先看 [CLI 參數](cli-options.md)，再依需要看表格或段落規則。
3. 維護規則：先看 [開發者維護筆記](developer-notes.md)，再對照相關專題頁與程式碼。
