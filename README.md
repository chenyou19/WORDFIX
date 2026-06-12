# Word DOCX Fixer / WORDFIX

WORDFIX 是一個用來整理 Word `.docx` 文件格式的工具，提供 GUI 與 CLI 兩種使用方式。它主要協助修正表格版面、表格底色、段落大綱階層與縮排，並在需要時透過 Word COM 補救 Word 實際開啟後才容易判斷的版面差異。

## 快速啟動

啟動 GUI：

```powershell
python main.py
```

使用 CLI：

```powershell
python main.py input.docx output.docx --table --color --paragraph
```

若 CLI 沒有指定任何主要處理項目，會啟用預設處理動作。CLI 預設會輸出 log；GUI 內建預設不輸出 log。

## 主要功能

- 修正表格版面，依欄數分成一般表格與特殊表格處理。
- 修正表格底色，保留合規底色，必要時改成預設灰色或清除。
- 修正段落大綱階層與縮排，支援手動編號、自動編號與樣式編號。
- 可獨立處理序言區段縮排與大綱階層。
- 預設保護巢狀表格，也就是表格內含表格或表格本身位於其他表格內。
- 針對「參、價格形成之主要因素分析」提供三個獨立保護選項：表格版面不調整、表格顏色不調整、縮排不調整。
- 可輸出 `process_log`、`table_log`、`heading_suffix_log` 追查處理結果。
- GUI 可保存目前勾選狀態與縮排設定，下次啟動時沿用。

## 詳細文件

- [GUI 選項](docs/gui-options.md)
- [CLI 參數](docs/cli-options.md)
- [表格處理規則](docs/table-rules.md)
- [巢狀表格處理](docs/nested-tables.md)
- [參章保護邏輯](docs/chapter-three-protection.md)
- [段落、大綱與縮排規則](docs/paragraph-indent-rules.md)
- [預設設定檔](docs/default-settings.md)
- [Log 輸出說明](docs/logs.md)
- [EXE 與設定攜帶](docs/exe-build-and-settings.md)
- [開發者維護筆記](docs/developer-notes.md)

文件索引頁見 [docs/README.md](docs/README.md)。

## GUI 預設重點

GUI 第一頁勾選預設以 `docx_fixer/gui_defaults.py` 的 `built_in_gui_defaults()` 為準。目前內建預設包含：

- `skip_chapter_three_table_layout = True`
- `skip_chapter_three_table_color = True`
- `skip_chapter_three_indents = False`

也就是說，GUI 內建預設會保護參章表格版面與表格顏色，但不會預設勾選參章縮排不調整。

## EXE 與設定檔

預設設定檔是 `indent_defaults.json`。

- 原始碼執行時，設定檔位於目前工作目錄。
- EXE 執行時，設定檔位於 EXE 同層。
- 設定檔內的 `indent_settings` 儲存縮排設定，`gui_defaults` 儲存 GUI 第一頁勾選狀態。

若要把 EXE 帶到另一台電腦並沿用設定，請把 EXE 與 `indent_defaults.json` 放在同一層。

## 注意事項

- 輸入與輸出路徑不能是同一個 `.docx` 檔。
- 批次處理前建議先備份原始文件。
- 特殊格式文件建議先用少量樣本測試。
- Word COM 相關功能需要 Windows 上可用的 Microsoft Word 環境。
