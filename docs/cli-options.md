# CLI 參數

CLI 由 `docx_fixer/cli.py` 負責。當 `main.py` 同時收到輸入與輸出檔路徑時會進入 CLI 模式：

```powershell
python main.py input.docx output.docx --table --color --paragraph
```

## 預設行為

- CLI 預設會輸出 log。除非使用 `--no-log` 或 `--skip-log-output`，否則處理完成後會寫出 `process_log`、`table_log`、`heading_suffix_log`。
- 如果沒有指定任何主要處理項目，CLI 會啟用預設處理動作：修正表格版面、修正表格底色、修正段落大綱階層與縮排。
- 啟用預設處理動作時，`remove_all_outline`、`indent_preface`、`outline_preface` 不會被自動打開。
- `--quiet` 只關閉進度輸出，不會關閉完成後的摘要輸出。

主要處理項目指 `--table`、`--color`、`--paragraph`、`--remove-all-outline`、`--indent-preface`、`--outline-preface`。

## 主要處理參數

| 參數 | 狀態 | 說明 |
| --- | --- | --- |
| `--table` | 目前參數 | 修正表格版面。 |
| `--color` | 目前參數 | 修正表格底色。 |
| `--paragraph` | 目前參數 | 修正本文段落的大綱階層與縮排。 |
| `--remove-all-outline` | 目前參數 | 移除處理範圍內的段落大綱階層。 |
| `--indent-preface` | 目前參數 | 對序言區段套用縮排。 |
| `--outline-preface` | 目前參數 | 對序言區段套用 Word 大綱階層。 |

## 段落與 Word COM 參數

| 參數 | 狀態 | 說明 |
| --- | --- | --- |
| `--level1-level2-body-first-line-indent` | 新版參數 | 第 1、2 層標題下方普通內文套用 `560` twips 首行縮排。 |
| `--level2-body-first-line-indent` | 相容舊名稱 | 與 `--level1-level2-body-first-line-indent` 相同。 |
| `--word-com-check-body-font` | 目前參數 | XML 判斷內文字號不是 14pt 時，改由 Word COM 再確認是否應套用內文縮排。 |

## 參章保護參數

新版文件主推拆分後的三個參數：

| 參數 | 狀態 | 預設 | 說明 |
| --- | --- | --- | --- |
| `--skip-chapter-three-table-layout` | 新版參數 | 開啟 | 跳過「參、價格形成之主要因素分析」章內表格版面處理。 |
| `--no-skip-chapter-three-table-layout` | 新版參數 |  | 允許該章內表格版面處理。 |
| `--skip-chapter-three-table-color` | 新版參數 | 開啟 | 跳過該章內表格底色處理。 |
| `--no-skip-chapter-three-table-color` | 新版參數 |  | 允許該章內表格底色處理。 |
| `--skip-chapter-three-indents` | 新版參數 | 開啟 | 跳過該章內段落縮排與 Word COM 內文縮排補救。 |
| `--no-skip-chapter-three-indents` | 新版參數 |  | 允許該章內段落縮排處理。 |

相容舊名稱仍可使用，但不建議新文件或新腳本主推：

| 參數 | 狀態 | 說明 |
| --- | --- | --- |
| `--skip-chapter-three-tables` | 相容舊名稱 | 同時跳過參章表格版面與表格底色。 |
| `--no-skip-chapter-three-tables` | 相容舊名稱 | 同時允許參章表格版面與表格底色。 |
| `--skip-all-under-chapter-three` | 相容舊名稱 | 同時跳過參章表格版面、表格底色與縮排。 |
| `--no-skip-all-under-chapter-three` | 相容舊名稱 | 同時允許參章表格版面、表格底色與縮排。 |
| `--skip-special-layout-under-chapter-three` | 相容舊名稱 | 舊的合併跳過選項，會把表格版面、表格底色與縮排都設為跳過。 |

CLI 與 GUI 的參章縮排預設不同：CLI 的 `--skip-chapter-three-indents` 預設是開啟；GUI 內建的 `skip_chapter_three_indents` 預設是不勾選。

## 巢狀表格參數

| 參數 | 狀態 | 預設 | 說明 |
| --- | --- | --- | --- |
| `--skip-nested-tables` | 目前參數 | 開啟 | 跳過表格內含表格，或本身位於其他表格內的表格。 |
| `--no-skip-nested-tables` | 目前參數 |  | 允許處理上述巢狀表格。 |

## 表格顏色參數

| 參數 | 狀態 | 預設 | 說明 |
| --- | --- | --- | --- |
| `--table-keep-colors DDEBF7,FFFFFF` | 目前參數 | 讀取保存的 `table_color_settings` | 逗號分隔 HEX 清單，命中時底色保留。 |
| `--table-gray-colors BFBFBF,C0C0C0` | 目前參數 | 讀取保存的 `table_color_settings` | 逗號分隔 HEX 清單，命中時改成目標灰色。 |
| `--table-gray-target D9D9D9` | 目前參數 | 讀取保存的 `table_color_settings` | 轉灰色時使用的目標灰色。 |
| `--special-color-skip-colors FFFF00,FF0000` | 目前參數 | 讀取保存的 `table_color_settings` | 表格任一格底色命中清單就整張表跳過。 |
| `--skip-special-color-tables` | 目前參數 | 關閉 | 啟用特殊顏色表格整張跳過。 |
| `--no-skip-special-color-tables` | 目前參數 |  | 停用特殊顏色表格跳過。 |
| `--clear-special-colors-after-skip` | 目前參數 | 關閉 | 跳過後只把命中指定清單的儲存格底色清成無色。 |
| `--no-clear-special-colors-after-skip` | 目前參數 |  | 跳過後保留命中的指定顏色。 |

CLI 沒指定顏色清單參數時，會讀取 `indent_defaults.json` 內保存的 `table_color_settings`；檔案不存在時使用內建預設。

## Log 與輸出參數

| 參數 | 狀態 | 說明 |
| --- | --- | --- |
| `--no-log` | 目前參數 | 不寫出 process、table、heading suffix log。 |
| `--skip-log-output` | 相容同義名稱 | 與 `--no-log` 相同。 |
| `--quiet` | 目前參數 | 不顯示進度 callback 輸出。 |

## 隱藏相容參數

`docx_fixer/cli.py` 仍保留 `--normalize-body-style-to-none`、`--no-normalize-body-style-to-none`、`--no-normalize-body-style-to-default-text`，但 help 文字被隱藏，不屬於一般使用者文件主推項目。
