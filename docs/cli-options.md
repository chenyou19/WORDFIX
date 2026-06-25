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
| `--skip-chapter-three-numbering-suffix-cleanup` | 新版參數 | 開啟 | 「參、不要清理編號後綴 tab/space」：不清理參章 numbering definition 的 `w:suff`／`w:pPr/w:tabs`／`w:lvlText` 結尾空白。只控制 numbering 後綴清理，不影響版面、顏色、縮排。 |
| `--no-skip-chapter-three-numbering-suffix-cleanup` | 新版參數 |  | 允許清理參章 numbering definition 的後綴 tab/space。 |

相容舊名稱仍可使用，但不建議新文件或新腳本主推：

| 參數 | 狀態 | 說明 |
| --- | --- | --- |
| `--skip-chapter-three-tables` | 相容舊名稱 | 同時跳過參章表格版面與表格底色。 |
| `--no-skip-chapter-three-tables` | 相容舊名稱 | 同時允許參章表格版面與表格底色。 |
| `--skip-all-under-chapter-three` | 相容舊名稱 | 同時跳過參章表格版面、表格底色與縮排。 |
| `--no-skip-all-under-chapter-three` | 相容舊名稱 | 同時允許參章表格版面、表格底色與縮排。 |
| `--skip-special-layout-under-chapter-three` | 相容舊名稱 | 舊的合併跳過選項，會把表格版面、表格底色與縮排都設為跳過。 |

CLI 與 GUI 的參章縮排預設不同：CLI 的 `--skip-chapter-three-indents` 預設是開啟；GUI 內建的 `skip_chapter_three_indents` 預設是不勾選。

## 表格內註記搬移參數（GUI 已隱藏並強制關閉）

> 表格註記搬移已從 GUI 隱藏並強制關閉（GUI／設定檔載入一律 `move_table_notes_below=False`、`skip_chapter_three_table_notes=False`）。下列 CLI 參數仍保留未刪除，預設皆為「不搬移」；GUI 不受 CLI 影響。

| 參數 | 狀態 | 預設 | 說明 |
| --- | --- | --- | --- |
| `--move-table-notes-below` | 相容保留 | 關閉 | 把每張表格中以 `註：`、`註1：`、`註一、` 等開頭的儲存格搬到表格正下方，成為標楷體 10pt 段落。預設關閉。 |
| `--no-move-table-notes-below` | 相容保留 | | 保留表格內的註記儲存格。 |
| `--skip-chapter-three-table-notes` | 相容保留 | 開啟 | 「參、不要表格註記搬移」：啟用註記搬移時，正文中「參、」章節（以章節編號第 3 章判斷，目錄不會觸發）內的表格不搬移註記。只控制註記搬移。 |
| `--no-skip-chapter-three-table-notes` | 相容保留 | | 連「參、」章節內的表格也搬移註記。 |

## 表格最後一列說明格式化參數

| 參數 | 狀態 | 預設 | 說明 |
| --- | --- | --- | --- |
| `--enable-table-footer-source-format` | 目前參數 | 關閉 | 「表格最後一列說明格式化」：對符合 footer eligibility 的表格，依序套用全表 11pt、外圍黑色雙線、第一列單 cell title row 黑色雙線外框、底部連續說明列格式（「基期：」「資料來源：」以及符合 `^註(?:\d+)?[：:]` 的註記），最後依底部是否為 footer 分流：一般資料表底部為黑色雙線，footer 說明區底部為無線且左右外側不補雙線。獨立功能，不依賴註記搬移或黑色雙線外框，也不混入顏色處理，且不搬移／刪除／新增任何 cell。 |
| `--table-footer-source-format` | 別名 | | 與 `--enable-table-footer-source-format` 相同。 |
| `--no-enable-table-footer-source-format` | 目前參數 | | 不套用此格式（別名 `--no-table-footer-source-format`）。 |

CLI 與 GUI 都把同一個布林值 `enable_table_footer_source_format` 傳入核心流程；CLI / 核心 `ProcessOptions` 預設關閉，GUI 內建預設勾選並以設定檔 `gui_defaults` 為準。詳細處理順序與跳過規則見 [表格處理規則](table-rules.md)。

## 隱藏參數

| 參數 | 狀態 | 預設 | 說明 |
| --- | --- | --- | --- |
| `--force-note-paragraph-left-alignment` | 隱藏參數 | 關閉 | 開啟後才會把「註…」開頭段落強制靠左（`force_note_paragraph_left_alignment_in_docx`）。預設關閉，help 不顯示，GUI 也不提供。 |
| `--no-force-note-paragraph-left-alignment` | 隱藏參數 | | 維持預設關閉。 |
| `--enable-double-black-table-borders` | 隱藏參數 | 關閉 | 開啟後才會對一般表格與特殊表格套用黑色雙線外框（`enable_double_black_table_borders`）。預設關閉，不改 `w:tblBorders`；help 不顯示，GUI 也不提供。 |
| `--no-enable-double-black-table-borders` | 隱藏參數 | | 維持預設關閉。 |

舊版的「參、不要調整」整章保護參數 `--skip-chapter-three-adjustments`（別名 `--protect-section-three`）仍保留為相容舊腳本用途，但不再是主推功能，且不再控制表格註記搬移；新腳本請改用上方的 `--skip-chapter-three-table-notes`。

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
