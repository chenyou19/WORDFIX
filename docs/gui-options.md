# GUI 選項

GUI 由 `main.py` 在沒有輸入、輸出檔參數時啟動，主要介面在 `docx_fixer/gui_app.py`。第一頁勾選狀態會從 `docx_fixer/gui_defaults.py` 的 `built_in_gui_defaults()` 或 `indent_defaults.json` 裡的 `gui_defaults` 讀取。

## 內建勾選預設

以下預設值以 `built_in_gui_defaults()` 為準：

| GUI 選項 | 設定鍵 | 內建預設 |
| --- | --- | --- |
| 修正表格版面 | `fix_table` | 勾選 |
| 修正表格底色 | `fix_color` | 勾選 |
| 修正段落大綱階層與縮排 | `fix_paragraph` | 勾選 |
| 移除所有段落大綱階層 | `remove_all_outline` | 勾選 |
| 前言段落套用縮排 | `indent_preface` | 不勾選 |
| 前言段落套用大綱階層 | `outline_preface` | 不勾選 |
| 階層 1、2 標題下方普通內文首行縮排兩個中文字 | `level1_level2_body_first_line_indent` | 勾選 |
| XML 判斷非 14pt 時使用 Word COM 確認內文字號 | `word_com_check_body_font` | 不勾選 |
| 不要輸出 log 檔 | `skip_log_output` | 勾選 |
| 表格中有表格不調整 | `skip_nested_tables` | 勾選 |
| 參、價格形成之主要因素分析：表格版面不調整 | `skip_chapter_three_table_layout` | 勾選 |
| 參、價格形成之主要因素分析：表格顏色不調整 | `skip_chapter_three_table_color` | 勾選 |
| 參、價格形成之主要因素分析：縮排不調整 | `skip_chapter_three_indents` | 不勾選 |
| 跳過特殊顏色表格（第三頁） | `skip_special_color_tables` | 不勾選 |
| 跳過後將指定顏色改回無色彩（第三頁） | `clear_special_colors_after_skip` | 不勾選 |

注意：GUI 內建預設不是「參章縮排不調整」。參章保護已拆成表格版面、表格顏色、縮排三個選項，其中只有表格版面與表格顏色預設勾選。

## 處理選項

- **修正表格版面**：啟用 `ProcessOptions.fix_table_layout`，表格會依欄數走一般表格或特殊表格版面規則。
- **修正表格底色**：啟用 `ProcessOptions.fix_color`，依 `docx_fixer/shading.py` 判斷底色要保留、改成預設灰色或清除。
- **移除所有段落大綱階層**：啟用 `remove_all_outline_levels`，會將處理範圍內段落大綱層級改回本文層級，但受參章縮排保護的段落會再恢復原本標題層級。
- **修正段落大綱階層與縮排**：啟用 `fix_paragraph`，主要在 `word/document.xml` 依手動編號、自動編號與樣式編號修正標題大綱與縮排。
- **前言段落套用縮排**：啟用 `indent_preface_paragraphs`，主本文開始前的序言段落可套用獨立的前言縮排表。
- **前言段落套用大綱階層**：啟用 `outline_preface_paragraphs`，主本文開始前的序言段落可寫入 Word outline level。
- **XML 判斷非 14pt 時使用 Word COM 確認內文字號**：啟用 `word_com_check_body_font_when_xml_not_14`。XML 判斷內文字號不是 14pt 時，先記錄待確認項目，後續交給 Word COM 以實際開啟後的字號判斷是否補套內文縮排。
- **不要輸出 log 檔**：啟用 `skip_log_output`。GUI 內建預設不輸出 log，取消勾選才會寫出 `process_log`、`table_log`、`heading_suffix_log`。
- **階層 1、2 標題下方普通內文首行縮排兩個中文字**：啟用 `enable_level1_level2_body_first_line_indent`。第 1、2 層標題下方普通內文會套用 `560` twips 的首行縮排。

## 保護選項

- **表格中有表格不調整**：預設勾選，對應 `skip_nested_tables=True`。表格本身在另一張表格內，或表格內含另一張表格，都會跳過。
- **參、價格形成之主要因素分析：表格版面不調整**：只停用該章內表格版面處理，不影響底色。
- **參、價格形成之主要因素分析：表格顏色不調整**：只停用該章內表格底色處理，不影響版面。
- **參、價格形成之主要因素分析：縮排不調整**：只停用該章內段落縮排、firstLine、hanging、tabs、字元縮排清理與 Word COM 內文縮排補救；真正標題的 outline level 仍會恢復。

## 表格顏色設定（第三頁）

第三頁「表格顏色設定」控制底色規則與特殊顏色跳過，分成三區：

- **保留顏色（不調整）**：HEX 色碼清單，命中時底色保留。每行一個或逗號分隔，可輸入 `#DDEBF7` 或 `DDEBF7`，內部統一轉成大寫 6 碼。內建預設 `DDEBF7`。
- **轉成灰色的顏色**：HEX 色碼清單，命中時改成「目標灰色」（預設 `D9D9D9`）。內建清單：`BFBFBF`、`C0C0C0`、`A6A6A6`、`808080`；比目標灰色更深的灰色不在清單內也會用既有規則轉灰。
- **指定顏色跳過整張表**：勾選「跳過特殊顏色表格」後，表格中任一格底色命中指定清單，整張表跳過版面與顏色處理（`special_color_skipped_table`）。再勾選「跳過後將指定顏色改回無色彩」則只把命中清單的儲存格底色清成無色，其他顏色不動。

第三頁底部按鈕：

- **套用顏色設定**：只套用到目前程式執行中的設定，不保存。
- **保存成預設顏色設定**：寫入 `indent_defaults.json` 的 `table_color_settings` 區塊，下次開 GUI 自動載入。
- **還原內建顏色設定**：回復 `built_in_table_color_settings()`。

兩個 checkbox 的勾選狀態屬於 `gui_defaults`，由第一頁「保存目前勾選為預設方案」保存。

## 預設方案按鈕

- **保存目前勾選為預設方案**：將 GUI 勾選狀態（含第三頁兩個 checkbox）寫入 `indent_defaults.json` 的 `gui_defaults` 區塊。
- **還原內建勾選預設**：把 GUI 勾選狀態還原成 `built_in_gui_defaults()`。

縮排設定頁另有「保存成預設樣式」與「還原內建預設」，那兩個按鈕影響 `indent_settings`；表格顏色設定頁的按鈕影響 `table_color_settings`，皆不等同於 `gui_defaults`。
