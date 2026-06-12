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

## 預設方案按鈕

- **保存目前勾選為預設方案**：將 GUI 第一頁勾選狀態寫入 `indent_defaults.json` 的 `gui_defaults` 區塊。
- **還原內建勾選預設**：把 GUI 第一頁勾選狀態還原成 `built_in_gui_defaults()`。

縮排設定頁另有「保存成預設樣式」與「還原內建預設」，那兩個按鈕影響 `indent_settings`，不等同於 GUI 第一頁的 `gui_defaults`。
