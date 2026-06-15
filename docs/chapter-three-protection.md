# 參章保護邏輯

本文件說明「參、價格形成之主要因素分析」的保護邏輯。偵測主要在 `docx_fixer/protected_region.py`，實際套用在 `docx_fixer/docx_processor.py` 與 `docx_fixer/table_pipeline.py`。

## 保護目標

目標標題是：

```text
參、價格形成之主要因素分析
```

程式內部也保留標題文字核心：

```text
價格形成之主要因素分析
```

## 偵測方式

`is_chapter_three_start_marker()` 支援兩種判斷：

- 可見文字開頭比對：段落壓縮空白後，以 `參、價格形成之主要因素分析` 開頭即視為目標章節。
- 編號判斷：如果文字本身以 `價格形成之主要因素分析` 開頭，會再檢查自動編號、樣式編號或手動編號是否為第一階層標題。

自動編號與樣式編號會透過 numbering/style lookup 判斷 outline level；必要時會依中文法律數字或中文數字推算第一階層章節前綴。

## 保護範圍

`collect_chapter_three_paragraph_ids()` 會從目標章節開始收集段落，直到下一個第一階層標題前為止。TOC 目錄段落會排除，不拿來判斷或保護。

參章保護目前只針對 `word/document.xml`。header、footer、footnotes、endnotes 不使用這個章節範圍。

## 拆分後的保護選項

現在表格保護與縮排保護已拆開：

- `skip_chapter_three_table_layout`
- `skip_chapter_three_table_color`
- `skip_chapter_three_indents`

表格版面保護只會跳過該章內表格版面處理；若底色選項允許，底色仍可處理。

表格顏色保護只會跳過該章內表格底色處理；若版面選項允許，版面仍可處理。

縮排保護只會跳過該章內：

- 段落縮排處理
- `firstLine`
- `hanging`
- tabs
- 字元縮排屬性清理
- Word COM 內文縮排補救
- styles/numbering 中被判定為參章使用的縮排定義清理

真正標題的 outline level 仍會恢復。文件不應把這個選項描述成整章完全什麼都不動。

## 「參、不要調整」整章保護

`skip_chapter_three_adjustments`（GUI：「參、不要調整（整個參、章節都不調整）」）是整章保護開關。啟用後：

- `ProcessOptions.__post_init__` 會強制 `skip_chapter_three_table_layout`、`skip_chapter_three_table_color`、`skip_chapter_three_indents` 全部為真，所以參、章內表格在 `table_pipeline.py` 會走「版面與顏色都跳過」分支而完全略過：不調整版面、不調整顏色、不調整字體、不套黑色雙線外框、不搬移表格內註記、也不加入 Word COM AutoFit 清單。
- 段落縮排、標題大綱、標題後方分隔符清理、Word COM 內文縮排補救、最終硬清理都沿用既有的參章保護路徑跳過受保護段落。

偵測方式與一般 `skip_chapter_three_*` 不同：

- 一般選項使用 `is_chapter_three_start_marker()`，綁定特定標題「參、價格形成之主要因素分析」。
- 「參、不要調整」使用 `is_section_three_chapter_marker()`，以章節編號判斷，凡第一階層章節編號解析為「參」（壹貳參的第 3 章）即觸發，標題文字可以是任何內容。

兩者都重用同一套第一階層標題判斷（已排除 TOC 目錄），所以目錄裡的「參、」不會觸發保護，只有正文的「參、」會。`collect_chapter_three_paragraph_ids()` 接受 `start_marker` 參數，整章保護啟用時改用泛用偵測器；保護範圍一樣是從「參、」標題收集到下一個第一階層標題前。

`ProtectedRegionContext` 會記錄 `section_three_protection_enabled` 與 `section_three_detection_source`（`generic_section_three_chapter_參` 或 `title_specific_價格形成之主要因素分析`），並寫進 `numbering_xml_logs` 的 `CHAPTER_THREE_SKIP_IDS` 行。

## 與移除大綱階層的關係

當 `remove_all_outline_levels` 和 `skip_chapter_three_indents` 同時啟用時，程式會先處理大綱移除，再對受保護的參章段落呼叫 `restore_outline_levels_for_protected_paragraphs()`。因此參章真正標題仍保留標題大綱層級。

## 表格 log

參章表格被保護時，`table_log` 會分別記錄：

- `chapter_three_table_layout_skipped`
- `chapter_three_table_color_skipped`

原因文字會依實際情況顯示：

- `chapter three protected table; layout and color skipped`
- `chapter three protected table; layout skipped; color allowed`
- `chapter three protected table; layout allowed; color skipped`

啟用「參、不要調整」時，`table_log` 另記：

- `skip_section_three_adjustments_enabled`
- `in_section_three_protected`
- `section_three_detection_source`
- `skipped_by_section_three_protection`

## 相容舊名稱

以下選項仍存在是為了相容舊腳本或舊呼叫方式，不建議新文件主推：

- `skip_chapter_three_tables`：舊的合併表格選項，現在代表同時跳過表格版面與表格顏色。
- `skip_all_under_chapter_three`：同時跳過表格版面、表格顏色與縮排，行為等同 `skip_chapter_three_adjustments`，但使用標題特定偵測；新文件請改用「參、不要調整」。
- `skip_special_layout_under_chapter_three`：CLI 舊名稱，會把表格版面、表格顏色與縮排都設為跳過。
