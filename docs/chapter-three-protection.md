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
- `skip_chapter_three_numbering_suffix_cleanup`

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

## 「參、不要清理編號後綴 tab/space」

`skip_chapter_three_numbering_suffix_cleanup`（GUI：「參、不要清理編號後綴 tab/space」，**預設勾選**）只控制 `word/numbering.xml` 的後綴清理，不影響表格版面、表格顏色或段落縮排。

**只用精準 (numId, ilvl) pair，不用整包 abstractNumId。** Word 的 `abstractNumId` 是一整套編號定義，「壹、貳、參、肆」常共用同一個 `abstractNumId`；若整個 `abstractNumId` 排除，會連帶讓共用該定義的其他章節 / 其他 level 都不清理。因此參章保護**只使用 `chapter_three_numbering_pairs`（參章實際用到的 `(numId, ilvl)`）**，不使用 `chapter_three_num_ids` 或 `chapter_three_abstract_ids` 整包排除。TOC 仍維持原本的 abstractId / numId / pair 排除。

兩個清理階段都據此精準保護：

- `apply_numbering_outline_format()`（`docx_fixer/numbering.py`）：**先判斷 `should_skip_numbering()` 再 sanitize**（被跳過的 level 完全不呼叫 `sanitize_numbering_level_suffix_tabs_and_text()`，`w:suff`／`w:pPr/w:tabs`／`w:lvlText` 結尾空白全保留）。`./w:abstractNum/w:lvl` 沒有 numId，函式因此把 `excluded_numbering_pairs` 透過 `numId→abstractNumId` 轉成 `protected_abstract_levels`（`(abstractNumId, ilvl)`），只跳過對應的 ilvl，不會排除整個 abstractNumId。
- final `force_clean_numbering_suffix_tabs_in_docx()`：參章 pair 以新參數 `protected_numbering_pairs` 傳入，作為「硬保護」在 `should_skip_level()` 最前面判斷，勝過 body-heading 的「重新納入」；body-heading re-include 集合維持原本完整內容（不再整包扣除參章）。函式同樣把 pair 轉成 `(abstractNumId, ilvl)` 只保護該 level。

`docx_processor.py` 以新選項決定集合：

- 啟用：`numbering_suffix_excluded_numbering_pairs` = TOC pairs ∪ 參章 pairs；`numbering_suffix_excluded_num_ids/abstract_ids` = **只含 TOC**；final cleanup 額外傳 `protected_numbering_pairs` = 參章 pairs。並寫 `CHAPTER_THREE_NUMBERING_SUFFIX_CLEANUP_SKIP enabled=true protected_pairs=... protected_abstract_levels=... protected_abstractIds_not_used_for_chapter_three=true`。
- 停用：排除集合只含 TOC；不傳 `protected_numbering_pairs`。寫 `CHAPTER_THREE_NUMBERING_SUFFIX_CLEANUP_SKIP enabled=false`。

結果：同一個 `abstractNumId` 內，只有參章實際用到的 ilvl 被保留，其餘 ilvl 仍會清理；其他章節用不同 ilvl 的編號照常清理。唯一無法區分的情況是某章節與參章用到**完全相同的 `(abstractNumId, ilvl)`**（同一個 lvl element）——此時該 level 會一併保留，因為定義本身是共用的。此選項不會單獨觸發表格版面/顏色或段落縮排保護，也不影響 numbering.xml 的字元縮排清理（char-indent 仍用 TOC-only 排除）。

## 「參、不要表格註記搬移」

`skip_chapter_three_table_notes`（GUI：「參、不要表格註記搬移」，預設勾選）只在啟用「將表格內註記儲存格移至表格下方」（`move_table_notes_below`）時生效，而且**只控制表格註記搬移**，不會跳過表格版面、表格顏色、段落縮排、外框線或其他處理：

- 正文中「參、」章節內的表格不搬移註記；章節外的表格仍會搬移。
- 不影響版面／顏色／外框：受此選項保護的參、表格仍會照常套用版面、顏色與黑色雙線外框（除非另外用 `skip_chapter_three_table_layout` / `_color` 保護）。
- 舊的「參、表格版面不調整」「參、表格顏色不調整」不會擋住表格註記搬移，只有這個選項會。

偵測使用獨立、泛用的 `is_section_three_chapter_marker()`，以章節編號判斷（第一階層編號解析為「參」即第 3 章），標題文字不限；目錄中的「參、」不會觸發。這個註記保護區段獨立收集（`document_section_three_note_paragraph_ids` / `is_table_in_section_three_for_notes()`），與版面／顏色保護的 `is_table_protected()` 區段分開，因此勾選它不會改變哪些表格的版面或顏色被保護。

`ProtectedRegionContext.from_document(collect_section_three_note_region=...)` 在 `move_table_notes_below and skip_chapter_three_table_notes` 時建立此區段，並寫進 `numbering_xml_logs` 的 `SECTION_THREE_TABLE_NOTE_SKIP_IDS` 行。

## 已移除的「參、不要調整」整章保護（相容說明）

舊版 GUI 曾有「參、不要調整（整個參、章節都不調整）」選項，現已從 GUI 移除。`skip_chapter_three_adjustments` 欄位仍保留在 `ProcessOptions` 與 CLI（`--skip-chapter-three-adjustments` / `--protect-section-three`）作為相容舊腳本用途：啟用時 `__post_init__` 會強制三個 `skip_chapter_three_*` 為真，並改用泛用 `is_section_three_chapter_marker()` 偵測，`ProtectedRegionContext` 記錄 `section_three_detection_source`。它**不再控制表格註記搬移**（改由 `skip_chapter_three_table_notes` 負責）。舊設定檔若仍含 `skip_chapter_three_adjustments` 欄位，GUI 載入時會忽略，不會報錯。

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

表格註記搬移相關，`table_log` 另記：

- `move_table_notes_below_enabled`
- `skip_chapter_three_table_notes_enabled`
- `table_notes_skipped_by_chapter_three`（因「參、不要表格註記搬移」而未搬移該表格註記）

`skip_chapter_three_adjustments`（相容）啟用時另記：

- `skip_section_three_adjustments_enabled`
- `in_section_three_protected`
- `section_three_detection_source`
- `skipped_by_section_three_protection`

## 相容舊名稱

以下選項仍存在是為了相容舊腳本或舊呼叫方式，不建議新文件主推：

- `skip_chapter_three_tables`：舊的合併表格選項，現在代表同時跳過表格版面與表格顏色。
- `skip_all_under_chapter_three`：同時跳過表格版面、表格顏色與縮排，使用標題特定偵測。
- `skip_chapter_three_adjustments`：舊「參、不要調整」整章保護開關，GUI 已移除，僅保留為相容舊腳本（見上方說明）。
- `skip_special_layout_under_chapter_three`：CLI 舊名稱，會把表格版面、表格顏色與縮排都設為跳過。
