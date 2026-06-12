# 表格處理規則

表格流程主要在 `docx_fixer/table_pipeline.py`、`docx_fixer/table_format.py`、`docx_fixer/table_word_com.py`，由 `docx_fixer/docx_processor.py` 對每個可處理 XML part 呼叫。

## 處理範圍

表格處理會跑在 `should_process_part()` 回傳 `True` 的 XML part：`word/document.xml`、header、footer、footnotes、endnotes。參章保護與第一階層標題判斷只針對 `word/document.xml`。

## 跳過順序

每張表格會依下列順序判斷：

1. `word/document.xml` 的第一張表格會跳過，`table_type = skipped_first_table`，原因是 `first table in word/document.xml`。
2. 若 `skip_nested_tables=True`，表格本身在另一張表格內，或表格內含另一張表格，會跳過。
3. 若表格位在參章保護範圍，且版面與底色兩者都因選項被停用，會跳過。
4. 儲存格數 `cell_count <= 4` 的表格會跳過，`table_type = skipped_small_table`。
5. 其餘表格依欄數走特殊表格、一般表格或只處理顏色。

## 巢狀表格

`table_pipeline.py` 使用兩個條件判斷：

- `is_nested_table(tbl)`：表格本身有 `ancestor::w:tbl`，代表位於另一張表格內。
- `contains_nested_table(tbl)`：表格內部有 `.//w:tc//w:tbl`，代表表格中含有另一張表格。

預設會跳過這兩種表格。詳細說明見 [巢狀表格處理](nested-tables.md)。

## 特殊表格與一般表格

- `column_count <= 4` 時視為特殊表格。
- `column_count > 4` 時視為一般表格。
- 若只啟用底色、不啟用表格版面，會成為 `color_only_table`。

特殊表格會嘗試解析幾何資料：

1. 找表格前方最近的有效段落，空白段落會跳過，但有自動編號或樣式編號的段落仍可作為錨點。
2. 依段落層級、段落縮排、首行或凸排推算文字起點。
3. 讀取表格所在 section 的頁寬與左右邊界，算出頁面文字寬度 `text_width = page_width - left_margin - right_margin`。section 判斷以表格後方最近的 `w:sectPr` 優先（OOXML 的 sectPr 放在該節結尾），找不到才用 `w:body/w:sectPr`，避免多 section 文件抓到上一節的紙張設定。
4. 文字起點會 clamp 到 `0 <= left_indent < text_width`；表格寬度為 `width = text_width - left_indent`，保證 `tblInd + tblW` 不超過頁面文字寬度。
5. 表格左側對齊上一個有效段落的文字起點，右側精準對齊頁面右邊界（`page_width - right_margin`）。

若特殊表格無法解析幾何資料（找不到錨點段落、找不到 section 設定，或 `width <= 0`），會退回右對齊與 autofit 的安全方案，不會輸出超頁表格。

## 一般表格版面

一般表格會套用 `apply_table_format()`：

- 表格置中：`w:jc val=center`。
- 表格寬度：`w:tblW type=pct w=5000`。
- 清除固定寬度限制：移除 `w:tblGrid`，並移除各儲存格 `w:tcW`。
- 表格配置：`w:tblLayout type=autofit`。
- 後續會記錄進 Word COM AutoFit 清單。

Word COM AutoFit 在 `table_word_com.py` 執行，順序是：

1. `AutoFitContent`
2. `AutoFitWindow`

成功套用後，`table_log` 會將 `word_com_autofit_applied` 標成 `true`，`word_com_autofit_sequence` 會是 `content_then_window`。

### Word COM 失敗時的 XML fallback

`apply_table_autofit_with_word_com()` 回傳 `(logs, applied_indices, failed_indices)`；只要某張表格沒有出現 `WORD_COM_TABLE_AUTOFIT_APPLIED` 確認（包含單表錯誤、找不到表格、COM 例外、PowerShell 失敗、Word 未安裝、runner 例外），就會列入 `failed_indices`。

`docx_processor.py` 在 Word COM 跑完後，若 `failed_indices` 不為空，會立刻呼叫 `table_fallback.fallback_normal_table_autofit_in_docx()` 直接修復 `output_docx`：

- 只處理 `part_name == "word/document.xml"` 的 normal_table，用 `table_index` 定位表格，避免 header/footer 的 index 跟 Word COM `doc.Tables.Item` 對不上。
- 對每張失敗表格重新套用 `apply_table_format()` 的安全視窗寬度格式：移除 `w:tblGrid` 與所有 `w:tcW`、`jc=center`、`tblW type=pct w=5000`、`tblLayout type=autofit`，保留既有框線與 11pt 文字設定。
- log 會輸出 `WORD_COM_TABLE_AUTOFIT_FALLBACK_STARTED failed_records_count=N`、每張 `WORD_COM_TABLE_AUTOFIT_FALLBACK_APPLIED global_table_index=...`、`WORD_COM_TABLE_AUTOFIT_FALLBACK_DONE applied=N`；fallback 本身失敗時輸出 `WORD_COM_TABLE_AUTOFIT_FALLBACK_ERROR`。

`table_log` 每張表格會記錄：

- `word_com_autofit_applied`：Word COM 是否成功。
- `word_com_autofit_fallback_applied`：是否由 XML fallback 修復。
- `word_com_autofit_status`：`word_com`（COM 成功）、`xml_fallback`（COM 失敗但已自行修復，不算整體失敗）、`failed`（COM 與 fallback 都失敗）、`not_needed`（非 normal_table 或未啟用 Word COM）。

CLI 會輸出 `word_com_table_autofit_applied_count`、`word_com_table_autofit_fallback_count`、`word_com_table_autofit_failed_count`；GUI 即使勾選不輸出 log 檔，也會在畫面 log 與完成訊息顯示這三個數字，`failed` 大於 0 時顯示明確警告。

## 特殊表格版面

特殊表格解析到幾何資料時會套用 `apply_special_table_format()`：

- 先清除舊寬度限制：移除 `w:tblGrid` 與所有儲存格 `w:tcW`，避免 Word 沿用原本欄寬把表格撐出頁面。
- 表格左對齊：`w:jc val=left`。
- 表格寬度：`w:tblW type=dxa w=<可用寬度減文字起點>`。
- 固定配置：`w:tblLayout type=fixed`。
- 表格縮排：`w:tblInd type=dxa w=<文字起點>`。
- 重建 `w:tblGrid`：以 `table_column_count()` 取得欄數，將 `tblW` 平均分配到各欄（餘數分給前面的欄），欄寬總和不超過 `tblW`。
- 重設各儲存格 `w:tcW type=dxa`；有 `gridSpan` 的儲存格寬度為跨欄欄寬總和。

特殊表格套用幾何資料時，`table_log` 會記錄：

- `special_left_indent_twips`：表格左縮排。
- `special_width_twips`：表格寬度。
- `special_text_width_twips`：頁面文字寬度。
- `special_right_edge_twips`：`left_indent + width`。
- `special_overflow_twips`：`max(0, right_edge - text_width)`，正常應為 0。

走 fallback 時這些欄位為 `none`。

特殊表格無法解析幾何資料時會套用 `apply_autofit_contents_right_format()`：

- 表格右對齊：`w:jc val=right`。
- 表格寬度：`w:tblW type=auto w=0`。
- 表格配置：`w:tblLayout type=autofit`。

## 表格內文字與列格式

`_apply_table_content_format()` 會套用在一般表格、特殊表格與右對齊 autofit 表格：

- 每列列高：`w:trHeight val=340 hRule=atLeast`。
- 儲存格垂直置中：`w:vAlign val=center`。
- 表格內段落置中：`w:jc val=center`。
- 表格內段落間距：before `0`、after `0`、line `240`、lineRule `auto`。
- 表格內 run 字級：`w:sz val=22`、`w:szCs val=22`，也就是 11pt。

## 表格底色

底色處理由 `table_format.apply_table_color()` 與 `shading.py` 負責：

- 沒有底色、`AUTO`、`NONE` 且沒有 theme 色彩時保留。
- 灰色且比預設灰 `D9D9D9` 更深時，改成 `D9D9D9`。
- 灰色且比預設灰更淺或相同時保留。
- 無法解析的 theme 色彩保留。
- 明確的非灰色十六進位底色會清除成無色。

`table_log` 會記錄 `changed_to_gray`、`cleared_colors` 與 `shading_debug`。
