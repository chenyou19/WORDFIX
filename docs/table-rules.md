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
3. 讀取所在 section 的頁寬與左右邊界，算出頁面可用寬度。
4. 表格左側對齊上一個有效段落的文字起點，右側延伸到頁面可用寬度。

若特殊表格無法解析幾何資料，會改用右對齊與 autofit。

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

## 特殊表格版面

特殊表格解析到幾何資料時會套用 `apply_special_table_format()`：

- 表格左對齊：`w:jc val=left`。
- 表格寬度：`w:tblW type=dxa w=<可用寬度減文字起點>`。
- 固定配置：`w:tblLayout type=fixed`。
- 表格縮排：`w:tblInd type=dxa w=<文字起點>`。

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
