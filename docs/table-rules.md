# 表格處理規則

表格流程主要在 `docx_fixer/table_pipeline.py`、`docx_fixer/table_format.py`、`docx_fixer/table_word_com.py`，由 `docx_fixer/docx_processor.py` 對每個可處理 XML part 呼叫。

## 處理範圍

表格處理會跑在 `should_process_part()` 回傳 `True` 的 XML part：`word/document.xml`、header、footer、footnotes、endnotes。參章保護與第一階層標題判斷只針對 `word/document.xml`。

## 跳過順序

每張表格會依下列順序判斷：

1. `word/document.xml` 的第一張表格會跳過，`table_type = skipped_first_table`，原因是 `first table in word/document.xml`。
2. 若 `skip_nested_tables=True`，表格本身在另一張表格內，或表格內含另一張表格，會跳過。
3. 若表格位在參章保護範圍，且版面與底色兩者都因選項被停用，會跳過。
4. 若啟用「將表格內註記儲存格移至表格下方」，先搬移註記並刪除 cell / row（受保護的參、表格不會到這步）。
5. 搬移後重新計算 `cell_count`、`column_count` 與一般 / 特殊表格類型。
6. 命中「指定顏色跳過整張表」清單的表格會跳過。
7. 儲存格數 `cell_count <= 4` 的表格會跳過，`table_type = skipped_small_table`。
8. 其餘表格依欄數走特殊表格、一般表格或只處理顏色，套用版面、字體、顏色後，**若隱藏選項 `enable_double_black_table_borders` 為 True**，會對一般表格與特殊表格套用黑色雙線外框（預設關閉，不套用）；接著**若選項 `enable_table_footer_source_format` 為 True 且該表格版面有被調整**，套用「表格基期/資料來源格式化」。

註記搬移、黑色雙線外框與表格基期/資料來源格式化的詳細規則見下方專節。

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

底色處理由 `table_format.apply_table_color()` 與 `shading.py` 負責，規則不再寫死，改由 GUI 第三頁「表格顏色設定」、CLI 參數或 `table_color_settings.py` 的設定決定（`ProcessOptions` 的 `table_keep_colors`、`table_gray_colors`、`table_gray_target`）。

`get_shading_decision()` 判斷順序：

1. 沒有底色、`AUTO`、`NONE` 且沒有 theme 色彩時保留。
2. 底色在保留顏色清單（`keep_colors`）時保留。
3. 底色在轉灰色清單（`gray_colors`）時改成目標灰色（`gray_target`）。
4. 灰色且比目標灰色更深時，改成目標灰色。
5. 灰色且比目標灰色更淺或相同時保留。
6. 無法解析的 theme 色彩保留。
7. fill hex 不可用時保留。
8. 其餘明確的非灰色十六進位底色清除成無色。

內建預設：保留 `D9D9D9`、`F2F2F2`；轉灰色清單 `BFBFBF`、`C0C0C0`、`A6A6A6`、`808080`；目標灰色 `D9D9D9`。其他顏色不在預設保留清單內，明確的非灰色底色預設會被清除。

`table_log` 會記錄 `changed_to_gray`、`cleared_colors`、`shading_debug`，以及生效中的 `table_keep_colors`、`table_gray_colors`、`table_gray_target`。

## 特殊顏色表格跳過

啟用「跳過特殊顏色表格」（`skip_special_color_tables`）並設定指定顏色清單（`special_color_skip_colors`）後，只要表格中任一格底色命中清單，整張表就跳過：

- 判斷順序：第一張表格 → 巢狀表格保護 → 參章保護 → **特殊顏色表格** → 小表格（cell_count <= 4）→ 特殊/一般表格。
- 命中時不做任何版面調整、不套用一般底色規則、不加入 Word COM AutoFit 清單。
- `table_type = special_color_skipped_table`、`action = skipped_special_color_table`、`reason = matched special color skip list`。
- 若勾選「跳過後將指定顏色改回無色彩」（`clear_special_colors_after_skip`），只把命中指定清單的儲存格底色清成無色，其他顏色完全不動；清除格數記在 `special_color_cleared_count`。
- `table_log` 另記 `special_color_skip_matched` 與命中的 `special_color_skip_colors`。

顏色設定保存在與縮排設定共用的 `indent_defaults.json`（key 為 `table_color_settings`），EXE 版會放在執行檔同資料夾，可攜帶到其他電腦。

## 將表格內註記儲存格移至表格下方（已從 GUI 隱藏並強制關閉）

> 此功能目前已從 GUI 隱藏，且 GUI／設定檔載入一律強制 `move_table_notes_below=False`、`skip_chapter_three_table_notes=False`（見 [GUI 選項](gui-options.md) 的「已隱藏並強制關閉的選項」）。底層函式 `move_table_note_cells_below()` 與 CLI 參數 `--move-table-notes-below` 仍保留未刪除，預設亦為關閉；以下說明保留作為該函式的行為參考。若需在最後一列就地格式化「註：」cell（不搬移），請改用「表格最後一列說明格式化」。

由 GUI／CLI 的 `move_table_notes_below` 控制（預設關閉），實作在 `docx_fixer/table_notes.py` 的 `move_table_note_cells_below()`。這是**完全獨立的功能**，與表格版面、表格顏色、特殊顏色跳過不耦合：

- **可單獨執行**：`docx_processor.py` 在 `fix_table_layout or fix_color or move_table_notes_below` 時就會跑表格流程；只勾「將表格內註記儲存格移至表格下方」（不勾版面、顏色）也會搬移註記。GUI 的「尚未選擇處理項目」判斷與 CLI 的預設動作判斷都已納入 `move_table_notes_below`。
- 註記判斷使用嚴格規則 `note_detection.is_note_cell_text()`，正規式 `^註(?:\d+|[一二三四五六七八九十]+)?\s*[：:、.．]`。會處理 `註：`、`註1：`、`註2.`、`註一：`；不會處理 `註冊資料`、`註銷登記`、`註明事項`、`本表註1如下`。
- 掃描順序為由上到下、由左到右。命中的儲存格：若該列其他 cell 都空白則刪整列（`delete_row`），否則只刪該 cell（`delete_cell`）。
- 搬出的註記依掃描順序，各自成為一個段落插入在表格正下方。段落格式：標楷體（eastAsia `標楷體`、ascii/hAnsi `DFKai-SB`）、10pt（`w:sz`/`w:szCs` 值 `20`）、`outlineLvl` 9（本文）、不帶編號、不設首行縮排。
- 搬移在版面／字體／顏色／外框之前執行，之後才重算欄數、格數與表格類型，確保刪 cell / row 後新邊界仍能正確套用外框。
- **「參、不要表格註記搬移」**（`skip_chapter_three_table_notes`，預設開啟）：正文中「參、」章節內的表格不搬移註記（以泛用章節編號偵測，目錄不會觸發），章節外仍會搬移。這個選項只影響註記搬移，不會擋掉該表格的版面、顏色或外框。舊的「參、表格版面／顏色不調整」不會擋住註記搬移。`table_log` 以 `skip_chapter_three_table_notes_enabled`、`table_notes_skipped_by_chapter_three` 記錄。

## 表格外框線黑色雙線（隱藏功能，預設關閉）

黑色雙線外框是**隱藏功能**，由 `ProcessOptions.enable_double_black_table_borders` 控制，**預設 `False`**。GUI 不顯示此選項、也不保存它；一般 GUI 使用者預設不會更動表格外框線。開發者可用隱藏的 CLI 參數 `--enable-double-black-table-borders` 啟用。

`table_format.apply_double_black_table_borders()` 會重建 `w:tblBorders`，把 `top`、`left`、`bottom`、`right`、`insideH`、`insideV` 都設成 `w:val="double"`、`w:color="000000"`、`w:sz="4"`、`w:space="0"`，等於黑色雙線外框與內框。

- **預設關閉時**：不新增或改寫 `w:tblBorders`；`table_log` 的 `double_border_applied` 為 `false`、`double_border_tables` 為 `0`。不會因為表格註記搬移、表格版面或表格顏色而自動套外框。
- **隱藏選項為 True 時**：一般表格與特殊表格在套完版面／字體／顏色後才套用，發生在註記搬移與 cell / row 刪除「之後」，避免刪除後新邊界沒有雙線。只對實際走過版面／顏色處理的表格套用。
- 受參、保護的表格不套外框（保留原框線）。第一張表格、巢狀表格、小表格、特殊顏色跳過表格不在此處理。
- 黑色雙線外框只由 `enable_double_black_table_borders` 控制，與「將表格內註記儲存格移至表格下方」「參、不要表格註記搬移」完全獨立，互不耦合。
- `table_log` 以 `double_border_enabled`（隱藏選項是否啟用）與 `double_border_applied`（該表是否實際套用）記錄，摘要以 `套用黑色雙線外框的表格數` 統計。

## 表格最後一列說明格式化（獨立功能，預設關閉）

由 `ProcessOptions.enable_table_footer_source_format` 控制（**預設 `False`**），GUI 勾選項「表格最後一列說明格式化」、CLI 參數 `--enable-table-footer-source-format`（別名 `--table-footer-source-format`）皆可啟用，勾選狀態存於 `indent_defaults.json` 的 `gui_defaults`。這是**完全獨立的功能**：不依賴（已隱藏的）表格註記搬移、不依賴黑色雙線外框、不混入顏色處理，且**不搬移、不刪除、不新增任何 cell／列／段落，不改變表格結構**，只格式化最後一列中符合條件的 cell。實作在 `table_format.apply_table_footer_source_format()`。

啟用且**該表格版面有被調整**時（見「與既有跳過邏輯的關係」），依固定順序處理：

1. 整張表格內所有 run 字級設為 11pt（`w:sz`/`w:szCs` 值 `22`）。
2. 表格外圍框線（`top`/`bottom`/`left`/`right`）設為黑色雙線（`w:val="double"`、`w:color="000000"`、`w:sz="4"`、`w:space="0"`），內框 `insideH`/`insideV` 不動。
3. 若第一列只有一個儲存格，覆蓋該 cell 邊框：`top`/`left`/`right` 設為無邊框（`w:val="nil"`），`bottom` 設為黑色雙線。
4. 對最後一列每個儲存格取可見文字（合併段落、去除前後空白與零寬／cell 結尾控制字元、收斂換行與全/半形空白後）判斷。以 cell 的 XML element 去重，合併儲存格只處理一次：
   - 以「基期：」開頭：cell 內所有文字 10pt（值 `20`）、段落靠左、`left`/`right`/`bottom` 無邊框、`top` 黑色雙線。
   - 以「資料來源：」開頭：cell 內所有文字 10pt、段落靠右、`left`/`right`/`bottom` 無邊框、`top` 黑色雙線。
   - 符合 `^註(?:\d+)?[：:]`（如「註：」「註:」「註1：」「註10:」；不含「備註：」「註記：」「本註：」「說明註：」）：cell 內所有文字 10pt、段落靠左、`left`/`right`/`bottom` 無邊框、`top` 黑色雙線。
   - 不符合的儲存格不更動字級、對齊與邊框。

註記規則整合在同一個最後一列 footer 流程（`_classify_footer_cell()`），與基期/資料來源共用同一套字級、邊框處理，避免多個重疊的最後一列處理函式。註記判斷 regex `FOOTER_NOTE_PREFIX_PATTERN` 定義於 `docx_fixer/table_format.py`。

優先權（後者覆蓋前者）：最後一列 footer 規則（基期/資料來源/註記）＞ 第一列單 cell 規則 ＞ 表格外圍黑色雙線 ＞ 全表 11pt。因此最後一列的 10pt 不會被全表 11pt 蓋掉，最後一列的左右下無邊框、第一列單 cell 的上左右無邊框也不會被外圍雙線蓋掉。

邊框採**局部更新**（`set_border_double_black()`／`set_border_nil()`）：只改指定那一邊，不重建整個 `tcBorders`，因此其他 cell 原本已有的黑色雙線不會被誤清。允許消失的邊只有規則明確指定為無邊框的邊（第一列單 cell 的 `top`/`left`/`right`；命中的最後一列 cell 的 `left`/`right`/`bottom`）。

### 與既有跳過邏輯的關係

本功能被歸類為「表格版面格式」，只在該表格的版面實際被調整（`effective_fix_table_layout` 為 True）時執行：

- 第一張表格、巢狀表格、小表格（cell_count ≤ 4）、特殊顏色跳過表格：在到達本步驟前就已跳過，本功能不執行。
- 「參、不要調整」整章保護的表格：整張跳過，本功能不執行。
- 參、表格**版面**不調整：版面未被調整，本功能不執行（`table_footer_note_source_format_skipped_reason = layout not adjusted for this table`）。
- 參、表格**顏色**不調整、但版面仍調整：本功能照常執行（顏色跳過不會擋住本功能）。

### table_log 欄位

- `table_footer_note_source_format_enabled`：選項是否啟用。
- `table_footer_note_source_format_applied`：該表是否實際套用。
- `outer_double_border_applied_by_footer_source_format`：是否套用外圍黑色雙線。
- `first_row_single_cell_border_adjusted`：第一列單 cell 是否被調整。
- `footer_note_cells_adjusted`：最後一列命中的 cell 數。
- `footer_note_cell_matches`：命中類型（`note`／`base_period`／`source`）。
- `footer_note_cell_debug`：命中 cell 文字前 50 字與套用動作。
- `table_footer_note_source_format_skipped_reason`：未套用時的原因（`feature_disabled`／`layout not adjusted for this table`／各跳過原因）。

另外，每張表格 log 都會記錄已隱藏的表格註記搬移狀態：`table_note_move_gui_hidden`（恆 `true`）、`table_note_move_forced_false`（`move_table_notes_below` 是否為關）、`skip_chapter_three_table_note_move_forced_false`（`skip_chapter_three_table_notes` 是否為關）。

摘要以「套用最後一列說明格式化的表格數」統計（`summary.table_footer_source_format_tables`）。
