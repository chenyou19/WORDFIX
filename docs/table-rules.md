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
8. 其餘表格依欄數走特殊表格、一般表格或只處理顏色，套用版面、字體、顏色後，**若隱藏選項 `enable_double_black_table_borders` 為 True**，會對一般表格與特殊表格套用黑色雙線外框（預設關閉，不套用）；接著**若選項 `enable_table_footer_source_format` 為 True 且該表格符合 footer eligibility**，僅在此處「記錄」該表格，實際的「表格最後一列說明格式化」延後到 Word COM AutoFit 與 fallback 之後的最後 post-process 才套用（見下方專節）。一般表格需有表格版面處理；參、表格若只是被 `skip_chapter_three_table_layout` 保護而關掉 effective layout，但全域仍有要求表格版面處理，也會被記錄。

註記搬移、黑色雙線外框與表格最後一列說明格式化的詳細規則見下方專節。

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

## 表格最後一列說明格式化（獨立功能，核心預設關閉、GUI 內建勾選）

由 `ProcessOptions.enable_table_footer_source_format` 控制；核心 `ProcessOptions` 與 CLI 預設為 `False`，GUI 內建預設 `gui_defaults["enable_table_footer_source_format"] = True`，勾選狀態存於 `indent_defaults.json` 的 `gui_defaults`。GUI 勾選項「表格最後一列說明格式化」、CLI 參數 `--enable-table-footer-source-format`（別名 `--table-footer-source-format`）皆會傳入同一個布林值。這是**完全獨立的功能**：不依賴（已隱藏的）表格註記搬移、不依賴黑色雙線外框、不混入顏色處理，且**不搬移、不刪除、不新增任何 cell／列／段落，不改變表格結構**。它會先格式化表格底部連續說明列，再依最後一列是否為 footer 分流：一般資料表底部保留黑色雙線；footer 說明區底部不顯示框線。元素層級實作在 `table_format.apply_table_footer_source_format()`。

**執行時機（重要）**：本功能是**最後一個表格格式化步驟**。XML pipeline（`table_pipeline.py`）只在處理該表格時把它「記錄」到 `summary.table_footer_source_format_records`，實際格式化延後到 `docx_processor.py` 於 **Word COM AutoFit 與 XML fallback 之後、final note alignment 之前**呼叫 `table_footer_postprocess.apply_table_footer_source_format_in_docx()` 才套用。

原因：一般表格會被加入 Word COM AutoFit 清單，AutoFit 會重存整份文件；若 AutoFit 失敗，fallback `fallback_normal_table_autofit_in_docx()` 會重跑 `apply_table_format()`，把整表 run 改回 11pt、段落改回置中。若 footer 在 XML pipeline 階段就套用（Word COM 之前），一般表格的 footer 會被 AutoFit／fallback 覆蓋（特殊表格不進 Word COM 所以原本正常）。改成最後 post-process 後，一般表格與特殊表格都正確。post-process 只用記錄中的 `(part_name, table_index)` 精準重新定位該表格（與 fallback 相同的定位方式），不會無差別掃全部表格。

啟用且該表符合 eligibility 時（見「與既有跳過邏輯的關係」），post-process 對每張記錄到的表格依固定順序處理：

1. 整張表格內所有 run 字級設為 11pt（`w:sz`/`w:szCs` 值 `22`）。
2. 表格外圍框線（`top`/`bottom`/`left`/`right`）設為黑色雙線（`w:val="double"`、`w:color="000000"`、`w:sz="4"`、`w:space="0"`），內框 `insideH`/`insideV` 不動。
3. 若第一列只有一個儲存格，覆蓋該 cell 邊框：`top`/`left`/`right` 設為無邊框（`w:val="nil"`），`bottom` 設為黑色雙線。
4. **從表格底部連續往上掃描 footer 列**。對每一列逐 cell 取可見文字（合併段落、去除前後空白與零寬／cell 結尾控制字元、收斂換行與全/半形空白後）判斷，以 cell 的 XML element 去重，合併儲存格每列只處理一次：
   - 該列只要**任一** cell 命中以下任一規則，即視為 footer 列，格式化該列**所有命中**的 cell，並繼續往上一列檢查。
   - 命中規則與套用格式：
     - 以「基期：」開頭：cell 內所有文字 10pt（值 `20`）、段落靠左、`left`/`right`/`bottom` 無邊框。
     - 以「資料來源：」開頭：cell 內所有文字 10pt、段落靠右、`left`/`right`/`bottom` 無邊框。
     - 符合 `^註(?:\d+)?[：:]`（如「註：」「註:」「註1：」「註10:」；不含「備註：」「註記：」「本註：」「說明註：」）：cell 內所有文字 10pt、段落靠左、`left`/`right`/`bottom` 無邊框。
     - footer block 最上方那一列的**所有實體 cell** 會套 `top` 黑色雙線，作為資料區與 footer 區的分隔線；多列 footer 時，後續 footer 列的 `top` 會設為 `nil`，列與列之間不顯示線。
     - 該列其餘不命中的儲存格不更動字級與對齊，但若位於最後一列 footer，也會在最後分流中清掉 `bottom`。
   - **停止條件**：某一列若沒有任何 cell 命中 footer 規則（含空白列），立即停止往上掃描。因此只處理表格底部**連續**的說明列；中間夾著資料列或空白列就會中斷，非連續的註記列不會被處理。不會掃整張表。
5. 最後依 `footer_rows_detected` 分流：
   - `footer_rows_detected=False`：一般資料表底部使用 `data_double` mode。`w:tblBorders/w:bottom` 與最後可見底邊 cell 的 `w:tcBorders/w:bottom` 都設為黑色雙線（`w:val="double"`、`w:sz="4"`、`w:space="0"`、`w:color="000000"`）。若最後列有 `gridSpan`，以實體 cell 承擔跨欄底線；若最後列含 `vMerge=continue`，會同時處理 continuation cell 與對應 restart owner，且不破壞合併 XML。
   - `footer_rows_detected=True`：footer 說明區使用 `footer_none` mode。`w:tblBorders/w:bottom` 與最後一個 footer 實體列的所有實體 cell `w:tcBorders/w:bottom` 都設為 `nil`，避免「註：」「基期：」「資料來源：」下方被補回雙線。資料區與 footer 區之間只保留 footer block 最上方列的 `top` 黑色雙線。

邊框容器本身也會維持 WordprocessingML schema 位置：`w:tblBorders` 會放在 `tblLayout` 等後續 table property 前，`w:tcBorders` 會放在 `vAlign` 等後續 cell property 前。若輸入檔或舊版輸出已經有 `tblLayout -> tblBorders` 或 `vAlign -> tcBorders` 的錯序，本步驟會 relocate 既有 border 容器，不會新增第二個容器，也不會重建整張 `tblBorders` 而誤改 `insideH`/`insideV`。

註記規則整合在同一個 footer 流程（`_classify_footer_cell()`），與基期/資料來源共用同一套字級、邊框處理。註記判斷 regex `FOOTER_NOTE_PREFIX_PATTERN` 定義於 `docx_fixer/table_format.py`。

優先權（後者覆蓋前者）：最後底邊 mode（`data_double` 或 `footer_none`，只覆蓋 bottom）＞ footer 列規則（基期/資料來源/註記）＞ 第一列單 cell 規則 ＞ 表格外圍黑色雙線 ＞ 全表 11pt。因此 footer cell 的 10pt 不會被全表 11pt 蓋掉，footer cell 的左右無邊框、第一列單 cell 的上左右無邊框也不會被外圍雙線蓋掉；最後一列 footer cell 的 bottom 會在 `footer_none` mode 中保持 `nil`。

邊框採**局部更新**（`set_border_double_black()`／`set_border_nil()`）：只改指定那一邊，不重建整個 `tcBorders`，因此其他 cell 原本已有的黑色雙線不會被誤清。允許消失的邊只有規則明確指定為無邊框的邊（第一列單 cell 的 `top`/`left`/`right`；命中的 footer cell 的 `left`/`right`；footer block 內部列的 `top`；以及 `footer_none` mode 中 footer 最末列所有實體 cell 的 `bottom`）。

### 與既有跳過邏輯的關係

本功能仍以「有要求表格版面處理」作為安全前提，但參、表格的版面／顏色保護不得阻擋它：

- 第一張表格、巢狀表格、小表格（cell_count ≤ 4）、特殊顏色跳過表格：在到達本步驟前就已跳過，本功能不執行。
- 全域 `fix_table_layout=False`：不會因為本功能啟用而自動套 footer（`table_footer_note_source_format_skipped_reason = layout not adjusted for this table`）。
- 一般表格：`effective_fix_table_layout=True` 時才記錄並在 final post-process 套用。
- 參、表格：若全域 `fix_table_layout=True`，即使 `skip_chapter_three_table_layout=True` 使 `effective_fix_table_layout=False`，仍會記錄並在 final post-process 套用；`skip_chapter_three_table_color=True` 也不會阻擋。這只套 footer 格式，不會呼叫一般 `process_table()`，因此參、表格的 width、layout、欄寬與顏色保護仍維持。

### table_log 欄位

- `table_footer_note_source_format_enabled`：選項是否啟用。
- `table_footer_note_source_format_should_apply`：XML pipeline 判定該表需要 footer 格式化（已記錄、待 post-process 套用）。
- `table_footer_note_source_format_applied`：final post-process 是否實際套用（由 post-process 回寫）。
- `outer_double_border_applied_by_footer_source_format`：是否套用外圍黑色雙線。
- `table_bottom_border_mode`：最後底邊決策。`data_double` 表示一般資料表底部套黑色雙線；`footer_none` 表示 footer 說明區底部清為無線；`not_applied` 表示未套用。
- `table_bottom_border_cell_count`：本次 mode 實際處理的底邊 cell 數。
- `table_bottom_border_xml_verified`：XML 是否符合該 mode 的預期；這只表示最終 DOCX XML 條件通過，不等於 Word 實際畫面或像素級驗證。
- `table_bottom_border_verify_detail`：XML 驗證摘要，包含 `tbl_bottom`、`last_row_tc_bottoms`、`table_border_schema_order_valid`、`tblPr_child_order`、`last_row_tcPr_child_orders`、`last_row_grid_span_sum`、`last_row_vmerge_states` 等。
- `table_bottom_double_border_applied`：舊欄位，只有 `data_double` mode 時才會是 `true`；`footer_none` 不會把「無線」記成雙線成功。
- `table_bottom_double_border_cell_count`：舊欄位，`data_double` mode 的底邊 cell 數；`footer_none` 為 `0`。
- `table_bottom_double_border_xml_verified`：舊欄位，僅驗證 `data_double`；`footer_none` 為 `false`。
- `table_bottom_double_border_verify_detail`：舊欄位，保留相同 XML 摘要供相容讀取。
- `footer_terminal_bottom_none_applied`：footer mode 是否已清除 footer 最末列下方底線。
- `footer_terminal_bottom_none_cell_count`：footer 最末列實際清除 bottom 的實體 cell 數。
- `last_row_physical_cell_count`、`last_row_grid_span_sum`、`last_row_vmerge_states`、`last_row_bottom_edge_target_count`：最後列實體 cell、邏輯欄寬與垂直合併診斷。
- `table_border_schema_order_valid`、`tblPr_child_order`、`last_row_tcPr_child_orders`：`tblBorders`/`tcBorders` 是否位於合法 OOXML schema 位置，以及實際 child order 摘要。
- `first_row_single_cell_border_adjusted`：第一列單 cell 是否被調整。
- `footer_row_count`：底部連續 footer 列實際處理的列數。
- `footer_cell_matches`：每列命中的類型（由上往下，例如 `note | base_period,source`）。
- `footer_note_cells_adjusted`：所有 footer 列命中並格式化的 cell 總數。
- `footer_note_cell_matches`：所有命中 cell 的類型（`note`／`base_period`／`source`）。
- `footer_note_cell_debug`：每個命中 cell 文字前 50 字與套用動作。
- `table_footer_note_source_format_skipped_reason`：未套用時的原因（`feature_disabled`／`layout not adjusted for this table`／各跳過原因）。

另外，每張表格 log 都會記錄已隱藏的表格註記搬移狀態：`table_note_move_gui_hidden`（恆 `true`）、`table_note_move_forced_false`（`move_table_notes_below` 是否為關）、`skip_chapter_three_table_note_move_forced_false`（`skip_chapter_three_table_notes` 是否為關）。

摘要以「套用最後一列說明格式化的表格數」統計（`summary.table_footer_source_format_tables`）。
