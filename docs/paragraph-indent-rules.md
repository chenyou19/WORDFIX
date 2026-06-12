# 段落、大綱與縮排規則

段落處理主要在 `docx_fixer/docx_processor.py` 與 `docx_fixer/outline.py`。縮排預設來自 `docx_fixer/constants.py`，可由 `indent_defaults.json` 覆寫。

## 處理範圍

- 主要段落階層邏輯只在 `word/document.xml` 執行。
- header、footer、footnotes、endnotes 仍可能被其他 XML 清理流程處理，但不跑本文段落階層判斷。
- 表格內段落會被跳過，避免把表格文字當成本文標題或內文。
- TOC 目錄段落會被跳過，避免誤改 Word 自動產生的目錄內容。

## 目錄保護

`outline.py` 會收集兩類 TOC：

- Word 欄位或樣式形成的目錄，例如 `TOC1`、`TOCHeading`。
- 可見文字是「目錄」、「目录」或 `TOC` 的純文字目錄區段，直到主本文開始標記前。

這些段落不會套用一般大綱或縮排格式。

## 編號來源與判斷順序

段落層級支援三種來源，判斷順序如下：

1. **自動編號**：段落本身有 `w:pPr/w:numPr` 時，先以 `numId` + `ilvl` 查 numbering.xml（`detect_valid_auto_heading_level()`）。只有 `numFmt` + `lvlText` 明確符合支援的標題格式（例如 `decimal` + `（%1）` → 階層 5）才算有效標題；bullet、缺 `numFmt`/`lvlText`、或查不到對應格式時一律無效，**不可只靠 `ilvl` fallback 成標題階層**。Word 自動編號的標號不會出現在 `paragraph_text(p)` 裡，因此自動編號不要求可見文字有前綴。
2. **可見文字前綴**（手動編號）：自動編號無效時，再檢查 `paragraph_text(p)` 是否以 `壹、`、`一、`、`（一）`、`1.`、`（1）`、`A.`、`（A）`、`a.`、`（a）` 等支援格式開頭。殘留無效 `numPr` 但文字有手動前綴的段落，會移除 `numPr` 並以手動前綴決定層級。
3. **樣式編號**：前兩步都沒有結果時，由段落樣式的 numbering identity 查 level lookup；lookup 只含格式驗證過的項目，無法確認就不當標題。

以上都沒有結果時，段落視為一般內文（依最近標題層級套用內文縮排）或 unknown；殘留 `numPr` 的內文不得被誤判成階層大綱。自動編號有效/無效都會寫入 log（`auto numbering valid: ...` / `auto numbering skipped: missing or unsupported numbering format; ...`）。

殘留無效 `numPr` 的內文走 body indent 時，除了移除段落層 `w:numPr`，若段落的 paragraph style 在 style numbering lookup 中（代表 styles.xml 直接或經 basedOn 繼承綁定 numbering），會**強制**呼叫 `normalize_paragraph_style_to_none()` 移除 `w:pStyle`，並把殘留的 `outlineLvl 0~8` 改回 9——否則 Word 重新開啟時會從樣式層級自行補回「1.」之類的編號。此行為不依賴 GUI 的 `normalize_body_style_to_none` 選項，log 會記錄 `invalid auto numbering body paragraph: removed paragraph numPr and numbering style; ...; style_numbering_removed=True; reason=style-level numbering would reappear in Word`。style 不帶 numbering 的內文則保留原樣式。有效的 auto numbering 標題不受影響，樣式與編號都保留。

因為移除帶 numbering 的 `w:pStyle` 會讓段落失去原本從樣式繼承的字體與字級，強制移除樣式後會接著呼叫 `apply_body_font_after_numbering_cleanup()`，把字型直接寫回段落：

- 段落層 `w:pPr/w:rPr` 與每個有可見文字的 run 都寫入 `w:rFonts`（ascii/eastAsia/hAnsi/cs = 標楷體）、`w:sz val=28`、`w:szCs val=28`（14pt）。
- 只覆寫 rFonts、sz、szCs；粗體、底線、顏色等其他 run 屬性保留。
- 只在「無效 auto numbering 內文且實際移除了帶 numbering 的樣式」時觸發，不會無條件改所有內文字體；只移除 numPr、樣式不帶 numbering 的內文不受影響。
- log 記錄 `invalid auto numbering body paragraph font normalized: font=標楷體; size=14pt; ...; reason=numbering/style cleanup removed inherited formatting`；body indent debug record 另含 `font_normalized_after_numbering_cleanup`、`normalized_font_name`、`normalized_font_size_pt`。

特殊情況：段落同時有「有效自動編號」與「可見手動標題前綴」（例如手動標題殘留清單編號）時，維持手動標題優先並移除殘留 `numPr`。

過長且不像標題的註記型段落會跳過樣式編號推論，避免誤判成標題。

## 本文開始

主本文由 `is_processing_start_marker()` 判斷。可見文字支援 `壹、序言`、`壹、前言`，或文字核心是 `序言` 且層級為第一階層。程式需要看到兩次處理起始標記後，才把後續段落視為主要本文。

## 本文標題與內文

進入主本文後：

- 偵測到標題層級時，會套用本文縮排與 Word outline level。
- 本文縮排使用 `TEMPLATE_OUTLINE_INDENTS`。
- 手動編號段落會移除段落自身的 `numPr`，並清理編號後方多餘空白或 tab。
- 編號段落會記錄 measurement 與 debug log。
- 普通內文段落會依最近一個標題層級決定縮排。
- 普通內文只在字號判定允許時直接套用；XML 判斷不是 14pt 時可依選項交給 Word COM 補救。

`level1_level2_body_first_line_indent` 會讓第 1、2 層標題下方普通內文套用 `560` twips 首行縮排，約等於兩個中文字。未啟用時，普通內文的 `firstLine` 會清成 `0`。

## 序言區段

主本文開始前可獨立套用序言規則：

- `indent_preface_paragraphs`：對序言段落套用縮排。
- `outline_preface_paragraphs`：對序言段落套用 outline level。

序言前縮排使用 `PREFACE_OUTLINE_INDENTS`。偵測到的層級會映射到前言縮排表，例如偵測層級 1 會使用前言第 1 階。

## GUI 第二頁欄位

GUI 的縮排設定頁把每一階拆成三個欄位：

- 標號起點 cm：對應 `number_start_cm`。
- 懸掛 cm / 凸排距離：對應 `hanging_cm`。
- 內文起點 cm：對應 `body_left_cm`。

保存後會寫入 `indent_defaults.json` 的 `indent_settings` 區塊。載入後會覆寫執行中的 `TEMPLATE_OUTLINE_INDENTS` 與 `PREFACE_OUTLINE_INDENTS`。

## Word COM 字號補救

`word_com_check_body_font_when_xml_not_14` 是 XML 字號判斷不足時的補救流程：

1. XML 先判斷普通內文的主要字號。
2. 若 XML 判斷內文字號不是 14pt，且選項開啟，先把段落記錄成待 Word COM 確認。
3. Word COM 開啟輸出 docx 後，檢查實際 dominant font size。
4. 只有 Word COM 判斷為 14pt 時，才把內文縮排補寫回 docx XML。

如果選項未開啟，XML 判斷不是 14pt 的普通內文會直接略過內文縮排。
