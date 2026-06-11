# Word DOCX Fixer / WORDFIX

WORDFIX 是一個用來批次修正 Word `.docx` 文件版面的工具，提供 GUI 與 CLI 兩種使用方式。專案重點在於協助整理文件中的表格版面、表格底色、段落縮排與大綱層級，並在需要時搭配 Word COM 進一步驗證內文字號與補救縮排結果。

## 功能簡介

- 修正表格版面，依欄數與章節位置套用既有規則。
- 修正表格底色，將不符合規則的底色轉灰或清除。
- 修正段落大綱層級與編號段落縮排。
- 可選擇對序言區段另外套用縮排或大綱設定。
- 可輸出段落處理記錄 `process_log` 與表格處理記錄 `table_log`，方便追查每次修正內容。

## GUI 使用方式

在專案目錄執行：

```powershell
python main.py
```

啟動後可在 GUI 中：

- 選擇輸入檔與輸出檔。
- 勾選是否修正表格版面、表格底色與段落。
- 開啟縮排設定視窗，調整各層級的縮排預設值。
- 視需要啟用序言處理、Word COM 字號確認等選項。
- GUI 預設勾選「不要輸出 log 檔」，因此處理完成後只會輸出修正後的 docx。
- 如果需要除錯或檢查表格處理細節，請取消勾選「不要輸出 log 檔」，程式才會額外輸出 process log、table log、heading suffix log。

GUI 預設勾選 `參、價格形成之主要因素分析：表格不調整` 與 `參、價格形成之主要因素分析：縮排不調整`，兩者可獨立切換。

## CLI 使用方式

基本範例：

```powershell
python main.py input.docx output.docx --table --color --paragraph
```

常見參數：

- `--table`：修正表格版面。
- `--color`：修正表格底色。
- `--paragraph`：修正本文段落的大綱層級與縮排。
- `--remove-all-outline`：移除文件中的大綱層級設定。
- `--indent-preface`：對序言區段套用縮排規則。
- `--outline-preface`：對序言區段套用 Word 大綱層級。
- `--level1-level2-body-first-line-indent`：在第 1、2 層標題下的內文，額外套用首行縮排。
- `--level2-body-first-line-indent`：相容舊參數名稱，作用同上。
- `--word-com-check-body-font`：當 XML 判斷內文字號不是 14pt 時，改用 Word COM 再確認。
- `--skip-chapter-three-tables` / `--no-skip-chapter-three-tables`：控制 `參、價格形成之主要因素分析` 區段表格版面與底色是否跳過，預設跳過。
- `--skip-chapter-three-indents` / `--no-skip-chapter-three-indents`：控制該區段段落縮排、tabs、首行縮排與 Word COM 內文縮排補救是否跳過，預設跳過。
- `--skip-all-under-chapter-three`、`--no-skip-all-under-chapter-three`、`--skip-special-layout-under-chapter-three`：相容舊參數名稱，會映射到上述兩個新選項。
- `--no-log` / `--skip-log-output`：不輸出 process log、table log、heading suffix log。CLI 預設仍會輸出 log 檔。
- `--quiet`：減少終端輸出訊息。

## 主要處理選項說明

- 表格修正與段落修正可獨立啟用，也可同時使用。
- `--paragraph` 主要作用於本文段落，不會直接把所有 XML 區域都視為本文處理。
- `--indent-preface` 與 `--outline-preface` 只影響序言區段，不改變本文既有邏輯。
- `--word-com-check-body-font` 是 XML 字號判斷不足時的補救流程，避免誤套用本文縮排。

## 表格處理規則

- 第一個本文表格會依現有邏輯略過，不直接套用一般修正。
- 儲存格數量過少的表格會依既有規則略過。
- 欄數較少的表格會使用專案既有的特殊版面處理。
- `參、價格形成之主要因素分析` 區段的表格與縮排保護可獨立控制。表格保護只跳過表格版面與底色修正；縮排保護只跳過段落縮排、firstLine、hanging、tabs、字元縮排清理與 Word COM 內文縮排補救，真正標題的 outline level 仍會恢復。

## 段落 / 大綱 / 縮排處理規則

- 本文段落依既有規則辨識手動編號、自動編號與樣式編號。
- 會根據編號層級套用對應的縮排與 Word `outlineLvl`。
- 本文內一般段落會依最近一個標題層級決定內文縮排。
- 序言區段可依選項決定是否套用縮排與大綱層級。
- 目錄（TOC）相關段落會跳過，避免誤改 Word 自動產生的目錄內容。
- numbering.xml 的 `suffix=nothing` 等既有處理邏輯維持不變。

## 預設設定檔位置與用途

- 預設設定共用單一檔案：`indent_defaults.json`。
- `indent_defaults.json` 內的 `indent_settings` 區塊是縮排預設樣式；GUI 縮排設定頁按「保存成預設樣式」後會寫入這個區塊，GUI 與 CLI 都會載入使用。
- `indent_defaults.json` 內的 `gui_defaults` 區塊是 GUI 第一頁勾選狀態的預設方案；按「保存目前勾選為預設方案」後會寫入這個區塊。
- 一般以原始碼執行時，`indent_defaults.json` 放在目前工作目錄，也就是啟動 `python main.py` 時所在的資料夾。
- 打包成 EXE 後，`indent_defaults.json` 放在 EXE 同層目錄。
- 若檔案存在，程式會優先載入其中設定；若不存在，則使用程式內建預設。
- 要把 EXE 帶到另一台電腦沿用設定時，只要把 EXE 與 `indent_defaults.json` 放在同一層目錄即可。

## `process_log` 與 `table_log` 輸出說明

- `process_log`：記錄段落、大綱、縮排、Word COM 驗證與其他文字處理過程。
- `table_log`：記錄每個表格的所在 XML、表格索引、欄數、處理動作、略過原因與底色處理結果。
- 這些 log 可用來追查某一段落或某一張表格為何被修改、略過或改用補救流程。

## 注意事項

- 不要直接覆寫原始 `.docx` 檔案，輸入與輸出路徑應分開。
- 建議在批次處理前先備份原始文件。
- 若使用 EXE 版本，可在 EXE 同層保留 `indent_defaults.json`，方便攜帶自訂縮排與 GUI 勾選設定。
- 對於特殊格式文件，建議先用少量樣本測試，再批次處理正式檔案。
