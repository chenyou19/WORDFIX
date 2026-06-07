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

`gui_app.py` 目前介面文字已正常，這次不特別更動 GUI 顯示內容。

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
- `--skip-all-under-chapter-three`：從 `參、` 標題本身開始，到下一個同級第一層標題前，完全不調整該區段內容。
- `--skip-special-layout-under-chapter-three`：相容舊參數名稱，作用同 `--skip-all-under-chapter-three`。
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
- `參、` 區段可透過 `--skip-all-under-chapter-three` 完全跳過處理；從 `參、` 標題本身開始，到下一個同級第一層標題前，段落、表格、底色、縮排、大綱、Word COM 補救與字元縮排清理都不會套用。

## 段落 / 大綱 / 縮排處理規則

- 本文段落依既有規則辨識手動編號、自動編號與樣式編號。
- 會根據編號層級套用對應的縮排與 Word `outlineLvl`。
- 本文內一般段落會依最近一個標題層級決定內文縮排。
- 序言區段可依選項決定是否套用縮排與大綱層級。
- 目錄（TOC）相關段落會跳過，避免誤改 Word 自動產生的目錄內容。
- numbering.xml 的 `suffix=nothing` 等既有處理邏輯維持不變。

## `indent_defaults.json` 位置與用途

- 檔名固定為 `indent_defaults.json`。
- 一般以原始碼執行時，預設位置在目前工作目錄。
- 打包成 EXE 後，預設位置在 EXE 同層目錄。
- 這個檔案用來覆寫內建的縮排預設值，供 GUI 與 CLI 共用。
- 若檔案存在，程式會優先載入其中設定；若不存在，則使用程式內建預設。

## `process_log` 與 `table_log` 輸出說明

- `process_log`：記錄段落、大綱、縮排、Word COM 驗證與其他文字處理過程。
- `table_log`：記錄每個表格的所在 XML、表格索引、欄數、處理動作、略過原因與底色處理結果。
- 這些 log 可用來追查某一段落或某一張表格為何被修改、略過或改用補救流程。

## 注意事項

- 不要直接覆寫原始 `.docx` 檔案，輸入與輸出路徑應分開。
- 建議在批次處理前先備份原始文件。
- 若使用 EXE 版本，可在 EXE 同層保留 `indent_defaults.json`，方便攜帶自訂縮排設定。
- 對於特殊格式文件，建議先用少量樣本測試，再批次處理正式檔案。
