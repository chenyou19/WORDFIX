# Word DOCX 快速整理工具

這個工具可整理 Word `.docx` 文件中的表格格式、儲存格底色、文件編號/大綱階層，以及 `numbering.xml` 自動編號縮排與編號後方留白。

可另外勾選或指定「去除所有大綱階層」，先移除文件中既有的 `w:outlineLvl`，再依所選功能執行段落與表格處理。這個全文件清除功能預設不啟用，避免未指定時意外移除既有大綱階層。

段落處理會保留並修正符合文件編號規則的段落，包含段落本身的自動編號與段落樣式中定義的自動編號；長篇說明文字或 `※` 註記開頭段落不會只因套用帶編號的樣式而加入大綱階層。不屬於文件編號、但已帶有大綱階層的段落，會改成 Word 的「本文」階層。項目符號也會視為本文，不套用文件編號的大綱階層。

段落大綱階層處理會以文件中第一次出現的 `壹、序言` 為界線：從 `壹、序言` 該段開始，只要啟用「調整段落」，就依正式文件規則同時調整縮排並加入 Word 大綱階層。`壹、序言` 前的段落預設不處理；只有另外啟用「縮排壹、序言前」或「壹、序言前加入大綱階層」時，才會對第一次 `壹、序言` 前的編號段落套用對應處理。

正式段落內建縮排預設為：`壹、` 文字起點 1.11 cm、編號起點 -0.04 cm；`一、` 文字起點 1.8 cm、編號起點 0.69 cm；`（一）` 文字起點 2.32 cm、編號起點 1.32 cm；`1.` 文字起點 3.79 cm、編號起點 3.05 cm；`（1）` 文字起點 4.76 cm、編號起點 3.53 cm；`A.` 文字起點 5.27 cm、編號起點 4.52 cm；`（A）` 文字起點 6.26 cm、編號起點 5.02 cm；`a.` 文字起點 6.96 cm、編號起點 6.2 cm；`（a）` 文字起點 8.96 cm、編號起點 7.72 cm。

`壹、序言` 前有專用縮排設定：`一、` 文字起點 1.11 cm、懸掛 1.15 cm；`（一）` 文字起點 1.54 cm、懸掛 0.85 cm；`1.` 文字起點 3.01 cm、懸掛 1.01 cm；`（1）` 文字起點 4.02 cm、懸掛 1.72 cm。此設定只有啟用「縮排壹、序言前」時才會套用；「壹、序言前加入大綱階層」可單獨啟用，單獨啟用時只設定 Word 大綱階層，不修改縮排。

GUI 第二頁「縮排預設」可分別調整 `壹、序言` 前後各階的文字起點與編號起點。若沒有本機 `indent_defaults.json`，GUI 會顯示上述內建預設；若工作目錄或執行檔旁已有 `indent_defaults.json`，開啟 GUI 或使用 CLI 時會自動載入該檔並覆蓋內建值。按「套用目前設定」會立即套用到本次處理；按「保存成預設樣式」會寫入 `indent_defaults.json`，保存功能維持不變。

標題下方的非編號內文段落會自動把左縮排對齊上一個標題的文字起點；遇到下一個標題時，內文縮排基準會更新成新的標題文字起點。

文件編號偵測需符合指定標點格式：`壹、`、`一、`、`（一）`/`(一)`、`1.`、`（1）`/`(1)`、`A.`、`（A）`/`(A)`、`a.`、`（a）`/`(a)`。未帶標點的 `1`、`一` 不會被視為文件編號。

每次處理完成會在輸出檔旁產生 `輸出檔名_log.txt`，記錄實際量測到的手動編號格式、文字起點、編號起點、編號大小、字型大小，以及每個被修改的大綱階層段落、修改原因、修改前後階層與段落文字。

## 執行 GUI

```powershell
python main.py
```

## 執行 CLI

```powershell
python main.py input.docx output.docx --table --color --paragraph
```

可用參數：

- `--table`：整理表格版面。
- `--color`：整理表格儲存格底色。
- `--paragraph`：整理文件編號段落與大綱階層，並將非文件編號與項目符號的大綱階層改成本文。
- `--remove-all-outline`：去除文件中所有既有 Word 大綱階層。
- `--indent-preface`：啟用第一次 `壹、序言` 前的編號段落縮排，預設關閉。
- `--outline-preface`：啟用第一次 `壹、序言` 前的編號段落 Word 大綱階層，預設關閉。
- `--paragraph-in-tables`：段落大綱階層處理包含表格內段落。
- `--quiet`：命令列模式不輸出進度。

若 CLI 沒有指定 `--table`、`--color`、`--paragraph`、`--remove-all-outline`、`--indent-preface`、`--outline-preface` 任一項，會維持 GUI 既有預設行為：表格、顏色與正式段落處理啟用，但不會預設啟用「去除所有大綱階層」、「縮排壹、序言前」或「壹、序言前加入大綱階層」。

## 模組用途

- `main.py`：程式入口，依是否提供 input/output 決定走 CLI 或 GUI。
- `docx_fixer/constants.py`：預設後綴、灰色值、底色規則、Word XML namespace、編號文字集合、範本縮排設定。
- `docx_fixer/models.py`：處理選項與處理結果統計資料模型。
- `docx_fixer/exceptions.py`：停止處理例外。
- `docx_fixer/stop_controller.py`：停止處理控制器。
- `docx_fixer/xml_utils.py`：Word XML 共用工具。
- `docx_fixer/path_utils.py`：路徑正規化與同檔判斷。
- `docx_fixer/shading.py`：儲存格底色判斷與修改。
- `docx_fixer/numbering.py`：`numbering.xml` 自動編號階層判斷與格式修正。
- `docx_fixer/indent_settings.py`：GUI/CLI 共用的縮排預設讀寫、套用與公分/twips 轉換。
- `docx_fixer/outline.py`：文件編號偵測、非文件編號與項目符號改本文、TOC 排除、段落大綱階層與縮排套用。
- `docx_fixer/table_format.py`：表格版面與表格底色處理。
- `docx_fixer/docx_processor.py`：DOCX zip/XML 主處理流程。
- `docx_fixer/gui_app.py`：Tkinter GUI。
- `docx_fixer/cli.py`：命令列參數與 CLI 執行流程。
