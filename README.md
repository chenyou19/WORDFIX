# Word DOCX 快速整理工具

這個工具可整理 Word `.docx` 文件中的表格格式、儲存格底色、文件編號/大綱階層，以及 `numbering.xml` 自動編號縮排與編號後方留白。

段落處理會保留並修正符合文件編號規則的段落，包含段落本身的自動編號與段落樣式中定義的自動編號；長篇說明文字或 `※` 註記開頭段落不會只因套用帶編號的樣式而加入大綱階層。不屬於文件編號、但已帶有大綱階層的段落，會改成 Word 的「本文」階層。項目符號也會視為本文，不套用文件編號的大綱階層。

段落大綱階層處理會以文件中第一次出現的 `壹、序言` 為界線：在此之前的段落只調整文件編號縮排，不加入 Word 大綱階層，且前置段落不使用 `壹、` 階層，所以 `一、` 會視為縮排階層 0、`（一）` 視為縮排階層 1，依序往後；從 `壹、序言` 該段開始，才同時調整縮排並加入 Word 大綱階層。

`壹、序言` 前有專用縮排設定：`一、` 文字起點 1.11 cm、懸掛 1.15 cm；`（一）` 文字起點 1.54 cm、懸掛 0.85 cm；`1.` 文字起點 3.01 cm、懸掛 1.01 cm；`（1）` 文字起點 4.02 cm、懸掛 1.72 cm。

文件編號偵測需符合指定標點格式：`壹、`、`一、`、`（一）`/`(一)`、`1.`、`（1）`/`(1)`、`A.`、`（A）`/`(A)`、`a.`、`（a）`/`(a)`。未帶標點的 `1`、`一` 不會被視為文件編號。

每次處理完成會在輸出檔旁產生 `輸出檔名_log.txt`，記錄每個被修改的大綱階層段落、修改原因、修改前後階層與段落文字。

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
- `--remove-preface-outline`：移除第一次 `壹、序言` 前既有的 Word 大綱階層；搭配 `--paragraph` 時，前置編號段落只套用縮排，不加入大綱階層。
- `--paragraph-in-tables`：段落大綱階層處理包含表格內段落。
- `--quiet`：命令列模式不輸出進度。

若 CLI 沒有指定 `--table`、`--color`、`--paragraph`、`--remove-preface-outline` 任一項，會維持 GUI 預設行為：四種方案全部啟用。

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
- `docx_fixer/outline.py`：文件編號偵測、非文件編號與項目符號改本文、TOC 排除、段落大綱階層與縮排套用。
- `docx_fixer/table_format.py`：表格版面與表格底色處理。
- `docx_fixer/docx_processor.py`：DOCX zip/XML 主處理流程。
- `docx_fixer/gui_app.py`：Tkinter GUI。
- `docx_fixer/cli.py`：命令列參數與 CLI 執行流程。
