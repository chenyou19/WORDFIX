# 開發者維護筆記

這份文件列出主要檔案責任，方便維護時快速定位。

## 入口與介面

- `main.py`：程式入口。若提供輸入與輸出檔走 CLI，否則啟動 Tkinter GUI。
- `docx_fixer/cli.py`：命令列參數、相容舊參數映射、CLI 預設處理動作、CLI log 寫出。
- `docx_fixer/gui_app.py`：Tkinter GUI、檔案選擇、第一頁處理選項、縮排設定頁、背景處理執行與 GUI log 寫出。
- `docx_fixer/gui_defaults.py`：GUI 第一頁勾選狀態的內建預設、載入、正規化與保存。

## 設定與常數

- `docx_fixer/indent_settings.py`：`indent_defaults.json` 路徑、縮排設定載入與保存、cm/twips 轉換、舊格式相容。
- `docx_fixer/constants.py`：預設灰色、XML namespace、本文與序言內建縮排表、twips 生成工具。
- `docx_fixer/models.py`：`ProcessOptions` 與 `ProcessSummary`。舊的參章合併選項也在 `ProcessOptions.__post_init__()` 做相容。

## 核心處理流程

- `docx_fixer/docx_processor.py`：docx zip 讀寫、可處理 XML part 判斷、TOC/參章保護 context 建立、段落與表格流程串接、Word COM 補救流程、最後 numbering suffix cleanup。
- `docx_fixer/protected_region.py`：TOC 排除、參章偵測、參章段落集合、參章表格判斷、第一階層標題查找。
- `docx_fixer/outline.py`：手動編號偵測、自動/樣式編號層級判斷、本文/序言縮排與 outline level、TOC 跳過、普通內文縮排、Word COM 字號待確認紀錄。
- `docx_fixer/numbering.py`：numbering.xml 層級辨識、編號縮排套用、suffix/tab 清理、styles outline 格式套用。
- `docx_fixer/numbering_cleanup.py`：處理完後再次清理 docx 中 numbering suffix 與 tab。
- `docx_fixer/indent_sanitizer.py`：清除 styles/numbering 中的字元縮排屬性，並避開 TOC 或參章保護定義。

## 表格與顏色

- `docx_fixer/table_pipeline.py`：逐張表格處理順序、第一張表格跳過、巢狀表格版面保護與 color-only 分支、參章表格保護、特殊/一般表格判斷、table log record。
- `docx_fixer/table_format.py`：一般表格格式、特殊表格格式、右對齊 autofit 格式、表格內文字與列高格式、底色處理呼叫。
- `docx_fixer/table_word_com.py`：一般表格的 Word COM AutoFit 流程，順序是 `AutoFitContent` 再 `AutoFitWindow`。
- `docx_fixer/shading.py`：底色判斷，包含預設灰、無色、灰色深淺、theme 色彩與非灰色清除。

## Word COM 與外部流程

- `docx_fixer/word_com_indent.py`：Word COM 內文縮排補救、字號確認、PowerShell fallback、核准後寫回 docx XML。
- `docx_fixer/process_runner.py`：執行 PowerShell script/file，支援停止控制與 timeout。
- `docx_fixer/stop_controller.py`：GUI 停止處理用的控制物件。
- `docx_fixer/table_cross_page.py`：跨頁表格統計與外部調整流程的彙整函式，目前主流程未直接呼叫。

## Log 與工具

- `docx_fixer/process_log.py`：`process_log`、`table_log`、`heading_suffix_log` 格式化與寫檔。
- `docx_fixer/style_resolver.py`：解析 styles.xml 中的字號，供段落字號判斷使用。
- `docx_fixer/xml_utils.py`：XML QName、元素建立、段落文字讀取、字元縮排屬性清除。
- `docx_fixer/path_utils.py`：比較輸入與輸出檔是否為同一路徑。
- `docx_fixer/exceptions.py`：處理停止例外。
