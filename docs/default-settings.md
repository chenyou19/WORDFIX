# 預設設定檔

WORDFIX 的共用預設設定檔名稱是：

```text
indent_defaults.json
```

相關程式在 `docx_fixer/indent_settings.py` 與 `docx_fixer/gui_defaults.py`。

## 檔案位置

- 原始碼執行時：放在目前工作目錄，也就是啟動 `python main.py` 時所在的資料夾。
- EXE 執行時：放在 EXE 同層目錄。

位置由 `get_indent_settings_path()` 判斷：若 `sys.frozen` 為真，使用 `sys.executable` 同層；否則使用 `Path.cwd()`。

## 主要區塊

`indent_defaults.json` 內含兩個主要區塊：

```json
{
  "indent_settings": {},
  "gui_defaults": {}
}
```

- `indent_settings`：儲存縮排設定，包含 `body` 與 `preface`。
- `gui_defaults`：儲存 GUI 第一頁勾選狀態。

## 載入規則

- 若檔案不存在，縮排使用 `constants.py` 的內建值，GUI 勾選使用 `built_in_gui_defaults()`。
- 若只存在其中一個區塊，另一個區塊仍使用 code 內建預設。
- 舊格式若直接包含 `body`、`preface`，會被視為舊版縮排設定並轉成 `indent_settings`。
- `gui_defaults` 只會接受既有鍵；缺少的鍵會補成 `built_in_gui_defaults()` 的值。

## GUI 按鈕對應

- 縮排設定頁「保存成預設樣式」：寫入 `indent_settings`。
- 縮排設定頁「還原內建預設」：套回 code 內建縮排。
- 第一頁「保存目前勾選為預設方案」：寫入 `gui_defaults`。
- 第一頁「還原內建勾選預設」：套回 `built_in_gui_defaults()`。

## EXE 設定攜帶

如果要把 EXE 帶到另一台電腦沿用設定，請把下列檔案放在同一層：

- EXE
- `indent_defaults.json`

只把 EXE 帶走不會自動帶走自訂設定。設定檔是外部 JSON，程式啟動時才讀取。
