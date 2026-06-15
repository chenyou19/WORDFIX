# EXE 與設定攜帶

## 執行 GUI

原始碼執行 GUI：

```powershell
python main.py
```

若同時提供輸入與輸出檔，`main.py` 會改走 CLI：

```powershell
python main.py input.docx output.docx --table --color --paragraph
```

## 封裝 EXE

專案根目錄已有 `DocxFixer.spec`，入口是 `main.py`，EXE 名稱是 `DocxFixer`，且 `console=False`。

常見封裝指令：

```powershell
pyinstaller DocxFixer.spec
```

或在已有 PyInstaller 的環境中依專案流程執行同等指令。封裝完成後，輸出通常會在 `dist` 目錄。

## 隱藏 PowerShell 視窗

`console=False` 只隱藏主 EXE 的 console，不會自動隱藏 Word COM / 表格 fallback 流程啟動的 `powershell.exe` 子程序。隱藏子程序視窗的邏輯集中在 `docx_fixer/process_runner.py` 的 `_run_powershell_process()`：

- Windows（`os.name == "nt"`）下對 `subprocess.Popen` 加上 `creationflags=CREATE_NO_WINDOW`，並設定 `STARTUPINFO`（`STARTF_USESHOWWINDOW` + `SW_HIDE`），避免黑色視窗閃出。
- 指令另外帶 `-WindowStyle Hidden` 作為輔助，但主要靠 `Popen` 參數，避免視窗閃一下。
- 非 Windows 環境不套用上述參數，`Popen` 行為與原本相同，Linux/macOS 測試不受影響。
- stdout/stderr 擷取、timeout、stop/cancel、`CODEX_STOP_PATH`、`PYTHONUTF8`、`PYTHONIOENCODING` 與所有 `WORD_COM_*` log 均維持原樣。

## 設定檔位置

EXE 執行時，`indent_defaults.json` 要放在 EXE 同層。程式不會把執行後保存的設定寫回 EXE 內部，而是讀寫外部 JSON。

不能直接把設定寫進 EXE 內部的原因：

- EXE 是封裝後的執行檔，不適合在使用者每次改設定時重寫。
- 不同電腦、不同使用者需要自己的設定。
- 外部 JSON 比較容易備份、複製與手動檢查。

## 從 A 電腦帶到 B 電腦

要沿用設定，請一起帶：

- EXE
- `indent_defaults.json`
- 必要時的測試 docx

到 B 電腦後，把 EXE 與 `indent_defaults.json` 放在同一個資料夾，再執行 EXE 即可讀到原本設定。
