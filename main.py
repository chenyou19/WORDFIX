from __future__ import annotations

import sys

import tkinter as tk

from docx_fixer.cli import parse_args, run_cli
from docx_fixer.gui_app import DocxFixerApp


def main() -> int:
    args = parse_args(sys.argv[1:])

    # 有 input/output 參數時走 CLI，否則啟動 GUI。
    if args.input_docx and args.output_docx:
        return run_cli(args)

    root = tk.Tk()
    # Windows 下設定 Tk 預設中文字型，避免介面字體顯示異常。
    try:
        root.option_add("*Font", ("Microsoft JhengHei UI", 10))
    except Exception:
        pass

    DocxFixerApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
