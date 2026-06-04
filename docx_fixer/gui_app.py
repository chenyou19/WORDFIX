from __future__ import annotations

import datetime
import os
import queue
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .constants import DEFAULT_GRAY, DEFAULT_SUFFIX
from .docx_processor import fix_docx_fast
from .exceptions import ProcessStopped
from .indent_settings import (
    apply_indent_settings,
    built_in_indent_settings,
    current_indent_settings,
    format_cm,
    get_indent_settings_path,
    load_saved_indent_settings,
    save_indent_settings,
)
from .models import ProcessOptions
from .path_utils import is_same_file_path
from .process_log import write_process_log
from .stop_controller import StopController

DEFAULT_WINDOW_GEOMETRY = "1080x760"
MIN_WINDOW_SIZE = (980, 680)

class DocxFixerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Word DOCX 快速整理工具")
        self.root.geometry(DEFAULT_WINDOW_GEOMETRY)
        self.root.minsize(*MIN_WINDOW_SIZE)

        self.stop_controller = StopController()
        self.worker_thread: threading.Thread | None = None
        self.ui_queue: queue.Queue = queue.Queue()
        self.current_temp_output: Path | None = None

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.indent_vars: dict[str, list[tuple[int, str, tk.StringVar, tk.StringVar, tk.StringVar]]] = {
            "preface": [],
            "body": [],
        }
        self.indent_settings_load_error: str | None = None

        try:
            load_saved_indent_settings()
        except Exception as exc:
            self.indent_settings_load_error = str(exc)

        self.fix_table_var = tk.BooleanVar(value=True)
        self.fix_color_var = tk.BooleanVar(value=True)
        self.fix_paragraph_var = tk.BooleanVar(value=True)
        self.remove_all_outline_var = tk.BooleanVar(value=False)
        self.indent_preface_var = tk.BooleanVar(value=False)
        self.outline_preface_var = tk.BooleanVar(value=False)
        self.paragraph_in_tables_var = tk.BooleanVar(value=False)

        self.status_var = tk.StringVar(value="請選擇 .docx 檔案。")
        self.progress_var = tk.DoubleVar(value=0)

        self._build_ui()
        if self.indent_settings_load_error:
            messagebox.showwarning(
                "縮排預設載入失敗",
                f"已改用內建縮排設定。\n\n{self.indent_settings_load_error}",
            )
        self._poll_queue()
        self.root.protocol("WM_DELETE_WINDOW", self.exit_app)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)

        title = ttk.Label(outer, text="Word DOCX 快速整理工具", font=("Microsoft JhengHei UI", 16, "bold"))
        title.pack(anchor="w", pady=(0, 10))

        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True)

        process_tab = ttk.Frame(notebook, padding=10)
        indent_tab = ttk.Frame(notebook, padding=10)
        notebook.add(process_tab, text="修改文件")
        notebook.add(indent_tab, text="縮排預設")

        # 檔案選擇
        file_frame = ttk.LabelFrame(process_tab, text="檔案")
        file_frame.pack(fill="x", pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="來源檔案：").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(file_frame, textvariable=self.input_var).grid(row=0, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(file_frame, text="瀏覽...", command=self.browse_input).grid(row=0, column=2, padx=8, pady=8)

        ttk.Label(file_frame, text="輸出檔案：").grid(row=1, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(file_frame, textvariable=self.output_var).grid(row=1, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(file_frame, text="另存為...", command=self.browse_output).grid(row=1, column=2, padx=8, pady=8)

        hint = ttk.Label(file_frame, text=f"輸出檔案可留空；留空時會自動輸出為：原檔名{DEFAULT_SUFFIX}.docx")
        hint.grid(row=2, column=1, padx=8, pady=(0, 8), sticky="w")

        # 功能選項
        option_frame = ttk.LabelFrame(process_tab, text="可勾選的處理方案")
        option_frame.pack(fill="x", pady=(0, 10))

        ttk.Checkbutton(
            option_frame,
            text="調整表格格式：置中、列高、字體 11、單行間距、外框雙黑線、寬度 100%",
            variable=self.fix_table_var,
        ).pack(anchor="w", padx=8, pady=(8, 4))

        ttk.Checkbutton(
            option_frame,
            text="調整顏色：BFBFBF／A6A6A6／808080 改 D9D9D9；F2F2F2 保持；其他顏色改無色彩",
            variable=self.fix_color_var,
        ).pack(anchor="w", padx=8, pady=4)

        ttk.Checkbutton(
            option_frame,
            text="去除所有大綱階層",
            variable=self.remove_all_outline_var,
        ).pack(anchor="w", padx=8, pady=4)

        ttk.Checkbutton(
            option_frame,
            text="調整段落：依範本階層處理壹／一／（一）/(一)／1／（1）/(1)／A／（A）/(A)／a／（a）/(a) 的縮排，並加上 Word 大綱階層",
            variable=self.fix_paragraph_var,
        ).pack(anchor="w", padx=8, pady=4)

        ttk.Checkbutton(
            option_frame,
            text="縮排壹、序言前",
            variable=self.indent_preface_var,
        ).pack(anchor="w", padx=8, pady=4)

        ttk.Checkbutton(
            option_frame,
            text="壹、序言前加入大綱階層",
            variable=self.outline_preface_var,
        ).pack(anchor="w", padx=8, pady=4)

        ttk.Checkbutton(
            option_frame,
            text="段落縮排也處理表格內文字",
            variable=self.paragraph_in_tables_var,
        ).pack(anchor="w", padx=28, pady=(0, 8))

        # 進度與狀態
        progress_frame = ttk.LabelFrame(process_tab, text="進度")
        progress_frame.pack(fill="x", pady=(0, 10))

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
        )
        self.progress_bar.pack(fill="x", padx=8, pady=(8, 4))

        ttk.Label(progress_frame, textvariable=self.status_var).pack(anchor="w", padx=8, pady=(0, 8))

        # 按鈕
        button_frame = ttk.Frame(process_tab)
        button_frame.pack(fill="x", pady=(0, 10))

        self.start_button = ttk.Button(button_frame, text="開始修改", command=self.start_process)
        self.start_button.pack(side="left", padx=(0, 8))

        self.stop_button = ttk.Button(button_frame, text="停止修改", command=self.stop_process, state="disabled")
        self.stop_button.pack(side="left", padx=8)

        ttk.Button(button_frame, text="終止程式", command=self.exit_app).pack(side="right")

        # 訊息紀錄
        log_frame = ttk.LabelFrame(process_tab, text="訊息")
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_frame, height=8, wrap="word")
        self.log_text.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)

        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y", padx=(0, 8), pady=8)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self._build_indent_settings_tab(indent_tab)

    def _build_indent_settings_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.columnconfigure(0, weight=1)
        scroll_frame.columnconfigure(1, weight=1)

        window_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        def update_scroll_region(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def update_window_width(event) -> None:
            canvas.itemconfigure(window_id, width=event.width)

        scroll_frame.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", update_window_width)

        settings = current_indent_settings()
        self.indent_vars = {"preface": [], "body": []}

        self._build_indent_section(
            scroll_frame,
            title="壹、序言前",
            section="preface",
            rows=settings["preface"],
            row=0,
            column=0,
        )
        self._build_indent_section(
            scroll_frame,
            title="壹、序言後",
            section="body",
            rows=settings["body"],
            row=0,
            column=1,
        )

        button_frame = ttk.Frame(parent)
        button_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        ttk.Button(button_frame, text="套用目前設定", command=self.apply_indent_entries).pack(side="left")
        ttk.Button(button_frame, text="保存成預設樣式", command=self.save_indent_defaults).pack(side="left", padx=8)
        ttk.Button(button_frame, text="還原新版內建預設", command=self.restore_builtin_indent_defaults).pack(side="left", padx=8)

        path_label = ttk.Label(
            button_frame,
            text=f"預設檔：{get_indent_settings_path()}",
        )
        path_label.pack(side="right")

    def _build_indent_section(
        self,
        parent: ttk.Frame,
        title: str,
        section: str,
        rows: list[dict[str, float | int | str]],
        row: int,
        column: int,
    ) -> None:
        frame = ttk.LabelFrame(parent, text=title)
        frame.grid(row=row, column=column, sticky="nsew", padx=6, pady=6)
        parent.columnconfigure(column, weight=1)

        headers = ["階層", "編號格式", "標號起點 cm", "凸排距離 cm", "內文起點 cm"]
        for col, header in enumerate(headers):
            ttk.Label(frame, text=header).grid(row=0, column=col, padx=4, pady=(6, 4), sticky="w")

        for index, item in enumerate(rows, start=1):
            level = int(item["level"])
            label = str(item["label"])
            number_start_var = tk.StringVar(value=format_cm(float(item["number_start_cm"])))
            hanging_var = tk.StringVar(value=format_cm(float(item["hanging_cm"])))
            body_left_var = tk.StringVar(value=format_cm(float(item["body_left_cm"])))

            ttk.Label(frame, text=str(index)).grid(row=index, column=0, padx=4, pady=3, sticky="w")
            ttk.Label(frame, text=label).grid(row=index, column=1, padx=4, pady=3, sticky="w")
            ttk.Entry(frame, textvariable=number_start_var, width=9).grid(row=index, column=2, padx=4, pady=3)
            ttk.Entry(frame, textvariable=hanging_var, width=9).grid(row=index, column=3, padx=4, pady=3)
            ttk.Entry(frame, textvariable=body_left_var, width=9).grid(row=index, column=4, padx=4, pady=3)
            self.indent_vars[section].append((level, label, number_start_var, hanging_var, body_left_var))

    def collect_indent_entries(self) -> dict[str, list[dict[str, float | int | str]]]:
        settings: dict[str, list[dict[str, float | int | str]]] = {"preface": [], "body": []}
        section_names = {
            "preface": "壹、序言前",
            "body": "壹、序言後",
        }

        for section, rows in self.indent_vars.items():
            for display_index, (level, label, number_start_var, hanging_var, body_left_var) in enumerate(rows, start=1):
                try:
                    number_start_cm = float(number_start_var.get().strip())
                    hanging_cm = float(hanging_var.get().strip())
                    body_left_cm = float(body_left_var.get().strip())
                except ValueError as exc:
                    raise ValueError(
                        f"{section_names[section]}第 {display_index} 階請輸入數字。"
                    ) from exc

                settings[section].append({
                    "level": level,
                    "label": label,
                    "number_start_cm": number_start_cm,
                    "hanging_cm": hanging_cm,
                    "body_left_cm": body_left_cm,
                })

        return settings

    def apply_indent_entries(self) -> bool:
        try:
            apply_indent_settings(self.collect_indent_entries())
        except Exception as exc:
            messagebox.showerror("縮排設定錯誤", str(exc))
            return False

        self.status_var.set("已套用目前縮排設定。")
        return True

    def save_indent_defaults(self) -> None:
        try:
            path = save_indent_settings(self.collect_indent_entries())
        except Exception as exc:
            messagebox.showerror("縮排設定錯誤", str(exc))
            return

        self.status_var.set("已保存縮排預設樣式。")
        messagebox.showinfo("已保存", f"已保存成預設樣式：\n{path}")

    def restore_builtin_indent_defaults(self) -> None:
        settings = built_in_indent_settings()
        for section, rows in settings.items():
            by_level = {int(row["level"]): row for row in rows}
            for level, _label, number_start_var, hanging_var, body_left_var in self.indent_vars[section]:
                row = by_level[level]
                number_start_var.set(format_cm(float(row["number_start_cm"])))
                hanging_var.set(format_cm(float(row["hanging_cm"])))
                body_left_var.set(format_cm(float(row["body_left_cm"])))

        apply_indent_settings(settings)
        self.status_var.set("已還原新版內建縮排預設；按「保存成預設樣式」可覆蓋本機 indent_defaults.json。")

    def browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="選擇 Word DOCX 檔案",
            filetypes=[("Word 文件", "*.docx"), ("所有檔案", "*.*")],
        )
        if path:
            self.input_var.set(path)
            self.status_var.set("已選擇來源檔案。")

    def browse_output(self) -> None:
        input_path = self.input_var.get().strip()
        initialdir = None
        initialfile = "output.docx"

        if input_path:
            p = Path(input_path)
            initialdir = str(p.parent)
            initialfile = f"{p.stem}{DEFAULT_SUFFIX}.docx"

        path = filedialog.asksaveasfilename(
            title="選擇輸出檔案",
            defaultextension=".docx",
            initialdir=initialdir,
            initialfile=initialfile,
            filetypes=[("Word 文件", "*.docx")],
        )
        if path:
            input_text = self.input_var.get().strip()
            if input_text and is_same_file_path(input_text, path):
                messagebox.showerror(
                    "輸出檔案錯誤",
                    "修改後的檔案不可以跟原始檔案相同。\n\n"
                    "請改用不同檔名，例如：TT_已修改.docx。",
                )
                return
            self.output_var.set(path)

    def append_log(self, text: str) -> None:
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")

    def set_running_state(self, running: bool) -> None:
        self.start_button.configure(state="disabled" if running else "normal")
        self.stop_button.configure(state="normal" if running else "disabled")

    def start_progress_animation(self) -> None:
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start(10)

    def stop_progress_animation(self, final_value: float | None = None) -> None:
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        if final_value is not None:
            self.progress_var.set(final_value)

    def resolve_output_path(self, input_path: Path) -> Path:
        output_text = self.output_var.get().strip()

        if not output_text:
            return input_path.with_name(f"{input_path.stem}{DEFAULT_SUFFIX}{input_path.suffix}")

        output_path = Path(output_text)

        if output_path.exists() and output_path.is_dir():
            return output_path / f"{input_path.stem}{DEFAULT_SUFFIX}{input_path.suffix}"

        if not output_path.suffix:
            output_path = output_path.with_suffix(".docx")

        if not output_path.is_absolute():
            output_path = input_path.parent / output_path

        return output_path

    def validate_before_start(self):
        input_text = self.input_var.get().strip()
        if not input_text:
            messagebox.showwarning("缺少來源檔案", "請先選擇要修改的 .docx 檔案。")
            return None

        input_path = Path(input_text)
        if not input_path.exists():
            messagebox.showerror("找不到檔案", f"找不到來源檔案：\n{input_path}")
            return None

        if input_path.suffix.lower() != ".docx":
            messagebox.showerror("格式不支援", "目前只支援 .docx，不支援 .doc。")
            return None

        if not self.apply_indent_entries():
            return None

        options = ProcessOptions(
            fix_table_layout=self.fix_table_var.get(),
            fix_color=self.fix_color_var.get(),
            fix_paragraph=self.fix_paragraph_var.get(),
            include_tables_in_paragraph=self.paragraph_in_tables_var.get(),
            remove_all_outline_levels=self.remove_all_outline_var.get(),
            indent_preface_paragraphs=self.indent_preface_var.get(),
            outline_preface_paragraphs=self.outline_preface_var.get(),
        )

        if not (
            options.fix_table_layout
            or options.fix_color
            or options.fix_paragraph
            or options.remove_all_outline_levels
            or options.indent_preface_paragraphs
            or options.outline_preface_paragraphs
        ):
            messagebox.showwarning("尚未選擇方案", "請至少勾選一種修改方案。")
            return None

        output_path = self.resolve_output_path(input_path)

        if is_same_file_path(input_path, output_path):
            messagebox.showerror(
                "輸出檔案錯誤",
                "修改後的檔案不可以跟原始檔案相同。\n\n"
                "請重新選擇輸出檔名，例如：TT_已修改.docx。",
            )
            return None

        if output_path.exists():
            ok = messagebox.askyesno(
                "檔案已存在",
                f"輸出檔案已存在，是否覆蓋？\n\n{output_path}",
            )
            if not ok:
                return None

        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("無法建立輸出資料夾", str(exc))
            return None

        return input_path, output_path, options

    def start_process(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return

        validated = self.validate_before_start()
        if validated is None:
            return

        input_path, output_path, options = validated

        temp_output = output_path.with_name(f"{output_path.stem}.__tmp__{output_path.suffix}")
        self.current_temp_output = temp_output

        self.stop_controller.clear()
        self.progress_var.set(0)
        self.status_var.set("開始處理...")
        self.start_progress_animation()
        self.append_log(f"來源：{input_path}")
        self.append_log(f"輸出：{output_path}")
        self.append_log("開始修改。")
        self.set_running_state(True)

        self.worker_thread = threading.Thread(
            target=self._worker,
            args=(input_path, output_path, temp_output, options),
            daemon=True,
        )
        self.worker_thread.start()

    def _worker(self, input_path: Path, output_path: Path, temp_output: Path, options: ProcessOptions) -> None:
        def progress_callback(percent: float, message: str) -> None:
            self.ui_queue.put(("progress", max(0, min(100, percent)), message))

        try:
            if temp_output.exists():
                temp_output.unlink()

            summary = fix_docx_fast(
                input_docx=input_path,
                output_docx=temp_output,
                options=options,
                stop=self.stop_controller,
                progress_callback=progress_callback,
            )

            self.stop_controller.check()

            if is_same_file_path(input_path, output_path):
                raise ValueError("修改後的檔案不可以跟原始檔案相同，避免覆蓋原檔。")

            try:
                os.replace(temp_output, output_path)
                final_output = output_path
            except PermissionError:
                # Windows 會在目標檔案被 Word、檔案總管預覽窗格、OneDrive 等占用時拒絕覆蓋。
                # 為避免已處理好的暫存檔被刪掉，改存成帶時間戳的新檔名。
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                fallback_output = output_path.with_name(
                    f"{output_path.stem}_{timestamp}{output_path.suffix}"
                )
                os.replace(temp_output, fallback_output)
                final_output = fallback_output
                self.ui_queue.put((
                    "warning",
                    f"原輸出檔可能正在被 Word、檔案總管預覽窗格或同步程式占用，已改存為：{fallback_output}",
                ))

            log_path = None
            try:
                log_path = write_process_log(final_output, summary)
            except Exception as exc:
                self.ui_queue.put(("warning", f"處理紀錄檔寫入失敗：{exc}"))

            self.ui_queue.put(("done", final_output, summary, log_path))

        except ProcessStopped:
            try:
                if temp_output.exists():
                    temp_output.unlink()
            except Exception:
                pass
            self.current_temp_output = None
            self.ui_queue.put(("stopped",))

        except Exception as exc:
            try:
                if temp_output.exists():
                    temp_output.unlink()
            except Exception:
                pass
            self.current_temp_output = None
            self.ui_queue.put(("error", str(exc), traceback.format_exc()))

    def stop_process(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            self.stop_controller.stop()
            self.stop_button.configure(state="disabled")
            self.status_var.set("正在中止所有程序並釋放檔案鎖定...")
            self.append_log("已送出終止要求，正在關閉 Word/PowerShell 程序並清理暫存檔。")

    def exit_app(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            self.stop_controller.stop()
        self.root.destroy()

    def _poll_queue(self) -> None:
        try:
            while True:
                item = self.ui_queue.get_nowait()
                kind = item[0]

                if kind == "progress":
                    _, percent, message = item
                    self.status_var.set(message)

                elif kind == "warning":
                    _, message = item
                    self.status_var.set(message)
                    self.append_log("提醒：" + message)

                elif kind == "done":
                    _, output_path, summary, log_path = item
                    self.current_temp_output = None
                    self.stop_progress_animation(100)
                    self.status_var.set("完成！")
                    self.append_log("完成！")
                    self.append_log(f"已處理表格數：{summary.tables}")
                    self.append_log(f"跳過第 1 頁的表格數：{summary.skipped_first_page_tables}")
                    self.append_log(f"因格子數小於等於 4 而跳過的表格數：{summary.skipped_small_tables}")
                    self.append_log(f"偵測到跨頁的表格數：{summary.cross_page_tables}")
                    self.append_log(f"成功調整後不再跨頁的表格數：{summary.cross_page_resolved_tables}")
                    self.append_log(f"調整後仍跨頁的表格數：{summary.cross_page_still_split_tables}")
                    self.append_log(f"儲存格內距被調整的表格數：{summary.adjusted_cell_padding_tables}")
                    self.append_log(f"行距或段距被調整的表格數：{summary.adjusted_table_spacing_tables}")
                    self.append_log(f"列高改為自動高度的表格數：{summary.auto_height_tables}")
                    self.append_log(f"移到下一頁後成功不跨頁的表格數：{summary.moved_next_page_resolved_tables}")
                    self.append_log(f"不縮小字體下無法完全避免跨頁的表格數：{summary.cannot_avoid_cross_page_tables}")
                    self.append_log(f"處理失敗但已略過的表格數：{summary.failed_cross_page_tables}")
                    self.append_log(f"套用「內容大小＋靠右對齊」的表格數：{summary.special_autofit_right_tables}")
                    self.append_log(f"其他正常處理的表格數：{summary.normal_processed_tables}")
                    self.append_log(f"顏色調整總數：{summary.changed_colors}")
                    self.append_log(f"指定色碼改成 {DEFAULT_GRAY} 的儲存格數：{summary.changed_to_gray}")
                    self.append_log(f"其他顏色改成無色彩的儲存格數：{summary.cleared_colors}")
                    self.append_log(f"已套用階層縮排與大綱階層的段落數：{summary.paragraphs}")
                    self.append_log(f"總段落數：{summary.total_paragraphs}")
                    self.append_log(f"跳過目錄段落數：{summary.skipped_toc_paragraphs}")
                    self.append_log(f"跳過表格段落數：{summary.skipped_table_paragraphs}")
                    self.append_log(f"移除全文件既有大綱階層的段落數：{summary.removed_all_outline_paragraphs}")
                    self.append_log(f"套用壹、序言前縮排的段落數：{summary.indented_preface_paragraphs}")
                    self.append_log(f"套用壹、序言前大綱階層的段落數：{summary.outlined_preface_paragraphs}")
                    for level, count in enumerate(summary.paragraph_level_counts, start=1):
                        self.append_log(f"成功套用第 {level} 階數量：{count}")
                    self.append_log(f"無法判斷而跳過的段落數：{summary.unknown_paragraphs}")
                    self.append_log(f"輸出檔案：{output_path}")
                    if log_path is not None:
                        self.append_log(f"處理紀錄檔：{log_path}")
                    self.set_running_state(False)
                    messagebox.showinfo("完成", f"修改完成！\n\n輸出檔案：\n{output_path}")

                elif kind == "stopped":
                    self.current_temp_output = None
                    self.stop_progress_animation()
                    self.status_var.set("已停止，所有可控程序已中止，暫存檔已清理。")
                    self.append_log("已停止修改，已嘗試關閉 Word/PowerShell 程序並解除檔案鎖定。")
                    self.set_running_state(False)

                elif kind == "error":
                    _, err, tb = item
                    self.current_temp_output = None
                    self.stop_progress_animation()
                    self.status_var.set("發生錯誤。")
                    self.append_log("發生錯誤：" + err)
                    self.append_log(tb)
                    self.set_running_state(False)
                    messagebox.showerror("錯誤", err)

        except queue.Empty:
            pass

        self.root.after(120, self._poll_queue)
