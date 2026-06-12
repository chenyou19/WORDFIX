from __future__ import annotations

import datetime
import os
import queue
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .constants import DEFAULT_SUFFIX
from .docx_processor import fix_docx_fast
from .exceptions import ProcessStopped
from .gui_defaults import (
    built_in_gui_defaults,
    load_saved_gui_defaults,
    save_gui_defaults as write_gui_defaults,
)
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
from .process_log import write_heading_suffix_log_file, write_process_log, write_table_log_file
from .stop_controller import StopController

DEFAULT_WINDOW_GEOMETRY = "1080x760"
MIN_WINDOW_SIZE = (980, 680)
_BUILTIN_GUI_DEFAULTS = built_in_gui_defaults()
DEFAULT_SKIP_CHAPTER_THREE_TABLE_LAYOUT = _BUILTIN_GUI_DEFAULTS["skip_chapter_three_table_layout"]
DEFAULT_SKIP_CHAPTER_THREE_TABLE_COLOR = _BUILTIN_GUI_DEFAULTS["skip_chapter_three_table_color"]
DEFAULT_SKIP_CHAPTER_THREE_INDENTS = _BUILTIN_GUI_DEFAULTS["skip_chapter_three_indents"]


def write_logs_if_enabled(
    final_output: Path,
    summary,
    skip_log_output: bool,
    warning_callback=None,
) -> tuple[Path | None, Path | None, Path | None]:
    log_path = None
    table_log_path = None
    heading_suffix_log_path = None

    if skip_log_output:
        return log_path, table_log_path, heading_suffix_log_path

    try:
        log_path = write_process_log(final_output, summary)
    except Exception as exc:
        if warning_callback is not None:
            warning_callback(f"無法寫入處理 log：{exc}")

    try:
        table_log_path = write_table_log_file(final_output, summary)
    except Exception as exc:
        if warning_callback is not None:
            warning_callback(f"無法寫入表格 log：{exc}")

    try:
        heading_suffix_log_path = write_heading_suffix_log_file(final_output, summary)
    except Exception as exc:
        if warning_callback is not None:
            warning_callback(f"無法寫入標題後方分隔符 log：{exc}")

    return log_path, table_log_path, heading_suffix_log_path


class DocxFixerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Word DOCX Fixer")
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
        self.gui_defaults_load_error: str | None = None

        try:
            load_saved_indent_settings()
        except Exception as exc:
            self.indent_settings_load_error = str(exc)

        try:
            gui_defaults = load_saved_gui_defaults()
        except Exception as exc:
            self.gui_defaults_load_error = str(exc)
            gui_defaults = built_in_gui_defaults()

        self.fix_table_var = tk.BooleanVar(value=gui_defaults["fix_table"])
        self.fix_color_var = tk.BooleanVar(value=gui_defaults["fix_color"])
        self.fix_paragraph_var = tk.BooleanVar(value=gui_defaults["fix_paragraph"])
        self.remove_all_outline_var = tk.BooleanVar(value=gui_defaults["remove_all_outline"])
        self.indent_preface_var = tk.BooleanVar(value=gui_defaults["indent_preface"])
        self.outline_preface_var = tk.BooleanVar(value=gui_defaults["outline_preface"])
        self.level1_level2_body_first_line_indent_var = tk.BooleanVar(
            value=gui_defaults["level1_level2_body_first_line_indent"]
        )
        self.word_com_check_body_font_var = tk.BooleanVar(value=gui_defaults["word_com_check_body_font"])
        self.skip_log_output_var = tk.BooleanVar(value=gui_defaults["skip_log_output"])
        self.skip_nested_tables_var = tk.BooleanVar(value=gui_defaults["skip_nested_tables"])
        self.skip_chapter_three_table_layout_var = tk.BooleanVar(
            value=gui_defaults["skip_chapter_three_table_layout"]
        )
        self.skip_chapter_three_table_color_var = tk.BooleanVar(
            value=gui_defaults["skip_chapter_three_table_color"]
        )
        self.skip_chapter_three_indents_var = tk.BooleanVar(value=gui_defaults["skip_chapter_three_indents"])

        self.status_var = tk.StringVar(value="請先選擇 .docx 檔案")
        self.progress_var = tk.DoubleVar(value=0)

        self._build_ui()
        if self.indent_settings_load_error:
            messagebox.showwarning(
                "縮排設定載入失敗",
                "無法讀取既有的縮排設定，將改用目前內建預設。\n"
                f"{self.indent_settings_load_error}",
            )
        if self.gui_defaults_load_error:
            messagebox.showwarning(
                "GUI 預設勾選方案載入失敗",
                "無法讀取既有的 GUI 預設勾選方案，將改用內建預設。\n"
                f"{self.gui_defaults_load_error}",
            )
        self._poll_queue()
        self.root.protocol("WM_DELETE_WINDOW", self.exit_app)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)

        title = ttk.Label(outer, text="Word DOCX Fixer", font=("Microsoft JhengHei UI", 16, "bold"))
        title.pack(anchor="w", pady=(0, 10))

        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True)

        process_tab = ttk.Frame(notebook, padding=10)
        indent_tab = ttk.Frame(notebook, padding=10)
        notebook.add(process_tab, text="處理")
        notebook.add(indent_tab, text="縮排設定")

        file_frame = ttk.LabelFrame(process_tab, text="檔案")
        file_frame.pack(fill="x", pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="輸入檔案：").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(file_frame, textvariable=self.input_var).grid(row=0, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(file_frame, text="瀏覽...", command=self.browse_input).grid(row=0, column=2, padx=8, pady=8)

        ttk.Label(file_frame, text="輸出檔案：").grid(row=1, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(file_frame, textvariable=self.output_var).grid(row=1, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(file_frame, text="另存...", command=self.browse_output).grid(row=1, column=2, padx=8, pady=8)

        hint = ttk.Label(file_frame, text=f"未指定輸出檔時，預設加上 {DEFAULT_SUFFIX}.docx")
        hint.grid(row=2, column=1, padx=8, pady=(0, 8), sticky="w")

        option_frame = ttk.LabelFrame(process_tab, text="處理選項")
        option_frame.pack(fill="x", pady=(0, 10))
        option_frame.columnconfigure(0, weight=1)
        option_frame.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            option_frame,
            text="修正表格版面",
            variable=self.fix_table_var,
        ).grid(row=0, column=0, padx=(12, 16), pady=(10, 4), sticky="w")

        ttk.Checkbutton(
            option_frame,
            text="修正表格底色",
            variable=self.fix_color_var,
        ).grid(row=1, column=0, padx=(12, 16), pady=4, sticky="w")

        ttk.Checkbutton(
            option_frame,
            text="移除所有段落大綱階層",
            variable=self.remove_all_outline_var,
        ).grid(row=0, column=1, padx=(12, 16), pady=(10, 4), sticky="w")

        ttk.Checkbutton(
            option_frame,
            text="修正段落大綱階層與縮排",
            variable=self.fix_paragraph_var,
        ).grid(row=1, column=1, padx=(12, 16), pady=4, sticky="w")

        preface_option_frame = ttk.Frame(option_frame)
        preface_option_frame.grid(row=2, column=0, padx=(12, 16), pady=(4, 10), sticky="nw")

        ttk.Checkbutton(
            preface_option_frame,
            text="前言段落套用縮排",
            variable=self.indent_preface_var,
        ).grid(row=0, column=0, pady=4, sticky="w")

        ttk.Checkbutton(
            preface_option_frame,
            text="前言段落套用大綱階層",
            variable=self.outline_preface_var,
        ).grid(row=1, column=0, pady=4, sticky="w")

        ttk.Checkbutton(
            preface_option_frame,
            text="XML 判斷非 14pt 時使用 Word COM 確認內文字號",
            variable=self.word_com_check_body_font_var,
        ).grid(row=2, column=0, pady=4, sticky="w")

        ttk.Checkbutton(
            preface_option_frame,
            text="不要輸出 log 檔",
            variable=self.skip_log_output_var,
        ).grid(row=3, column=0, pady=(4, 0), sticky="w")

        advanced_option_frame = ttk.Frame(option_frame)
        advanced_option_frame.grid(row=2, column=1, padx=(12, 16), pady=(4, 10), sticky="nw")

        ttk.Checkbutton(
            advanced_option_frame,
            text="階層 1、2 標題下方普通內文首行縮排兩個中文字",
            variable=self.level1_level2_body_first_line_indent_var,
        ).grid(row=0, column=0, pady=4, sticky="w")

        ttk.Checkbutton(
            advanced_option_frame,
            text="表格中有表格不調整",
            variable=self.skip_nested_tables_var,
        ).grid(row=1, column=0, pady=4, sticky="w")

        ttk.Checkbutton(
            advanced_option_frame,
            text="參、價格形成之主要因素分析：表格版面不調整",
            variable=self.skip_chapter_three_table_layout_var,
        ).grid(row=2, column=0, pady=4, sticky="w")

        ttk.Checkbutton(
            advanced_option_frame,
            text="參、價格形成之主要因素分析：表格顏色不調整",
            variable=self.skip_chapter_three_table_color_var,
        ).grid(row=3, column=0, pady=4, sticky="w")

        ttk.Checkbutton(
            advanced_option_frame,
            text="參、價格形成之主要因素分析：縮排不調整",
            variable=self.skip_chapter_three_indents_var,
        ).grid(row=4, column=0, pady=4, sticky="w")

        defaults_button_frame = ttk.Frame(option_frame)
        defaults_button_frame.grid(row=3, column=0, columnspan=2, padx=(12, 16), pady=(0, 10), sticky="w")
        ttk.Button(
            defaults_button_frame,
            text="保存目前勾選為預設方案",
            command=self.save_gui_defaults,
        ).pack(side="left")
        ttk.Button(
            defaults_button_frame,
            text="還原內建勾選預設",
            command=self.restore_builtin_gui_defaults,
        ).pack(side="left", padx=8)

        progress_frame = ttk.LabelFrame(process_tab, text="處理進度")
        progress_frame.pack(fill="x", pady=(0, 10))

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
        )
        self.progress_bar.pack(fill="x", padx=8, pady=(8, 4))

        ttk.Label(progress_frame, textvariable=self.status_var).pack(anchor="w", padx=8, pady=(0, 8))

        button_frame = ttk.Frame(process_tab)
        button_frame.pack(fill="x", pady=(0, 10))

        self.start_button = ttk.Button(button_frame, text="開始處理", command=self.start_process)
        self.start_button.pack(side="left", padx=(0, 8))

        self.stop_button = ttk.Button(button_frame, text="停止處理", command=self.stop_process, state="disabled")
        self.stop_button.pack(side="left", padx=8)

        ttk.Button(button_frame, text="離開", command=self.exit_app).pack(side="right")

        log_frame = ttk.LabelFrame(process_tab, text="處理紀錄")
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
            title="前言縮排",
            section="preface",
            rows=settings["preface"],
            row=0,
            column=0,
        )
        self._build_indent_section(
            scroll_frame,
            title="內文縮排",
            section="body",
            rows=settings["body"],
            row=0,
            column=1,
        )

        button_frame = ttk.Frame(parent)
        button_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        ttk.Button(button_frame, text="套用目前設定", command=self.apply_indent_entries).pack(side="left")
        ttk.Button(button_frame, text="保存成預設樣式", command=self.save_indent_defaults).pack(side="left", padx=8)
        ttk.Button(button_frame, text="還原內建預設", command=self.restore_builtin_indent_defaults).pack(side="left", padx=8)

        path_label = ttk.Label(
            button_frame,
            text=f"設定檔位置：{get_indent_settings_path()}",
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

        headers = ["階層", "標號", "起點 cm", "懸掛 cm", "內文起點 cm"]
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
            "preface": "前言",
            "body": "內文",
        }

        for section, rows in self.indent_vars.items():
            for display_index, (level, label, number_start_var, hanging_var, body_left_var) in enumerate(rows, start=1):
                try:
                    number_start_cm = float(number_start_var.get().strip())
                    hanging_cm = float(hanging_var.get().strip())
                    body_left_cm = float(body_left_var.get().strip())
                except ValueError as exc:
                    raise ValueError(
                        f"{section_names[section]}第 {display_index} 列請輸入有效數字"
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

        self.status_var.set("已套用縮排設定")
        return True

    def save_indent_defaults(self) -> None:
        try:
            path = save_indent_settings(self.collect_indent_entries())
        except Exception as exc:
            messagebox.showerror("縮排設定錯誤", str(exc))
            return

        self.status_var.set("已保存並套用縮排預設")
        messagebox.showinfo("已儲存", f"縮排設定已儲存：\n{path}")

    def collect_gui_defaults(self) -> dict[str, bool]:
        return {
            "fix_table": self.fix_table_var.get(),
            "fix_color": self.fix_color_var.get(),
            "fix_paragraph": self.fix_paragraph_var.get(),
            "remove_all_outline": self.remove_all_outline_var.get(),
            "indent_preface": self.indent_preface_var.get(),
            "outline_preface": self.outline_preface_var.get(),
            "level1_level2_body_first_line_indent": self.level1_level2_body_first_line_indent_var.get(),
            "word_com_check_body_font": self.word_com_check_body_font_var.get(),
            "skip_log_output": self.skip_log_output_var.get(),
            "skip_nested_tables": self.skip_nested_tables_var.get(),
            "skip_chapter_three_table_layout": self.skip_chapter_three_table_layout_var.get(),
            "skip_chapter_three_table_color": self.skip_chapter_three_table_color_var.get(),
            "skip_chapter_three_indents": self.skip_chapter_three_indents_var.get(),
        }

    def save_gui_defaults(self) -> None:
        try:
            path = write_gui_defaults(self.collect_gui_defaults())
        except Exception as exc:
            messagebox.showerror("GUI 預設勾選方案錯誤", str(exc))
            return

        self.status_var.set("已保存 GUI 預設勾選方案")
        messagebox.showinfo("已儲存", f"GUI 預設勾選方案已儲存：\n{path}")

    def restore_builtin_gui_defaults(self) -> None:
        defaults = built_in_gui_defaults()
        self.fix_table_var.set(defaults["fix_table"])
        self.fix_color_var.set(defaults["fix_color"])
        self.fix_paragraph_var.set(defaults["fix_paragraph"])
        self.remove_all_outline_var.set(defaults["remove_all_outline"])
        self.indent_preface_var.set(defaults["indent_preface"])
        self.outline_preface_var.set(defaults["outline_preface"])
        self.level1_level2_body_first_line_indent_var.set(
            defaults["level1_level2_body_first_line_indent"]
        )
        self.word_com_check_body_font_var.set(defaults["word_com_check_body_font"])
        self.skip_log_output_var.set(defaults["skip_log_output"])
        self.skip_nested_tables_var.set(defaults["skip_nested_tables"])
        self.skip_chapter_three_table_layout_var.set(defaults["skip_chapter_three_table_layout"])
        self.skip_chapter_three_table_color_var.set(defaults["skip_chapter_three_table_color"])
        self.skip_chapter_three_indents_var.set(defaults["skip_chapter_three_indents"])
        self.status_var.set("已還原內建 GUI 預設勾選方案")

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
        self.status_var.set("已還原內建縮排設定")

    def browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="選擇 Word DOCX 檔案",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.input_var.set(path)
            self.status_var.set("已選擇輸入檔")

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
            filetypes=[("Word documents", "*.docx")],
        )
        if path:
            input_text = self.input_var.get().strip()
            if input_text and is_same_file_path(input_text, path):
                messagebox.showerror(
                    "輸出檔案錯誤",
                    "輸出檔案不能與輸入檔案相同。\n請選擇其他輸出位置或使用 *_fixed.docx。",
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
            messagebox.showwarning("尚未選擇輸入檔", "請選擇一個 .docx 檔案")
            return None

        input_path = Path(input_text)
        if not input_path.exists():
            messagebox.showerror("檔案不存在", f"輸入檔案不存在：\n{input_path}")
            return None

        if input_path.suffix.lower() != ".docx":
            messagebox.showerror("檔案格式錯誤", "輸入檔案必須是 .docx")
            return None

        if not self.apply_indent_entries():
            return None

        options = ProcessOptions(
            fix_table_layout=self.fix_table_var.get(),
            fix_color=self.fix_color_var.get(),
            fix_paragraph=self.fix_paragraph_var.get(),
            remove_all_outline_levels=self.remove_all_outline_var.get(),
            indent_preface_paragraphs=self.indent_preface_var.get(),
            outline_preface_paragraphs=self.outline_preface_var.get(),
            enable_level1_level2_body_first_line_indent=self.level1_level2_body_first_line_indent_var.get(),
            word_com_check_body_font_when_xml_not_14=self.word_com_check_body_font_var.get(),
            skip_chapter_three_table_layout=self.skip_chapter_three_table_layout_var.get(),
            skip_chapter_three_table_color=self.skip_chapter_three_table_color_var.get(),
            skip_chapter_three_indents=self.skip_chapter_three_indents_var.get(),
            skip_nested_tables=self.skip_nested_tables_var.get(),
            skip_log_output=self.skip_log_output_var.get(),
        )

        if not (
            options.fix_table_layout
            or options.fix_color
            or options.fix_paragraph
            or options.remove_all_outline_levels
            or options.indent_preface_paragraphs
            or options.outline_preface_paragraphs
        ):
            messagebox.showwarning("尚未選擇處理項目", "請至少勾選一個處理選項")
            return None

        output_path = self.resolve_output_path(input_path)

        if is_same_file_path(input_path, output_path):
            messagebox.showerror(
                "輸出檔案錯誤",
                "輸出檔案不能與輸入檔案相同。\n請選擇其他輸出位置或使用 *_fixed.docx。",
            )
            return None

        if output_path.exists():
            ok = messagebox.askyesno(
                "覆寫輸出檔案？",
                f"輸出檔案已存在，是否覆寫？\n\n{output_path}",
            )
            if not ok:
                return None

        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("輸出資料夾錯誤", str(exc))
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
        self.status_var.set("處理中...")
        self.start_progress_animation()
        self.append_log(f"輸入檔案: {input_path}")
        self.append_log(f"輸出檔案: {output_path}")
        self.append_log("開始處理")
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
                raise ValueError("Input and output paths must be different")

            try:
                os.replace(temp_output, output_path)
                final_output = output_path
            except PermissionError:
                # If the target is open in Word, save beside it with a timestamp.
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                fallback_output = output_path.with_name(
                    f"{output_path.stem}_{timestamp}{output_path.suffix}"
                )
                os.replace(temp_output, fallback_output)
                final_output = fallback_output
                self.ui_queue.put((
                    "warning",
                    f"無法覆寫原輸出檔，已改存為：{fallback_output}",
                ))

            log_path, table_log_path, heading_suffix_log_path = write_logs_if_enabled(
                final_output,
                summary,
                options.skip_log_output,
                warning_callback=lambda message: self.ui_queue.put(("warning", message)),
            )

            self.ui_queue.put(("done", final_output, summary, log_path, table_log_path, heading_suffix_log_path))

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
            self.status_var.set("已送出停止請求，正在收尾...")
            self.append_log("已要求停止處理")

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
                    self.append_log("警告: " + message)

                elif kind == "done":
                    _, output_path, summary, log_path, table_log_path, heading_suffix_log_path = item
                    self.current_temp_output = None
                    self.stop_progress_animation(100)
                    self.status_var.set("處理完成")
                    self.append_log("處理完成")
                    self.append_log(f"tables={summary.tables}")
                    self.append_log(f"paragraphs_changed={summary.paragraphs}")
                    self.append_log(f"total_paragraphs={summary.total_paragraphs}")
                    self.append_log(f"skipped_toc_paragraphs={summary.skipped_toc_paragraphs}")
                    self.append_log(f"skipped_table_paragraphs={summary.skipped_table_paragraphs}")
                    self.append_log(f"removed_all_outline_paragraphs={summary.removed_all_outline_paragraphs}")
                    self.append_log(f"unknown_paragraphs={summary.unknown_paragraphs}")
                    self.append_log(f"output={output_path}")
                    if log_path is not None:
                        self.append_log(f"process_log={log_path}")
                    if table_log_path is not None:
                        self.append_log(f"table_log={table_log_path}")
                    if heading_suffix_log_path is not None:
                        self.append_log(f"heading_suffix_log={heading_suffix_log_path}")
                    self.set_running_state(False)
                    messagebox.showinfo(
                        "處理完成",
                        "DOCX 處理完成。\n\n"
                        f"輸出檔案：\n{output_path}\n\n"
                        f"處理 log：\n{log_path if log_path is not None else '未產生'}\n\n"
                        f"表格 log：\n{table_log_path if table_log_path is not None else '未產生'}\n\n"
                        f"標題後方分隔符 log：\n{heading_suffix_log_path if heading_suffix_log_path is not None else '未產生'}",
                    )

                elif kind == "stopped":
                    self.current_temp_output = None
                    self.stop_progress_animation()
                    self.status_var.set("已停止")
                    self.append_log("處理已停止")
                    self.set_running_state(False)

                elif kind == "error":
                    _, err, tb = item
                    self.current_temp_output = None
                    self.stop_progress_animation()
                    self.status_var.set("發生錯誤")
                    self.append_log("錯誤: " + err)
                    self.append_log(tb)
                    self.set_running_state(False)
                    messagebox.showerror("處理錯誤", err)

        except queue.Empty:
            pass

        self.root.after(120, self._poll_queue)
