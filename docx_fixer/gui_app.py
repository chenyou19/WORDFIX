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
from .process_log import write_process_log, write_table_log_file
from .stop_controller import StopController

DEFAULT_WINDOW_GEOMETRY = "1080x760"
MIN_WINDOW_SIZE = (980, 680)

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

        try:
            load_saved_indent_settings()
        except Exception as exc:
            self.indent_settings_load_error = str(exc)

        self.fix_table_var = tk.BooleanVar(value=False)
        self.fix_color_var = tk.BooleanVar(value=False)
        self.fix_paragraph_var = tk.BooleanVar(value=True)
        self.remove_all_outline_var = tk.BooleanVar(value=True)
        self.indent_preface_var = tk.BooleanVar(value=False)
        self.outline_preface_var = tk.BooleanVar(value=False)
        self.level2_body_first_line_indent_var = tk.BooleanVar(value=False)
        self.word_com_check_body_font_var = tk.BooleanVar(value=False)

        self.status_var = tk.StringVar(value="Choose a .docx file")
        self.progress_var = tk.DoubleVar(value=0)

        self._build_ui()
        if self.indent_settings_load_error:
            messagebox.showwarning(
                "蝮格??身頛憭望?",
                f"撌脫?典撱箇葬?身摰n\n{self.indent_settings_load_error}",
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
        notebook.add(process_tab, text="靽格?辣")
        notebook.add(indent_tab, text="蝮格??身")

        # 瑼??豢?
        file_frame = ttk.LabelFrame(process_tab, text="瑼?")
        file_frame.pack(fill="x", pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="Input file:").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(file_frame, textvariable=self.input_var).grid(row=0, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(file_frame, text="?汗...", command=self.browse_input).grid(row=0, column=2, padx=8, pady=8)

        ttk.Label(file_frame, text="Output file:").grid(row=1, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(file_frame, textvariable=self.output_var).grid(row=1, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(file_frame, text="?血???..", command=self.browse_output).grid(row=1, column=2, padx=8, pady=8)

        hint = ttk.Label(file_frame, text=f"Default output suffix: {DEFAULT_SUFFIX}.docx")
        hint.grid(row=2, column=1, padx=8, pady=(0, 8), sticky="w")

        # ??賊?
        option_frame = ttk.LabelFrame(process_tab, text="?臬?貊????寞?")
        option_frame.pack(fill="x", pady=(0, 10))

        ttk.Checkbutton(
            option_frame,
            text="隤踵銵冽?澆?嚗蔭銝准?擃?擃?11?銵?頝?獢?暺??祝摨?100%",
            variable=self.fix_table_var,
        ).pack(anchor="w", padx=8, pady=(8, 4))

        ttk.Checkbutton(
            option_frame,
            text="Fix table colors",
            variable=self.fix_color_var,
        ).pack(anchor="w", padx=8, pady=4)

        ttk.Checkbutton(
            option_frame,
            text="Remove all outline levels",
            variable=self.remove_all_outline_var,
        ).pack(anchor="w", padx=8, pady=4)

        ttk.Checkbutton(
            option_frame,
            text="隤踵畾菔嚗?蝭?惜??憯對?銝嚗?銝嚗?(銝)嚗?嚗?1嚗?(1)嚗嚗?A嚗?(A)嚗嚗?a嚗?(a) ?葬??銝血?銝?Word 憭抒雇?惜",
            variable=self.fix_paragraph_var,
        ).pack(anchor="w", padx=8, pady=4)

        ttk.Checkbutton(
            option_frame,
            text="Indent preface paragraphs",
            variable=self.indent_preface_var,
        ).pack(anchor="w", padx=8, pady=4)

        ttk.Checkbutton(
            option_frame,
            text="Outline preface paragraphs",
            variable=self.outline_preface_var,
        ).pack(anchor="w", padx=8, pady=4)

        ttk.Checkbutton(
            option_frame,
            text="Apply 560 twips first-line indent to plain body text under level 2 headings",
            variable=self.level2_body_first_line_indent_var,
        ).pack(anchor="w", padx=28, pady=4)

        ttk.Checkbutton(
            option_frame,
            text="When XML body font is not 14pt, ask Word COM to verify before applying body indent",
            variable=self.word_com_check_body_font_var,
        ).pack(anchor="w", padx=28, pady=(0, 8))

        # ?脣漲????
        progress_frame = ttk.LabelFrame(process_tab, text="?脣漲")
        progress_frame.pack(fill="x", pady=(0, 10))

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
        )
        self.progress_bar.pack(fill="x", padx=8, pady=(8, 4))

        ttk.Label(progress_frame, textvariable=self.status_var).pack(anchor="w", padx=8, pady=(0, 8))

        # ??
        button_frame = ttk.Frame(process_tab)
        button_frame.pack(fill="x", pady=(0, 10))

        self.start_button = ttk.Button(button_frame, text="??靽格", command=self.start_process)
        self.start_button.pack(side="left", padx=(0, 8))

        self.stop_button = ttk.Button(button_frame, text="?迫靽格", command=self.stop_process, state="disabled")
        self.stop_button.pack(side="left", padx=8)

        ttk.Button(button_frame, text="蝯迫蝔?", command=self.exit_app).pack(side="right")

        # 閮蝝??
        log_frame = ttk.LabelFrame(process_tab, text="閮")
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
            title="Preface indents",
            section="preface",
            rows=settings["preface"],
            row=0,
            column=0,
        )
        self._build_indent_section(
            scroll_frame,
            title="Body indents",
            section="body",
            rows=settings["body"],
            row=0,
            column=1,
        )

        button_frame = ttk.Frame(parent)
        button_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        ttk.Button(button_frame, text="憟?桀?閮剖?", command=self.apply_indent_entries).pack(side="left")
        headers = ["Level", "Label", "Number start cm", "Hanging cm", "Body left cm"]
        ttk.Button(button_frame, text="???啁??批遣?身", command=self.restore_builtin_indent_defaults).pack(side="left", padx=8)

        path_label = ttk.Label(
            button_frame,
            text=f"?身瑼?{get_indent_settings_path()}",
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

        headers = ["?惜", "蝺刻??澆?", "璅?韏琿? cm", "?豢?頝 cm", "?扳?韏琿? cm"]
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
            "preface": "Preface",
            "body": "Body",
        }

        for section, rows in self.indent_vars.items():
            for display_index, (level, label, number_start_var, hanging_var, body_left_var) in enumerate(rows, start=1):
                try:
                    number_start_cm = float(number_start_var.get().strip())
                    hanging_cm = float(hanging_var.get().strip())
                    body_left_cm = float(body_left_var.get().strip())
                except ValueError as exc:
                    raise ValueError(
                        f"{section_names[section]} level {display_index}: please enter numeric cm values"
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
            messagebox.showerror("蝮格?閮剖??航炊", str(exc))
            return False

        self.status_var.set("Indent settings applied")
        return True

    def save_indent_defaults(self) -> None:
        try:
            path = save_indent_settings(self.collect_indent_entries())
        except Exception as exc:
            messagebox.showerror("蝮格?閮剖??航炊", str(exc))
            return

        messagebox.showinfo("Saved", f"Indent settings saved:\n{path}")


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
        self.status_var.set("Restored built-in indent settings")

    def browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="?豢? Word DOCX 瑼?",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.input_var.set(path)
            self.status_var.set("Input file selected")

    def browse_output(self) -> None:
        input_path = self.input_var.get().strip()
        initialdir = None
        initialfile = "output.docx"

        if input_path:
            p = Path(input_path)
            initialdir = str(p.parent)
            initialfile = f"{p.stem}{DEFAULT_SUFFIX}.docx"

        path = filedialog.asksaveasfilename(
            title="?豢?頛詨瑼?",
            defaultextension=".docx",
            initialdir=initialdir,
            initialfile=initialfile,
            filetypes=[("Word ?辣", "*.docx")],
        )
        if path:
            input_text = self.input_var.get().strip()
            if input_text and is_same_file_path(input_text, path):
                messagebox.showerror(
                    "頛詨瑼??航炊",
                    "靽格敺?瑼?銝隞亥???瑼??詨??n\n"
                    "Choose a different output path, for example *_fixed.docx.",
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
            messagebox.showwarning("Missing input", "Please choose an input .docx file")
            return None

        input_path = Path(input_text)
        if not input_path.exists():
            messagebox.showerror("File not found", f"Input file does not exist:\n{input_path}")
            return None

        if input_path.suffix.lower() != ".docx":
            messagebox.showerror("Invalid file type", "Input file must be a .docx file")
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
            enable_level2_body_first_line_indent=self.level2_body_first_line_indent_var.get(),
            word_com_check_body_font_when_xml_not_14=self.word_com_check_body_font_var.get(),
        )

        if not (
            options.fix_table_layout
            or options.fix_color
            or options.fix_paragraph
            or options.remove_all_outline_levels
            or options.indent_preface_paragraphs
            or options.outline_preface_paragraphs
        ):
            messagebox.showwarning("No options selected", "Please select at least one processing option")
            return None

        output_path = self.resolve_output_path(input_path)

        if is_same_file_path(input_path, output_path):
            messagebox.showerror(
                "頛詨瑼??航炊",
                "靽格敺?瑼?銝隞亥???瑼??詨??n\n"
                "Choose a different output path, for example *_fixed.docx.",
            )
            return None

        if output_path.exists():
            ok = messagebox.askyesno(
                "Overwrite output file?",
                f"頛詨瑼?撌脣??剁??臬閬?嚗n\n{output_path}",
            )
            if not ok:
                return None

        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Output folder error", str(exc))
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
        self.status_var.set("Processing...")
        self.start_progress_animation()
        self.append_log(f"Input: {input_path}")
        self.append_log(f"Output: {output_path}")
        self.append_log("Started processing")
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
                # Windows ??格?瑼?鋡?Word??獢蜇蝞⊿?閬賜??潦neDrive 蝑??冽???閬???
                # ?粹?歇??憟賜??怠?瑼◤?芣?嚗摮?撣嗆???瑼???
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                fallback_output = output_path.with_name(
                    f"{output_path.stem}_{timestamp}{output_path.suffix}"
                )
                os.replace(temp_output, fallback_output)
                final_output = fallback_output
                self.ui_queue.put((
                    "warning",
                    f"Could not replace output file; wrote fallback: {fallback_output}",
                ))

            log_path = None
            table_log_path = None
            try:
                log_path = write_process_log(final_output, summary)
            except Exception as exc:
                self.ui_queue.put(("warning", f"Could not write process log: {exc}"))
            try:
                table_log_path = write_table_log_file(final_output, summary)
            except Exception as exc:
                self.ui_queue.put(("warning", f"Could not write table log: {exc}"))

            self.ui_queue.put(("done", final_output, summary, log_path, table_log_path))

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
            self.status_var.set("甇?銝剜迫???摨蒂?瑼???...")
            self.append_log("Stop requested; waiting for current operation to finish")

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
                    self.append_log("Warning: " + message)

                elif kind == "done":
                    _, output_path, summary, log_path, table_log_path = item
                    self.current_temp_output = None
                    self.stop_progress_animation(100)
                    self.status_var.set("Done")
                    self.append_log("Done")
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
                    self.set_running_state(False)
                    messagebox.showinfo("摰?", f"靽格摰?嚗n\n頛詨瑼?嚗n{output_path}")

                elif kind == "stopped":
                    self.current_temp_output = None
                    self.stop_progress_animation()
                    self.status_var.set("Stopped")
                    self.append_log("Processing stopped")
                    self.set_running_state(False)

                elif kind == "error":
                    _, err, tb = item
                    self.current_temp_output = None
                    self.stop_progress_animation()
                    self.status_var.set("Error")
                    self.append_log("Error: " + err)
                    self.append_log(tb)
                    self.set_running_state(False)
                    messagebox.showerror("?航炊", err)

        except queue.Empty:
            pass

        self.root.after(120, self._poll_queue)
