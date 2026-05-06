"""Tkinter GUI for AutoGraph."""

import os
import queue
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from autograph_service.service import (
    DATA_KIND_S2P,
    DATA_KIND_TXT,
    DEFAULT_CHARTS_SHEET,
    DEFAULT_TEMPLATE_SHEET,
    JobConfig,
    JobValidationError,
    SOURCE_KIND_DIRECTORY,
    SOURCE_KIND_FILE,
    run_job,
    validate_job_config,
)


class AutoGraphApp(ttk.Frame):
    """Single-window GUI for running AutoGraph jobs."""

    def __init__(self, master=None, runner=run_job):
        super().__init__(master, padding=12)
        self.master = master
        self.runner = runner
        self.queue = queue.Queue()
        self.worker_thread = None
        self.interactive_widgets = []
        self.field_rows = {}
        self.field_order = []

        self.data_kind_var = tk.StringVar(value=DATA_KIND_TXT)
        self.source_kind_var = tk.StringVar(value=SOURCE_KIND_FILE)
        self.template_mode_var = tk.BooleanVar(value=True)
        self.use_template_charts_var = tk.BooleanVar(value=True)

        self.input_path_var = tk.StringVar()
        self.excel_path_var = tk.StringVar()
        self.sheet_name_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.template_sheet_var = tk.StringVar(value=DEFAULT_TEMPLATE_SHEET)
        self.charts_sheet_var = tk.StringVar(value=DEFAULT_CHARTS_SHEET)
        self.charts_sheet_template_var = tk.StringVar(value=DEFAULT_CHARTS_SHEET)

        self.input_label_var = tk.StringVar()
        self.excel_label_var = tk.StringVar()
        self.s2p_note_var = tk.StringVar()

        self.grid(sticky="nsew")
        self._configure_root()
        self._build_ui()
        self._bind_state_updates()
        self._refresh_form()

    def get_visible_field_keys(self):
        """Return visible field keys for the current mode."""
        keys = ["input_path", "excel_path"]
        template_mode = self.template_mode_var.get()
        single_file = self.source_kind_var.get() == SOURCE_KIND_FILE

        if self.data_kind_var.get() == DATA_KIND_S2P:
            keys.extend(["template_sheet", "charts_sheet_template", "use_template_charts"])
            if single_file:
                keys.append("output_path")
            return keys

        if template_mode:
            keys.extend(["template_sheet", "charts_sheet"])
            if single_file:
                keys.append("output_path")
            return keys

        if single_file:
            keys.extend(["sheet_name", "output_path"])
        return keys

    def clear_form(self):
        """Reset form inputs and clear logs."""
        self.data_kind_var.set(DATA_KIND_TXT)
        self.source_kind_var.set(SOURCE_KIND_FILE)
        self.template_mode_var.set(True)
        self.use_template_charts_var.set(True)
        self.input_path_var.set("")
        self.excel_path_var.set("")
        self.sheet_name_var.set("")
        self.output_path_var.set("")
        self.template_sheet_var.set(DEFAULT_TEMPLATE_SHEET)
        self.charts_sheet_var.set(DEFAULT_CHARTS_SHEET)
        self.charts_sheet_template_var.set(DEFAULT_CHARTS_SHEET)
        self._set_log_text("")
        self._refresh_form()

    def collect_config(self):
        """Build a JobConfig instance from the form state."""
        template_mode = self.template_mode_var.get()
        if self.data_kind_var.get() == DATA_KIND_S2P:
            template_mode = True

        return JobConfig(
            data_kind=self.data_kind_var.get(),
            source_kind=self.source_kind_var.get(),
            input_path=self.input_path_var.get(),
            excel_path=self.excel_path_var.get(),
            template_mode=template_mode,
            sheet_name=self.sheet_name_var.get(),
            output_path=self.output_path_var.get(),
            template_sheet=self.template_sheet_var.get(),
            charts_sheet=self.charts_sheet_var.get(),
            charts_sheet_template=self.charts_sheet_template_var.get(),
            use_template_charts=self.use_template_charts_var.get(),
        )

    def start_run(self):
        """Validate the form and start the job in a worker thread."""
        try:
            config = validate_job_config(self.collect_config())
        except JobValidationError as error:
            messagebox.showerror("Ошибка проверки", str(error), parent=self.master)
            return

        self._set_busy(True)
        self._append_log("Запуск обработки...")

        self.worker_thread = threading.Thread(
            target=self._run_in_worker,
            args=(config,),
            daemon=True,
        )
        self.worker_thread.start()
        self.after(100, self._poll_queue)

    def _configure_root(self):
        self.master.title("AutoGraph")
        self.master.geometry("920x700")
        self.master.minsize(760, 620)
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)

    def _build_ui(self):
        title = ttk.Label(
            self,
            text="AutoGraph",
            font=("Segoe UI", 16, "bold"),
        )
        title.grid(row=0, column=0, sticky="w")

        subtitle = ttk.Label(
            self,
            text="Перенос измерительных данных из TXT и S2P в Excel-шаблоны",
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(2, 12))

        mode_frame = ttk.LabelFrame(self, text="Режим")
        mode_frame.grid(row=2, column=0, sticky="ew")
        mode_frame.columnconfigure(0, weight=1)
        mode_frame.columnconfigure(1, weight=1)

        type_frame = ttk.Frame(mode_frame)
        type_frame.grid(row=0, column=0, sticky="w", padx=(8, 16), pady=8)
        ttk.Label(type_frame, text="Тип данных:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        txt_rb = ttk.Radiobutton(
            type_frame,
            text="TXT",
            value=DATA_KIND_TXT,
            variable=self.data_kind_var,
        )
        txt_rb.grid(row=0, column=1, sticky="w")
        s2p_rb = ttk.Radiobutton(
            type_frame,
            text="S2P",
            value=DATA_KIND_S2P,
            variable=self.data_kind_var,
        )
        s2p_rb.grid(row=0, column=2, sticky="w", padx=(8, 0))

        source_frame = ttk.Frame(mode_frame)
        source_frame.grid(row=0, column=1, sticky="w", padx=8, pady=8)
        ttk.Label(source_frame, text="Источник:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        single_rb = ttk.Radiobutton(
            source_frame,
            text="Один файл",
            value=SOURCE_KIND_FILE,
            variable=self.source_kind_var,
        )
        single_rb.grid(row=0, column=1, sticky="w")
        directory_rb = ttk.Radiobutton(
            source_frame,
            text="Папка",
            value=SOURCE_KIND_DIRECTORY,
            variable=self.source_kind_var,
        )
        directory_rb.grid(row=0, column=2, sticky="w", padx=(8, 0))

        template_check = ttk.Checkbutton(
            mode_frame,
            text="Шаблонный режим",
            variable=self.template_mode_var,
        )
        template_check.grid(row=1, column=0, sticky="w", padx=8, pady=(0, 8))
        self.template_check = template_check

        s2p_note = ttk.Label(
            mode_frame,
            textvariable=self.s2p_note_var,
            foreground="#8B5A00",
        )
        s2p_note.grid(row=1, column=1, sticky="w", padx=8, pady=(0, 8))

        form_frame = ttk.LabelFrame(self, text="Параметры")
        form_frame.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
        form_frame.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=0)
        self.columnconfigure(0, weight=1)

        self._create_path_row(
            form_frame,
            "input_path",
            self.input_label_var,
            self.input_path_var,
            self._browse_input_path,
        )
        self._create_path_row(
            form_frame,
            "excel_path",
            self.excel_label_var,
            self.excel_path_var,
            self._browse_excel_path,
        )
        self._create_path_row(
            form_frame,
            "output_path",
            "Выходной файл",
            self.output_path_var,
            self._browse_output_path,
        )
        self._create_entry_row(form_frame, "sheet_name", "Имя листа", self.sheet_name_var)
        self._create_entry_row(form_frame, "template_sheet", "Лист-шаблон", self.template_sheet_var)
        self._create_entry_row(form_frame, "charts_sheet", "Лист графиков", self.charts_sheet_var)
        self._create_entry_row(
            form_frame,
            "charts_sheet_template",
            "Лист графиков в шаблоне",
            self.charts_sheet_template_var,
        )
        self._create_check_row(
            form_frame,
            "use_template_charts",
            "Использовать графики из шаблона",
            self.use_template_charts_var,
        )

        buttons_frame = ttk.Frame(self)
        buttons_frame.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        buttons_frame.columnconfigure(0, weight=1)

        run_button = ttk.Button(buttons_frame, text="Запустить", command=self.start_run)
        run_button.grid(row=0, column=0, sticky="w")
        clear_button = ttk.Button(buttons_frame, text="Очистить", command=self.clear_form)
        clear_button.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.run_button = run_button
        self.clear_button = clear_button
        self.interactive_widgets.extend([txt_rb, s2p_rb, single_rb, directory_rb, template_check, run_button, clear_button])

        log_frame = ttk.LabelFrame(self, text="Лог выполнения")
        log_frame.grid(row=5, column=0, sticky="nsew", pady=(12, 0))
        self.rowconfigure(5, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=16, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _bind_state_updates(self):
        self.data_kind_var.trace_add("write", self._on_state_change)
        self.source_kind_var.trace_add("write", self._on_state_change)
        self.template_mode_var.trace_add("write", self._on_state_change)

    def _on_state_change(self, *_args):
        self._refresh_form()

    def _refresh_form(self):
        is_s2p = self.data_kind_var.get() == DATA_KIND_S2P
        if is_s2p:
            if not self.template_mode_var.get():
                self.template_mode_var.set(True)
            self.template_check.state(["disabled"])
            self.s2p_note_var.set("S2P доступен только в шаблонном режиме.")
        else:
            self.template_check.state(["!disabled"])
            self.s2p_note_var.set("")

        if self.source_kind_var.get() == SOURCE_KIND_DIRECTORY:
            kind_text = "Папка с S2P" if is_s2p else "Папка с TXT"
        else:
            kind_text = "S2P-файл" if is_s2p else "TXT-файл"
        self.input_label_var.set(kind_text)

        if self.template_mode_var.get() or is_s2p:
            self.excel_label_var.set("Excel-шаблон")
        else:
            self.excel_label_var.set("Excel-файл")

        visible_keys = set(self.get_visible_field_keys())
        for key in self.field_order:
            row_frame = self.field_rows[key]
            row_frame.grid_remove()

        row_index = 0
        for key in self.field_order:
            if key in visible_keys:
                self.field_rows[key].grid(row=row_index, column=0, sticky="ew", padx=8, pady=4)
                row_index += 1

    def _create_path_row(self, parent, key, label_text, textvariable, browse_command):
        row = ttk.Frame(parent)
        row.columnconfigure(1, weight=1)
        label = ttk.Label(row, textvariable=label_text) if isinstance(label_text, tk.StringVar) else ttk.Label(row, text=label_text)
        label.grid(row=0, column=0, sticky="w", padx=(0, 12))
        entry = ttk.Entry(row, textvariable=textvariable)
        entry.grid(row=0, column=1, sticky="ew")
        button = ttk.Button(row, text="Обзор...", command=browse_command)
        button.grid(row=0, column=2, sticky="e", padx=(8, 0))

        self.field_rows[key] = row
        self.field_order.append(key)
        self.interactive_widgets.extend([entry, button])

    def _create_entry_row(self, parent, key, label_text, textvariable):
        row = ttk.Frame(parent)
        row.columnconfigure(1, weight=1)
        label = ttk.Label(row, text=label_text)
        label.grid(row=0, column=0, sticky="w", padx=(0, 12))
        entry = ttk.Entry(row, textvariable=textvariable)
        entry.grid(row=0, column=1, sticky="ew")

        self.field_rows[key] = row
        self.field_order.append(key)
        self.interactive_widgets.append(entry)

    def _create_check_row(self, parent, key, label_text, variable):
        row = ttk.Frame(parent)
        checkbox = ttk.Checkbutton(row, text=label_text, variable=variable)
        checkbox.grid(row=0, column=0, sticky="w")

        self.field_rows[key] = row
        self.field_order.append(key)
        self.interactive_widgets.append(checkbox)

    def _browse_input_path(self):
        if self.source_kind_var.get() == SOURCE_KIND_DIRECTORY:
            selected = filedialog.askdirectory(parent=self.master, title="Выберите папку с данными")
        else:
            if self.data_kind_var.get() == DATA_KIND_S2P:
                filetypes = [("S2P files", "*.s2p"), ("All files", "*.*")]
            else:
                filetypes = [("TXT files", "*.txt"), ("All files", "*.*")]
            selected = filedialog.askopenfilename(
                parent=self.master,
                title="Выберите файл данных",
                filetypes=filetypes,
            )
        if selected:
            self.input_path_var.set(selected)

    def _browse_excel_path(self):
        selected = filedialog.askopenfilename(
            parent=self.master,
            title="Выберите Excel-файл",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if selected:
            self.excel_path_var.set(selected)

    def _browse_output_path(self):
        initial_name = self.output_path_var.get() or "result.xlsx"
        selected = filedialog.asksaveasfilename(
            parent=self.master,
            title="Выберите выходной файл",
            defaultextension=".xlsx",
            initialfile=os.path.basename(initial_name),
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if selected:
            self.output_path_var.set(selected)

    def _run_in_worker(self, config):
        def reporter(message):
            self.queue.put(("log", message))

        try:
            result = self.runner(config, reporter)
        except Exception as error:  # pragma: no cover - exercised in integration flow
            self.queue.put(("error", str(error), traceback.format_exc()))
            return
        self.queue.put(("done", result))

    def _poll_queue(self):
        finished = False

        while True:
            try:
                item = self.queue.get_nowait()
            except queue.Empty:
                break

            kind = item[0]
            if kind == "log":
                self._append_log(item[1])
                continue

            finished = True
            self._set_busy(False)
            if kind == "done":
                result = item[1]
                summary = (
                    f"Обработка завершена. Файлов: {result.processed_files}. "
                    f"Результатов: {len(result.output_paths)}."
                )
                self._append_log(summary)
                messagebox.showinfo("Готово", summary, parent=self.master)
            elif kind == "error":
                message = item[1]
                details = item[2]
                self._append_log("Ошибка выполнения:")
                self._append_log(details.rstrip())
                messagebox.showerror("Ошибка", message, parent=self.master)

        if not finished and self.worker_thread and self.worker_thread.is_alive():
            self.after(100, self._poll_queue)
        elif not finished:
            self._set_busy(False)

    def _set_busy(self, busy):
        state = "disabled" if busy else "normal"
        for widget in self.interactive_widgets:
            try:
                widget.configure(state=state)
            except tk.TclError:
                pass

        if self.data_kind_var.get() == DATA_KIND_S2P and not busy:
            self.template_check.state(["disabled"])

    def _append_log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_log_text(self, value):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        if value:
            self.log_text.insert("1.0", value)
        self.log_text.configure(state="disabled")


def launch_gui():
    """Create the Tk root window and run the GUI main loop."""
    root = tk.Tk()
    AutoGraphApp(master=root)
    root.mainloop()


def main():
    """Package entrypoint for launching the GUI."""
    launch_gui()


if __name__ == "__main__":
    main()
