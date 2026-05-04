import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:  # pragma: no cover - optional desktop dependency
    DND_FILES = None
    TkinterDnD = None


MODEL_PRICES_PER_1M = {
    "gpt-4.1-mini": {"input": 0.40, "output": 1.60},
    "gpt-4.1": {"input": 2.00, "output": 8.00},
}

MODE_CALLS_PER_PAIR = {
    "classic": 0,
    "classic_llm": 2,
    "ai": 2,
    "hybrid": 2,
    "ai_validated": 3,
    "triangulated": 4,
}

LANGUAGE_PRESETS = {
    "hu": {"name": "Hungarian", "ocr": "hun"},
    "sk": {"name": "Slovak", "ocr": "slk"},
    "sl": {"name": "Slovenian", "ocr": "slv"},
    "cz": {"name": "Czech", "ocr": "ces"},
    "hr": {"name": "Croatian", "ocr": "hrv"},
    "ro": {"name": "Romanian", "ocr": "ron"},
}

MODE_PRESETS = {
    "Fast review": {
        "engine": "classic",
        "matching": "positional",
        "pairs": 1,
        "description": "No OpenAI calls. Best for checking image pairing and raw Tesseract evidence.",
    },
    "Balanced AI": {
        "engine": "ai_validated",
        "matching": "llm_objects",
        "pairs": 1,
        "description": "AI OCR primary extraction with raw OCR corroboration and semantic validation.",
    },
    "Best quality": {
        "engine": "triangulated",
        "matching": "llm_objects",
        "pairs": 1,
        "description": "Classic OCR, LLM OCR-normalization, AI image OCR, triangulation, and object matching.",
    },
}


BaseTk = TkinterDnD.Tk if TkinterDnD is not None else tk.Tk


class OCRReportGUI(BaseTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("String Image OCR Report")
        self.geometry("1040x760")
        self.minsize(920, 660)
        self.configure(bg="#eef2f6")
        self.output_queue: queue.Queue[str] = queue.Queue()
        self.proc: subprocess.Popen[str] | None = None

        cwd = Path.cwd()
        self.root_var = tk.StringVar(value=str(cwd))
        self.target_lang_var = tk.StringVar(value="hu")
        self.target_ocr_var = tk.StringVar(value="hun")
        self.source_ocr_var = tk.StringVar(value="eng")
        self.engine_var = tk.StringVar(value="ai_validated")
        self.matching_var = tk.StringVar(value="llm_objects")
        self.ai_model_var = tk.StringVar(value="gpt-4.1-mini")
        self.llm_model_var = tk.StringVar(value="gpt-4.1-mini")
        self.match_model_var = tk.StringVar(value="gpt-4.1-mini")
        self.detail_var = tk.StringVar(value="high")
        self.image_name_var = tk.StringVar(value="enis01ct009b.eps")
        self.output_var = tk.StringVar(value=str(cwd / "report_gui.xlsx"))
        self.log_var = tk.StringVar(value=str(cwd / "report_gui.log"))
        self.audit_var = tk.StringVar(value=str(cwd / "report_gui_ai_audit.html"))
        self.write_rendered_var = tk.BooleanVar(value=False)
        self.fallback_var = tk.BooleanVar(value=True)
        self.audit_enabled_var = tk.BooleanVar(value=True)
        self.pair_count_var = tk.IntVar(value=1)
        self.input_tokens_var = tk.IntVar(value=5000)
        self.output_tokens_var = tk.IntVar(value=1200)
        self.cost_var = tk.StringVar(value="")
        self.call_count_var = tk.StringVar(value="")
        self.preset_hint_var = tk.StringVar(value=MODE_PRESETS["Balanced AI"]["description"])
        self.drop_hint_var = tk.StringVar(value="Drop a project folder here, or drop an EPS image to set the filename filter")
        self.status_var = tk.StringVar(value="Ready")

        self._configure_styles()
        self._build_ui()
        self._wire_traces()
        self._recalculate_cost()
        self.after(150, self._drain_output_queue)

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#eef2f6")
        style.configure("Card.TFrame", background="#ffffff", relief="flat")
        style.configure("TLabel", background="#eef2f6", foreground="#172033", font=("Segoe UI", 9))
        style.configure("Card.TLabel", background="#ffffff", foreground="#172033", font=("Segoe UI", 9))
        style.configure("Muted.Card.TLabel", background="#ffffff", foreground="#667085", font=("Segoe UI", 9))
        style.configure("Title.TLabel", background="#172033", foreground="#ffffff", font=("Segoe UI Semibold", 18))
        style.configure("Subtitle.TLabel", background="#172033", foreground="#cbd5e1", font=("Segoe UI", 9))
        style.configure("Accent.TButton", font=("Segoe UI Semibold", 10), padding=(16, 8))
        style.configure("Quiet.TButton", padding=(10, 6))
        style.configure("TNotebook", background="#eef2f6", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(16, 8), font=("Segoe UI", 9))
        style.configure("Horizontal.TProgressbar", troughcolor="#d9e2ec", background="#2563eb")

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        header = tk.Frame(self, bg="#172033", height=88)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="String Image OCR Report", style="Title.TLabel").grid(row=0, column=0, sticky="w", padx=22, pady=(16, 0))
        ttk.Label(
            header,
            text="Triangulated OCR, AI image reading, bilingual GUI matching, and review-ready Excel output",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", padx=24, pady=(2, 14))

        quick = ttk.Frame(self, style="Card.TFrame", padding=14)
        quick.grid(row=1, column=0, sticky="ew", padx=16, pady=(14, 10))
        quick.columnconfigure(1, weight=1)
        quick.columnconfigure(2, weight=1)
        ttk.Label(quick, text="Quick Run", style="Card.TLabel", font=("Segoe UI Semibold", 11)).grid(row=0, column=0, sticky="w")
        preset_frame = ttk.Frame(quick, style="Card.TFrame")
        preset_frame.grid(row=0, column=1, sticky="w", padx=16)
        for idx, name in enumerate(MODE_PRESETS):
            ttk.Button(preset_frame, text=name, command=lambda n=name: self._apply_preset(n), style="Quiet.TButton").grid(row=0, column=idx, padx=(0, 8))
        ttk.Label(quick, textvariable=self.preset_hint_var, style="Muted.Card.TLabel", wraplength=360).grid(row=0, column=2, sticky="ew")

        main = ttk.Frame(self)
        main.grid(row=2, column=0, sticky="ew", padx=16)
        main.columnconfigure(0, weight=2)
        main.columnconfigure(1, weight=1)

        paths = ttk.Frame(main, style="Card.TFrame", padding=14)
        paths.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        paths.columnconfigure(1, weight=1)
        ttk.Label(paths, text="Inputs and Output", style="Card.TLabel", font=("Segoe UI Semibold", 11)).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        self._path_row(paths, 1, "Project", self.root_var, self._choose_root)
        self._path_row(paths, 2, "Workbook", self.output_var, self._choose_output)
        self._path_row(paths, 3, "Log", self.log_var, self._choose_log)
        self._path_row(paths, 4, "AI audit", self.audit_var, self._choose_audit)

        spend = ttk.Frame(main, style="Card.TFrame", padding=14)
        spend.grid(row=0, column=1, sticky="nsew")
        spend.columnconfigure(0, weight=1)
        ttk.Label(spend, text="Spend Guard", style="Card.TLabel", font=("Segoe UI Semibold", 11)).grid(row=0, column=0, sticky="w")
        ttk.Label(spend, textvariable=self.cost_var, style="Card.TLabel", font=("Segoe UI Semibold", 18), foreground="#0f766e").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Label(spend, textvariable=self.call_count_var, style="Muted.Card.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Progressbar(spend, mode="determinate", maximum=1, value=1).grid(row=3, column=0, sticky="ew", pady=10)
        ttk.Button(spend, text="Recalculate", command=self._recalculate_cost, style="Quiet.TButton").grid(row=4, column=0, sticky="w")

        drop = tk.Frame(self, bg="#dbeafe", highlightbackground="#93c5fd", highlightthickness=1)
        drop.grid(row=3, column=0, sticky="ew", padx=16, pady=(10, 10))
        drop.columnconfigure(0, weight=1)
        tk.Label(drop, textvariable=self.drop_hint_var, bg="#dbeafe", fg="#1e3a8a", font=("Segoe UI Semibold", 10), pady=12).grid(row=0, column=0, sticky="ew")
        drop.bind("<Button-1>", lambda _e: self._choose_root())
        if DND_FILES is not None:
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind("<<Drop>>", self._handle_drop)

        notebook = ttk.Notebook(self)
        notebook.grid(row=4, column=0, sticky="nsew", padx=16, pady=(0, 10))
        self.rowconfigure(4, weight=1)
        run_tab = ttk.Frame(notebook, padding=12)
        advanced_tab = ttk.Frame(notebook, padding=12)
        log_tab = ttk.Frame(notebook, padding=8)
        notebook.add(run_tab, text="Run")
        notebook.add(advanced_tab, text="Advanced")
        notebook.add(log_tab, text="Log")
        self._build_run_tab(run_tab)
        self._build_advanced_tab(advanced_tab)
        self._build_log_tab(log_tab)

        footer = ttk.Frame(self, padding=(16, 0, 16, 14))
        footer.grid(row=5, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        self.run_button = ttk.Button(footer, text="Run Report", command=self._run, style="Accent.TButton")
        self.run_button.grid(row=0, column=1, padx=8)
        self.stop_button = ttk.Button(footer, text="Stop", command=self._stop, state="disabled")
        self.stop_button.grid(row=0, column=2)

    def _build_run_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(3, weight=1)
        self._combo(parent, "Target language", self.target_lang_var, list(LANGUAGE_PRESETS), 0, 0)
        self._combo(parent, "OCR engine", self.engine_var, list(MODE_CALLS_PER_PAIR), 0, 2)
        self._combo(parent, "Matching", self.matching_var, ["llm_objects", "positional", "llm_text"], 1, 0)
        self._entry(parent, "Image filename filter", self.image_name_var, 1, 2)
        ttk.Checkbutton(parent, text="Fallback to classic if AI fails", variable=self.fallback_var).grid(row=2, column=0, columnspan=2, sticky="w", pady=10)
        ttk.Checkbutton(parent, text="Write rendered PNGs", variable=self.write_rendered_var).grid(row=2, column=2, columnspan=2, sticky="w", pady=10)
        ttk.Checkbutton(parent, text="Write readable AI request/response audit HTML", variable=self.audit_enabled_var).grid(row=3, column=0, columnspan=4, sticky="w", pady=(0, 10))
        ttk.Separator(parent).grid(row=4, column=0, columnspan=4, sticky="ew", pady=8)
        self._spin(parent, "Pairs to estimate", self.pair_count_var, 1, 10000, 5, 0)
        self._spin(parent, "Input tokens / call", self.input_tokens_var, 500, 250000, 5, 2)
        self._spin(parent, "Output tokens / call", self.output_tokens_var, 100, 50000, 6, 0)

    def _build_advanced_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(3, weight=1)
        self._entry(parent, "Source OCR language", self.source_ocr_var, 0, 0)
        self._entry(parent, "Target OCR language", self.target_ocr_var, 0, 2)
        self._combo(parent, "LLM OCR model", self.llm_model_var, list(MODEL_PRICES_PER_1M), 1, 0)
        self._combo(parent, "AI OCR model", self.ai_model_var, list(MODEL_PRICES_PER_1M), 1, 2)
        self._combo(parent, "Match model", self.match_model_var, list(MODEL_PRICES_PER_1M), 2, 0)
        self._combo(parent, "Image detail", self.detail_var, ["high", "low", "auto", "original"], 2, 2)
        ttk.Label(
            parent,
            text="Tip: keep gpt-4.1-mini for routine batches. Use gpt-4.1 only when a difficult sample needs extra accuracy.",
            wraplength=780,
        ).grid(row=3, column=0, columnspan=4, sticky="w", pady=(14, 0))

    def _build_log_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        self.log_text = tk.Text(parent, height=16, wrap="word", bg="#0f172a", fg="#e2e8f0", insertbackground="#ffffff", relief="flat", font=("Consolas", 9))
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(parent, orient="vertical", command=self.log_text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

    def _path_row(self, parent: ttk.Frame, row: int, label: str, var: tk.StringVar, cmd) -> None:
        ttk.Label(parent, text=label, style="Card.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text="Browse", command=cmd, style="Quiet.TButton").grid(row=row, column=2, padx=(8, 0), pady=4)

    def _combo(self, parent: ttk.Frame, label: str, var: tk.StringVar, values: list[str], row: int, col: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=6, pady=6)
        ttk.Combobox(parent, textvariable=var, values=values, state="readonly").grid(row=row, column=col + 1, sticky="ew", padx=6, pady=6)

    def _entry(self, parent: ttk.Frame, label: str, var: tk.StringVar, row: int, col: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=6, pady=6)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=col + 1, sticky="ew", padx=6, pady=6)

    def _spin(self, parent: ttk.Frame, label: str, var: tk.IntVar, low: int, high: int, row: int, col: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=6, pady=6)
        spin = ttk.Spinbox(parent, from_=low, to=high, textvariable=var, command=self._recalculate_cost, width=14)
        spin.grid(row=row, column=col + 1, sticky="w", padx=6, pady=6)
        spin.bind("<KeyRelease>", lambda _e: self._recalculate_cost())

    def _wire_traces(self) -> None:
        for var in (
            self.engine_var,
            self.matching_var,
            self.ai_model_var,
            self.llm_model_var,
            self.match_model_var,
            self.pair_count_var,
            self.input_tokens_var,
            self.output_tokens_var,
        ):
            var.trace_add("write", lambda *_: self._recalculate_cost())
        self.target_lang_var.trace_add("write", lambda *_: self._sync_language())

    def _apply_preset(self, name: str) -> None:
        preset = MODE_PRESETS[name]
        self.engine_var.set(preset["engine"])
        self.matching_var.set(preset["matching"])
        self.pair_count_var.set(preset["pairs"])
        self.preset_hint_var.set(preset["description"])

    def _sync_language(self) -> None:
        preset = LANGUAGE_PRESETS.get(self.target_lang_var.get())
        if preset:
            self.target_ocr_var.set(preset["ocr"])

    def _choose_root(self) -> None:
        path = filedialog.askdirectory(initialdir=self.root_var.get())
        if path:
            self.root_var.set(path)
            self._refresh_default_outputs(Path(path))

    def _choose_output(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.output_var.set(path)

    def _choose_log(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".log", filetypes=[("Log file", "*.log")])
        if path:
            self.log_var.set(path)

    def _choose_audit(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML report", "*.html")])
        if path:
            self.audit_var.set(path)

    def _refresh_default_outputs(self, root: Path) -> None:
        self.output_var.set(str(root / "report_gui.xlsx"))
        self.log_var.set(str(root / "report_gui.log"))
        self.audit_var.set(str(root / "report_gui_ai_audit.html"))

    def _handle_drop(self, event) -> None:
        raw = event.data.strip()
        items = self.tk.splitlist(raw)
        if not items:
            return
        dropped = Path(items[0])
        if dropped.is_dir():
            self.root_var.set(str(dropped))
            self._refresh_default_outputs(dropped)
            self.drop_hint_var.set(f"Project folder set: {dropped}")
        else:
            self.image_name_var.set(dropped.name)
            project = self._find_project_root(dropped)
            if project:
                self.root_var.set(str(project))
                self._refresh_default_outputs(project)
            self.drop_hint_var.set(f"Image filter set: {dropped.name}")

    def _find_project_root(self, path: Path) -> Path | None:
        for parent in [path.parent, *path.parents]:
            if (parent / "ACROSS_MM_OM9AM08E(EN)").exists() or (parent / "src" / "image_ocr_match_report.py").exists():
                return parent
        return None

    def _safe_int(self, var: tk.IntVar) -> int:
        try:
            return int(var.get())
        except (tk.TclError, ValueError):
            return 0

    def _recalculate_cost(self) -> None:
        calls = MODE_CALLS_PER_PAIR.get(self.engine_var.get(), 0)
        if self.matching_var.get() == "llm_objects" and self.engine_var.get() != "classic":
            calls += 1
        pair_count = max(0, self._safe_int(self.pair_count_var))
        input_tokens = max(0, self._safe_int(self.input_tokens_var))
        output_tokens = max(0, self._safe_int(self.output_tokens_var))
        model = self.ai_model_var.get() if self.engine_var.get() in {"ai", "hybrid", "triangulated", "ai_validated"} else self.llm_model_var.get()
        price = MODEL_PRICES_PER_1M.get(model, MODEL_PRICES_PER_1M["gpt-4.1-mini"])
        call_count = pair_count * calls
        estimate = call_count * ((input_tokens / 1_000_000) * price["input"] + (output_tokens / 1_000_000) * price["output"])
        self.cost_var.set(f"${estimate:.4f}")
        self.call_count_var.set(f"{call_count} estimated OpenAI calls")

    def _build_command(self) -> list[str]:
        script = Path(__file__).with_name("image_ocr_match_report.py")
        cmd = [
            sys.executable,
            str(script),
            "--root",
            self.root_var.get(),
            "--target-lang",
            self.target_lang_var.get(),
            "--target-ocr-lang",
            self.target_ocr_var.get(),
            "--source-ocr-lang",
            self.source_ocr_var.get(),
            "--ocr-engine",
            self.engine_var.get(),
            "--report-format",
            "multi_sheet",
            "--matching-mode",
            self.matching_var.get(),
            "--llm-ocr-normalization-model",
            self.llm_model_var.get(),
            "--ai-ocr-model",
            self.ai_model_var.get(),
            "--openai-model",
            self.match_model_var.get(),
            "--ai-ocr-detail",
            self.detail_var.get(),
            "--output",
            self.output_var.get(),
            "--log-file",
            self.log_var.get(),
            "--verbose",
        ]
        if self.engine_var.get() == "triangulated":
            cmd.append("--expert-debug-mode")
        if self.audit_enabled_var.get():
            cmd.extend(["--ai-audit-html", self.audit_var.get()])
        else:
            cmd.append("--disable-ai-audit-html")
        if self.fallback_var.get():
            cmd.append("--fallback-to-classic")
        if self.write_rendered_var.get():
            cmd.append("--write-rendered-images")
        if self.image_name_var.get().strip():
            cmd.extend(["--image-name", self.image_name_var.get().strip()])
        return cmd

    def _run(self) -> None:
        if self.proc and self.proc.poll() is None:
            return
        self._recalculate_cost()
        if self.engine_var.get() != "classic" and not os.environ.get("OPENAI_API_KEY"):
            if not messagebox.askyesno("No API key", "OPENAI_API_KEY is not set. Run with classic fallback only?"):
                return
        self.log_text.delete("1.0", "end")
        self.status_var.set("Running...")
        self.run_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        threading.Thread(target=self._worker, args=(self._build_command(),), daemon=True).start()

    def _worker(self, cmd: list[str]) -> None:
        self.output_queue.put("> " + " ".join(f'"{x}"' if " " in x else x for x in cmd) + "\n")
        try:
            self.proc = subprocess.Popen(
                cmd,
                cwd=self.root_var.get(),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            assert self.proc.stdout is not None
            for line in self.proc.stdout:
                self.output_queue.put(line)
            code = self.proc.wait()
            self.output_queue.put(f"\nProcess exited with code {code}\n")
        except Exception as exc:
            self.output_queue.put(f"\nFailed to start report: {exc}\n")
        finally:
            self.output_queue.put("__DONE__")

    def _stop(self) -> None:
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()
            self.status_var.set("Stopping...")

    def _drain_output_queue(self) -> None:
        try:
            while True:
                item = self.output_queue.get_nowait()
                if item == "__DONE__":
                    self.run_button.configure(state="normal")
                    self.stop_button.configure(state="disabled")
                    self.status_var.set("Ready")
                else:
                    self.log_text.insert("end", item)
                    self.log_text.see("end")
        except queue.Empty:
            pass
        self.after(150, self._drain_output_queue)


if __name__ == "__main__":
    OCRReportGUI().mainloop()
