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
    "ai_validated": 3,
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
    "Fast cached AI": {
        "engine": "ai_validated",
        "matching": "llm_objects",
        "pairs": 1,
        "description": "AI OCR with raw OCR corroboration + semantic validation, optimized for cache-first reruns.",
    }
}


BaseTk = TkinterDnD.Tk if TkinterDnD is not None else tk.Tk


class OCRReportGUI(BaseTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("String Image OCR Report")
        self.geometry("1040x760")
        self.minsize(920, 660)
        self.configure(bg="#eef2f6")
        self._companion_path_autosync = True
        self.output_queue: queue.Queue[str] = queue.Queue()
        self.proc: subprocess.Popen[str] | None = None

        cwd = Path.cwd()
        self.root_var = tk.StringVar(value=str(cwd))
        self.target_lang_var = tk.StringVar(value="hu")
        self.target_ocr_var = tk.StringVar(value="hun")
        self.source_ocr_var = tk.StringVar(value="eng")
        self.engine_var = tk.StringVar(value="ai_validated")
        self.ai_model_var = tk.StringVar(value="gpt-4.1-mini")
        self.match_model_var = tk.StringVar(value="gpt-4.1-mini")
        self.detail_var = tk.StringVar(value="high")
        self.image_name_var = tk.StringVar(value="enis01ct009b.eps")
        self.output_var = tk.StringVar(value=str(cwd / "report_gui.xlsx"))
        self.log_var = tk.StringVar(value=str(cwd / "report_gui.log"))
        self.audit_var = tk.StringVar(value=str(cwd / "report_gui_ai_audit.html"))
        self.om_strings_var = tk.StringVar(value=str(cwd / "input" / "OM_strings_EN.xlsx"))
        self.openai_cache_only_var = tk.BooleanVar(value=False)
        self.use_full_cache_var = tk.BooleanVar(value=True)
        self.llm_workers_var = tk.IntVar(value=2)
        self.pair_count_var = tk.IntVar(value=1)
        self.input_tokens_var = tk.IntVar(value=5000)
        self.output_tokens_var = tk.IntVar(value=1200)
        self.cost_var = tk.StringVar(value="")
        self.call_count_var = tk.StringVar(value="")
        self.preset_hint_var = tk.StringVar(value=MODE_PRESETS["Fast cached AI"]["description"])
        self.drop_hint_var = tk.StringVar(value="Drop a project folder here, or drop an EPS image to set the filename filter")
        self.status_var = tk.StringVar(value="Ready")
        self._refresh_default_outputs(cwd)

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
        self.rowconfigure(0, weight=1)

        outer = ttk.Frame(self)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)

        canvas = tk.Canvas(outer, bg="#eef2f6", highlightthickness=0)
        vscroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vscroll.grid(row=0, column=1, sticky="ns")

        content = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure("all", width=e.width))

        def _on_mousewheel(event):
            delta = -1 * int(event.delta / 120) if event.delta else 0
            canvas.yview_scroll(delta, "units")

        self.bind_all("<MouseWheel>", _on_mousewheel)

        content.columnconfigure(0, weight=1)
        content.rowconfigure(4, weight=1)

        header = tk.Frame(content, bg="#172033", height=88)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="String Image OCR Report", style="Title.TLabel").grid(row=0, column=0, sticky="w", padx=22, pady=(16, 0))
        ttk.Label(
            header,
            text="AI-validated OCR, raw evidence corroboration, semantic checks, and review-ready Excel output",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", padx=24, pady=(2, 14))

        quick = ttk.Frame(content, style="Card.TFrame", padding=14)
        quick.grid(row=1, column=0, sticky="ew", padx=16, pady=(14, 10))
        quick.columnconfigure(0, weight=1)
        quick.columnconfigure(2, weight=1)
        ttk.Label(quick, text="Quick Run", style="Card.TLabel", font=("Segoe UI Semibold", 11)).grid(row=0, column=0, sticky="w")
        ttk.Label(quick, text="Mode: AI Validated (only)", style="Card.TLabel", font=("Segoe UI Semibold", 10), foreground="#0f766e").grid(row=0, column=1, sticky="w", padx=12)
        ttk.Label(quick, textvariable=self.preset_hint_var, style="Muted.Card.TLabel", wraplength=420).grid(row=0, column=2, sticky="ew")

        main = ttk.Frame(content)
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
        self._path_row(paths, 5, "OM strings", self.om_strings_var, self._choose_om_strings)

        spend = ttk.Frame(main, style="Card.TFrame", padding=14)
        spend.grid(row=0, column=1, sticky="nsew")
        spend.columnconfigure(0, weight=1)
        ttk.Label(spend, text="Spend Guard", style="Card.TLabel", font=("Segoe UI Semibold", 11)).grid(row=0, column=0, sticky="w")
        ttk.Label(spend, textvariable=self.cost_var, style="Card.TLabel", font=("Segoe UI Semibold", 18), foreground="#0f766e").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Label(spend, textvariable=self.call_count_var, style="Muted.Card.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Progressbar(spend, mode="determinate", maximum=1, value=1).grid(row=3, column=0, sticky="ew", pady=10)
        ttk.Button(spend, text="Recalculate", command=self._recalculate_cost, style="Quiet.TButton").grid(row=4, column=0, sticky="w")

        drop = tk.Frame(content, bg="#dbeafe", highlightbackground="#93c5fd", highlightthickness=1)
        drop.grid(row=3, column=0, sticky="ew", padx=16, pady=(10, 10))
        drop.columnconfigure(0, weight=1)
        tk.Label(drop, textvariable=self.drop_hint_var, bg="#dbeafe", fg="#1e3a8a", font=("Segoe UI Semibold", 10), pady=12).grid(row=0, column=0, sticky="ew")
        drop.bind("<Button-1>", lambda _e: self._choose_root())
        if DND_FILES is not None:
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind("<<Drop>>", self._handle_drop)

        notebook = ttk.Notebook(content)
        notebook.grid(row=4, column=0, sticky="nsew", padx=16, pady=(0, 10))
        run_tab = ttk.Frame(notebook, padding=12)
        advanced_tab = ttk.Frame(notebook, padding=12)
        log_tab = ttk.Frame(notebook, padding=8)
        notebook.add(run_tab, text="Run")
        notebook.add(advanced_tab, text="Advanced")
        notebook.add(log_tab, text="Log")
        self._build_run_tab(run_tab)
        self._build_advanced_tab(advanced_tab)
        self._build_log_tab(log_tab)

        footer = ttk.Frame(content, padding=(16, 0, 16, 14))
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
        ttk.Label(parent, text="OCR engine").grid(row=0, column=2, sticky="w", padx=6, pady=6)
        ttk.Label(parent, text="ai_validated", font=("Segoe UI Semibold", 10), foreground="#0f766e").grid(row=0, column=3, sticky="w", padx=6, pady=6)
        self._entry(parent, "Image filename filter", self.image_name_var, 1, 0)
        ttk.Label(parent, text="Run behavior: fail-fast on AI errors, temporary rendered PNGs, always-on AI audit HTML.").grid(
            row=2, column=0, columnspan=4, sticky="w", padx=6, pady=10
        )
        ttk.Separator(parent).grid(row=4, column=0, columnspan=4, sticky="ew", pady=8)
        self._spin(parent, "Pairs to estimate", self.pair_count_var, 1, 10000, 5, 0)
        self._spin(parent, "Input tokens / call", self.input_tokens_var, 500, 250000, 5, 2)
        self._spin(parent, "Output tokens / call", self.output_tokens_var, 100, 50000, 6, 0)

    def _build_advanced_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(3, weight=1)
        self._combo(parent, "AI OCR model", self.ai_model_var, list(MODEL_PRICES_PER_1M), 0, 0)
        self._combo(parent, "Match model", self.match_model_var, list(MODEL_PRICES_PER_1M), 0, 2)
        self._combo(parent, "Image detail", self.detail_var, ["high", "low", "auto", "original"], 1, 0)
        ttk.Checkbutton(parent, text="Use cache only (no OpenAI calls)", variable=self.openai_cache_only_var, command=self._recalculate_cost).grid(
            row=1, column=2, columnspan=2, sticky="w", padx=6, pady=6
        )
        ttk.Checkbutton(parent, text="Enable full cache restore/write (recommended)", variable=self.use_full_cache_var).grid(
            row=2, column=0, columnspan=2, sticky="w", padx=6, pady=(10, 6)
        )
        self._spin(parent, "Parallel workers", self.llm_workers_var, 1, 4, 2, 2)
        ttk.Label(
            parent,
            text="Tip: keep gpt-4.1-mini for routine batches. Use gpt-4.1 only for hard samples.",
            wraplength=780,
        ).grid(row=4, column=0, columnspan=4, sticky="w", pady=(14, 0))

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
            self.ai_model_var,
            self.match_model_var,
            self.pair_count_var,
            self.input_tokens_var,
            self.output_tokens_var,
            self.openai_cache_only_var,
        ):
            var.trace_add("write", lambda *_: self._recalculate_cost())
        self.target_lang_var.trace_add("write", lambda *_: self._sync_language())
        self.output_var.trace_add("write", lambda *_: self._sync_companion_paths_from_output())

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

    def _choose_om_strings(self) -> None:
        path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.om_strings_var.set(path)

    def _refresh_default_outputs(self, root: Path) -> None:
        out_dir = root / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        self.output_var.set(str(out_dir / "report_gui.xlsx"))
        self.log_var.set(str(out_dir / "report_gui.log"))
        self.audit_var.set(str(out_dir / "report_gui_ai_audit.html"))

    def _sync_companion_paths_from_output(self) -> None:
        if not self._companion_path_autosync:
            return
        out = Path(self.output_var.get().strip() or "")
        if not out:
            return
        if out.suffix.lower() != ".xlsx":
            out = out.with_suffix(".xlsx")
        stem = out.stem
        parent = out.parent if str(out.parent) not in {"", "."} else Path.cwd() / "output"
        parent.mkdir(parents=True, exist_ok=True)
        self.log_var.set(str(parent / f"{stem}.log"))
        self.audit_var.set(str(parent / f"{stem}_ai_audit.html"))

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
        engine = "ai_validated"
        if self.openai_cache_only_var.get():
            self.cost_var.set("$0.0000")
            self.call_count_var.set("0 live OpenAI calls (cache-only rerun)")
            return
        calls = MODE_CALLS_PER_PAIR.get(engine, 0)
        calls += 1  # llm_objects matching
        if engine == "ai_validated":
            # semantic check now runs for all eligible matches (no user cap);
            # use a fixed planning baseline for spend estimate.
            calls += 8
        pair_count = max(0, self._safe_int(self.pair_count_var))
        input_tokens = max(0, self._safe_int(self.input_tokens_var))
        output_tokens = max(0, self._safe_int(self.output_tokens_var))
        model = self.ai_model_var.get()
        price = MODEL_PRICES_PER_1M.get(model, MODEL_PRICES_PER_1M["gpt-4.1-mini"])
        call_count = pair_count * calls
        estimate = call_count * ((input_tokens / 1_000_000) * price["input"] + (output_tokens / 1_000_000) * price["output"])
        self.cost_var.set(f"${estimate:.4f}")
        note = " (includes semantic QA estimate)" if engine == "ai_validated" else ""
        self.call_count_var.set(f"{call_count} estimated OpenAI calls{note}")

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
            "ai_validated",
            "--report-format",
            "multi_sheet",
            "--matching-mode",
            "llm_objects",
            "--ai-ocr-model",
            self.ai_model_var.get(),
            "--openai-model",
            self.match_model_var.get(),
            "--ai-ocr-detail",
            self.detail_var.get(),
            "--llm-workers",
            str(max(1, self._safe_int(self.llm_workers_var))),
            "--output",
            self.output_var.get(),
            "--log-file",
            self.log_var.get(),
            "--om-strings-xlsx",
            self.om_strings_var.get(),
            "--verbose",
        ]
        if self.use_full_cache_var.get():
            cmd.extend([
                "--restore-object-match-from-db-cache",
                "--restore-semantic-check-from-db-cache",
            ])
        if self.openai_cache_only_var.get():
            cmd.append("--openai-cache-only")
        cmd.extend(["--ai-audit-html", self.audit_var.get()])
        if self.image_name_var.get().strip():
            cmd.extend(["--image-name", self.image_name_var.get().strip()])
        return cmd

    def _run(self) -> None:
        if self.proc and self.proc.poll() is None:
            return
        self._recalculate_cost()
        if not self.openai_cache_only_var.get() and not os.environ.get("OPENAI_API_KEY"):
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
