# Changelog

## 2.2.0 - 2026-05-05

- Hardened `ai_validated` matching for split/merged label cases (for example `Edit` + `Favourites`) and improved pair-level consistency in final matches.
- Added duplicate object-id normalization before matching to prevent target reuse artifacts from ambiguous AI OCR object identifiers.
- Improved `ai_validated` identity normalization in final rows so `pair_id`, source image, and target image metadata stay stable across recovery and merge paths.
- Optimized cold-run performance with parallel AI OCR + semantic check execution while keeping SQLite-backed audit/cache writes thread-safe.
- Improved stage-level telemetry for long runs (clearer per-stage OpenAI call start/done logging with elapsed times).
- Added/expanded regression coverage for `ai_validated` matching edge cases and cache-backed execution paths.

## 2.1.0 - 2026-05-04

- Fixed `ai_validated` matching for short translatable warning/status messages, including safety prompts like `BRAKE!` / `FÉK!`.
- Relaxed raw OCR corroboration to allow punctuation-insensitive evidence matches when Tesseract drops punctuation but preserves the visible word.
- Added `--openai-cache-only` for no-spend reruns from cached AI OCR results, plus a GUI checkbox for cache-only runs.
- Added SQLite OpenAI audit/cache database support (`--openai-audit-db`) with run tracking and structured API event logging.
- Added DB-backed stage cache restore/write for AI OCR, LLM OCR-normalization, object matching, and semantic quality checks.
- Added stage-level replay controls: `--restore-object-match-from-db-cache`, `--restore-semantic-check-from-db-cache`, and `--disable-db-stage-cache-write`.
- Tightened `ai_validated` quality filtering to suppress icon-inference-only objects (no raw OCR evidence + icon-like rationale) before final matching.
- Reduced false `needs_review` outcomes for high-confidence, raw-confirmed semantic pairs by downgrading strict semantic mismatches to warnings in trusted cases.
- Added source/target image path columns to `Final Matches` and clickable filename hyperlinks to open image files directly from Excel.
- Fixed Excel multi-sheet export crashes caused by illegal worksheet characters by sanitizing values before writing.
- Added deeper quality analysis documentation with focused noisy/uncertain EN-HU case studies.
- Added report output glossary documenting every sheet and column in `report_gui.xlsx`, including generation logic.
- Moved evaluation and glossary docs under `docs/` for clearer project structure.
- Added repository `.gitignore` for virtualenv, caches, generated reports/logs, and release artifacts.
- Added new default OCR engine mode: `ai_validated` (AI OCR primary + raw OCR corroboration + semantic quality checks).
- Kept legacy `triangulated` flow available behind expert/debug gate (`--expert-debug-mode`), without removing existing logic.
- Added pre-semantic object lanes (`translatable_gui`, `value_only`, `masked_text`, `map_background`, `decorative_status`, `unknown_review`) to suppress noisy/value-only matching.
- Added raw OCR corroboration fields in final outputs: `raw_ocr_validation_status` and `raw_ocr_validation_reason`.
- Added semantic QA fields in final outputs: `semantic_check_status`, `semantic_check_confidence`, `semantic_check_reason`.
- Added source-guided conflict diagnostics: `source_guided_conflict` and `suggested_target_candidate` (hint-only, no silent overwrite).
- Enforced one-to-one target assignment by downgrading reused target matches.
- Added blocking quality gate to prevent accepted matches when semantic status is `mismatch` or raw corroboration is `not_confirmed`.
- Updated Excel formatting to highlight semantic/raw-evidence conflicts in red.
- Added targeted blocking regression tests for Pair A/B/C behaviors and global no-reuse/no-accept gates.
- Installed and verified runtime dependencies (`pandas`, `openpyxl`, `Pillow`, `pytesseract`, `pydantic`, `openai`) for local execution.
- Added `--llm-workers` to parallelize independent source/target AI OCR + LLM normalization steps per pair (when SQLite audit DB is disabled).
- Improved OpenAI logging with stage/title/model/attempt/elapsed-time entries so long-running bottlenecks are traceable by pipeline step.
- Suppressed low-value transport noise from `httpx/openai` logs in favor of descriptive stage-level telemetry.
- Added `OM strings` worksheet generation (from `input/OM_strings_EN.xlsx` by default) as the first sheet in `ai_validated` workbooks, mapped against `Final Matches`.
- Tightened text-only behavior in `ai_validated`: only translatable GUI objects with raw OCR corroboration are admitted to matching.
- Simplified GUI to an `ai_validated`-only workflow and removed legacy mode controls from the main user path.
- Added OM strings input picker to GUI and wired `--om-strings-xlsx` into report generation.
- Updated GUI defaults for speed + caching: cache-restore flags enabled, `--llm-workers 2`, and streamlined advanced options.
- Simplified GUI settings further: removed fallback toggle, rendered-PNG toggle, and semantic-check-count control; pipeline now runs fail-fast and always writes AI audit HTML.
- Added automatic cleanup of temporary rendered PNG batches when persistent PNG export is not enabled.
- Made SQLite OpenAI audit/cache layer thread-safe (`check_same_thread=False` + internal lock), enabling parallel LLM workers together with audit DB logging.
- Parallelized semantic pair quality checks inside `ai_validated` pair processing, reducing cold-run latency while keeping quality gates intact.

## 2.0.0 - 2026-05-02

- Added Phase 1 classic OCR evidence collection while preserving the legacy `extract_text_lines()` path.
- Added Pydantic schemas for image pairs, classic OCR tokens/lines, normalized GUI objects, triangulation rows, and final GUI matches.
- Added multi-sheet Excel reporting with image pairs, raw OCR lines/tokens, LLM-normalized objects, AI image OCR objects, triangulation, final matches, and summary sheets.
- Added `--ocr-engine classic|classic_llm|ai|hybrid|triangulated`, `--report-format`, `--matching-mode`, cache, rendering, fallback, and model-selection CLI options.
- Added OpenAI SDK + Pydantic structured-output paths for LLM OCR-normalization, AI image OCR, and object-based bilingual matching.
- Added EPS-to-PNG rendering helper and image hashing for AI OCR cache keys.
- Added triangulation selection and review flags comparing classic OCR, LLM OCR-normalization, and AI image OCR evidence.
- Added object-based LLM matching with deterministic confidence recalculation and positional fallback.
- Added a Windows-oriented Tkinter GUI launcher with mode/model controls, output/log selection, rendered-image toggle, classic fallback toggle, and editable OpenAI spend estimation.
- Redesigned the GUI with a cleaner Windows 11-style layout: colored header, quick-run presets, prominent spend guard, tabbed advanced settings, dark run log, status footer, and optional drag-and-drop support via `tkinterdnd2`.
- Changed LLM OCR-normalization evidence packaging so full raw OCR text from every Tesseract pass is sent as primary evidence; token/line breakup remains available for audit sheets but is no longer the main LLM input.
- Added user-friendly AI audit HTML generation for multi-sheet runs, tracking LLM OCR-normalization, AI image OCR, object matching requests, cached responses, live responses, and errors without writing API keys or binary image payloads.
- Added unit tests for OCR candidate packaging, triangulation selection, and multi-sheet workbook creation.
- Verified the sample pair `enis01ct009b.eps` / `huis01ct009b.eps` in both classic multi-sheet and triangulated modes.
