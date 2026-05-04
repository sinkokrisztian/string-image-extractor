# Changelog

## 2.1.0 - 2026-05-04

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
