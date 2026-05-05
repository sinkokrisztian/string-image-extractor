# String Image Extractor

OCR and bilingual GUI-string matching pipeline for automotive screenshot assets (EPS/bitmap), with classic OCR, LLM normalization, AI vision OCR, triangulation, and Excel reporting.

## Main entry points

- CLI: `src/image_ocr_match_report.py`
- GUI (Windows): `src/ocr_report_gui.py`
- VS Code launcher: `.vscode/launch.json`

## Quick start

1. Create/activate a Python 3.11+ environment.
2. Install dependencies:
   - `pip install -r requirements.txt`
3. Set OpenAI key (for AI modes):
   - PowerShell: `$env:OPENAI_API_KEY=\"...\"`
4. Run GUI:
   - `python src/ocr_report_gui.py`
5. Or run CLI example (triangulated EN->HU):
   - `python src/image_ocr_match_report.py --root . --source-prefix enis --target-lang hu --source-ocr-lang eng --target-ocr-lang hun --ocr-engine triangulated --report-format multi_sheet --matching-mode llm_objects --llm-ocr-normalization-model gpt-4.1-mini --ai-ocr-model gpt-4.1-mini --openai-model gpt-4.1-mini --fallback-to-classic --output report_gui.xlsx --log-file report_gui.log`

## Documentation

- Plan: `docs/codex_triangulated_ai_ocr_implementation_plan.md`
- Simplified AI-validated plan: `docs/codex_ai_ocr_simplified_implementation_plan.md`
- Quality evaluation: `docs/quality_evaluation.md`
- Output glossary: `docs/output_glossary.md`
- Windows installer build: `docs/installer_build.md`
- Change history: `CHANGELOG.md`
