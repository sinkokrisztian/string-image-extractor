# Codex Implementation Plan: Triangulated OCR + AI Image OCR + LLM OCR-Normalization for Vehicle Infotainment Screenshots

## 0. Purpose

This document is a concrete implementation plan for extending the current `image_ocr_match_report.py` script.

The key requirement is to keep the previous **LLM OCR-normalization** stage as a first-class path **in parallel with** the new **AI image OCR** path.

The final design must support three different sources of OCR/GUI-string evidence:

```text
1. Classic OCR
   Tesseract reads rendered images and produces raw tokens/lines.

2. LLM OCR-normalization
   Tesseract OCR candidates are passed to an LLM adjudication/normalization prompt.
   This stage repairs OCR damage using raw OCR evidence, coordinates, confidence values, and vehicle GUI context.

3. AI image OCR
   A vision-capable AI model reads the rendered screenshot image directly and extracts GUI objects.
```

The strongest production mode should be a **triangulated mode**:

```text
Tesseract raw OCR
    + LLM-normalized OCR from Tesseract evidence
    + AI image OCR from rendered screenshot
    → compare, merge, review-flag, match bilingually
```

Official API references to consider during implementation:

- OpenAI Images and Vision guide: <https://developers.openai.com/api/docs/guides/images-vision>
- OpenAI Structured Outputs guide: <https://developers.openai.com/api/docs/guides/structured-outputs>

---

## 1. Current Script Summary

The current script already performs these steps:

```text
1. Discover project image folder.
2. Discover language folders such as Hungarian(hu), Slovak(sk), Slovenian(sl).
3. Infer English source images by source filename prefix, usually enis.
4. Match English source images to target-language images by filename normalization.
5. Render EPS files through Ghostscript.
6. Run Tesseract OCR with several preprocessing/PSM combinations.
7. Pick a best OCR line set using heuristic line quality scoring.
8. Optionally run LLM matching on final OCR text blocks.
9. Export a single-sheet Excel report.
```

The current weak points are:

```text
1. Non-selected OCR passes are mostly discarded.
2. Raw OCR is not easily reviewable per image pair.
3. OCR normalization is not a separate auditable step.
4. Existing LLM matching prompt is too restrictive and does not use semantic equivalence well enough.
5. Positional line alignment is fragile for multilingual GUI screenshots.
6. The current report does not show enough evidence behind each decision.
```

---

## 2. Required Final Architecture

The revised pipeline should be:

```text
Input EPS/raster images
    ↓
Discover image pairs
    ↓
Rasterize EPS to PNG
    ↓
Classic Tesseract OCR, keeping all raw tokens and lines
    ↓
LLM OCR-normalization from Tesseract candidates
    ↓
AI image OCR directly from rendered screenshot
    ↓
Triangulated OCR comparison and object selection
    ↓
Object-based bilingual EN-target matching
    ↓
Deterministic validation and confidence scoring
    ↓
Multi-sheet Excel report
```

Important: the LLM OCR-normalization stage and the AI image OCR stage are **not the same thing**.

```text
LLM OCR-normalization:
    Input = OCR candidates from Tesseract.
    Role = adjudicate and normalize OCR evidence.
    It does not see the screenshot directly unless optionally provided later.

AI image OCR:
    Input = rendered screenshot image.
    Role = directly read visible GUI strings from the image.
    It does not depend on Tesseract OCR.
```

Both should produce the same kind of normalized GUI-object structure so that they can be compared and merged.

---

## 3. New OCR Engine Modes

Add a richer engine selector.

```bash
--ocr-engine classic|classic_llm|ai|hybrid|triangulated
```

Default:

```bash
classic
```

Meaning:

```text
classic
    Tesseract only.

classic_llm
    Tesseract raw OCR candidates → LLM OCR-normalization.

ai
    Rendered screenshot PNG → AI image OCR directly.

hybrid
    Tesseract raw OCR + AI image OCR.
    No LLM OCR-normalization unless explicitly enabled.

triangulated
    Tesseract raw OCR + LLM OCR-normalization + AI image OCR.
    This is the recommended high-quality mode.
```

Recommended production candidate:

```bash
--ocr-engine triangulated
```

---

## 4. New CLI Options

### 4.1 OCR engine and report format

```bash
--ocr-engine classic|classic_llm|ai|hybrid|triangulated
--report-format legacy|multi_sheet
```

Default:

```bash
--ocr-engine classic
--report-format legacy
```

If `--ocr-engine` is `classic_llm`, `ai`, `hybrid`, or `triangulated`, strongly recommend or automatically default to:

```bash
--report-format multi_sheet
```

unless the user explicitly requests legacy output.

---

### 4.2 Classic OCR options

Keep existing options:

```bash
--target-ocr-lang hun
--source-ocr-lang eng
--tesseract-cmd "C:\Program Files\Tesseract-OCR\tesseract.exe"
--ghostscript-cmd "C:\Program Files\gs\gs10.07.0\bin\gswin64c.exe"
--eps-dpi 600
--use-temp-local-copy
```

Add optional classic OCR debug options:

```bash
--classic-ocr-keep-all-passes
--classic-ocr-min-conf 45
```

In all non-legacy modes, keep all passes automatically.

---

### 4.3 LLM OCR-normalization options

```bash
--use-llm-ocr-normalization
--llm-ocr-normalization-model gpt-4.1
--llm-ocr-normalization-timeout-sec 60
--llm-ocr-normalization-max-retries 2
--llm-ocr-normalization-cache-dir .ocr_llm_norm_cache
--llm-ocr-normalization-no-cache
```

Default behavior:

```text
classic       → disabled
classic_llm   → enabled
ai            → disabled
hybrid        → disabled unless --use-llm-ocr-normalization is passed
triangulated  → enabled
```

Prompt version constant:

```python
LLM_OCR_NORMALIZATION_PROMPT_VERSION = "llm_ocr_normalization_gui_v1"
```

---

### 4.4 AI image OCR options

```bash
--ai-ocr-model gpt-4.1
--ai-ocr-detail low|high|auto|original
--ai-ocr-timeout-sec 60
--ai-ocr-max-retries 2
--ai-ocr-cache-dir .ocr_ai_cache
--ai-ocr-no-cache
```

Default recommendation:

```bash
--ai-ocr-model gpt-4.1
--ai-ocr-detail high
```

Notes:

```text
- Keep model configurable.
- Render EPS to PNG before AI OCR.
- Do not send EPS directly to the API.
- If detail=original is unsupported by the selected model/API, retry with high or auto and log the fallback.
```

Prompt version constant:

```python
AI_IMAGE_OCR_PROMPT_VERSION = "ai_image_ocr_gui_v1"
```

---

### 4.5 OCR comparison and fallback options

```bash
--compare-ocr-engines
--fallback-to-classic
--write-rendered-images
--rendered-image-dir rendered_images
```

Recommended defaults:

```text
--compare-ocr-engines = true for hybrid and triangulated modes
--fallback-to-classic = true for batch production
--write-rendered-images = false unless requested
```

---

### 4.6 Matching options

Keep existing line/LLM options if needed, but add:

```bash
--matching-mode positional|llm_text|llm_objects
```

Default:

```bash
positional
```

Recommended for new AI/triangulated pipeline:

```bash
llm_objects
```

If `--ocr-engine` is `classic_llm`, `ai`, `hybrid`, or `triangulated`, log a recommendation if the user leaves `--matching-mode positional`.

---

## 5. Data Models

Use Pydantic models if available. If avoiding new dependencies, use dataclasses plus explicit validation helpers.

Pydantic is preferred because AI responses should be validated against strict schemas.

---

### 5.1 ImagePair

```python
class ImagePair(BaseModel):
    pair_id: str
    source_image_filename: str
    source_image_path: str
    target_language_code: str
    target_language_name: str
    target_image_filename: str | None
    target_image_path: str | None
    image_match_method: str
    image_match_confidence: float
    image_match_status: str
    alternatives: list[str] = []
    notes: str = ""
```

---

### 5.2 OCRToken

```python
class OCRToken(BaseModel):
    pair_id: str
    side: Literal["source", "target"]
    image_filename: str
    ocr_engine: Literal["classic"]
    ocr_pass: str
    preprocessing: str
    psm: str
    block_num: int
    par_num: int
    line_num: int
    word_num: int
    token_text: str
    token_confidence: float
    left: int
    top: int
    width: int
    height: int
```

---

### 5.3 OCRLine

```python
class OCRLine(BaseModel):
    pair_id: str
    side: Literal["source", "target"]
    image_filename: str
    ocr_engine: Literal["classic"]
    ocr_pass: str
    preprocessing: str
    psm: str
    line_no: int
    raw_line: str
    cleaned_line: str
    avg_tesseract_confidence: float
    bbox_left: int
    bbox_top: int
    bbox_right: int
    bbox_bottom: int
    kept_or_filtered: Literal["kept", "filtered"]
    filter_reason: str = ""
    selected_pass: bool = False
    line_quality_score: float = 0.0
```

---

### 5.4 NormalizedGUIObject

Use the same object structure for both LLM OCR-normalization and AI image OCR so they can be compared.

```python
class NormalizedGUIObject(BaseModel):
    object_id: str
    normalized_text: str
    raw_evidence: list[str] = []
    raw_visual_evidence: str = ""
    gui_role: Literal[
        "title",
        "breadcrumb",
        "menu_label",
        "description",
        "value",
        "button",
        "tab",
        "map_label",
        "status_bar",
        "other",
    ]
    row_group: int
    screen_area: Literal[
        "top",
        "middle",
        "bottom",
        "left",
        "right",
        "center",
        "unknown",
    ]
    reading_order: int
    is_translatable_gui_string: bool
    correction_type: Literal[
        "none",
        "whitespace",
        "accent",
        "character",
        "line_join",
        "line_split",
        "semantic_ocr_repair",
        "vision_read",
        "uncertain",
    ] = "none"
    ocr_evidence_confidence: float = Field(ge=0.0, le=1.0)
    normalization_confidence: float = Field(ge=0.0, le=1.0)
    needs_review: bool
    review_reason: str = ""
    rationale: str = ""
```

---

### 5.5 OCRNormalizedResult

This result is used by both LLM OCR-normalization and AI image OCR.

```python
class OCRNormalizedResult(BaseModel):
    image_filename: str
    language_code: str
    language_name: str
    side: Literal["source", "target"]
    source_engine: Literal["llm_ocr_normalization", "ai_image_ocr", "classic_fallback"]
    model: str
    prompt_version: str
    image_hash: str = ""
    cached: bool = False
    ui_objects: list[NormalizedGUIObject]
    ignored_items: list[dict] = []
    warnings: list[str] = []
```

---

### 5.6 OCRTriangulationRow

```python
class OCRTriangulationRow(BaseModel):
    pair_id: str
    side: Literal["source", "target"]
    image_filename: str
    classic_raw_line: str = ""
    llm_object_id: str = ""
    llm_normalized_text: str = ""
    ai_object_id: str = ""
    ai_normalized_text: str = ""
    classic_vs_llm_similarity: float = 0.0
    classic_vs_ai_similarity: float = 0.0
    llm_vs_ai_similarity: float = 0.0
    agreement_status: Literal[
        "all_agree",
        "llm_ai_agree",
        "classic_llm_agree",
        "classic_ai_agree",
        "partial",
        "disagree",
        "classic_only",
        "llm_only",
        "ai_only",
    ]
    selected_final_text: str
    selected_source: Literal[
        "classic",
        "llm_ocr_normalization",
        "ai_image_ocr",
        "merged",
        "none",
    ]
    selected_confidence: float
    needs_review: bool
    review_reason: str = ""
```

---

### 5.7 GUIMatch

```python
class GUIMatch(BaseModel):
    pair_id: str
    source_image_filename: str
    target_image_filename: str
    source_object_id: str
    target_object_id: str | None
    source_text: str
    target_text: str
    match_type: Literal[
        "semantic_translation",
        "same_text",
        "value_equivalent",
        "unmatched",
        "uncertain",
    ]
    semantic_confidence: float = Field(ge=0.0, le=1.0)
    layout_confidence: float = Field(ge=0.0, le=1.0)
    overall_confidence: float = Field(ge=0.0, le=1.0)
    status: Literal["matched", "unmatched", "uncertain", "needs_review"]
    rationale: str
    review_reason: str = ""
```

---

## 6. Classic OCR Refactor

Refactor existing `extract_text_lines()` so that it can return all OCR evidence, not only selected final lines.

Current OCR passes:

```text
psm6_gray
psm11_gray
psm6_bw
psm11_bw
```

New behavior:

```text
1. Run all OCR passes.
2. Store token-level output for every pass.
3. Store line-level output for every pass.
4. Apply existing noise filters, but record filter decisions and reasons.
5. Score every pass.
6. Mark selected pass.
7. Keep all data for Excel and LLM normalization input.
```

Do not delete the existing cleaning functions immediately. They remain useful for candidate preparation and comparison.

Required output from classic OCR:

```text
classic_tokens: list[OCRToken]
classic_lines: list[OCRLine]
selected_lines: list[str]
selected_block: str
warnings: list[str]
```

---

## 7. EPS Rasterization

Keep Ghostscript as the primary EPS renderer.

Add a reusable function:

```python
def render_image_for_ocr(
    image_path: Path,
    ghostscript_cmd: str | None,
    eps_dpi: int,
    output_dir: Path | None = None,
    keep_rendered: bool = False,
) -> tuple[Path, dict]:
    ...
```

Behavior:

```text
- If input is EPS, render to PNG.
- If input is already raster, optionally copy/normalize to PNG.
- Return rendered PNG path.
- Return metadata: original path, rendered path, DPI, width, height, image hash.
```

The rendered PNG should be the input to AI image OCR.

---

## 8. LLM OCR-Normalization Stage

This is the stage the user explicitly wants to preserve.

It receives OCR candidates produced from Tesseract and produces normalized GUI objects.

### 8.1 Input to LLM OCR-normalization

Build a compact JSON package per image:

```json
{
  "image_filename": "huis01ct009b.eps",
  "language_code": "hu",
  "language_name": "Hungarian",
  "side": "target",
  "domain": "vehicle infotainment / car information display",
  "ocr_passes": [
    {
      "pass_name": "psm6_gray",
      "preprocessing": "gray",
      "psm": "6",
      "selected_pass": true,
      "line_quality_score": 12.4,
      "lines": [
        {
          "line_no": 1,
          "raw_line": "Hanger6",
          "cleaned_line": "Hanger6",
          "avg_confidence": 83.1,
          "bbox": [100, 40, 420, 85],
          "kept_or_filtered": "kept",
          "filter_reason": ""
        }
      ]
    }
  ]
}
```

Include enough OCR evidence for the LLM to normalize responsibly, but avoid sending excessive token data for every image if not needed. For debug mode, allow more verbose packages.

---

### 8.2 LLM OCR-normalization prompt

Prompt version:

```python
LLM_OCR_NORMALIZATION_PROMPT_VERSION = "llm_ocr_normalization_gui_v1"
```

System/developer prompt:

```text
You are an OCR adjudication and normalization engine for vehicle infotainment screenshots.

You will receive OCR candidates produced from a screenshot of a car information display.
The screenshot may contain navigation, audio, phone, vehicle settings, warning settings, display settings, and other GUI text.

Your task:
1. Identify meaningful GUI text objects.
2. Normalize OCR damage when strongly supported.
3. Preserve the actual language of the screenshot.
4. Do not translate.
5. Do not invent text that is not OCR-evidence supported.
6. Keep separate GUI objects separate:
   - page title
   - menu label
   - description/help text
   - setting value
   - button text
   - breadcrumb/navigation chrome
7. Ignore icons, time, signal indicators, slider graphics, decorative marks, and map background labels unless they are meaningful GUI text.

Normalization rules:
- You may correct OCR character errors, missing accents, broken words, and bad line breaks.
- You may join wrapped lines if they form one GUI text object.
- You may split merged OCR lines if they contain separate GUI objects.
- You must not replace the text with a translation.
- You must not silently guess. If uncertain, mark needs_review=true.
- For every normalized string, include the raw OCR evidence that supports it.
- Prefer high precision over high recall.

Controlled context:
- Domain: vehicle infotainment / car information display
- Text types: settings, audio, navigation, driver assistance, warnings, display, phone, connectivity

Return strict JSON only. No Markdown. No commentary outside JSON.
```

User prompt template:

```text
FILENAME: {image_filename}
LANGUAGE CODE: {language_code}
LANGUAGE NAME: {language_name}
SIDE: {source_or_target}
DOMAIN: vehicle infotainment / car information display

TASK:
Normalize OCR candidates into meaningful GUI text objects.
Do not translate.
Use only OCR evidence and controlled vehicle-GUI context.

OCR_CANDIDATES_JSON:
{ocr_candidates_json}
```

---

### 8.3 LLM OCR-normalization output schema

Use the same normalized object schema as AI image OCR, with source engine:

```json
{
  "image_filename": "string",
  "language_code": "string",
  "language_name": "string",
  "side": "source | target",
  "source_engine": "llm_ocr_normalization",
  "ui_objects": [
    {
      "object_id": "string",
      "normalized_text": "string",
      "raw_evidence": ["string"],
      "raw_visual_evidence": "",
      "gui_role": "title | breadcrumb | menu_label | description | value | button | tab | map_label | status_bar | other",
      "row_group": 0,
      "screen_area": "top | middle | bottom | left | right | center | unknown",
      "reading_order": 0,
      "is_translatable_gui_string": true,
      "correction_type": "none | whitespace | accent | character | line_join | line_split | semantic_ocr_repair | uncertain",
      "ocr_evidence_confidence": 0.0,
      "normalization_confidence": 0.0,
      "needs_review": false,
      "review_reason": "string",
      "rationale": "string"
    }
  ],
  "ignored_items": [],
  "warnings": []
}
```

---

## 9. AI Image OCR Stage

AI image OCR reads the rendered screenshot directly.

It is different from LLM OCR-normalization because it uses image input, not only OCR text candidates.

### 9.1 AI image OCR input

Input:

```text
rendered PNG image
image filename
language code
language name
source/target side
vehicle infotainment context
```

EPS must be rendered first. Do not send EPS directly.

---

### 9.2 AI image OCR prompt

Prompt version:

```python
AI_IMAGE_OCR_PROMPT_VERSION = "ai_image_ocr_gui_v1"
```

System/developer prompt:

```text
You are an expert OCR and GUI-string extraction engine for vehicle infotainment screenshots.

The image is from a car information display / infotainment system. It may show navigation, audio, phone, vehicle settings, safety settings, warning settings, display settings, or other in-car UI screens.

Your task is to extract meaningful visible GUI text from the screenshot.

You are not translating. You are reading the screenshot.

Rules:
1. Preserve the screenshot language.
2. Extract translatable GUI text objects:
   - page titles
   - setting labels
   - menu items
   - descriptions/help text
   - values such as On/Off/Low/High/Auto
   - buttons
   - tabs
   - warnings
3. Keep separate GUI objects separate.
4. Join wrapped visual lines only if they form one sentence or one GUI object.
5. Split merged text if it actually contains separate UI objects.
6. Ignore icons, time, signal indicators, decorative marks, and non-text graphics.
7. In navigation/map screens, distinguish real GUI text from map background labels.
8. Do not invent text.
9. Do not translate.
10. You may normalize obvious character/spacing/diacritic errors only if the correction is strongly supported by the visible image and by the vehicle UI context.
11. If uncertain, keep the best reading but set needs_review=true and explain why.
12. Return strict JSON only. No Markdown. No explanation outside JSON.
```

User prompt template:

```text
FILENAME: {image_filename}
LANGUAGE CODE: {language_code}
LANGUAGE NAME: {language_name}
SIDE: {source_or_target}
DOMAIN: vehicle infotainment / car information display

TASK:
Extract visible GUI strings from this screenshot.
Do not translate.
Return structured GUI objects only.
```

---

### 9.3 AI image OCR output schema

Use the same normalized result schema, with source engine:

```json
{
  "image_filename": "string",
  "language_code": "string",
  "language_name": "string",
  "side": "source | target",
  "source_engine": "ai_image_ocr",
  "ui_objects": [
    {
      "object_id": "string",
      "normalized_text": "string",
      "raw_evidence": [],
      "raw_visual_evidence": "string",
      "gui_role": "title | breadcrumb | menu_label | description | value | button | tab | map_label | status_bar | other",
      "row_group": 0,
      "screen_area": "top | middle | bottom | left | right | center | unknown",
      "reading_order": 0,
      "is_translatable_gui_string": true,
      "correction_type": "vision_read | whitespace | accent | character | line_join | line_split | uncertain",
      "ocr_evidence_confidence": 0.0,
      "normalization_confidence": 0.0,
      "needs_review": false,
      "review_reason": "string",
      "rationale": "string"
    }
  ],
  "ignored_items": [],
  "warnings": []
}
```

For AI image OCR:

```text
ocr_evidence_confidence can be interpreted as visual evidence confidence.
normalization_confidence can be the confidence in the final normalized text.
```

---

## 10. OpenAI API Implementation Notes

Use a vision-capable model and image input.

Implementation should use:

```text
- rendered PNG image input
- configurable model name
- configurable detail level
- strict JSON schema / Structured Outputs where possible
```

If using the OpenAI Python SDK, prefer a modern Responses API implementation. If the current environment already uses raw `urllib.request`, implement an adapter so the API transport can be swapped later.

Pseudo-helper:

```python
def image_to_data_url(path: Path) -> str:
    import base64, mimetypes
    mime = mimetypes.guess_type(path.name)[0] or "image/png"
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{data}"
```

Supported image input should be raster format. For this project, use PNG.

---

## 11. Caching

Both LLM OCR-normalization and AI image OCR need caching.

### 11.1 LLM OCR-normalization cache key

Include:

```text
OCR candidates hash
prompt version
model name
language code
image filename
script version if available
```

Suggested file:

```text
.ocr_llm_norm_cache/{hash}__{model}__{lang}__llm_ocr_normalization_gui_v1.json
```

---

### 11.2 AI image OCR cache key

Include:

```text
rendered image hash
prompt version
model name
detail level
language code
image filename
script version if available
```

Suggested file:

```text
.ocr_ai_cache/{hash}__{model}__{detail}__{lang}__ai_image_ocr_gui_v1.json
```

---

### 11.3 Cache behavior

Log:

```text
LLM OCR normalization cache HIT/MISS/WRITE
AI image OCR cache HIT/MISS/WRITE
```

If cached JSON fails schema validation:

```text
- ignore cache
- log warning
- rerun API request
```

---

## 12. Triangulation / Merge Policy

In `triangulated` mode, compare:

```text
classic raw OCR selected lines
LLM-normalized OCR objects
AI image OCR objects
```

### 12.1 Similarity helper

Implement:

```python
def normalize_for_similarity(text: str, accent_insensitive: bool = True) -> str:
    ...

def normalized_text_similarity(a: str, b: str) -> float:
    ...
```

Use:

```text
Unicode NFC/NFKD normalization
casefold
whitespace collapse
optional accent stripping
punctuation-light comparison
SequenceMatcher or rapidfuzz if available
```

---

### 12.2 Agreement statuses

```text
all_agree
llm_ai_agree
classic_llm_agree
classic_ai_agree
partial
disagree
classic_only
llm_only
ai_only
```

Suggested thresholds:

```text
>= 0.85 → agree
0.60–0.84 → partial
< 0.60 → disagree
```

---

### 12.3 Final object selection policy

Use this policy as a starting point:

```text
If LLM-normalized OCR and AI image OCR agree:
    accept merged object with high confidence.

If all three agree:
    accept with very high confidence.

If AI image OCR finds text missing from Tesseract:
    accept if AI confidence is high, but selected_source=ai_image_ocr and needs_review may be true depending on confidence.

If Tesseract evidence is strong and LLM normalization repairs it, and AI image OCR supports the repair:
    accept LLM/AI normalized text.

If LLM normalization suggests a correction but AI image OCR does not support it:
    keep LLM result but mark needs_review unless confidence is very high.

If AI image OCR and LLM-normalized OCR disagree strongly:
    keep both candidates in comparison sheet and mark needs_review.

If only classic OCR exists:
    use classic fallback object with needs_review=true unless confidence is strong and mode allows fallback acceptance.
```

Do not silently discard disagreement.

---

## 13. Object-Based Bilingual Matching

The final bilingual matcher should consume selected normalized GUI objects, not flat OCR lines.

Prompt version:

```python
GUI_OBJECT_MATCH_PROMPT_VERSION = "gui_object_match_v1"
```

### 13.1 Matching prompt

System/developer prompt:

```text
You are a bilingual GUI string alignment specialist for vehicle infotainment screenshots.

You receive structured GUI text objects from two screenshots:
- Source screenshot: English
- Target screenshot: {target_language}

The screenshots show the same car information display screen in different languages.

Your task:
1. Match each English GUI object to the corresponding target-language GUI object.
2. Use semantic equivalence, not line position alone.
3. Use GUI role, row group, screen area, and reading order as supporting evidence.
4. Match title with title, menu label with menu label, description with description, value with value, and button with button.
5. Do not invent target text.
6. Do not translate the output yourself.
7. Use translation knowledge only to decide whether two already-extracted strings correspond.
8. If no reliable counterpart exists, return status "unmatched".
9. If multiple target objects could match, return status "uncertain".
10. Prefer high precision over high recall.

Confidence rubric:
- 0.95–1.00: exact semantic counterpart, same GUI role, same row/group
- 0.85–0.94: strong semantic counterpart, minor OCR/layout uncertainty
- 0.70–0.84: probable counterpart, but role/layout or OCR evidence is imperfect
- 0.40–0.69: weak/uncertain
- below 0.40: do not match

Return strict JSON only.
```

User prompt template:

```text
PAIR ID: {pair_id}
SOURCE IMAGE: {source_image_filename}
TARGET IMAGE: {target_image_filename}
SOURCE LANGUAGE: English
TARGET LANGUAGE CODE: {target_language_code}
TARGET LANGUAGE NAME: {target_language_name}

SOURCE_OBJECTS:
{source_objects_json}

TARGET_OBJECTS:
{target_objects_json}

TASK:
Match source GUI objects to target GUI objects.
```

---

### 13.2 Matching output schema

```json
{
  "pair_id": "string",
  "source_image": "string",
  "target_image": "string",
  "matches": [
    {
      "source_object_id": "string",
      "target_object_id": "string or null",
      "source_text": "string",
      "target_text": "string",
      "match_type": "semantic_translation | same_text | value_equivalent | unmatched | uncertain",
      "semantic_confidence": 0.0,
      "layout_confidence": 0.0,
      "overall_confidence": 0.0,
      "status": "matched | unmatched | uncertain | needs_review",
      "rationale": "string",
      "review_reason": "string"
    }
  ]
}
```

---

## 14. Deterministic Validation

Validate all AI outputs.

### 14.1 Validate normalized OCR objects

Checks:

```text
required fields exist
confidence values are 0..1
normalized_text is not empty unless object is ignored/warning
valid enum values
reading_order is numeric
object IDs are unique per image/result
language fields match expected language
```

If invalid:

```text
retry if API output invalid
fall back if configured
otherwise mark image as failed and continue batch
```

---

### 14.2 Validate GUI matches

Checks:

```text
source_object_id exists
if matched, target_object_id exists
target object not used twice unless explicitly allowed
source_text equals selected source object text
target_text equals selected target object text
unmatched rows have no target object
confidence values are valid
status is valid enum
```

If mismatch:

```text
mark needs_review
log warning
keep row for audit
```

---

## 15. Confidence Policy

The script should calculate or verify a deterministic final confidence.

Suggested formula for final matches:

```text
overall_confidence =
  0.25 * source_object_confidence
+ 0.25 * target_object_confidence
+ 0.25 * semantic_match_confidence
+ 0.15 * layout_confidence
+ 0.10 * triangulation_agreement_score
```

Triangulation agreement score:

```text
all_agree          → 1.00
llm_ai_agree       → 0.95
classic_llm_agree  → 0.85
classic_ai_agree   → 0.80
partial            → 0.65
disagree           → 0.30
single-source only → based on source confidence, but capped
```

Status thresholds:

| Overall confidence | Status |
|---:|---|
| `>= 0.90` | `matched` |
| `0.75–0.89` | `needs_review` |
| `0.50–0.74` | `uncertain` |
| `< 0.50` | `unmatched` or `needs_review` depending on context |

---

## 16. Excel Report Redesign

Add multi-sheet report.

### 16.1 Sheet: `Image pairs`

One row per image pair.

Columns:

```text
pair_id
source_image_filename
source_image_path
target_language_code
target_language_name
target_image_filename
target_image_path
image_match_method
image_match_confidence
image_match_status
alternative_candidate_filenames
source_rendered_png
target_rendered_png
classic_ocr_status
llm_ocr_normalization_status
ai_image_ocr_status
pair_status
notes
```

---

### 16.2 Sheet: `Raw OCR - Classic Lines`

One row per classic OCR line candidate.

Columns:

```text
pair_id
side
image_filename
ocr_pass
preprocessing
psm
line_no
raw_line
cleaned_line
avg_tesseract_confidence
bbox_left
bbox_top
bbox_right
bbox_bottom
kept_or_filtered
filter_reason
selected_pass
line_quality_score
```

---

### 16.3 Sheet: `Raw OCR - Classic Tokens`

One row per Tesseract token.

Columns:

```text
pair_id
side
image_filename
ocr_pass
preprocessing
psm
block_num
par_num
line_num
word_num
token_text
token_confidence
left
top
width
height
```

---

### 16.4 Sheet: `LLM OCR Normalized Objects`

One row per LLM-normalized GUI object.

Columns:

```text
pair_id
side
image_filename
language_code
language_name
model
prompt_version
object_id
normalized_text
raw_evidence
gui_role
row_group
screen_area
reading_order
is_translatable_gui_string
correction_type
ocr_evidence_confidence
normalization_confidence
needs_review
review_reason
rationale
cached
warnings
```

---

### 16.5 Sheet: `AI Image OCR Objects`

One row per AI image OCR GUI object.

Columns:

```text
pair_id
side
image_filename
language_code
language_name
model
prompt_version
object_id
normalized_text
raw_visual_evidence
gui_role
row_group
screen_area
reading_order
is_translatable_gui_string
correction_type
visual_evidence_confidence
normalization_confidence
needs_review
review_reason
rationale
cached
warnings
```

---

### 16.6 Sheet: `OCR Triangulation`

One row per compared/selected OCR object.

Columns:

```text
pair_id
side
image_filename
classic_raw_line
llm_object_id
llm_normalized_text
ai_object_id
ai_normalized_text
classic_vs_llm_similarity
classic_vs_ai_similarity
llm_vs_ai_similarity
agreement_status
selected_final_text
selected_source
selected_confidence
needs_review
review_reason
```

---

### 16.7 Sheet: `Final Matches`

One row per source-target GUI string match.

Columns:

```text
pair_id
source_image_filename
target_image_filename
source_object_id
target_object_id
source_text
target_text
source_selected_source
target_selected_source
source_gui_role
target_gui_role
source_row_group
target_row_group
source_screen_area
target_screen_area
match_type
semantic_confidence
layout_confidence
triangulation_agreement_score
overall_confidence
status
rationale
review_reason
```

---

### 16.8 Sheet: `Summary`

Include:

```text
run timestamp
script version
ocr_engine
matching_mode
classic OCR enabled yes/no
LLM OCR-normalization enabled yes/no
LLM OCR-normalization model
LLM OCR-normalization prompt version
AI image OCR enabled yes/no
AI image OCR model
AI image OCR detail
AI image OCR prompt version
GUI match prompt version
source image count
matched image pair count
missing image pair count
classic OCR error count
LLM OCR-normalization error count
AI image OCR error count
triangulation disagreement count
final matched string count
needs_review count
uncertain count
unmatched count
LLM normalization cache hit count
LLM normalization cache miss count
AI OCR cache hit count
AI OCR cache miss count
```

---

## 17. Formatting Excel Output

Apply:

```text
freeze top row on all sheets
auto-filter on all sheets
wrap text on long-text columns
reasonable column widths
conditional formatting for status/review columns
```

Recommended colors:

```text
matched             → no fill or light green
needs_review        → light yellow
uncertain           → light orange
unmatched/no match  → light red
error               → stronger red or purple
```

Do not over-format so heavily that Excel becomes slow.

---

## 18. Error Handling

### 18.1 Missing API key

If an AI stage is required and no API key exists:

```text
classic_llm → clear error unless fallback allowed
ai          → clear error unless fallback allowed
hybrid      → downgrade to classic if --fallback-to-classic
triangulated → downgrade to classic if --fallback-to-classic, but log that LLM/AI stages were skipped
```

Error message:

```text
AI processing requires an OpenAI API key. Pass --openai-api-key or set OPENAI_API_KEY.
```

---

### 18.2 Invalid AI JSON

If invalid JSON/schema failure:

```text
1. retry up to configured retries
2. save raw response to debug file if possible
3. mark image/stage as failed
4. fallback if enabled
5. continue batch
```

---

### 18.3 AI request failure

For HTTP 429/5xx/timeouts:

```text
retry with exponential backoff
then fallback if enabled
otherwise mark failed and continue
```

For 400-level non-retryable errors:

```text
log clear error
fallback if possible
```

---

## 19. Logging Requirements

Add clear log lines:

```text
OCR engine selected
matching mode selected
Tesseract path
Ghostscript path
rendered image path
rendered image hash
LLM OCR-normalization enabled/model/prompt version
AI image OCR enabled/model/detail/prompt version
cache hit/miss/write for each AI stage
object counts per image and per engine
triangulation agreement/disagreement counts
final match counts
review counts
fallback events
```

Example:

```text
INFO | OCR engine: triangulated
INFO | Matching mode: llm_objects
INFO | Rendered EPS: enis01ct009b.eps -> rendered_images/enis01ct009b.png
INFO | Classic OCR extracted 54 tokens and 10 selected lines from enis01ct009b.eps
INFO | LLM OCR-normalization cache MISS: enis01ct009b.eps
INFO | LLM OCR-normalization extracted 10 GUI objects from enis01ct009b.eps
INFO | AI image OCR cache HIT: enis01ct009b.eps
INFO | AI image OCR extracted 10 GUI objects from enis01ct009b.eps
WARNING | OCR triangulation disagreement: pair=0001 side=target llm='Rendszerhang' ai='Rendszer hang'
INFO | Final matches for pair=0001: matched=10, needs_review=1, unmatched=0
```

---

## 20. Implementation Phases for Codex

### Phase 1 — Refactor classic OCR without behavior change

Goal:

```text
Prepare the codebase for multi-engine OCR while preserving current output.
```

Tasks:

```text
1. Add internal data models for image pairs, OCR tokens, OCR lines.
2. Refactor extract_text_lines() so it can optionally return full OCR pass evidence.
3. Keep old behavior as the default legacy path.
4. Add tests around filename matching and current OCR extraction.
```

Acceptance:

```text
Existing classic command still works and produces equivalent output.
```

---

### Phase 2 — Multi-sheet raw OCR report

Goal:

```text
Allow raw OCR review per image pair.
```

Tasks:

```text
1. Add --report-format legacy|multi_sheet.
2. Add Image pairs, Raw OCR - Classic Lines, Raw OCR - Classic Tokens, Final Matches, Summary.
3. Keep legacy output available.
4. Add Excel formatting.
```

Acceptance:

```text
Classic OCR with --report-format multi_sheet produces reviewable raw OCR lines/tokens.
```

---

### Phase 3 — LLM OCR-normalization

Goal:

```text
Preserve and implement the original OCR adjudication/normalization prompt as a separate engine path.
```

Tasks:

```text
1. Add classic_llm engine mode.
2. Add LLM OCR-normalization prompt and schema.
3. Build OCR candidate JSON from classic OCR results.
4. Add cache.
5. Add validation.
6. Add LLM OCR Normalized Objects sheet.
```

Acceptance:

```text
Running --ocr-engine classic_llm writes both raw Tesseract evidence and LLM-normalized GUI objects.
```

---

### Phase 4 — AI image OCR

Goal:

```text
Add direct screenshot-reading AI OCR.
```

Tasks:

```text
1. Add ai engine mode.
2. Add image rendering helper and image hashing.
3. Add AI image OCR prompt and schema.
4. Add image API call wrapper.
5. Add cache.
6. Add AI Image OCR Objects sheet.
```

Acceptance:

```text
Running --ocr-engine ai sends rendered PNG screenshots to a vision-capable model and extracts structured GUI objects.
```

---

### Phase 5 — Triangulated OCR mode

Goal:

```text
Compare Tesseract, LLM-normalized OCR, and AI image OCR.
```

Tasks:

```text
1. Add triangulated engine mode.
2. Implement similarity helpers.
3. Implement agreement statuses.
4. Implement final object selection policy.
5. Add OCR Triangulation sheet.
```

Acceptance:

```text
Running --ocr-engine triangulated produces raw OCR, LLM-normalized objects, AI image OCR objects, and triangulated selected objects.
```

---

### Phase 6 — Object-based bilingual matching

Goal:

```text
Replace fragile positional alignment with structured GUI-object matching.
```

Tasks:

```text
1. Add --matching-mode llm_objects.
2. Add object-based matching prompt and schema.
3. Validate matching output deterministically.
4. Add Final Matches sheet based on selected GUI objects.
5. Add confidence policy.
```

Acceptance:

```text
Sample pair enis01ct009b.eps / huis01ct009b.eps correctly aligns labels/descriptions/values.
```

---

### Phase 7 — Tests and hardening

Goal:

```text
Make the workflow safe enough for batch production.
```

Tasks:

```text
1. Add tests for filename matching.
2. Add tests for OCR candidate package creation.
3. Add tests for AI schema validation.
4. Add tests for cache key generation.
5. Add tests for fallback behavior.
6. Add tests for triangulation agreement logic.
7. Add tests for object matching validation.
8. Add tests for multi-sheet Excel creation.
```

Acceptance:

```text
pytest passes.
AI/API failures do not crash full batch when fallback is enabled.
Sample pair produces expected final matches.
```

---

## 21. Example Commands

### 21.1 Classic legacy run

```bash
python image_ocr_match_report.py ^
  --root C:\dev\string-image-extractor ^
  --target-lang hu ^
  --target-ocr-lang hun ^
  --source-ocr-lang eng ^
  --ocr-engine classic ^
  --output report_hu_classic.xlsx ^
  --verbose
```

---

### 21.2 Classic run with raw OCR review

```bash
python image_ocr_match_report.py ^
  --root C:\dev\string-image-extractor ^
  --target-lang hu ^
  --target-ocr-lang hun ^
  --source-ocr-lang eng ^
  --ocr-engine classic ^
  --report-format multi_sheet ^
  --output report_hu_classic_raw_review.xlsx ^
  --verbose
```

---

### 21.3 LLM OCR-normalization from Tesseract candidates

```bash
python image_ocr_match_report.py ^
  --root C:\dev\string-image-extractor ^
  --target-lang hu ^
  --target-ocr-lang hun ^
  --source-ocr-lang eng ^
  --ocr-engine classic_llm ^
  --llm-ocr-normalization-model gpt-4.1 ^
  --report-format multi_sheet ^
  --output report_hu_classic_llm_norm.xlsx ^
  --verbose
```

---

### 21.4 AI image OCR only

```bash
python image_ocr_match_report.py ^
  --root C:\dev\string-image-extractor ^
  --target-lang hu ^
  --ocr-engine ai ^
  --ai-ocr-model gpt-4.1 ^
  --ai-ocr-detail high ^
  --matching-mode llm_objects ^
  --report-format multi_sheet ^
  --fallback-to-classic ^
  --output report_hu_ai_image_ocr.xlsx ^
  --verbose
```

---

### 21.5 Triangulated production candidate

```bash
python image_ocr_match_report.py ^
  --root C:\dev\string-image-extractor ^
  --target-lang hu ^
  --target-ocr-lang hun ^
  --source-ocr-lang eng ^
  --ocr-engine triangulated ^
  --llm-ocr-normalization-model gpt-4.1 ^
  --ai-ocr-model gpt-4.1 ^
  --ai-ocr-detail high ^
  --matching-mode llm_objects ^
  --compare-ocr-engines ^
  --fallback-to-classic ^
  --report-format multi_sheet ^
  --use-temp-local-copy ^
  --output report_hu_triangulated_ai_review.xlsx ^
  --log-file report_hu_triangulated_ai_review.log ^
  --verbose
```

---

## 22. Expected Result for Sample Pair

For:

```text
enis01ct009b.eps
huis01ct009b.eps
```

Expected final matches should include:

| English source | Hungarian target |
|---|---|
| Volume | Hangerő |
| System voice | Rendszerhang |
| Voice alerts for Toyota Assistant (also Siri and Google Assistant when available) | Hangfigyelmeztetések a Toyota Assistant funkcióhoz (valamint a Siri és a Google Assistant funkcióhoz, ha elérhető) |
| Driving assist | Vezetéstámogatás |
| Voice alerts for active safety features | Biztonságifunkció-hangfigyelmeztetések |
| Navigation guidance | Navigációs útmutatás |
| Voice alerts for directions and traffic | Irány és közlekedési hangfigyelmeztetések |
| Automatic sound leveling | Automatikus hangszintszabályozás |
| Adjusts volume according to vehicle speed to compensate for increased road noise, wind noise, or other noises while driving | Módosítja a hangerőt a járművön kívüli zaj kompenzálására. |
| Low | Alacsony |

Important:

```text
- Do not achieve this by positional line alignment alone.
- Use structured GUI objects and semantic matching.
- Keep evidence from classic OCR, LLM OCR-normalization, and AI image OCR.
- Review flags must be visible in Excel.
```

---

## 23. Non-Goals

Do not implement unless separately requested:

```text
full external terminology QA
translation memory updates
PDF annotation
GUI desktop application
automatic screenshot editing
automatic rewriting of EPS/source images
```

This task is only about:

```text
OCR extraction
LLM OCR-normalization
AI image OCR
triangulated OCR comparison
bilingual GUI-string matching
Excel reporting
```

---

## 24. Final Acceptance Criteria

Implementation is acceptable when all of these are true:

```text
1. Existing classic OCR workflow still works by default.
2. User can choose --ocr-engine classic, classic_llm, ai, hybrid, or triangulated.
3. EPS files are rasterized to PNG before AI image OCR.
4. Classic OCR preserves all raw tokens and lines in multi-sheet output.
5. LLM OCR-normalization is implemented as its own first-class stage.
6. AI image OCR is implemented as a separate screenshot-reading stage.
7. Triangulated mode compares Tesseract, LLM-normalized OCR, and AI image OCR.
8. All AI outputs are validated with strict schema logic.
9. AI responses are cached.
10. Excel includes raw OCR, LLM OCR-normalized objects, AI image OCR objects, OCR triangulation, final matches, and summary.
11. Object-based matching avoids the previous positional-line mismatch problem.
12. Sample pair enis01ct009b.eps / huis01ct009b.eps produces correct EN-HU matches.
13. Failures are logged and do not crash full batch when fallback is enabled.
14. Every important decision is auditable in the output workbook.
```

---

## 25. Codex Working Instructions

When implementing this plan:

```text
1. Work in small, reviewable patches.
2. Preserve default behavior first.
3. Add tests alongside changes.
4. Keep prompts versioned as constants.
5. Keep schemas close to prompt definitions.
6. Do not silently discard OCR evidence.
7. Do not silently accept AI disagreements.
8. Prefer explicit needs_review flags over overconfident matches.
9. Keep Excel output practical and review-friendly.
10. Log enough information for debugging production runs.
```
