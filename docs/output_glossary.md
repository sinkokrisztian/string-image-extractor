# `report_gui.xlsx` Output Glossary

This document explains every tab and every column in `report_gui.xlsx`: what it means and how it is generated.

Note: the current default runtime is `--ocr-engine ai_validated`.  
Legacy triangulated outputs remain available in expert/debug mode.

---

## 1) Tab: `Image pairs`

Purpose: one row per source-target image pair candidate after filename-level pairing.

- `pair_id`: Sequential pair identifier from `pipeline_report` loop (`0001`, `0002`, ...).
- `source_image_filename`: Source filename (typically English side).
- `source_image_path`: Full source file path.
- `target_language_code`: Selected target language code (for example `hu`).
- `target_language_name`: Language folder display name (for example `Hungarian`).
- `target_image_filename`: Matched target filename (empty if no match).
- `target_image_path`: Full target file path (empty if no match).
- `image_match_method`: Filename pairing method from `match_source_to_target`:
  - `exact_normalized`, `variant_suffix`, `fuzzy`, etc.
- `image_match_confidence`: Confidence assigned by filename matcher.
- `image_match_status`: Pairing status (`OK`, `LOW CONFIDENCE`, `NO MATCH`).
- `alternatives`: Alternative filename candidates (mostly for fuzzy/ambiguous matching).
- `notes`: Pair-level notes and errors (render/OCR/matching notes appended during run).
- `source_rendered_png`: Path to rendered PNG used for AI OCR (if rendering enabled).
- `target_rendered_png`: Same for target side.
- `classic_ocr_status`: Classic OCR stage status (`ok` or error summary).
- `llm_ocr_normalization_status`: LLM normalization stage status.
- `ai_image_ocr_status`: Vision OCR stage status.
- `pair_status`: Pair review status (`ok` or `needs_review`) derived from triangulation review flags.

---

## 2) Tab: `Raw OCR - Classic Lines`

Purpose: line/object-level classic OCR evidence after preprocessing, filtering, and split heuristics.

- `pair_id`: Numeric pair ID.
- `side`: `source` or `target`.
- `image_filename`: Image filename.
- `ocr_engine`: Always `classic` in this sheet.
- `ocr_pass`: OCR pass ID (for example `psm6_gray`, `psm11_bw`).
- `preprocessing`: `gray` or `bw` based on pass.
- `psm`: Tesseract page segmentation mode used in that pass.
- `line_no`: Sequence number of kept/split line objects in that pass.
- `raw_line`: Raw concatenated OCR line text before heavy cleanup.
- `cleaned_line`: Sanitized line text (`clean_line`).
- `avg_tesseract_confidence`: Mean token confidence for the original line group.
- `bbox_left`, `bbox_top`, `bbox_right`, `bbox_bottom`: Bounding box around contributing tokens.
- `kept_or_filtered`: `kept` or `filtered` after noise rules.
- `filter_reason`: Why it was filtered (for example `noise_line`).
- `selected_pass`: `True` if this row came from the best-scoring pass for that image side.
- `line_quality_score`: Pass-level quality score used to choose best pass.

How generated:
- `extract_classic_ocr_evidence` runs multiple pass configs, builds lines from token groups, filters/splits, scores each pass, and marks selected pass.

---

## 3) Tab: `Raw OCR - Classic Tokens`

Purpose: token-level Tesseract evidence used to build line/object candidates.

- `pair_id`: Numeric pair ID.
- `side`: `source` or `target`.
- `image_filename`: Image filename.
- `ocr_engine`: Always `classic`.
- `ocr_pass`: Pass ID (same definitions as above).
- `preprocessing`: `gray` or `bw`.
- `psm`: Tesseract PSM.
- `block_num`, `par_num`, `line_num`, `word_num`: Tesseract structural indices.
- `token_text`: OCR token text.
- `token_confidence`: Tesseract token confidence.
- `left`, `top`, `width`, `height`: Token bounding box geometry.

How generated:
- From `pytesseract.image_to_data` inside `extract_classic_ocr_evidence`.

---

## 4) Tab: `LLM OCR Normalized Objects`

Purpose: structured UI objects produced by LLM normalization over classic OCR evidence.

- `side`: `source` / `target`.
- `image_filename`: Image filename.
- `language_code`: Language code for that side.
- `language_name`: Language name for that side.
- `model`: LLM model used.
- `prompt_version`: Prompt/schema version string.
- `object_id`: Object ID from model output.
- `normalized_text`: Final normalized GUI text.
- `raw_evidence`: Joined evidence snippets used by model.
- `raw_visual_evidence`: Free-form visual evidence note (if model provides it).
- `gui_role`: Role label (`title`, `button`, `value`, etc.).
- `row_group`: Logical row grouping index.
- `screen_area`: Coarse area (`top`, `middle`, `left`, etc.).
- `reading_order`: Reading order index.
- `is_translatable_gui_string`: Whether object is translatable GUI text.
- `correction_type`: Normalization/correction type (`none`, `line_join`, `semantic_ocr_repair`, etc.).
- `ocr_evidence_confidence`: Confidence in OCR evidence quality (0..1).
- `normalization_confidence`: Confidence in normalized text (0..1).
- `needs_review`: Model marks item for review.
- `review_reason`: Reason for review.
- `rationale`: Model rationale.
- `cached`: Whether loaded from cache.
- `warnings`: Warnings returned by model output.

How generated:
- `normalize_ocr_with_llm` -> `call_openai_structured` with `OCRNormalizedResult` schema.

---

## 5) Tab: `AI Image OCR Objects`

Purpose: structured UI objects produced by vision model directly from rendered image.

Columns are the same schema as `LLM OCR Normalized Objects` and mean the same.

How generated:
- `ai_image_ocr` -> `call_openai_structured` with the same `OCRNormalizedResult` schema, but image input included.

---

## 6) Tab: `OCR Triangulation`

Purpose: side-by-side reconciliation of classic vs LLM-normalized vs AI-vision OCR at object/line index level.

- `pair_id`: Numeric pair ID.
- `side`: `source` / `target`.
- `image_filename`: Image filename.
- `classic_raw_line`: Selected classic OCR line at this index (if available).
- `llm_object_id`: LLM object ID at this index.
- `llm_normalized_text`: LLM normalized text.
- `ai_object_id`: AI OCR object ID at this index.
- `ai_normalized_text`: AI OCR text.
- `classic_vs_llm_similarity`: String similarity (accent-insensitive) classic vs LLM.
- `classic_vs_ai_similarity`: Similarity classic vs AI.
- `llm_vs_ai_similarity`: Similarity LLM vs AI.
- `agreement_status`: Agreement class:
  - `all_agree`, `llm_ai_agree`, `classic_llm_agree`, `classic_ai_agree`,
  - `partial`, `disagree`, `classic_only`, `llm_only`, `ai_only`.
- `selected_final_text`: Chosen text forwarded downstream.
- `selected_source`: Which engine won (`classic`, `llm_ocr_normalization`, `ai_image_ocr`, `merged`, `none`).
- `selected_confidence`: Confidence attached to selected text.
- `needs_review`: Review flag derived from agreement and confidence.
- `review_reason`: Reason for review.
- `gui_role`: Role carried from chosen object when available.
- `row_group`: Chosen row group or fallback index.
- `screen_area`: Chosen screen area or fallback.
- `reading_order`: Chosen reading order or fallback.

How generated:
- `triangulate_side` computes similarities, assigns agreement, picks final text/source, and flags review.

---

## 7) Tab: `Final Matches`

Purpose: final source-to-target string mapping result after object matching (LLM object matching / fallback modes) plus confidence fusion.

- `pair_id`: Match row ID (can be expanded like `0001_003`, `0010-001` depending on matching stage output).
- `source_image_filename`: Source image filename.
- `target_image_filename`: Target image filename.
- `source_object_id`: Source selected object ID.
- `target_object_id`: Target selected object ID (empty for unmatched).
- `source_text`: Source GUI string.
- `target_text`: Target GUI string candidate.
- `match_type`: Match type (`semantic_translation`, `same_text`, `value_equivalent`, `uncertain`, `unmatched`).
- `semantic_confidence`: Semantic confidence from matching stage.
- `layout_confidence`: Positional/layout confidence from matching stage.
- `overall_confidence`: Final fused confidence:
  - blend of source/target OCR confidence + semantic + layout + triangulation agreement.
- `status`: Final status (`matched`, `needs_review`, `uncertain`, `unmatched`).
- `rationale`: Matching rationale text.
- `review_reason`: Why review is needed / why unmatched.
- `source_selected_source`: Which OCR source produced selected source text.
- `target_selected_source`: Which OCR source produced selected target text.
- `source_gui_role`: Source role.
- `target_gui_role`: Target role.
- `source_row_group`: Source row group.
- `target_row_group`: Target row group.
- `source_screen_area`: Source screen area.
- `target_screen_area`: Target screen area.
- `triangulation_agreement_score`: Agreement-derived score used in confidence fusion.
- `source_object_lane`: Pre-semantic lane classification for source object (`translatable_gui`, `value_only`, `masked_text`, `map_background`, `decorative_status`, `unknown_review`).
- `target_object_lane`: Same lane classification for target object.
- `raw_ocr_validation_status`: Corroboration result against target-side unfiltered `full_text_by_pass` (`confirmed`, `weakly_confirmed`, `not_confirmed`).
- `raw_ocr_validation_reason`: Reason/details for corroboration status.
- `semantic_check_status`: Pair-level semantic QA classifier (`ok`, `warning`, `mismatch`).
- `semantic_check_confidence`: Confidence of semantic QA classifier.
- `semantic_check_reason`: Short rationale from semantic QA checker.
- `source_guided_conflict`: True when source-context conflict is detected (review-only handling).
- `suggested_target_candidate`: Optional hint candidate shown for review; never a silent overwrite.

How generated:
- Usually via `match_objects_with_llm` + `validate_gui_matches` + confidence/status recalculation.
- Fallback modes: `deterministic_object_matches` or `legacy_llm_text_matches`.

---

## 8) Tab: `Summary`

Purpose: run-level metrics snapshot.

- `metric`: Metric name.
- `value`: Metric value.

Typical entries include:
- run metadata (`run timestamp`, `script version`, models, prompt versions),
- counts (`source image count`, `matched image pair count`),
- error counts (classic/LLM/AI),
- quality counts (`final matched string count`, `needs_review count`, `uncertain count`, `unmatched count`),
- cache hit/miss counts.

How generated:
- `summary` dictionary built in `pipeline_report`, then exported as key-value rows.

---

## Notes on ID Formats

- `Image pairs.pair_id` and `OCR Triangulation.pair_id` are base numeric pair IDs.
- `Final Matches.pair_id` may include sub-identifiers (for example `0001_003`, `0010-001`) depending on matching output granularity.
- For cross-tab debugging, use the leading numeric part of `Final Matches.pair_id` as the base pair key.
