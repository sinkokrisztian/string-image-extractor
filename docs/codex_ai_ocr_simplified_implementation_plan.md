# Codex AI-OCR Simplified Implementation Plan

Date: 2026-05-04  
Status: Draft for review before implementation

## 1) Objective

Simplify the extraction/matching pipeline to reduce noisy over-extraction and semantic hallucinations while preserving strong bilingual GUI matching.

### Target architecture

- Primary extractor: **AI OCR (vision model)**
- Validation evidence: **Classic OCR raw full text per pass**
- New default runtime mode: `--ocr-engine ai_validated`
- Keep legacy path available behind expert/debug mode:
  - existing triangulated path remains available (for diagnostics/comparison)
  - not used by default

## 2) Core principles

1. No hardcoded whitelist of allowed labels.
2. No blind translation-based replacement.
3. Source-side context can guide target quality checks, but target output must remain evidence-backed.
4. Prefer warning/review over silent incorrect acceptance.

## 3) Scope changes

## 3.1 Keep

- EPS/image rendering for OCR
- AI OCR extraction (`ai_image_ocr`)
- Full raw OCR collection per pass (`full_text_by_pass`)
- Final bilingual matching stage (with semantic checker)
- Excel multi-sheet reporting

## 3.2 Remove from active flow

- LLM normalization stage (`normalize_ocr_with_llm`) in **default** runtime path
- Triangulation stage (`triangulate_side`) and its agreement-based selection in **default** runtime path

Note: both remain implemented and callable via expert/debug flag(s), not deleted.

## 4) New pipeline design

1. Run classic OCR once per side and retain:
   - raw full text by pass
   - optional token/line evidence for diagnostics only
2. Run AI OCR on source and target images as primary object extractor.
3. Run **raw OCR corroboration** for each AI OCR object:
   - verify object text exists in raw OCR evidence (exact/normalized/fuzzy support)
   - if unsupported -> `needs_review` with `raw_ocr_validation_status=not_confirmed`
   - corroboration source must be **unfiltered `full_text_by_pass`**
   - filtered classic line objects are diagnostics only (never primary corroboration evidence)
4. Run source-target object matching (semantic + layout).
5. Run **final semantic quality check** LLM on matched pairs:
   - classify `ok` / `warning` / `mismatch`
   - if significant divergence -> red warning in final report

## 4.1 Object lanes (pre-semantic routing)

Before semantic matching, each extracted object is assigned to one lane:

- `translatable_gui`
- `value_only`
- `masked_text`
- `map_background`
- `decorative/status`
- `unknown_review`

Lane policy:
- semantic matching for translation quality is primarily run on `translatable_gui`
- non-translatable lanes are suppressed or review-routed as configured
- this prevents distances/times/masked phone strings/map labels from competing with real GUI labels

## 5) Prompt strategy update (AI OCR)

Revise AI OCR prompt to favor:
- translatable GUI labels/instructions/buttons/menus/tabs/titles

And to de-prioritize:
- values only (times/distances/numbers/units),
- background map/place labels,
- masked/noise text fragments,
- purely decorative overlays.

No whitelists in prompt or code.

## 6) Source-guided conflict safeguard (new)

### Motivation example

For `enis01ct013a.eps` vs `huis01ct013a.eps`, target AI OCR may hallucinate `Forgalom` where source has `Weather`.
Raw OCR on target contains weather-like fragment (`járás`), suggesting `Időjárás` family rather than `Forgalom`.

### Rule

When all are true:
1. Source text is high-confidence and clear (e.g., `Weather`),
2. Target AI OCR text is low-confidence or uncertain,
3. Source-target semantic distance is high,
4. Target raw OCR contains partial corroboration aligned with expected meaning,

Then:
- Do **not** silently accept the target AI OCR candidate.
- Mark row as `needs_review` with `source_guided_conflict=true`.
- Attach a `suggested_target_candidate` (e.g., `Időjárás`) only as a hint.
- Keep final extracted target text evidence-backed; no blind overwrite.
- Source-guided conflict handling is strictly review-only (never silent overwrite).

## 7) Reporting changes

## 7.1 New/updated fields in `Final Matches`

- `raw_ocr_validation_status` (`confirmed` | `weakly_confirmed` | `not_confirmed`)
- `raw_ocr_validation_reason`
- `semantic_check_status` (`ok` | `warning` | `mismatch`)
- `semantic_check_confidence`
- `semantic_check_reason`
- `source_guided_conflict` (bool)
- `suggested_target_candidate` (optional hint)

## 7.2 Excel formatting

- Red fill for:
  - `semantic_check_status in {warning, mismatch}`
  - `raw_ocr_validation_status=not_confirmed`
- Amber fill for weak corroboration / review-required rows.

## 8) Quality gates

1. No target label should be auto-accepted if:
   - semantically conflicts with source and
   - raw OCR fails corroboration.
2. One-to-one target assignment is required:
   - a `target_object_id` must not be used by multiple non-unmatched final rows
   - if collision occurs, all but one must be downgraded to `needs_review`/`uncertain`.
2. Pair A (`enis03ct033b` / `huis03ct033b`):
   - expected: no relevant translatable strings emitted.
3. Pair B (`enis05ct016c` / `huis05ct016b`):
   - expected retained labels:
     - `Favourites`
     - `Recents`
     - `Contacts`
     - `Keypad`
     - `Message`
     - `New Message`
   - others should be treated as values/noise/review.
4. Pair C (`enis01ct013a` / `huis01ct013a`):
   - `Forgalom` must not pass as clean match for `Weather`.
   - should surface as source-guided conflict with review warning.
5. Global gate: no target reuse in accepted rows.
6. Global gate: no accepted semantic match when:
   - `semantic_check_status=mismatch`, or
   - `raw_ocr_validation_status=not_confirmed`.

## 9) Regression test set (targeted)

1. `test_pair_a_values_only_suppression`
2. `test_pair_b_label_focus_without_whitelist`
3. `test_pair_c_weather_vs_forgalom_source_guided_conflict`
4. `test_raw_ocr_corroboration_blocks_unconfirmed_ai_text`
5. `test_semantic_warning_red_flag_in_final_sheet`
6. `test_no_target_reuse_in_accepted_rows`
7. `test_no_accept_when_mismatch_or_not_confirmed`

## 10) Implementation phases

### Phase 1: Simplify runtime pipeline
- add new default engine mode `ai_validated`
- keep triangulated path behind expert/debug mode
- keep AI OCR + classic raw OCR evidence collection

### Phase 2: Raw OCR corroboration and reporting fields
- add corroboration checker
- expose validation fields in final output

### Phase 3: Semantic quality checker
- add LLM pair-level meaning check and warning states
- apply red/amber formatting in Excel

### Phase 4: Source-guided conflict handling
- implement conflict rule and suggestion hint fields
- ensure no blind translation overwrite

### Phase 5: Regression tests + docs refresh
- add 3 targeted pair tests + utility tests
- update `docs/output_glossary.md` and `docs/quality_evaluation.md`

## 11) Risks and mitigations

- Risk: under-extraction after stricter filtering.
  - Mitigation: keep weakly-confirmed review lane instead of hard drop.
- Risk: source-guided logic becomes implicit translation.
  - Mitigation: enforce evidence gate + hint-only suggestion behavior.
- Risk: model drift in semantic checker.
  - Mitigation: strict schema outputs + confidence thresholds + regression tests.

## 12) Definition of done

1. Simplified AI-first pipeline runs end-to-end in GUI and CLI.
2. Pair A/B/C behaviors meet quality-gate expectations.
3. Final report surfaces semantic/raw-evidence warnings clearly.
4. No active dependence on triangulation or LLM normalization.
5. Docs and tests updated and passing.
