# Matching Quality Evaluation (`report_gui.xlsx`)

Date: 2026-05-03  
Scope: Full English -> Hungarian triangulated run (`matching_mode=llm_objects`)

## Executive Summary

The pipeline is producing many strong matches, but overall reliability is not yet production-grade without human review.

- Strong signal: 432 high-confidence `matched` rows (43.5% of final rows), with no low-confidence auto-matches.
- Major concern: very high review burden (`needs_review` + `uncertain` + `unmatched` = 561 rows, 56.5%).
- Core bottleneck: OCR triangulation disagreement is extremely high (`disagree` + `ai_only` dominate), which cascades into matching uncertainty.

Bottom line: the system is promising for assisted localization QA, but not yet robust enough for unattended matching.

## Data Reviewed

- Workbook: `C:\dev\string-image-extractor\report_gui.xlsx`
- Sheets analyzed:
  - `Image pairs`
  - `OCR Triangulation`
  - `Final Matches`
  - `Summary`

## Key Metrics

### 1) Image-pair coverage and pairing quality

- Source image count: **112**
- Matched image pairs: **112**
- Missing pairs: **0**
- Filename pairing quality:
  - `exact_normalized`: 100
  - `variant_suffix`: 12
  - `image_match_status=OK`: 112/112
  - Mean image match confidence: **0.998**

Assessment: file-level pairing is excellent and not the main issue.

### 2) Final match outcomes (`Final Matches`, n=993)

- `matched`: **432** (43.5%)
- `needs_review`: **269** (27.1%)
- `uncertain`: **121** (12.2%)
- `unmatched`: **171** (17.2%)
- Rows with blank `target_text`: **16.0%**

Confidence:
- Mean `overall_confidence`: **0.794** (median **0.895**)
- Mean `semantic_confidence`: **0.761**
- Mean `layout_confidence`: **0.741**
- Mean triangulation agreement score in final rows: **0.521** (low)

### 3) Triangulation stability (`OCR Triangulation`, n=2025)

- `needs_review=True`: **74.96%**
- Agreement status distribution:
  - `ai_only`: **671**
  - `disagree`: **617**
  - `llm_ai_agree`: 221
  - `classic_llm_agree`: 205
  - `partial`: 110
  - `llm_only`: 104
  - `all_agree`: 52
  - `classic_ai_agree`: 30
  - `classic_only`: 15

Selected source dominance:
- `ai_image_ocr`: **1582**
- `merged`: 273
- `llm_ocr_normalization`: 132
- `classic`: 15

Assessment: quality is heavily dependent on AI image OCR; classic/LLM normalization frequently disagree or fail to align, driving review volume.

## Quality Findings

## What is working well

1. Strong deterministic image pairing and full pair coverage.
2. High-quality top matches: best `matched` rows are semantically and positionally clean (`Settings -> Beállítások`, `Navigation settings -> Navigációs beállítások`, etc.).
3. Auto-match thresholding is conservative: no `matched` rows below 0.90 overall confidence.

## Main failure modes

1. **Triangulation disagreement overload**
   - `disagree` + `ai_only` = 1288/2025 triangulation rows (63.6%).
   - 679 rows are `disagree`/`partial` even with `selected_confidence >= 0.85`, suggesting confidence inflation despite cross-engine conflict.

2. **Target object reuse (collision)**
   - 15 duplicate target-object groups (30 rows involved).
   - Reuse appears in repeated labels (`Mobile`, distance/time clusters, long legal blocks), inflating `uncertain`.

3. **Noisy OCR tokens become semantic mismatches**
   - Examples: `Parking -> Pakistan`, `Q Search -> *****`, malformed route rows.
   - Symbol/control/noise artifacts still leak into object text and confuse matching.

4. **Long text blocks collapse into ambiguous one-to-many mappings**
   - Legal disclaimer and route-result blocks often merge/split inconsistently between EN/HU, then get reused or downgraded.

5. **Breadcrumb/navigation chrome handling still inconsistent**
   - Example unmatched: `Previous page -> Általános` with relatively high overall confidence for an unmatched row.

## Overall Quality Rating

- Pairing quality: **A**
- OCR triangulation reliability: **C-**
- Final semantic matching quality (no manual review): **C**
- Final semantic matching quality (with manual review workflow): **B-**

## Recommended Improvements (Priority Order)

1. **Calibrate confidence with disagreement penalties**
   - Reduce/clip `selected_confidence` and final `overall_confidence` when `agreement_status in {disagree, partial, ai_only, llm_only}`.
   - This will better separate trustworthy auto-matches from review-needed rows.

2. **Enforce one-to-one target assignment**
   - Add a post-matching bipartite optimization (or greedy lock with backtracking) to prevent duplicate target reuse unless explicitly justified.

3. **Improve OCR noise gating before matching**
   - Hard-filter control chars, repeated punctuation, masked strings (`*****`), and map-number clutter before object matching.

4. **Special handling for long paragraph/legal blocks**
   - Detect long disclaimer-like objects and match them with dedicated logic (length/keyword anchors), not general short-label matcher.

5. **Strengthen UI-role constraints**
   - Penalize cross-role mappings (e.g., breadcrumb/title/action confusion) unless semantic confidence is extremely strong.

6. **Add targeted regression tests from observed failures**
   - Cases: `Parking/Pakistan`, `Previous page/Általános`, duplicate `Mobile`, long HERE disclaimer reuse.

## Practical Readout for Current Workbook

Use current output as: **human-in-the-loop QA report**, not final auto-approved mapping.  
Safe subset for near-automatic consumption: rows where:
- `status == matched`
- `overall_confidence >= 0.93`
- no `review_reason`
- and no duplicate `target_object_id` in same `pair_id`.

That subset should give high precision while controlling risk.

## Deep Dive: Two Problem Pairs

Below are two concrete pairs with high uncertainty/noise, chosen to represent different failure profiles.

### Pair A: `enis03ct033b.eps -> huis03ct033b.eps` (base pair `59`, `exact_normalized`)

Why selected:
- Very high uncertainty concentration.
- Strong OCR noise artifacts (`**`, `***`) mixed with valid values.
- Duplicate target reuse in final matching.

Observed behavior:
- Final rows (14 total): `matched=3`, `needs_review=2`, `uncertain=8`, `unmatched=1`.
- Typical uncertain rows:
  - `** -> 54 perc`, `** -> 16:45`, `** -> 2.7 km`, `** -> 2.4 km`
  - `54 min -> 54 perc` marked uncertain due to reused target object.
- Triangulation profile:
  - `disagree=16`, `ai_only=7`, `partial=3`, `llm_ai_agree=1`.
  - Many selected texts are numeric distance/time values with adjacent masked tokens.

Interpretation:
- This looks like a route-options UI where OCR captures both meaningful values and nearby visual/noise fragments.
- The matcher can align some numeric strings correctly, but ambiguous masked tokens create one-to-many competition and target reuse.

Quality takeaway:
- Good value extraction potential, but insufficient noise suppression and assignment constraints for stable one-to-one matching.

---

### Pair B: `enis05ct016c.eps -> huis05ct016b.eps` (base pair `103`, `variant_suffix`)

Why selected:
- Uses `variant_suffix` filename fallback (not exact same variant letter).
- High masked-text ambiguity in call log style content.
- Mixed semantic success and unresolved obfuscated rows.

Observed behavior:
- Final rows (14 total): `matched=3`, `needs_review=6`, `uncertain=2`, `unmatched=3`.
- Low-confidence failures cluster around obfuscated strings:
  - `*(****) ****-**** -> (empty)` unmatched
  - `** ******* -> (empty)` unmatched
  - `***** ****** -> de. 9:05 Mobil` uncertain
- Strong but review-gated semantic rows:
  - `Contacts -> Névjegyek`
  - `Keypad -> Billentyűzet`
  - `Message -> Üzenet`
  - `Thursday · Unknown -> Csütörtök Ismeretlen`
- Triangulation profile:
  - `disagree=11`, `ai_only=7`, `llm_ai_agree=3`, `partial=2`, `classic_llm_agree=1`.

Interpretation:
- The variant fallback seems acceptable for image pairing, but OCR disagreement remains high.
- Obfuscated phone-related strings are expected to be hard; semantic UI labels are generally recoverable.

Quality takeaway:
- This pair demonstrates that filename fallback is not the root issue; text-obfuscation and OCR disagreement are.

---

## Focused Remediation for These Two Cases

1. Add an explicit `masked_text` class (`*`, `**`, `****-****`) and keep it out of normal semantic matching lanes.
2. Add one-to-one target locking with conflict resolution to eliminate target reuse side effects.
3. Introduce a route-value mode and a phone-log mode:
   - Route mode: prioritize numeric/time/distance token structures and separators (`•`, `km`, `min`, `perc`).
   - Phone-log mode: prioritize UI labels/date anchors while de-prioritizing obfuscated caller IDs.
4. Penalize `overall_confidence` when triangulation is `disagree`/`ai_only` unless a strict pattern-based equivalence is satisfied.
