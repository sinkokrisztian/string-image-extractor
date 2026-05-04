import argparse
import base64
import hashlib
import html
import difflib
import json
import logging
import math
import os
import re
import shutil
import subprocess
import tempfile
import time
import unicodedata
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Literal, List, Optional, Sequence, Tuple

import pandas as pd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from PIL import Image, ImageEnhance, ImageOps
from pydantic import BaseModel, Field, ValidationError

try:
    from openai import OpenAI
except Exception:  # pragma: no cover
    OpenAI = None

try:
    import pytesseract
except Exception:  # pragma: no cover
    pytesseract = None


LANG_FOLDER_PATTERN = re.compile(r"^(?P<name>.+)\((?P<code>[a-z]{2})\)$", re.IGNORECASE)
LANG_PREFIX_PATTERN = re.compile(r"^(?P<prefix>[a-z]{2})is(?P<rest>.+)$", re.IGNORECASE)
STRUCTURED_NAME_PATTERN = re.compile(
    r"^(?P<lang>[a-z]{2})is(?P<section>\d{2})ct(?P<item>\d{3})(?P<variant>[a-z])$",
    re.IGNORECASE,
)
SUPPORTED_EXTENSIONS = {".eps", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}

SCRIPT_VERSION = "2.0.1"
LLM_OCR_NORMALIZATION_PROMPT_VERSION = "llm_ocr_normalization_gui_v1"
AI_IMAGE_OCR_PROMPT_VERSION = "ai_image_ocr_gui_v1"
GUI_OBJECT_MATCH_PROMPT_VERSION = "gui_object_match_v1"
DEFAULT_AI_MODEL = "gpt-4.1-mini"

GUI_ROLES = (
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
)
SCREEN_AREAS = ("top", "middle", "bottom", "left", "right", "center", "unknown")
CORRECTION_TYPES = (
    "none",
    "whitespace",
    "accent",
    "character",
    "line_join",
    "line_split",
    "semantic_ocr_repair",
    "vision_read",
    "uncertain",
)


class ImagePair(BaseModel):
    pair_id: str
    source_image_filename: str
    source_image_path: str
    target_language_code: str
    target_language_name: str
    target_image_filename: Optional[str] = None
    target_image_path: Optional[str] = None
    image_match_method: str
    image_match_confidence: float
    image_match_status: str
    alternatives: List[str] = Field(default_factory=list)
    notes: str = ""
    source_rendered_png: str = ""
    target_rendered_png: str = ""
    classic_ocr_status: str = ""
    llm_ocr_normalization_status: str = ""
    ai_image_ocr_status: str = ""
    pair_status: str = ""


class OCRToken(BaseModel):
    pair_id: str
    side: Literal["source", "target"]
    image_filename: str
    ocr_engine: Literal["classic"] = "classic"
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


class OCRLine(BaseModel):
    pair_id: str
    side: Literal["source", "target"]
    image_filename: str
    ocr_engine: Literal["classic"] = "classic"
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


class NormalizedGUIObject(BaseModel):
    object_id: str
    normalized_text: str
    raw_evidence: List[str] = Field(default_factory=list)
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
    ] = "other"
    row_group: int = 0
    screen_area: Literal["top", "middle", "bottom", "left", "right", "center", "unknown"] = "unknown"
    reading_order: int = 0
    is_translatable_gui_string: bool = True
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
    ocr_evidence_confidence: float = Field(default=0.0, ge=0.0, le=1.0)
    normalization_confidence: float = Field(default=0.0, ge=0.0, le=1.0)
    needs_review: bool = False
    review_reason: str = ""
    rationale: str = ""


class IgnoredOCRItem(BaseModel):
    text: str = ""
    reason: str = ""


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
    ui_objects: List[NormalizedGUIObject] = Field(default_factory=list)
    ignored_items: List[IgnoredOCRItem] = Field(default_factory=list)
    warnings: List[str] = Field(default_factory=list)


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
    selected_source: Literal["classic", "llm_ocr_normalization", "ai_image_ocr", "merged", "none"]
    selected_confidence: float
    needs_review: bool
    review_reason: str = ""
    gui_role: str = "other"
    row_group: int = 0
    screen_area: str = "unknown"
    reading_order: int = 0


class GUIMatch(BaseModel):
    pair_id: str
    source_image_filename: str
    target_image_filename: str
    source_object_id: str
    target_object_id: Optional[str] = None
    source_text: str
    target_text: str = ""
    match_type: Literal["semantic_translation", "same_text", "value_equivalent", "unmatched", "uncertain"]
    semantic_confidence: float = Field(ge=0.0, le=1.0)
    layout_confidence: float = Field(ge=0.0, le=1.0)
    overall_confidence: float = Field(ge=0.0, le=1.0)
    status: Literal["matched", "unmatched", "uncertain", "needs_review"]
    rationale: str
    review_reason: str = ""
    source_selected_source: str = ""
    target_selected_source: str = ""
    source_gui_role: str = ""
    target_gui_role: str = ""
    source_row_group: int = 0
    target_row_group: int = 0
    source_screen_area: str = ""
    target_screen_area: str = ""
    triangulation_agreement_score: float = 0.0


class GUIObjectMatchResponse(BaseModel):
    pair_id: str
    source_image: str
    target_image: str
    matches: List[GUIMatch] = Field(default_factory=list)


@dataclass
class ClassicOCRResult:
    classic_tokens: List[OCRToken]
    classic_lines: List[OCRLine]
    selected_lines: List[str]
    selected_block: str
    warnings: List[str]
    full_text_by_pass: Dict[str, str] = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        if self.full_text_by_pass is None:
            self.full_text_by_pass = {}


@dataclass
class PipelineReport:
    image_pairs: List[ImagePair]
    classic_lines: List[OCRLine]
    classic_tokens: List[OCRToken]
    llm_results: List[OCRNormalizedResult]
    ai_results: List[OCRNormalizedResult]
    triangulation_rows: List[OCRTriangulationRow]
    final_matches: List[GUIMatch]
    summary: Dict[str, Any]


@dataclass
class AIAuditEvent:
    stage: str
    title: str
    model: str
    prompt_version: str
    request: Dict[str, Any]
    response: Optional[Dict[str, Any]] = None
    cached: bool = False
    error: str = ""
    timestamp: str = ""


class AIAuditLog:
    def __init__(self) -> None:
        self.events: List[AIAuditEvent] = []

    def add(
        self,
        stage: str,
        title: str,
        model: str,
        prompt_version: str,
        request: Dict[str, Any],
        response: Optional[Dict[str, Any]] = None,
        cached: bool = False,
        error: str = "",
    ) -> None:
        self.events.append(
            AIAuditEvent(
                stage=stage,
                title=title,
                model=model,
                prompt_version=prompt_version,
                request=request,
                response=response,
                cached=cached,
                error=error,
                timestamp=time.strftime("%Y-%m-%d %H:%M:%S"),
            )
        )

    def write_html(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        cards = []
        for idx, event in enumerate(self.events, start=1):
            status = "cached" if event.cached else "error" if event.error else "live"
            req = html.escape(json.dumps(event.request, ensure_ascii=False, indent=2))
            resp_payload = event.response if event.response is not None else {"error": event.error}
            resp = html.escape(json.dumps(resp_payload, ensure_ascii=False, indent=2))
            cards.append(
                f"""
                <section class="card">
                  <div class="card-head">
                    <div>
                      <span class="badge {status}">{html.escape(status.upper())}</span>
                      <h2>{idx}. {html.escape(event.title)}</h2>
                    </div>
                    <div class="meta">{html.escape(event.timestamp)}<br>{html.escape(event.model)}<br>{html.escape(event.prompt_version)}</div>
                  </div>
                  <p class="stage">{html.escape(event.stage)}</p>
                  <details open>
                    <summary>Request</summary>
                    <pre>{req}</pre>
                  </details>
                  <details open>
                    <summary>Response</summary>
                    <pre>{resp}</pre>
                  </details>
                </section>
                """
            )
        body = "\n".join(cards) if cards else '<section class="card"><p>No AI requests were recorded.</p></section>'
        path.write_text(
            f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>AI OCR Audit Report</title>
  <style>
    body {{ margin: 0; font-family: Segoe UI, Arial, sans-serif; background: #eef2f6; color: #172033; }}
    header {{ background: #172033; color: white; padding: 24px 32px; }}
    header p {{ color: #cbd5e1; margin: 6px 0 0; }}
    main {{ max-width: 1180px; margin: 24px auto; padding: 0 18px 32px; }}
    .card {{ background: white; border: 1px solid #d8e0ea; border-radius: 10px; margin-bottom: 18px; padding: 18px; box-shadow: 0 8px 24px rgba(15,23,42,.06); }}
    .card-head {{ display: flex; justify-content: space-between; gap: 18px; align-items: flex-start; }}
    h2 {{ margin: 8px 0 4px; font-size: 18px; }}
    .meta {{ color: #667085; text-align: right; font-size: 13px; line-height: 1.45; }}
    .stage {{ color: #475569; margin-top: 0; }}
    .badge {{ display: inline-block; font-size: 12px; font-weight: 700; border-radius: 999px; padding: 4px 9px; }}
    .badge.live {{ background: #dcfce7; color: #166534; }}
    .badge.cached {{ background: #dbeafe; color: #1e40af; }}
    .badge.error {{ background: #fee2e2; color: #991b1b; }}
    details {{ margin-top: 12px; }}
    summary {{ cursor: pointer; font-weight: 700; color: #334155; }}
    pre {{ white-space: pre-wrap; overflow-wrap: anywhere; background: #0f172a; color: #e2e8f0; border-radius: 8px; padding: 14px; font-size: 12px; line-height: 1.45; }}
  </style>
</head>
<body>
  <header>
    <h1>AI OCR Audit Report</h1>
    <p>{len(self.events)} tracked AI/cache event(s). API keys and binary image payloads are not written here.</p>
  </header>
  <main>{body}</main>
</body>
</html>
""",
            encoding="utf-8",
        )


@dataclass
class MatchResult:
    source_name: str
    target_lang: str
    target_name: Optional[str]
    method: str
    confidence: float
    alternatives: List[str]
    status: str
    notes: str


@dataclass
class LLMGuiMatch:
    source_string: str
    target_string: str
    confidence: float
    status: str
    notes: str
    source_span: str = ""
    target_span: str = ""
    rationale: str = ""


ROAD_LABEL_PATTERN = re.compile(
    r"\b(?:dr|rd|st|ln|ave|blvd|pkwy|trl|trail|hwy|way|ct|pl|hill|park)\b\.?$",
    re.IGNORECASE,
)


def configure_logging(log_file: Path, verbose: bool) -> None:
    log_file.parent.mkdir(parents=True, exist_ok=True)
    handlers = [logging.FileHandler(log_file, encoding="utf-8")]
    if verbose:
        handlers.append(logging.StreamHandler())
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=handlers,
    )


def _candidate_image_roots(root: Path) -> List[Path]:
    candidates: List[Path] = []
    seen = set()

    def add(p: Path) -> None:
        try:
            rp = p.resolve()
        except Exception:
            rp = p
        key = str(rp).lower()
        if key in seen:
            return
        seen.add(key)
        candidates.append(rp)

    add(root)
    if root.name.lower() == "image":
        add(root)
    add(root / "image")
    across_dirs = [p for p in root.glob("*(EN)") if p.is_dir()]
    for d in across_dirs:
        add(d / "image")
    for parent in root.parents:
        add(parent)
        add(parent / "image")
        across_parent_dirs = [p for p in parent.glob("*(EN)") if p.is_dir()]
        for d in across_parent_dirs:
            add(d / "image")
    return candidates


def _find_image_dir_from_root(root: Path) -> Tuple[Optional[Path], List[Path]]:
    checked: List[Path] = []
    for candidate in _candidate_image_roots(root):
        checked.append(candidate)
        if candidate.is_dir() and candidate.name.lower() == "image":
            return candidate, checked
        if candidate.is_dir():
            direct = candidate / "image"
            checked.append(direct)
            if direct.is_dir():
                return direct, checked
            deep = [p for p in candidate.rglob("*") if p.is_dir() and p.name.lower() == "image"]
            if deep:
                return deep[0], checked
    return None, checked


def discover_image_layout(root: Path) -> Tuple[Path, Dict[str, Path]]:
    image_root, checked = _find_image_dir_from_root(root)
    if image_root is None:
        checked_preview = "\n".join(f" - {p}" for p in checked[:12])
        if len(checked) > 12:
            checked_preview += f"\n - ... ({len(checked) - 12} more)"
        raise FileNotFoundError(
            "No folder named 'image' could be discovered.\n"
            f"Provided root: {root}\n"
            "Checked locations:\n"
            f"{checked_preview}\n"
            "Tip: pass --root to your project folder (the one containing ACROSS_* or image)."
        )

    lang_dirs: Dict[str, Path] = {}
    for child in image_root.iterdir():
        if child.is_dir():
            m = LANG_FOLDER_PATTERN.match(child.name)
            if m:
                lang_dirs[m.group("code").lower()] = child
    if not lang_dirs:
        raise FileNotFoundError(f"No language folders like 'Name(xx)' found under {image_root}")
    return image_root, lang_dirs


def list_images(folder: Path, extensions: Sequence[str]) -> List[Path]:
    ext_set = {e.lower() for e in extensions}
    files: List[Path] = []
    for p in folder.iterdir():
        if p.is_file() and p.suffix.lower() in ext_set:
            files.append(p)
    return sorted(files, key=lambda x: x.name.lower())


def language_agnostic_key(filename: str) -> str:
    stem, ext = os.path.splitext(filename.lower())
    m = LANG_PREFIX_PATTERN.match(stem)
    if m:
        stem = f"xxis{m.group('rest')}"
    return f"{stem}{ext}"


def normalize_basename_for_fuzzy(filename: str) -> str:
    stem, _ = os.path.splitext(filename.lower())
    m = LANG_PREFIX_PATTERN.match(stem)
    if m:
        stem = f"xxis{m.group('rest')}"
    return stem


def infer_source_images(image_root: Path, extensions: Sequence[str], source_prefix: Optional[str]) -> List[Path]:
    candidates = list_images(image_root, extensions)
    if source_prefix:
        pref = source_prefix.lower()
        filtered = [p for p in candidates if p.name.lower().startswith(pref)]
        return filtered

    buckets: Dict[str, List[Path]] = {}
    for p in candidates:
        m = LANG_PREFIX_PATTERN.match(p.stem.lower())
        if m:
            key = f"{m.group('prefix')}is"
            buckets.setdefault(key, []).append(p)

    if not buckets:
        return candidates

    source_key = max(buckets, key=lambda k: len(buckets[k]))
    logging.info("Auto-selected source prefix '%s' based on largest matching bucket.", source_key)
    return buckets[source_key]


def build_target_index(target_files: List[Path]) -> Dict[str, Path]:
    idx = {}
    for p in target_files:
        idx[language_agnostic_key(p.name)] = p
    return idx


def fuzzy_candidates(source_name: str, target_files: List[Path], top_n: int = 3) -> List[Tuple[Path, float]]:
    source_norm = normalize_basename_for_fuzzy(source_name)
    scored = []
    for p in target_files:
        ratio = difflib.SequenceMatcher(None, source_norm, normalize_basename_for_fuzzy(p.name)).ratio()
        scored.append((p, ratio))
    scored.sort(key=lambda x: x[1], reverse=True)
    return scored[:top_n]


def parse_structured_name(filename: str) -> Optional[Tuple[int, int, str]]:
    stem = Path(filename).stem.lower()
    m = STRUCTURED_NAME_PATTERN.match(stem)
    if not m:
        return None
    return (int(m.group("section")), int(m.group("item")), m.group("variant"))


def structured_variantless_key(filename: str) -> Optional[Tuple[int, int, str]]:
    parsed = parse_structured_name(filename)
    if not parsed:
        return None
    section, item, _variant = parsed
    return (section, item, Path(filename).suffix.lower())


def build_variantless_target_index(target_files: List[Path]) -> Dict[Tuple[int, int, str], List[Path]]:
    idx: Dict[Tuple[int, int, str], List[Path]] = {}
    for p in target_files:
        key = structured_variantless_key(p.name)
        if key is not None:
            idx.setdefault(key, []).append(p)
    for paths in idx.values():
        paths.sort(key=lambda x: x.name.lower())
    return idx


def structural_distance_score(source_name: str, target_name: str) -> float:
    src = parse_structured_name(source_name)
    tgt = parse_structured_name(target_name)
    if not src or not tgt:
        return 0.0
    sec_delta = abs(src[0] - tgt[0])
    item_delta = abs(src[1] - tgt[1])
    variant_penalty = 0.0 if src[2] == tgt[2] else 0.15
    # Strongly penalize crossing section/item; lightly penalize variant mismatch.
    return sec_delta * 2.0 + item_delta * 0.01 + variant_penalty


def ranked_fuzzy_candidates(source_name: str, target_files: List[Path], top_n: int = 5) -> List[Tuple[Path, float, float]]:
    source_norm = normalize_basename_for_fuzzy(source_name)
    scored: List[Tuple[Path, float, float]] = []
    for p in target_files:
        fuzz = difflib.SequenceMatcher(None, source_norm, normalize_basename_for_fuzzy(p.name)).ratio()
        struct = structural_distance_score(source_name, p.name)
        scored.append((p, fuzz, struct))
    scored.sort(key=lambda x: (-x[1], x[2], x[0].name.lower()))
    return scored[:top_n]


def match_source_to_target(
    source_files: List[Path],
    target_files: List[Path],
    fuzzy_threshold: float,
    allow_fuzzy: bool,
    reject_ambiguous_fuzzy: bool,
) -> List[MatchResult]:
    target_idx = build_target_index(target_files)
    target_variantless_idx = build_variantless_target_index(target_files)
    results: List[MatchResult] = []

    for src in source_files:
        key = language_agnostic_key(src.name)
        exact = target_idx.get(key)
        if exact:
            results.append(
                MatchResult(
                    source_name=src.name,
                    target_lang="",
                    target_name=exact.name,
                    method="exact_normalized",
                    confidence=1.0,
                    alternatives=[],
                    status="OK",
                    notes="",
                )
            )
            continue

        variantless_key = structured_variantless_key(src.name)
        variant_candidates = target_variantless_idx.get(variantless_key, []) if variantless_key else []
        if len(variant_candidates) == 1:
            variant_match = variant_candidates[0]
            results.append(
                MatchResult(
                    source_name=src.name,
                    target_lang="",
                    target_name=variant_match.name,
                    method="variant_suffix",
                    confidence=0.98,
                    alternatives=[],
                    status="OK",
                    notes="Matched by same section/item with different final variant suffix.",
                )
            )
            continue
        if len(variant_candidates) > 1:
            results.append(
                MatchResult(
                    source_name=src.name,
                    target_lang="",
                    target_name=variant_candidates[0].name,
                    method="variant_suffix_ambiguous",
                    confidence=0.75,
                    alternatives=[p.name for p in variant_candidates[1:]],
                    status="LOW CONFIDENCE",
                    notes="Multiple target files share the same section/item with different variant suffixes.",
                )
            )
            continue

        if not allow_fuzzy:
            results.append(
                MatchResult(
                    source_name=src.name,
                    target_lang="",
                    target_name=None,
                    method="exact_only_no_match",
                    confidence=0.0,
                    alternatives=[],
                    status="NO MATCH",
                    notes="No exact normalized match and fuzzy matching disabled.",
                )
            )
            continue

        cands = ranked_fuzzy_candidates(src.name, target_files, top_n=5)
        if not cands:
            results.append(
                MatchResult(
                    source_name=src.name,
                    target_lang="",
                    target_name=None,
                    method="none",
                    confidence=0.0,
                    alternatives=[],
                    status="NO MATCH",
                    notes="No candidates found.",
                )
            )
            continue

        best, score, struct_score = cands[0]
        alternatives = [f"{c.name} ({s:.3f}, d={d:.2f})" for c, s, d in cands[1:4]]
        ambiguous = False
        if len(cands) > 1:
            second = cands[1]
            if abs(score - second[1]) < 0.01 and abs(struct_score - second[2]) < 0.2:
                ambiguous = True
        if score >= fuzzy_threshold:
            if ambiguous and reject_ambiguous_fuzzy:
                results.append(
                    MatchResult(
                        source_name=src.name,
                        target_lang="",
                        target_name=None,
                        method="fuzzy_ambiguous_rejected",
                        confidence=score,
                        alternatives=[f"{best.name} ({score:.3f}, d={struct_score:.2f})"] + alternatives,
                        status="NO MATCH",
                        notes="Fuzzy fallback ambiguous; rejected by policy.",
                    )
                )
                continue
            status = "LOW CONFIDENCE" if (score < 0.9 or ambiguous) else "OK"
            results.append(
                MatchResult(
                    source_name=src.name,
                    target_lang="",
                    target_name=best.name,
                    method="fuzzy",
                    confidence=score,
                    alternatives=alternatives,
                    status=status,
                    notes="Fuzzy fallback used." + (" Ambiguous top candidates." if ambiguous else ""),
                )
            )
        else:
            results.append(
                MatchResult(
                    source_name=src.name,
                    target_lang="",
                    target_name=None,
                    method="fuzzy_below_threshold",
                    confidence=score,
                    alternatives=[f"{best.name} ({score:.3f})"] + alternatives,
                    status="NO MATCH",
                    notes="Best fuzzy score below threshold.",
                )
            )
    return results


def maybe_local_copy(path: Path, temp_root: Optional[Path]) -> Path:
    if temp_root is None:
        return path
    rel = path.name
    out = temp_root / rel
    if not out.exists():
        shutil.copy2(path, out)
    return out


def preprocess_image(img: Image.Image) -> Image.Image:
    if img.mode not in ("L", "RGB"):
        img = img.convert("RGB")
    gray = ImageOps.grayscale(img)
    gray = ImageOps.autocontrast(gray)
    gray = ImageEnhance.Contrast(gray).enhance(1.6)
    gray = ImageEnhance.Sharpness(gray).enhance(1.4)
    w, h = gray.size
    if max(w, h) < 1800:
        scale = 2
        gray = gray.resize((w * scale, h * scale), Image.Resampling.LANCZOS)
    return gray


def make_binary_text_image(gray: Image.Image) -> Image.Image:
    # Robust text-oriented binarization to suppress gradient UI backgrounds.
    hist = gray.histogram()
    total = sum(hist)
    csum = 0
    median_level = 128
    for i, v in enumerate(hist):
        csum += v
        if csum >= total / 2:
            median_level = i
            break
    # Slightly adaptive threshold around median.
    thr = max(120, min(190, int(median_level + 18)))
    bw = gray.point(lambda p: 255 if p > thr else 0, mode="1")
    return bw.convert("L")


def crop_for_ui_text(img: Image.Image) -> Image.Image:
    # Remove icon-heavy margins/status area while keeping breadcrumb/title/content.
    w, h = img.size
    left = int(w * 0.06)
    top = int(h * 0.02)
    right = int(w * 0.98)
    bottom = int(h * 0.97)
    return img.crop((left, top, right, bottom))


def resolve_executable(explicit_path: Optional[str], candidates: Sequence[str]) -> Optional[str]:
    if explicit_path:
        p = Path(explicit_path)
        if p.exists():
            return str(p)
    for name in candidates:
        found = shutil.which(name)
        if found:
            return found
    return None


def rasterize_eps_with_ghostscript(eps_path: Path, out_png: Path, ghostscript_cmd: str, dpi: int) -> Tuple[bool, Optional[str]]:
    cmd = [
        ghostscript_cmd,
        "-dNOPAUSE",
        "-dBATCH",
        "-dSAFER",
        "-dEPSCrop",
        f"-r{dpi}",
        "-sDEVICE=png16m",
        f"-sOutputFile={str(out_png)}",
        str(eps_path),
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
        if proc.returncode != 0:
            msg = (proc.stderr or proc.stdout or "Ghostscript failed").strip()
            return False, msg
        if not out_png.exists():
            return False, "Ghostscript succeeded but output PNG was not created."
        return True, None
    except Exception as exc:
        return False, str(exc)


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def render_image_for_ocr(
    image_path: Path,
    ghostscript_cmd: Optional[str],
    eps_dpi: int,
    output_dir: Optional[Path] = None,
    keep_rendered: bool = False,
) -> Tuple[Optional[Path], Dict[str, Any]]:
    output_root = output_dir or Path(tempfile.mkdtemp(prefix="ocr_render_"))
    output_root.mkdir(parents=True, exist_ok=True)
    out_png = output_root / f"{image_path.stem}.png"
    meta: Dict[str, Any] = {
        "original_path": str(image_path),
        "rendered_path": str(out_png),
        "dpi": eps_dpi,
        "width": 0,
        "height": 0,
        "image_hash": "",
        "kept": keep_rendered,
    }
    ext = image_path.suffix.lower()
    try:
        if ext == ".eps":
            if not ghostscript_cmd:
                return None, {**meta, "error": "Ghostscript is required to render EPS for OCR."}
            ok, err = rasterize_eps_with_ghostscript(image_path, out_png, ghostscript_cmd, eps_dpi)
            if not ok:
                return None, {**meta, "error": err or "Ghostscript EPS rasterization failed."}
        elif ext in {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}:
            with Image.open(image_path) as img:
                img.convert("RGB").save(out_png)
        else:
            return None, {**meta, "error": f"Unsupported image extension: {ext}"}

        with Image.open(out_png) as rendered:
            meta["width"], meta["height"] = rendered.size
        meta["image_hash"] = file_sha256(out_png)
        return out_png, meta
    except Exception as exc:
        return None, {**meta, "error": str(exc)}


def open_image_for_ocr(image_path: Path, ghostscript_cmd: Optional[str], eps_dpi: int) -> Tuple[Optional[Image.Image], Optional[str]]:
    ext = image_path.suffix.lower()
    if ext != ".eps":
        try:
            return Image.open(image_path), None
        except Exception as exc:
            return None, f"Open failed: {exc}"

    if ghostscript_cmd:
        tmp_png = Path(tempfile.gettempdir()) / f"ocr_eps_{next(tempfile._get_candidate_names())}.png"
        ok, err = rasterize_eps_with_ghostscript(image_path, tmp_png, ghostscript_cmd, eps_dpi)
        if ok:
            try:
                img = Image.open(tmp_png)
                return img, None
            except Exception as exc:
                return None, f"Rendered PNG open failed: {exc}"
        return None, f"Ghostscript EPS rasterization failed: {err}"

    try:
        with Image.open(image_path) as probe:
            probe.load(scale=8)
            return probe.copy(), None
    except Exception as exc:
        return None, f"EPS open failed (no Ghostscript cmd): {exc}"


def extract_text(
    image_path: Path,
    tesseract_lang: str,
    tesseract_cmd: str,
    ghostscript_cmd: Optional[str],
    eps_dpi: int,
) -> Tuple[str, Optional[str]]:
    if pytesseract is None:
        return "", "pytesseract is not installed."
    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    try:
        img, open_err = open_image_for_ocr(image_path, ghostscript_cmd, eps_dpi)
        if img is None:
            return "", open_err or "Image open failed."
        with img:
            proc = preprocess_image(crop_for_ui_text(img))
            text = pytesseract.image_to_string(proc, lang=tesseract_lang, config="--psm 6")
        return text.strip(), None
    except Exception as exc:
        return "", f"OCR failed: {exc}"


def is_noise_line(line: str, avg_conf: float) -> bool:
    s = re.sub(r"\s+", " ", line).strip()
    if not s:
        return True
    for pat in NAV_CHROME_PATTERNS:
        if re.fullmatch(pat, s, flags=re.IGNORECASE):
            return True
    # Keep potentially valid long words even at lower confidence.
    if avg_conf < 40:
        if not (" " not in s and s.isalpha() and len(s) >= 8):
            return True
    if re.fullmatch(r"[0-9:./-]+", s):
        return True
    if re.fullmatch(r"[0-9A-Z]{1,4}", s):
        return True
    if re.fullmatch(r"[^\w]+", s, flags=re.UNICODE):
        return True
    if re.search(r"\b\d{1,2}:\d{2}\b", s):
        return True
    if re.search(r"\b\d{1,2}:\d\b", s):
        return True
    if ROAD_LABEL_PATTERN.search(s):
        return True
    if " " not in s and any(ch.islower() for ch in s) and any(ch.isupper() for ch in s):
        # Likely map label/camelcase place names like BenallaDr.
        return True
    if re.search(r"\([A-Za-z0-9]\)$", s) and len(s) <= 10:
        return True
    # Typical OCR garbage from icon/shape regions: short lower-case single tokens.
    if " " not in s and s.isalpha() and s.islower() and len(s) <= 5 and avg_conf < 80:
        return True
    # Reject short standalone alpha artifacts unless they are known UI values/connectors.
    if " " not in s and s.isalpha() and len(s) <= 2:
        allowed_short = {"on", "off", "be", "ki", "és", "ha", "a"}
        if s.lower() not in allowed_short:
            return True
    letters = sum(ch.isalpha() for ch in s)
    digits = sum(ch.isdigit() for ch in s)
    punct = sum((not ch.isalnum() and not ch.isspace()) for ch in s)
    if letters < 2:
        return True
    if len(s) <= 6 and letters <= 4 and s.lower() not in VALUE_WORDS:
        return True
    if re.fullmatch(r"[A-Za-z]{1,2}\s+[A-Za-z]{1,2}", s) and s.lower() not in VALUE_WORDS:
        return True
    if letters > 0 and digits > letters * 1.2:
        return True
    if punct > max(3, letters):
        return True
    # repeated symbol-ish patterns from graphics
    if re.search(r"(2K|KKK|OOO|III|:::|;;;|---|___)", s, flags=re.IGNORECASE):
        return True
    tokens = [t for t in re.split(r"\s+", s) if t]
    if tokens:
        short_tokens = sum(1 for t in tokens if len(t) <= 2)
        if len(tokens) >= 3 and short_tokens / len(tokens) >= 0.6:
            return True
        if len(tokens) <= 2 and all(len(t) <= 3 for t in tokens) and s.lower() not in VALUE_WORDS:
            return True
    letters_only = "".join(ch.lower() for ch in s if ch.isalpha())
    if len(letters_only) >= 6:
        vowels = sum(ch in "aeiouáéíóöőúüű" for ch in letters_only)
        if vowels / max(1, len(letters_only)) < 0.18:
            return True
    return False


def clean_line(line: str) -> str:
    s = line.replace("|", " ").replace("—", " ").replace("_", " ")
    s = re.sub(r"[^\w\s:/.,%()+-]", " ", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_for_llm_guard(text: str) -> str:
    s = unicodedata.normalize("NFC", "" if text is None else str(text))
    return re.sub(r"\s+", " ", s).strip()


def remove_accents(text: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKD", text) if not unicodedata.combining(ch))


def consolidate_ocr_lines(lines: Sequence[str]) -> str:
    cleaned: List[str] = []
    seen = set()
    for line in lines:
        s = normalize_for_llm_guard(clean_line(line))
        if not s or s in seen:
            continue
        cleaned.append(s)
        seen.add(s)
    return "\n".join(cleaned)


def normalized_contains(block: str, value: str) -> bool:
    needle = normalize_for_llm_guard(value)
    if not needle:
        return False
    haystack = normalize_for_llm_guard(block)
    return needle in haystack


VALUE_WORDS = {
    "auto",
    "automatikus",
    "on",
    "off",
    "be",
    "ki",
    "low",
    "alacsony",
    "high",
    "magas",
    "done",
    "kész",
    "mm/dd/yy",
    "hh/nn/éé",
    "hh/mm/yy",
    "hh/nn/yy",
}

NAV_CHROME_PATTERNS = [
    r"^previous page$",
    r"^back$",
    r"^általános$",
    r"^keresés$",
]


TRAILING_VALUE_PATTERNS = [
    r"(Auto|Automatikus|On|Off|Be|Ki)$",
    r"(Low|Alacsony|High|Magas)$",
    r"(Done|Kész)$",
    r"(Connected|Csatlakoztatva|Disconnect|Leválasztás)$",
    r"(mm/dd/yy|hh/nn/éé|hh/mm/yy|hh/nn/yy)$",
]


def split_trailing_value_object(line: str) -> List[str]:
    s = clean_line(line)
    if not s or " " not in s:
        return [s] if s else []
    for pat in TRAILING_VALUE_PATTERNS:
        m = re.search(pat, s, flags=re.IGNORECASE)
        if not m:
            continue
        val = m.group(1).strip()
        left = s[: m.start(1)].strip()
        if left:
            return [left, val]
    return [s]


def extract_title_text(proc_img: Image.Image, tesseract_lang: str) -> str:
    w, h = proc_img.size
    band = proc_img.crop((int(w * 0.08), int(h * 0.08), int(w * 0.60), int(h * 0.24)))
    txt = pytesseract.image_to_string(band, lang=tesseract_lang, config="--psm 7")
    txt = clean_line(txt)
    # Filter breadcrumb-like line fragments.
    if not txt or len(txt) < 4:
        return ""
    if re.search(r"(previous page|általános|back)", txt, flags=re.IGNORECASE):
        return ""
    toks = [t for t in txt.split() if t]
    if not any(sum(ch.isalpha() for ch in t) >= 4 for t in toks):
        return ""
    if is_noise_line(txt, 80):
        return ""
    return txt


def merge_continuation_lines(lines: List[str]) -> List[str]:
    out: List[str] = []
    for line in lines:
        s = clean_line(line)
        if not s:
            continue
        if not out:
            out.append(s)
            continue
        low = s.lower()
        prev = out[-1]
        prev_low = prev.lower()
        cur_is_value = low in VALUE_WORDS
        # Wrap continuation lines.
        is_cont = bool(re.match(r"^(when|wind|and|or|a|az)\b", low))
        # Continuation if previous ends with comma/open paren.
        prev_incomplete = prev.endswith(",") or prev.endswith("(")
        # Do not append into value-only rows.
        prev_is_value = prev_low in VALUE_WORDS
        if (is_cont or prev_incomplete) and not cur_is_value:
            if not prev_is_value:
                out[-1] = f"{prev} {s}"
            else:
                # Attach continuation to nearest previous non-value object.
                merged = False
                for j in range(len(out) - 2, -1, -1):
                    if out[j].lower() not in VALUE_WORDS:
                        out[j] = f"{out[j]} {s}"
                        merged = True
                        break
                if not merged:
                    out.append(s)
        else:
            out.append(s)
    return out


def filter_map_overlay_lines(lines: List[str]) -> List[str]:
    """
    Special-case map reposition overlays:
    keep foreground instruction + action button, drop map background labels.
    """
    norm = [clean_line(x) for x in lines if clean_line(x)]
    has_overlay = any(
        re.search(r"(move the map to set your current position|mozgassa a térképet a jelenlegi pozíciójának beállításához)", s, flags=re.IGNORECASE)
        for s in norm
    )
    if not has_overlay:
        return norm

    kept: List[str] = []
    for s in norm:
        if re.search(r"(move the map to set your current position|mozgassa a térképet a jelenlegi pozíciójának beállításához)", s, flags=re.IGNORECASE):
            kept.append(s)
            continue
        if s.lower() in {"done", "kész"}:
            kept.append(s)
            continue
    # Deduplicate tiny map-overlay set while preserving order.
    out: List[str] = []
    for s in kept:
        if s not in out:
            out.append(s)
    return out


def extract_bottom_right_action(proc_img: Image.Image, tesseract_lang: str) -> str:
    w, h = proc_img.size
    # Action button is typically on the right side of the bottom overlay bar.
    roi = proc_img.crop((int(w * 0.70), int(h * 0.83), int(w * 0.98), int(h * 0.98)))
    txt = clean_line(pytesseract.image_to_string(roi, lang=tesseract_lang, config="--psm 7"))
    txt_low = txt.lower()
    if "done" in txt_low:
        return "Done"
    if "kész" in txt_low or "kesz" in txt_low:
        return "Kész"
    return ""


def split_line_into_objects(tokens: List[Tuple[int, int, str]]) -> List[str]:
    # Split by large horizontal gap (label/value separation in two-column UI rows).
    if not tokens:
        return []
    tokens = sorted(tokens, key=lambda t: t[0])
    gaps: List[int] = []
    for i in range(1, len(tokens)):
        prev_right = tokens[i - 1][1]
        cur_left = tokens[i][0]
        gaps.append(max(0, cur_left - prev_right))

    parts: List[List[str]]
    if not gaps:
        parts = [[t[2] for t in tokens]]
    else:
        max_gap = max(gaps)
        max_i = gaps.index(max_gap) + 1  # split index in tokens
        # If there is a strong row-internal separation, split into left/right text objects.
        if max_gap >= 70:
            parts = [
                [t[2] for t in tokens[:max_i]],
                [t[2] for t in tokens[max_i:]],
            ]
        else:
            parts = [[t[2] for t in tokens]]

    out = []
    for p in parts:
        s = clean_line(" ".join(p))
        if s:
            out.append(s)
    return out


def line_quality_score(lines: List[str]) -> float:
    if not lines:
        return -1e9
    good = 0.0
    for s in lines:
        t = clean_line(s)
        letters = sum(ch.isalpha() for ch in t)
        digits = sum(ch.isdigit() for ch in t)
        if letters >= 3:
            good += 2.0
        if " " in t and len(t) >= 10:
            good += 1.0
        if re.search(r"(2K|KKK|OOO|III)", t, flags=re.IGNORECASE):
            good -= 3.0
        if digits > letters:
            good -= 1.0
        if is_noise_line(t, 60):
            good -= 4.0
    return good


def _ocr_pass_parts(pass_name: str, config: str) -> Tuple[str, str]:
    preprocessing = "bw" if pass_name.endswith("_bw") else "gray"
    m = re.search(r"--psm\s+(\d+)", config)
    return preprocessing, m.group(1) if m else ""


def extract_classic_ocr_evidence(
    image_path: Path,
    tesseract_lang: str,
    tesseract_cmd: str,
    ghostscript_cmd: Optional[str],
    eps_dpi: int,
    min_conf: int = 45,
    pair_id: str = "",
    side: Literal["source", "target"] = "source",
) -> Tuple[ClassicOCRResult, Optional[str]]:
    empty = ClassicOCRResult([], [], [], "", [], {})
    if pytesseract is None:
        return empty, "pytesseract is not installed."
    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    try:
        img, open_err = open_image_for_ocr(image_path, ghostscript_cmd, eps_dpi)
        if img is None:
            return empty, open_err or "Image open failed."
        with img:
            base = preprocess_image(crop_for_ui_text(img))
            candidates = [
                ("psm6_gray", base, "--psm 6"),
                ("psm11_gray", base, "--psm 11"),
                ("psm6_bw", make_binary_text_image(base), "--psm 6"),
                ("psm11_bw", make_binary_text_image(base), "--psm 11"),
            ]

        best_lines: List[str] = []
        best_score = -1e18
        best_pass_name = ""
        last_proc = base
        best_proc = base
        all_tokens: List[OCRToken] = []
        all_lines: List[OCRLine] = []
        full_text_by_pass: Dict[str, str] = {}
        pass_line_indexes: Dict[str, List[int]] = {}

        for pass_name, proc, cfg in candidates:
            last_proc = proc
            preprocessing, psm = _ocr_pass_parts(pass_name, cfg)
            try:
                full_text_by_pass[pass_name] = pytesseract.image_to_string(proc, lang=tesseract_lang, config=cfg).strip()
            except Exception as exc:
                full_text_by_pass[pass_name] = f"[full OCR text failed: {exc}]"
            data = pytesseract.image_to_data(
                proc,
                lang=tesseract_lang,
                config=cfg,
                output_type=pytesseract.Output.DICT,
            )
            groups: Dict[Tuple[int, int, int], List[Tuple[int, int, int, int, str]]] = {}
            conf_groups: Dict[Tuple[int, int, int], List[float]] = {}
            n = len(data.get("text", []))
            for i in range(n):
                txt = (data["text"][i] or "").strip()
                if not txt:
                    continue
                try:
                    conf = int(float(data["conf"][i]))
                except Exception:
                    conf = -1
                left = int(data["left"][i])
                top = int(data["top"][i])
                width = int(data["width"][i])
                height = int(data["height"][i])
                all_tokens.append(
                    OCRToken(
                        pair_id=pair_id,
                        side=side,
                        image_filename=image_path.name,
                        ocr_pass=pass_name,
                        preprocessing=preprocessing,
                        psm=psm,
                        block_num=int(data["block_num"][i]),
                        par_num=int(data["par_num"][i]),
                        line_num=int(data["line_num"][i]),
                        word_num=int(data["word_num"][i]),
                        token_text=txt,
                        token_confidence=float(conf),
                        left=left,
                        top=top,
                        width=width,
                        height=height,
                    )
                )
                if conf < min_conf:
                    continue
                key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
                groups.setdefault(key, []).append((left, left + width, top, top + height, txt))
                conf_groups.setdefault(key, []).append(float(conf))

            lines: List[str] = []
            pass_line_indexes[pass_name] = []
            for raw_line_no, key in enumerate(sorted(groups), start=1):
                token_pairs = groups[key]
                raw = " ".join(t for _, _, _, _, t in token_pairs).strip()
                confs = conf_groups.get(key, [])
                avg_conf = sum(confs) / len(confs) if confs else 0.0
                cleaned = clean_line(raw)
                bbox_left = min((t[0] for t in token_pairs), default=0)
                bbox_top = min((t[2] for t in token_pairs), default=0)
                bbox_right = max((t[1] for t in token_pairs), default=0)
                bbox_bottom = max((t[3] for t in token_pairs), default=0)
                if is_noise_line(cleaned, avg_conf):
                    all_lines.append(
                        OCRLine(
                            pair_id=pair_id,
                            side=side,
                            image_filename=image_path.name,
                            ocr_pass=pass_name,
                            preprocessing=preprocessing,
                            psm=psm,
                            line_no=raw_line_no,
                            raw_line=raw,
                            cleaned_line=cleaned,
                            avg_tesseract_confidence=avg_conf,
                            bbox_left=bbox_left,
                            bbox_top=bbox_top,
                            bbox_right=bbox_right,
                            bbox_bottom=bbox_bottom,
                            kept_or_filtered="filtered",
                            filter_reason="noise_line",
                        )
                    )
                    continue

                split_input = [(left, right, txt) for left, right, _, _, txt in token_pairs]
                objects = split_line_into_objects(split_input)
                final_objects: List[str] = []
                for obj in (objects or [cleaned]):
                    final_objects.extend(split_trailing_value_object(obj))
                for obj in final_objects:
                    obj_clean = clean_line(obj)
                    obj_noise = is_noise_line(obj_clean, avg_conf)
                    all_lines.append(
                        OCRLine(
                            pair_id=pair_id,
                            side=side,
                            image_filename=image_path.name,
                            ocr_pass=pass_name,
                            preprocessing=preprocessing,
                            psm=psm,
                            line_no=len(pass_line_indexes[pass_name]) + 1,
                            raw_line=raw,
                            cleaned_line=obj_clean,
                            avg_tesseract_confidence=avg_conf,
                            bbox_left=bbox_left,
                            bbox_top=bbox_top,
                            bbox_right=bbox_right,
                            bbox_bottom=bbox_bottom,
                            kept_or_filtered="filtered" if obj_noise else "kept",
                            filter_reason="noise_line_after_split" if obj_noise else "",
                        )
                    )
                    if not obj_noise:
                        pass_line_indexes[pass_name].append(len(all_lines) - 1)
                        lines.append(obj_clean)

            lines = merge_continuation_lines(lines)
            lines = filter_map_overlay_lines(lines)
            score = line_quality_score(lines)
            for idx in pass_line_indexes.get(pass_name, []):
                all_lines[idx].line_quality_score = score
            if score > best_score:
                best_score = score
                best_lines = lines
                best_pass_name = pass_name
                best_proc = proc

        for idx in pass_line_indexes.get(best_pass_name, []):
            all_lines[idx].selected_pass = True

        title = extract_title_text(last_proc, tesseract_lang)
        if title and (not best_lines or best_lines[0].lower() != title.lower()):
            best_lines = [title] + best_lines

        has_overlay = any(
            re.search(
                r"(move the map to set your current position|mozgassa a terkepet a jelenlegi poziciojanak beallitasahoz)",
                remove_accents(s),
                flags=re.IGNORECASE,
            )
            for s in best_lines
        )
        if has_overlay and not any(remove_accents(s.lower()) in {"done", "kesz"} for s in best_lines):
            act = extract_bottom_right_action(best_proc, tesseract_lang)
            if act:
                best_lines.append(act)

        return ClassicOCRResult(all_tokens, all_lines, best_lines, consolidate_ocr_lines(best_lines), [], full_text_by_pass), None
    except Exception as exc:
        return empty, f"OCR failed: {exc}"


def extract_text_lines(
    image_path: Path,
    tesseract_lang: str,
    tesseract_cmd: str,
    ghostscript_cmd: Optional[str],
    eps_dpi: int,
    min_conf: int = 45,
) -> Tuple[List[str], Optional[str]]:
    if pytesseract is None:
        return [], "pytesseract is not installed."
    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    try:
        img, open_err = open_image_for_ocr(image_path, ghostscript_cmd, eps_dpi)
        if img is None:
            return [], open_err or "Image open failed."
        with img:
            base = preprocess_image(crop_for_ui_text(img))
            candidates = [
                ("psm6_gray", base, "--psm 6"),
                ("psm11_gray", base, "--psm 11"),
                ("psm6_bw", make_binary_text_image(base), "--psm 6"),
                ("psm11_bw", make_binary_text_image(base), "--psm 11"),
            ]

        best_lines: List[str] = []
        best_score = -1e18
        last_proc = base
        best_proc = base

        for _, proc, cfg in candidates:
            last_proc = proc
            data = pytesseract.image_to_data(
                proc,
                lang=tesseract_lang,
                config=cfg,
                output_type=pytesseract.Output.DICT,
            )
            groups: Dict[Tuple[int, int, int], List[Tuple[int, int, str]]] = {}
            conf_groups: Dict[Tuple[int, int, int], List[float]] = {}
            n = len(data.get("text", []))
            for i in range(n):
                txt = (data["text"][i] or "").strip()
                if not txt:
                    continue
                conf_raw = data["conf"][i]
                try:
                    conf = int(float(conf_raw))
                except Exception:
                    conf = -1
                if conf < min_conf:
                    continue
                key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
                left = int(data["left"][i])
                width = int(data["width"][i])
                right = left + width
                groups.setdefault(key, []).append((left, right, txt))
                conf_groups.setdefault(key, []).append(float(conf))

            lines: List[str] = []
            for key in sorted(groups):
                token_pairs = groups[key]
                raw = " ".join(t for _, _, t in token_pairs).strip()
                confs = conf_groups.get(key, [])
                avg_conf = sum(confs) / len(confs) if confs else 0.0
                cleaned = clean_line(raw)
                if not is_noise_line(cleaned, avg_conf):
                    objects = split_line_into_objects(token_pairs)
                    if objects:
                        final_objects: List[str] = []
                        for obj in objects:
                            final_objects.extend(split_trailing_value_object(obj))
                        for obj in final_objects:
                            if not is_noise_line(obj, avg_conf):
                                lines.append(obj)
                    else:
                        for obj in split_trailing_value_object(cleaned):
                            if not is_noise_line(obj, avg_conf):
                                lines.append(obj)

            lines = merge_continuation_lines(lines)
            lines = filter_map_overlay_lines(lines)
            score = line_quality_score(lines)
            if score > best_score:
                best_score = score
                best_lines = lines
                best_proc = proc

        title = extract_title_text(last_proc, tesseract_lang)
        if title and (not best_lines or best_lines[0].lower() != title.lower()):
            best_lines = [title] + best_lines

        # Map overlay special case: try to recover right-side action button text (Done/Kész).
        has_overlay = any(
            re.search(r"(move the map to set your current position|mozgassa a térképet a jelenlegi pozíciójának beállításához)", s, flags=re.IGNORECASE)
            for s in best_lines
        )
        if has_overlay and not any(s.lower() in {"done", "kész"} for s in best_lines):
            act = extract_bottom_right_action(best_proc, tesseract_lang)
            if act:
                best_lines.append(act)

        return best_lines, None
    except Exception as exc:
        return [], f"OCR failed: {exc}"


def align_lines(source_lines: List[str], target_lines: List[str]) -> List[Tuple[str, str]]:
    # UI rows are usually structurally aligned; pair by order with padding.
    n = max(len(source_lines), len(target_lines))
    pairs: List[Tuple[str, str]] = []
    for i in range(n):
        s = source_lines[i] if i < len(source_lines) else ""
        t = target_lines[i] if i < len(target_lines) else ""
        if not s and not t:
            continue
        pairs.append((s, t))
    return pairs


def build_llm_prompt(source_text_block: str, target_text_block: str, target_lang: str) -> str:
    return f"""You are matching GUI text extracted by OCR from two corresponding UI screenshots.

Rules:
- Use only the provided SOURCE_TEXT and TARGET_TEXT.
- Do not invent text.
- Do not paraphrase.
- Do not translate.
- Return exact strings as present in the text blocks, except trimming, collapsing repeated whitespace, and Unicode normalization.
- Detect GUI strings in SOURCE_TEXT and find corresponding target-language counterparts from TARGET_TEXT.
- If no reliable target counterpart exists, return target_string as an empty string and status "unmatched".
- Use status "matched", "unmatched", or "uncertain".
- confidence must be a number from 0 to 1.
- notes must be short.
- Return strict JSON only: a JSON array of objects.

Each object must have:
- source_string
- target_string
- confidence
- status
- notes

TARGET_LANGUAGE: {target_lang}

SOURCE_TEXT:
<<<SOURCE
{source_text_block}
SOURCE

TARGET_TEXT:
<<<TARGET
{target_text_block}
TARGET
"""


def parse_llm_match_response(response_text: str) -> List[LLMGuiMatch]:
    try:
        payload = json.loads(response_text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"LLM returned invalid JSON: {exc}") from exc
    if not isinstance(payload, list):
        raise ValueError("LLM response must be a JSON array.")

    matches: List[LLMGuiMatch] = []
    for idx, item in enumerate(payload):
        if not isinstance(item, dict):
            raise ValueError(f"LLM response item {idx} is not an object.")
        source_string = normalize_for_llm_guard(item.get("source_string", ""))
        target_string = normalize_for_llm_guard(item.get("target_string", ""))
        status = normalize_for_llm_guard(item.get("status", "")).lower()
        notes = normalize_for_llm_guard(item.get("notes", ""))
        if status not in {"matched", "unmatched", "uncertain"}:
            raise ValueError(f"LLM response item {idx} has invalid status: {status!r}")
        try:
            confidence = float(item.get("confidence", 0.0))
        except (TypeError, ValueError) as exc:
            raise ValueError(f"LLM response item {idx} has invalid confidence.") from exc
        confidence = max(0.0, min(1.0, confidence))
        if not source_string:
            raise ValueError(f"LLM response item {idx} has empty source_string.")
        if status == "matched" and not target_string:
            status = "unmatched"
            confidence = min(confidence, 0.2)
        if status == "unmatched" and target_string:
            status = "uncertain"
        matches.append(
            LLMGuiMatch(
                source_string=source_string,
                target_string=target_string,
                confidence=confidence,
                status=status,
                notes=notes,
                source_span=normalize_for_llm_guard(item.get("source_span", "")),
                target_span=normalize_for_llm_guard(item.get("target_span", "")),
                rationale=normalize_for_llm_guard(item.get("rationale", "")),
            )
        )
    return matches


def extract_openai_text_response(payload: Dict[str, Any]) -> str:
    if isinstance(payload.get("choices"), list) and payload["choices"]:
        message = payload["choices"][0].get("message", {})
        content = message.get("content", "")
        if isinstance(content, str):
            return content
        if isinstance(content, list):
            return "".join(part.get("text", "") for part in content if isinstance(part, dict))
    if isinstance(payload.get("output_text"), str):
        return payload["output_text"]
    raise ValueError("OpenAI response did not contain text content.")


def call_openai_chat_completion(
    prompt: str,
    model: str,
    api_key: str,
    timeout_sec: int,
    max_retries: int,
) -> str:
    body = {
        "model": model,
        "temperature": 0,
        "messages": [
            {
                "role": "system",
                "content": "Return only strict JSON. Do not include Markdown or commentary.",
            },
            {"role": "user", "content": prompt},
        ],
    }
    data = json.dumps(body).encode("utf-8")
    last_error: Optional[Exception] = None
    for attempt in range(max_retries + 1):
        req = urllib.request.Request(
            "https://api.openai.com/v1/chat/completions",
            data=data,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
                return extract_openai_text_response(json.loads(resp.read().decode("utf-8")))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            last_error = RuntimeError(f"OpenAI HTTP {exc.code}: {detail[:500]}")
            if exc.code < 500 and exc.code not in {408, 409, 429}:
                break
        except Exception as exc:
            last_error = exc
        if attempt < max_retries:
            time.sleep(min(2 ** attempt, 8))
    raise RuntimeError(f"OpenAI request failed after {max_retries + 1} attempt(s): {last_error}")


def validate_llm_matches(
    raw_matches: List[LLMGuiMatch],
    source_text_block: str,
    target_text_block: str,
    source_lines: Sequence[str],
) -> Tuple[List[LLMGuiMatch], List[str]]:
    validated: List[LLMGuiMatch] = []
    rejected: List[str] = []
    covered_sources = set()

    for item in raw_matches:
        if not normalized_contains(source_text_block, item.source_string):
            rejected.append(f"Rejected hallucinated source: {item.source_string}")
            continue
        if item.target_string and not normalized_contains(target_text_block, item.target_string):
            rejected.append(f"Rejected hallucinated target for source '{item.source_string}': {item.target_string}")
            item = LLMGuiMatch(
                source_string=item.source_string,
                target_string="",
                confidence=0.0,
                status="uncertain",
                notes=(item.notes + " | target not present in OCR text").strip(" |"),
                source_span=item.source_span,
                target_span="",
                rationale=item.rationale,
            )
        validated.append(item)
        covered_sources.add(normalize_for_llm_guard(item.source_string).lower())

    for source_line in source_lines:
        source = normalize_for_llm_guard(clean_line(source_line))
        if not source:
            continue
        if source.lower() not in covered_sources:
            validated.append(
                LLMGuiMatch(
                    source_string=source,
                    target_string="",
                    confidence=0.0,
                    status="unmatched",
                    notes="Not returned by LLM matcher.",
                )
            )
    return validated, rejected


def match_gui_strings_with_llm(
    source_text_block: str,
    target_text_block: str,
    target_lang: str,
    model: str,
    api_key: str,
    timeout_sec: int,
    max_retries: int,
    source_lines: Sequence[str],
) -> Tuple[List[LLMGuiMatch], List[str]]:
    prompt = build_llm_prompt(source_text_block, target_text_block, target_lang)
    response_text = call_openai_chat_completion(prompt, model, api_key, timeout_sec, max_retries)
    raw_matches = parse_llm_match_response(response_text)
    return validate_llm_matches(raw_matches, source_text_block, target_text_block, source_lines)


def model_to_dict(model: BaseModel) -> Dict[str, Any]:
    if hasattr(model, "model_dump"):
        return model.model_dump()
    return model.dict()  # type: ignore[attr-defined]


def model_to_json(model: BaseModel) -> str:
    if hasattr(model, "model_dump_json"):
        return model.model_dump_json()
    return model.json()  # type: ignore[attr-defined]


def cache_key_for_payload(*parts: str) -> str:
    h = hashlib.sha256()
    for part in parts:
        h.update(part.encode("utf-8", errors="replace"))
        h.update(b"\0")
    return h.hexdigest()


def read_cached_model(cache_dir: Path, key: str, model_cls: Any) -> Optional[BaseModel]:
    path = cache_dir / f"{key}.json"
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        return model_cls.model_validate(payload) if hasattr(model_cls, "model_validate") else model_cls.parse_obj(payload)
    except Exception as exc:
        logging.warning("Ignoring invalid cache file %s: %s", path, exc)
        return None


def write_cached_model(cache_dir: Path, key: str, model: BaseModel) -> None:
    cache_dir.mkdir(parents=True, exist_ok=True)
    (cache_dir / f"{key}.json").write_text(model_to_json(model), encoding="utf-8")


def ocr_candidates_package(
    image_filename: str,
    language_code: str,
    language_name: str,
    side: str,
    evidence: ClassicOCRResult,
) -> Dict[str, Any]:
    passes: Dict[str, Dict[str, Any]] = {}
    selected_pass_names = set()
    for line in evidence.classic_lines:
        if line.selected_pass:
            selected_pass_names.add(line.ocr_pass)
        p = passes.setdefault(
            line.ocr_pass,
            {
                "pass_name": line.ocr_pass,
                "preprocessing": line.preprocessing,
                "psm": line.psm,
                "selected_pass": False,
                "line_quality_score": line.line_quality_score,
                "full_ocr_text": evidence.full_text_by_pass.get(line.ocr_pass, ""),
                "lines": [],
            },
        )
        p["selected_pass"] = p["selected_pass"] or line.selected_pass
        p["line_quality_score"] = max(float(p["line_quality_score"]), line.line_quality_score)
        p["lines"].append(
            {
                "line_no": line.line_no,
                "raw_line": line.raw_line,
                "cleaned_line": line.cleaned_line,
                "avg_confidence": line.avg_tesseract_confidence,
                "bbox": [line.bbox_left, line.bbox_top, line.bbox_right, line.bbox_bottom],
                "kept_or_filtered": line.kept_or_filtered,
                "filter_reason": line.filter_reason,
            }
        )
    return {
        "image_filename": image_filename,
        "language_code": language_code,
        "language_name": language_name,
        "side": side,
        "domain": "vehicle infotainment / car information display",
        "selected_line_block_for_reference": evidence.selected_block,
        "full_ocr_text_by_pass": [
            {
                "pass_name": pass_name,
                "selected_pass": pass_name in selected_pass_names,
                "full_ocr_text": text,
            }
            for pass_name, text in evidence.full_text_by_pass.items()
        ],
        "ocr_passes": list(passes.values()),
    }


def system_prompt_llm_ocr_normalization() -> str:
    return """You are an OCR adjudication and normalization engine for vehicle infotainment screenshots.

You receive full raw Tesseract OCR text from several OCR passes for a screenshot of a car information display. Treat the full OCR text as the primary evidence. Use token/line metadata only as secondary support when it helps confidence or object separation. Decipher what the screenshot says from the complete OCR evidence, normalize OCR damage only when strongly supported by OCR evidence and vehicle-GUI context, preserve the screenshot language, and do not translate. Keep page titles, menu labels, descriptions, values, buttons, breadcrumbs, tabs, and status text separate. Ignore icons, time, signal indicators, slider graphics, decorative marks, and map background labels unless they are meaningful GUI text. Prefer high precision over high recall. Return data matching the provided schema."""


def system_prompt_ai_image_ocr() -> str:
    return """You are an expert OCR and GUI-string extraction engine for vehicle infotainment screenshots.

Read the screenshot directly. Extract meaningful visible GUI text objects in the screenshot language. Do not translate and do not invent text. Keep separate GUI objects separate, join wrapped visual lines only when they form one GUI object, split merged text when needed, ignore icons/time/signal/decorative graphics, and distinguish navigation map labels from actual GUI text. Mark uncertain readings with needs_review. Return data matching the provided schema."""


def system_prompt_gui_object_match(target_language: str) -> str:
    return f"""You are a bilingual GUI string alignment specialist for vehicle infotainment screenshots.

Match English GUI objects to corresponding {target_language} GUI objects using semantic equivalence, not line position alone. Use GUI role, row group, screen area, and reading order as supporting evidence. Do not invent target text and do not translate the output yourself. Return unmatched or uncertain when a reliable counterpart is not present. Return data matching the provided schema."""


def ensure_openai_client(api_key: str) -> Any:
    if OpenAI is None:
        raise RuntimeError("The openai package is not installed. Run pip install -r requirements.txt.")
    if not api_key:
        raise RuntimeError("AI processing requires an OpenAI API key. Pass --openai-api-key or set OPENAI_API_KEY.")
    return OpenAI(api_key=api_key)


def parse_response_output_model(response: Any, model_cls: Any) -> BaseModel:
    for output in getattr(response, "output", []) or []:
        for item in getattr(output, "content", []) or []:
            parsed = getattr(item, "parsed", None)
            if parsed is not None:
                return parsed
    output_text = getattr(response, "output_text", "")
    if not output_text:
        raise ValueError("OpenAI response did not contain parseable output.")
    payload = json.loads(output_text)
    return model_cls.model_validate(payload) if hasattr(model_cls, "model_validate") else model_cls.parse_obj(payload)


def call_openai_structured(
    api_key: str,
    model: str,
    system_prompt: str,
    user_content: Any,
    schema_model: Any,
    timeout_sec: int,
    max_retries: int,
    audit: Optional[AIAuditLog] = None,
    audit_stage: str = "",
    audit_title: str = "",
    audit_prompt_version: str = "",
    audit_request: Optional[Dict[str, Any]] = None,
) -> BaseModel:
    client = ensure_openai_client(api_key)
    last_error: Optional[Exception] = None
    request_for_audit = audit_request or {
        "instructions": system_prompt,
        "input": user_content,
        "schema": getattr(schema_model, "__name__", str(schema_model)),
    }
    for attempt in range(max_retries + 1):
        try:
            # SDK parse helper gives Pydantic validation where available.
            response = client.responses.parse(
                model=model,
                instructions=system_prompt,
                input=[{"role": "user", "content": user_content}],
                text_format=schema_model,
                timeout=timeout_sec,
            )
            parsed = parse_response_output_model(response, schema_model)
            if audit is not None:
                audit.add(
                    audit_stage,
                    audit_title,
                    model,
                    audit_prompt_version,
                    request_for_audit,
                    response=model_to_dict(parsed),
                )
            return parsed
        except TypeError:
            response = client.responses.create(
                model=model,
                instructions=system_prompt,
                input=[{"role": "user", "content": user_content}],
                timeout=timeout_sec,
                text={
                    "format": {
                        "type": "json_schema",
                        "name": schema_model.__name__,
                        "schema": schema_model.model_json_schema(),
                        "strict": True,
                    }
                },
            )
            parsed = parse_response_output_model(response, schema_model)
            if audit is not None:
                audit.add(
                    audit_stage,
                    audit_title,
                    model,
                    audit_prompt_version,
                    request_for_audit,
                    response=model_to_dict(parsed),
                )
            return parsed
        except Exception as exc:
            last_error = exc
            if attempt < max_retries:
                time.sleep(min(2 ** attempt, 8))
    if audit is not None:
        audit.add(
            audit_stage,
            audit_title,
            model,
            audit_prompt_version,
            request_for_audit,
            error=str(last_error),
        )
    raise RuntimeError(f"OpenAI structured request failed after {max_retries + 1} attempt(s): {last_error}")


def normalize_ocr_with_llm(
    image_filename: str,
    language_code: str,
    language_name: str,
    side: Literal["source", "target"],
    evidence: ClassicOCRResult,
    model: str,
    api_key: str,
    timeout_sec: int,
    max_retries: int,
    cache_dir: Path,
    use_cache: bool,
    audit: Optional[AIAuditLog] = None,
    pair_id: str = "",
) -> OCRNormalizedResult:
    package = ocr_candidates_package(image_filename, language_code, language_name, side, evidence)
    package_json = json.dumps(package, ensure_ascii=False, sort_keys=True)
    key = cache_key_for_payload(LLM_OCR_NORMALIZATION_PROMPT_VERSION, model, package_json)
    if use_cache:
        cached = read_cached_model(cache_dir, key, OCRNormalizedResult)
        if cached:
            result = cached  # type: ignore[assignment]
            result.cached = True
            if audit is not None:
                audit.add(
                    "LLM OCR-normalization",
                    f"{pair_id} {side} {image_filename}",
                    model,
                    LLM_OCR_NORMALIZATION_PROMPT_VERSION,
                    {
                        "cache_key": key,
                        "image_filename": image_filename,
                        "language": language_code,
                        "side": side,
                        "full_ocr_text_by_pass": package.get("full_ocr_text_by_pass", []),
                    },
                    response=model_to_dict(result),
                    cached=True,
                )
            return result
    user_text = (
        f"FILENAME: {image_filename}\n"
        f"LANGUAGE CODE: {language_code}\n"
        f"LANGUAGE NAME: {language_name}\n"
        f"SIDE: {side}\n"
        "DOMAIN: vehicle infotainment / car information display\n\n"
        "TASK: Normalize OCR candidates into meaningful GUI text objects. Do not translate.\n\n"
        "IMPORTANT: Use FULL_OCR_TEXT_BY_PASS as the main evidence. Do not over-trust line/token breakup.\n\n"
        f"OCR_CANDIDATES_JSON:\n{package_json}"
    )
    audit_request = {
        "filename": image_filename,
        "language_code": language_code,
        "language_name": language_name,
        "side": side,
        "system_prompt": system_prompt_llm_ocr_normalization(),
        "user_text": user_text,
        "ocr_candidates": package,
    }
    result = call_openai_structured(
        api_key,
        model,
        system_prompt_llm_ocr_normalization(),
        [{"type": "input_text", "text": user_text}],
        OCRNormalizedResult,
        timeout_sec,
        max_retries,
        audit=audit,
        audit_stage="LLM OCR-normalization",
        audit_title=f"{pair_id} {side} {image_filename}",
        audit_prompt_version=LLM_OCR_NORMALIZATION_PROMPT_VERSION,
        audit_request=audit_request,
    )
    assert isinstance(result, OCRNormalizedResult)
    result.source_engine = "llm_ocr_normalization"
    result.model = model
    result.prompt_version = LLM_OCR_NORMALIZATION_PROMPT_VERSION
    result.cached = False
    if use_cache:
        write_cached_model(cache_dir, key, result)
    return result


def encode_image_data_url(path: Path) -> str:
    mime = "image/png" if path.suffix.lower() == ".png" else "image/jpeg"
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{data}"


def ai_image_ocr(
    rendered_png: Path,
    image_filename: str,
    language_code: str,
    language_name: str,
    side: Literal["source", "target"],
    model: str,
    detail: str,
    api_key: str,
    timeout_sec: int,
    max_retries: int,
    cache_dir: Path,
    use_cache: bool,
    audit: Optional[AIAuditLog] = None,
    pair_id: str = "",
) -> OCRNormalizedResult:
    image_hash = file_sha256(rendered_png)
    key = cache_key_for_payload(AI_IMAGE_OCR_PROMPT_VERSION, model, detail, image_hash, image_filename, side, language_code)
    if use_cache:
        cached = read_cached_model(cache_dir, key, OCRNormalizedResult)
        if cached:
            result = cached  # type: ignore[assignment]
            result.cached = True
            if audit is not None:
                audit.add(
                    "AI image OCR",
                    f"{pair_id} {side} {image_filename}",
                    model,
                    AI_IMAGE_OCR_PROMPT_VERSION,
                    {
                        "cache_key": key,
                        "rendered_png": str(rendered_png),
                        "image_hash": image_hash,
                        "detail": detail,
                        "image_payload": f"omitted from HTML audit; data URL length={len(encode_image_data_url(rendered_png))}",
                    },
                    response=model_to_dict(result),
                    cached=True,
                )
            return result
    user_text = (
        f"FILENAME: {image_filename}\n"
        f"LANGUAGE CODE: {language_code}\n"
        f"LANGUAGE NAME: {language_name}\n"
        f"SIDE: {side}\n"
        "DOMAIN: vehicle infotainment / car information display\n\n"
        "TASK: Extract meaningful visible GUI text objects from the screenshot. Do not translate."
    )
    image_part: Dict[str, Any] = {"type": "input_image", "image_url": encode_image_data_url(rendered_png)}
    if detail:
        image_part["detail"] = detail
    audit_request = {
        "filename": image_filename,
        "language_code": language_code,
        "language_name": language_name,
        "side": side,
        "system_prompt": system_prompt_ai_image_ocr(),
        "user_text": user_text,
        "rendered_png": str(rendered_png),
        "image_hash": image_hash,
        "detail": detail,
        "image_payload": f"omitted from HTML audit; data URL length={len(image_part['image_url'])}",
    }
    result = call_openai_structured(
        api_key,
        model,
        system_prompt_ai_image_ocr(),
        [{"type": "input_text", "text": user_text}, image_part],
        OCRNormalizedResult,
        timeout_sec,
        max_retries,
        audit=audit,
        audit_stage="AI image OCR",
        audit_title=f"{pair_id} {side} {image_filename}",
        audit_prompt_version=AI_IMAGE_OCR_PROMPT_VERSION,
        audit_request=audit_request,
    )
    assert isinstance(result, OCRNormalizedResult)
    result.source_engine = "ai_image_ocr"
    result.model = model
    result.prompt_version = AI_IMAGE_OCR_PROMPT_VERSION
    result.image_hash = image_hash
    result.cached = False
    if use_cache:
        write_cached_model(cache_dir, key, result)
    return result


def classic_fallback_result(
    image_filename: str,
    language_code: str,
    language_name: str,
    side: Literal["source", "target"],
    evidence: ClassicOCRResult,
) -> OCRNormalizedResult:
    objects: List[NormalizedGUIObject] = []
    for idx, line in enumerate(evidence.selected_lines, start=1):
        conf = text_quality_score(line)
        objects.append(
            NormalizedGUIObject(
                object_id=f"classic_{idx:03d}",
                normalized_text=line,
                raw_evidence=[line],
                gui_role="other",
                row_group=idx,
                screen_area="unknown",
                reading_order=idx,
                correction_type="none",
                ocr_evidence_confidence=conf,
                normalization_confidence=conf,
                needs_review=conf < 0.75,
                review_reason="Classic OCR fallback only." if conf < 0.75 else "",
                rationale="Built from selected Tesseract OCR line.",
            )
        )
    return OCRNormalizedResult(
        image_filename=image_filename,
        language_code=language_code,
        language_name=language_name,
        side=side,
        source_engine="classic_fallback",
        model="classic",
        prompt_version="classic_fallback_v1",
        ui_objects=objects,
    )


def string_similarity(a: str, b: str) -> float:
    aa = remove_accents(clean_line(a)).lower()
    bb = remove_accents(clean_line(b)).lower()
    if not aa or not bb:
        return 0.0
    return difflib.SequenceMatcher(None, aa, bb).ratio()


def triangulation_score(status: str, selected_confidence: float) -> float:
    table = {
        "all_agree": 1.0,
        "llm_ai_agree": 0.95,
        "classic_llm_agree": 0.85,
        "classic_ai_agree": 0.80,
        "partial": 0.65,
        "disagree": 0.30,
    }
    if status in table:
        return table[status]
    return min(0.70, max(0.20, selected_confidence))


def triangulate_side(
    pair_id: str,
    side: Literal["source", "target"],
    image_filename: str,
    classic: ClassicOCRResult,
    llm_result: Optional[OCRNormalizedResult],
    ai_result: Optional[OCRNormalizedResult],
) -> List[OCRTriangulationRow]:
    llm_objects = llm_result.ui_objects if llm_result else []
    ai_objects = ai_result.ui_objects if ai_result else []
    classic_lines = classic.selected_lines
    max_len = max(len(classic_lines), len(llm_objects), len(ai_objects))
    rows: List[OCRTriangulationRow] = []
    for idx in range(max_len):
        classic_text = classic_lines[idx] if idx < len(classic_lines) else ""
        llm_obj = llm_objects[idx] if idx < len(llm_objects) else None
        ai_obj = ai_objects[idx] if idx < len(ai_objects) else None
        llm_text = llm_obj.normalized_text if llm_obj else ""
        ai_text = ai_obj.normalized_text if ai_obj else ""
        c_l = string_similarity(classic_text, llm_text)
        c_a = string_similarity(classic_text, ai_text)
        l_a = string_similarity(llm_text, ai_text)

        present = [bool(classic_text), bool(llm_text), bool(ai_text)]
        if all(present) and min(c_l, c_a, l_a) >= 0.88:
            status = "all_agree"
        elif llm_text and ai_text and l_a >= 0.88:
            status = "llm_ai_agree"
        elif classic_text and llm_text and c_l >= 0.88:
            status = "classic_llm_agree"
        elif classic_text and ai_text and c_a >= 0.88:
            status = "classic_ai_agree"
        elif sum(present) == 1:
            status = "classic_only" if classic_text else "llm_only" if llm_text else "ai_only"
        elif max(c_l, c_a, l_a) >= 0.55:
            status = "partial"
        else:
            status = "disagree"

        if llm_text and ai_text and l_a >= 0.80:
            selected_text = llm_text if (llm_obj and llm_obj.normalization_confidence >= (ai_obj.normalization_confidence if ai_obj else 0)) else ai_text
            selected_source = "merged" if status in {"all_agree", "llm_ai_agree"} else ("llm_ocr_normalization" if selected_text == llm_text else "ai_image_ocr")
            selected_conf = max(llm_obj.normalization_confidence if llm_obj else 0, ai_obj.normalization_confidence if ai_obj else 0)
            chosen_obj = llm_obj if selected_text == llm_text else ai_obj
        elif ai_text:
            selected_text = ai_text
            selected_source = "ai_image_ocr"
            selected_conf = ai_obj.normalization_confidence if ai_obj else 0.0
            chosen_obj = ai_obj
        elif llm_text:
            selected_text = llm_text
            selected_source = "llm_ocr_normalization"
            selected_conf = llm_obj.normalization_confidence if llm_obj else 0.0
            chosen_obj = llm_obj
        elif classic_text:
            selected_text = classic_text
            selected_source = "classic"
            selected_conf = text_quality_score(classic_text)
            chosen_obj = None
        else:
            selected_text = ""
            selected_source = "none"
            selected_conf = 0.0
            chosen_obj = None

        needs_review = status in {"partial", "disagree", "classic_only", "llm_only", "ai_only"} or selected_conf < 0.75
        review_reason = "" if not needs_review else f"OCR agreement status: {status}."
        rows.append(
            OCRTriangulationRow(
                pair_id=pair_id,
                side=side,
                image_filename=image_filename,
                classic_raw_line=classic_text,
                llm_object_id=llm_obj.object_id if llm_obj else "",
                llm_normalized_text=llm_text,
                ai_object_id=ai_obj.object_id if ai_obj else "",
                ai_normalized_text=ai_text,
                classic_vs_llm_similarity=round(c_l, 3),
                classic_vs_ai_similarity=round(c_a, 3),
                llm_vs_ai_similarity=round(l_a, 3),
                agreement_status=status,  # type: ignore[arg-type]
                selected_final_text=selected_text,
                selected_source=selected_source,  # type: ignore[arg-type]
                selected_confidence=round(selected_conf, 3),
                needs_review=needs_review,
                review_reason=review_reason,
                gui_role=chosen_obj.gui_role if chosen_obj else "other",
                row_group=chosen_obj.row_group if chosen_obj else idx + 1,
                screen_area=chosen_obj.screen_area if chosen_obj else "unknown",
                reading_order=chosen_obj.reading_order if chosen_obj else idx + 1,
            )
        )
    return rows


def selected_objects_from_triangulation(rows: Sequence[OCRTriangulationRow]) -> List[NormalizedGUIObject]:
    objects: List[NormalizedGUIObject] = []
    for idx, row in enumerate(rows, start=1):
        if not row.selected_final_text:
            continue
        objects.append(
            NormalizedGUIObject(
                object_id=f"selected_{idx:03d}",
                normalized_text=row.selected_final_text,
                raw_evidence=[row.classic_raw_line, row.llm_normalized_text, row.ai_normalized_text],
                gui_role=row.gui_role if row.gui_role in GUI_ROLES else "other",
                row_group=row.row_group,
                screen_area=row.screen_area if row.screen_area in SCREEN_AREAS else "unknown",
                reading_order=row.reading_order,
                ocr_evidence_confidence=row.selected_confidence,
                normalization_confidence=row.selected_confidence,
                needs_review=row.needs_review,
                review_reason=row.review_reason,
            )
        )
    return objects


def deterministic_object_matches(
    pair_id: str,
    source_image: str,
    target_image: str,
    source_rows: Sequence[OCRTriangulationRow],
    target_rows: Sequence[OCRTriangulationRow],
) -> List[GUIMatch]:
    source_objects = selected_objects_from_triangulation(source_rows)
    target_objects = selected_objects_from_triangulation(target_rows)
    aligned = align_lines([o.normalized_text for o in source_objects], [o.normalized_text for o in target_objects])
    matches: List[GUIMatch] = []
    for idx, (source_text, target_text) in enumerate(aligned, start=1):
        source_row = source_rows[idx - 1] if idx - 1 < len(source_rows) else None
        target_row = target_rows[idx - 1] if idx - 1 < len(target_rows) else None
        has_target = bool(target_text)
        semantic_conf = 0.55 if has_target else 0.0
        layout_conf = 0.70 if has_target else 0.0
        source_conf = source_row.selected_confidence if source_row else text_quality_score(source_text)
        target_conf = target_row.selected_confidence if target_row else text_quality_score(target_text)
        agreement = min(
            triangulation_score(source_row.agreement_status, source_conf) if source_row else 0.5,
            triangulation_score(target_row.agreement_status, target_conf) if target_row else 0.5,
        )
        overall = (
            0.25 * source_conf
            + 0.25 * target_conf
            + 0.25 * semantic_conf
            + 0.15 * layout_conf
            + 0.10 * agreement
        )
        status = "needs_review" if has_target and overall >= 0.50 else "unmatched"
        matches.append(
            GUIMatch(
                pair_id=pair_id,
                source_image_filename=source_image,
                target_image_filename=target_image,
                source_object_id=f"selected_{idx:03d}",
                target_object_id=f"selected_{idx:03d}" if has_target else None,
                source_text=source_text,
                target_text=target_text,
                match_type="uncertain" if has_target else "unmatched",
                semantic_confidence=semantic_conf,
                layout_confidence=layout_conf,
                overall_confidence=round(overall, 3),
                status=status,
                rationale="Deterministic positional fallback; use --matching-mode llm_objects for semantic matching.",
                review_reason="Semantic LLM object matching not used.",
                source_selected_source=source_row.selected_source if source_row else "",
                target_selected_source=target_row.selected_source if target_row else "",
                source_gui_role=source_row.gui_role if source_row else "",
                target_gui_role=target_row.gui_role if target_row else "",
                source_row_group=source_row.row_group if source_row else idx,
                target_row_group=target_row.row_group if target_row else idx,
                source_screen_area=source_row.screen_area if source_row else "",
                target_screen_area=target_row.screen_area if target_row else "",
                triangulation_agreement_score=round(agreement, 3),
            )
        )
    return matches


def legacy_llm_text_matches(
    pair_id: str,
    source_image: str,
    target_image: str,
    source_rows: Sequence[OCRTriangulationRow],
    target_rows: Sequence[OCRTriangulationRow],
    target_lang: str,
    model: str,
    api_key: str,
    timeout_sec: int,
    max_retries: int,
) -> List[GUIMatch]:
    source_lines = [r.selected_final_text for r in source_rows if r.selected_final_text]
    target_lines = [r.selected_final_text for r in target_rows if r.selected_final_text]
    source_block = consolidate_ocr_lines(source_lines)
    target_block = consolidate_ocr_lines(target_lines)
    raw_matches, _rejected = match_gui_strings_with_llm(
        source_block,
        target_block,
        target_lang,
        model,
        api_key,
        timeout_sec,
        max_retries,
        source_lines,
    )
    out: List[GUIMatch] = []
    for idx, item in enumerate(raw_matches, start=1):
        status = "matched" if item.status == "matched" and item.confidence >= 0.90 else "needs_review" if item.status == "matched" else item.status
        out.append(
            GUIMatch(
                pair_id=pair_id,
                source_image_filename=source_image,
                target_image_filename=target_image,
                source_object_id=f"text_{idx:03d}",
                target_object_id=f"text_{idx:03d}" if item.target_string else None,
                source_text=item.source_string,
                target_text=item.target_string,
                match_type="semantic_translation" if item.status == "matched" else item.status,
                semantic_confidence=item.confidence,
                layout_confidence=0.0,
                overall_confidence=item.confidence,
                status=status,  # type: ignore[arg-type]
                rationale=item.rationale or "Legacy LLM text-block matching.",
                review_reason=item.notes if status != "matched" else "",
            )
        )
    return out


def validate_gui_matches(
    result: GUIObjectMatchResponse,
    source_objects: Sequence[NormalizedGUIObject],
    target_objects: Sequence[NormalizedGUIObject],
) -> List[GUIMatch]:
    source_by_id = {o.object_id: o for o in source_objects}
    target_by_id = {o.object_id: o for o in target_objects}
    used_targets = set()
    validated: List[GUIMatch] = []
    for match in result.matches:
        source_obj = source_by_id.get(match.source_object_id)
        target_obj = target_by_id.get(match.target_object_id or "")
        if not source_obj:
            continue
        if match.target_object_id and match.target_object_id in used_targets:
            match.status = "needs_review"
            match.review_reason = (match.review_reason + " | target object reused").strip(" |")
        if match.target_object_id:
            used_targets.add(match.target_object_id)
        match.source_text = source_obj.normalized_text
        if target_obj:
            match.target_text = target_obj.normalized_text
        elif match.status == "matched":
            match.status = "needs_review"
            match.review_reason = (match.review_reason + " | matched target object missing").strip(" |")
        validated.append(match)
    return validated


def match_objects_with_llm(
    pair_id: str,
    source_image: str,
    target_image: str,
    target_language_code: str,
    target_language_name: str,
    source_rows: Sequence[OCRTriangulationRow],
    target_rows: Sequence[OCRTriangulationRow],
    model: str,
    api_key: str,
    timeout_sec: int,
    max_retries: int,
    audit: Optional[AIAuditLog] = None,
) -> List[GUIMatch]:
    source_objects = selected_objects_from_triangulation(source_rows)
    target_objects = selected_objects_from_triangulation(target_rows)
    if not source_objects:
        return []
    user_text = (
        f"PAIR ID: {pair_id}\n"
        f"SOURCE IMAGE: {source_image}\n"
        f"TARGET IMAGE: {target_image}\n"
        "SOURCE LANGUAGE: English\n"
        f"TARGET LANGUAGE CODE: {target_language_code}\n"
        f"TARGET LANGUAGE NAME: {target_language_name}\n\n"
        f"SOURCE_OBJECTS:\n{json.dumps([model_to_dict(o) for o in source_objects], ensure_ascii=False)}\n\n"
        f"TARGET_OBJECTS:\n{json.dumps([model_to_dict(o) for o in target_objects], ensure_ascii=False)}\n\n"
        "TASK: Match source GUI objects to target GUI objects."
    )
    audit_request = {
        "pair_id": pair_id,
        "source_image": source_image,
        "target_image": target_image,
        "target_language_code": target_language_code,
        "target_language_name": target_language_name,
        "system_prompt": system_prompt_gui_object_match(target_language_name),
        "source_objects": [model_to_dict(o) for o in source_objects],
        "target_objects": [model_to_dict(o) for o in target_objects],
        "user_text": user_text,
    }
    result = call_openai_structured(
        api_key,
        model,
        system_prompt_gui_object_match(target_language_name),
        [{"type": "input_text", "text": user_text}],
        GUIObjectMatchResponse,
        timeout_sec,
        max_retries,
        audit=audit,
        audit_stage="GUI object matching",
        audit_title=f"{pair_id} {source_image} -> {target_image}",
        audit_prompt_version=GUI_OBJECT_MATCH_PROMPT_VERSION,
        audit_request=audit_request,
    )
    assert isinstance(result, GUIObjectMatchResponse)
    matches = validate_gui_matches(result, source_objects, target_objects)
    source_row_by_id = {f"selected_{i:03d}": row for i, row in enumerate(source_rows, start=1)}
    target_row_by_id = {f"selected_{i:03d}": row for i, row in enumerate(target_rows, start=1)}
    for match in matches:
        source_row = source_row_by_id.get(match.source_object_id)
        target_row = target_row_by_id.get(match.target_object_id or "")
        source_conf = source_row.selected_confidence if source_row else 0.5
        target_conf = target_row.selected_confidence if target_row else 0.0
        agreement = min(
            triangulation_score(source_row.agreement_status, source_conf) if source_row else 0.5,
            triangulation_score(target_row.agreement_status, target_conf) if target_row else 0.5,
        )
        overall = (
            0.25 * source_conf
            + 0.25 * target_conf
            + 0.25 * match.semantic_confidence
            + 0.15 * match.layout_confidence
            + 0.10 * agreement
        )
        match.overall_confidence = round(overall, 3)
        match.triangulation_agreement_score = round(agreement, 3)
        if overall >= 0.90 and match.status == "matched":
            match.status = "matched"
        elif overall >= 0.75 and match.status == "matched":
            match.status = "needs_review"
            match.review_reason = (match.review_reason + " | confidence below auto-match threshold").strip(" |")
        elif overall >= 0.50 and match.status != "unmatched":
            match.status = "uncertain"
        elif overall < 0.50:
            match.status = "unmatched"
        match.source_selected_source = source_row.selected_source if source_row else ""
        match.target_selected_source = target_row.selected_source if target_row else ""
        match.source_gui_role = source_row.gui_role if source_row else ""
        match.target_gui_role = target_row.gui_role if target_row else ""
        match.source_row_group = source_row.row_group if source_row else 0
        match.target_row_group = target_row.row_group if target_row else 0
        match.source_screen_area = source_row.screen_area if source_row else ""
        match.target_screen_area = target_row.screen_area if target_row else ""
    return matches


def text_quality_score(text: str) -> float:
    s = clean_line("" if text is None else str(text))
    if not s:
        return 0.0
    letters = sum(ch.isalpha() for ch in s)
    digits = sum(ch.isdigit() for ch in s)
    ratio_letters = letters / max(1, len(s))
    score = 0.2
    if letters >= 3:
        score += 0.35
    if " " in s:
        score += 0.2
    if ratio_letters >= 0.45:
        score += 0.2
    if digits > letters:
        score -= 0.2
    if is_noise_line(s, 75):
        score -= 0.5
    return max(0.0, min(1.0, score))


def is_meaningful_text(text: str) -> bool:
    s = clean_line("" if text is None else str(text))
    if not s:
        return False
    return text_quality_score(s) >= 0.45


def set_excel_formatting(output_file: Path) -> None:
    from openpyxl import load_workbook

    wb = load_workbook(output_file)
    ws = wb.active
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        1: 32,
        2: 55,
        3: 60,
        4: 14,
        5: 36,
        6: 55,
        7: 60,
        8: 18,
        9: 14,
        10: 45,
        11: 24,
        12: 14,
        13: 14,
        14: 14,
        15: 18,
        16: 30,
        17: 30,
        18: 45,
    }
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    wrap_cols = [3, 7, 10, 11, 18]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for idx in wrap_cols:
            row[idx - 1].alignment = Alignment(wrap_text=True, vertical="top")

    no_match_fill = PatternFill(start_color="FFF4CCCC", end_color="FFF4CCCC", fill_type="solid")
    low_conf_fill = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")

    if ws.max_row >= 2:
        ws.conditional_formatting.add(
            f"A2:R{ws.max_row}",
            FormulaRule(formula=['$K2="NO MATCH"'], stopIfTrue=True, fill=no_match_fill),
        )
        ws.conditional_formatting.add(
            f"A2:R{ws.max_row}",
            FormulaRule(formula=['$K2="LOW CONFIDENCE"'], stopIfTrue=True, fill=low_conf_fill),
        )
    wb.save(output_file)


def make_report(
    source_files: List[Path],
    lang_code: str,
    target_dir: Path,
    tesseract_lang_target: str,
    tesseract_lang_source: str,
    tesseract_cmd: str,
    ghostscript_cmd: Optional[str],
    eps_dpi: int,
    use_temp_local: bool,
    fuzzy_threshold: float,
    progress_every: int,
    allow_fuzzy: bool,
    reject_ambiguous_fuzzy: bool,
    strict_text_only: bool,
    use_llm_matching: bool,
    openai_model: str,
    openai_api_key: Optional[str],
    llm_timeout_sec: int,
    llm_max_retries: int,
) -> pd.DataFrame:
    target_files = list_images(target_dir, SUPPORTED_EXTENSIONS)
    matches = match_source_to_target(
        source_files,
        target_files,
        fuzzy_threshold=fuzzy_threshold,
        allow_fuzzy=allow_fuzzy,
        reject_ambiguous_fuzzy=reject_ambiguous_fuzzy,
    )

    temp_dir_ctx = tempfile.TemporaryDirectory(prefix="ocr_local_") if use_temp_local else None
    temp_root = Path(temp_dir_ctx.name) if temp_dir_ctx else None
    if temp_root:
        logging.info("Using temporary local OCR copy folder: %s", temp_root)

    rows = []
    ocr_errors = 0
    low_conf = 0
    missing = 0
    matched = 0

    src_by_name = {p.name: p for p in source_files}
    trg_by_name = {p.name: p for p in target_files}

    def add_report_row(
        m: MatchResult,
        src_path: Path,
        trg_path: Optional[Path],
        lang_code: str,
        source_str: str,
        target_str: str,
        method: str,
        confidence: float,
        status_notes: str,
        llm_model: str = "",
        llm_source_span: str = "",
        llm_target_span: str = "",
        llm_rationale: str = "",
    ) -> None:
        src_q = text_quality_score(source_str)
        tgt_q = text_quality_score(target_str)
        row_q = round((src_q + tgt_q) / 2.0, 3)
        rows.append(
            {
                "Source image filename": m.source_name,
                "Source image path": str(src_path),
                "Source OCR text": source_str,
                "Target language": lang_code,
                "Matched target image filename": m.target_name or "",
                "Matched target image path": str(trg_path) if trg_path else "",
                "Target OCR text": target_str,
                "Match method": method,
                "Match confidence": round(confidence, 3),
                "Alternative candidate filenames": "; ".join(m.alternatives),
                "Status / notes": status_notes,
                "Source OCR confidence": round(src_q, 3),
                "Target OCR confidence": round(tgt_q, 3),
                "Row quality score": row_q,
                "LLM model": llm_model,
                "LLM source span": llm_source_span,
                "LLM target span": llm_target_span,
                "LLM rationale": llm_rationale,
            }
        )

    total = len(matches)
    for idx, m in enumerate(matches, start=1):
        src_path = src_by_name[m.source_name]
        trg_path = trg_by_name.get(m.target_name) if m.target_name else None

        src_lines: List[str] = []
        trg_lines: List[str] = []
        notes = m.notes

        src_local = maybe_local_copy(src_path, temp_root)
        src_lines, src_err = extract_text_lines(src_local, tesseract_lang_source, tesseract_cmd, ghostscript_cmd, eps_dpi)
        if src_err:
            ocr_errors += 1
            notes = f"{notes} | Source OCR: {src_err}".strip(" |")
            if ocr_errors <= 10:
                logging.warning("Source OCR error [%s]: %s", src_path.name, src_err)

        if trg_path is not None:
            matched += 1
            trg_local = maybe_local_copy(trg_path, temp_root)
            trg_lines, trg_err = extract_text_lines(trg_local, tesseract_lang_target, tesseract_cmd, ghostscript_cmd, eps_dpi)
            if trg_err:
                ocr_errors += 1
                notes = f"{notes} | Target OCR: {trg_err}".strip(" |")
                if ocr_errors <= 10:
                    logging.warning("Target OCR error [%s]: %s", trg_path.name, trg_err)
        else:
            missing += 1

        if m.status == "LOW CONFIDENCE":
            low_conf += 1

        if use_llm_matching and trg_path is not None:
            source_block = consolidate_ocr_lines(src_lines)
            target_block = consolidate_ocr_lines(trg_lines)
            try:
                llm_matches, rejected = match_gui_strings_with_llm(
                    source_text_block=source_block,
                    target_text_block=target_block,
                    target_lang=lang_code,
                    model=openai_model,
                    api_key=openai_api_key or "",
                    timeout_sec=llm_timeout_sec,
                    max_retries=llm_max_retries,
                    source_lines=src_lines,
                )
                for rejected_note in rejected:
                    logging.warning("LLM guard rejected [%s -> %s]: %s", src_path.name, trg_path.name, rejected_note)
                if not llm_matches:
                    llm_matches = [
                        LLMGuiMatch(
                            source_string=source_block,
                            target_string="",
                            confidence=0.0,
                            status="unmatched",
                            notes="LLM returned no matches.",
                        )
                    ]
                for llm_match in llm_matches:
                    if strict_text_only and not is_meaningful_text(llm_match.source_string):
                        continue
                    status = "OK" if llm_match.status == "matched" else llm_match.status.upper()
                    row_notes = f"{status}"
                    if notes:
                        row_notes = f"{row_notes} | {notes}"
                    if llm_match.notes:
                        row_notes = f"{row_notes} | {llm_match.notes}"
                    add_report_row(
                        m=m,
                        src_path=src_path,
                        trg_path=trg_path,
                        lang_code=lang_code,
                        source_str=llm_match.source_string,
                        target_str=llm_match.target_string,
                        method="llm_gui_match",
                        confidence=llm_match.confidence,
                        status_notes=row_notes,
                        llm_model=openai_model,
                        llm_source_span=llm_match.source_span,
                        llm_target_span=llm_match.target_span,
                        llm_rationale=llm_match.rationale,
                    )
            except Exception as exc:
                ocr_errors += 1
                err = f"LLM_ERROR: {exc}"
                logging.exception("LLM matching failed [%s -> %s]", src_path.name, trg_path.name if trg_path else "")
                add_report_row(
                    m=m,
                    src_path=src_path,
                    trg_path=trg_path,
                    lang_code=lang_code,
                    source_str=consolidate_ocr_lines(src_lines),
                    target_str="",
                    method="llm_gui_match",
                    confidence=0.0,
                    status_notes=f"{m.status} | {notes} | {err}".strip(" |"),
                    llm_model=openai_model,
                )
        else:
            aligned = align_lines(src_lines, trg_lines)
            if not aligned:
                aligned = [("", "")]
            for source_str, target_str in aligned:
                if strict_text_only and not (is_meaningful_text(source_str) and is_meaningful_text(target_str)):
                    continue
                add_report_row(
                    m=m,
                    src_path=src_path,
                    trg_path=trg_path,
                    lang_code=lang_code,
                    source_str=source_str,
                    target_str=target_str,
                    method=m.method,
                    confidence=m.confidence,
                    status_notes=m.status if not notes else f"{m.status} | {notes}",
                )
        if progress_every > 0 and (idx % progress_every == 0 or idx == total):
            logging.info("Progress: %d/%d image pairs processed.", idx, total)

    logging.info("Source image count: %d", len(source_files))
    logging.info("Matched image count: %d", matched)
    logging.info("Low-confidence matches: %d", low_conf)
    logging.info("Missing matches: %d", missing)
    logging.info("OCR errors: %d", ocr_errors)

    if temp_dir_ctx:
        temp_dir_ctx.cleanup()

    return pd.DataFrame(rows)


def language_name_from_dir(lang_code: str, lang_dir: Path) -> str:
    m = LANG_FOLDER_PATTERN.match(lang_dir.name)
    return m.group("name") if m else lang_code


def should_run_classic(ocr_engine: str, keep_optional_classic: bool = True) -> bool:
    return ocr_engine in {"classic", "classic_llm", "hybrid", "triangulated"} or keep_optional_classic


def should_run_llm_normalization(ocr_engine: str, explicit: bool) -> bool:
    return ocr_engine in {"classic_llm", "triangulated"} or explicit


def should_run_ai_ocr(ocr_engine: str) -> bool:
    return ocr_engine in {"ai", "hybrid", "triangulated"}


def pipeline_report(
    source_files: List[Path],
    lang_code: str,
    lang_name: str,
    target_dir: Path,
    tesseract_lang_target: str,
    tesseract_lang_source: str,
    tesseract_cmd: str,
    ghostscript_cmd: Optional[str],
    eps_dpi: int,
    use_temp_local: bool,
    fuzzy_threshold: float,
    progress_every: int,
    allow_fuzzy: bool,
    reject_ambiguous_fuzzy: bool,
    ocr_engine: str,
    matching_mode: str,
    openai_api_key: Optional[str],
    llm_ocr_model: str,
    llm_ocr_timeout_sec: int,
    llm_ocr_max_retries: int,
    llm_ocr_cache_dir: Path,
    llm_ocr_use_cache: bool,
    ai_ocr_model: str,
    ai_ocr_detail: str,
    ai_ocr_timeout_sec: int,
    ai_ocr_max_retries: int,
    ai_ocr_cache_dir: Path,
    ai_ocr_use_cache: bool,
    object_match_model: str,
    object_match_timeout_sec: int,
    object_match_max_retries: int,
    fallback_to_classic: bool,
    use_llm_ocr_normalization: bool,
    write_rendered_images: bool,
    rendered_image_dir: Path,
    classic_ocr_min_conf: int,
    image_name_filter: Optional[str] = None,
    audit_log: Optional[AIAuditLog] = None,
) -> PipelineReport:
    target_files = list_images(target_dir, SUPPORTED_EXTENSIONS)
    if image_name_filter:
        source_files = [p for p in source_files if p.name.lower() == image_name_filter.lower()]
        target_files = [p for p in target_files if p.name.lower() == image_name_filter.lower() or language_agnostic_key(p.name) == language_agnostic_key(image_name_filter)]
    matches = match_source_to_target(
        source_files,
        target_files,
        fuzzy_threshold=fuzzy_threshold,
        allow_fuzzy=allow_fuzzy,
        reject_ambiguous_fuzzy=reject_ambiguous_fuzzy,
    )

    run_classic = should_run_classic(ocr_engine, keep_optional_classic=True)
    run_llm = should_run_llm_normalization(ocr_engine, explicit=use_llm_ocr_normalization)
    run_ai = should_run_ai_ocr(ocr_engine)
    ai_required = ocr_engine in {"ai", "hybrid", "triangulated"}
    llm_required = ocr_engine in {"classic_llm", "triangulated"} or matching_mode == "llm_objects"
    if (ai_required or llm_required) and not openai_api_key:
        if fallback_to_classic:
            logging.warning("OpenAI API key missing; falling back to classic OCR where possible.")
            run_llm = False
            run_ai = False
            matching_mode = "positional"
        else:
            raise ValueError("AI processing requires an OpenAI API key. Pass --openai-api-key or set OPENAI_API_KEY.")

    temp_dir_ctx = tempfile.TemporaryDirectory(prefix="ocr_local_") if use_temp_local else None
    temp_root = Path(temp_dir_ctx.name) if temp_dir_ctx else None
    render_root = rendered_image_dir if write_rendered_images else Path(tempfile.mkdtemp(prefix="ocr_render_batch_"))
    render_root.mkdir(parents=True, exist_ok=True)

    image_pairs: List[ImagePair] = []
    all_lines: List[OCRLine] = []
    all_tokens: List[OCRToken] = []
    llm_results: List[OCRNormalizedResult] = []
    ai_results: List[OCRNormalizedResult] = []
    triangulation_rows: List[OCRTriangulationRow] = []
    final_matches: List[GUIMatch] = []
    cache_counts = {"llm_hit": 0, "llm_miss": 0, "ai_hit": 0, "ai_miss": 0}
    error_counts = {"classic": 0, "llm": 0, "ai": 0}

    src_by_name = {p.name: p for p in source_files}
    trg_by_name = {p.name: p for p in target_files}
    total = len(matches)
    for idx, m in enumerate(matches, start=1):
        pair_id = f"{idx:04d}"
        src_path = src_by_name[m.source_name]
        trg_path = trg_by_name.get(m.target_name) if m.target_name else None
        src_local = maybe_local_copy(src_path, temp_root)
        trg_local = maybe_local_copy(trg_path, temp_root) if trg_path else None

        pair = ImagePair(
            pair_id=pair_id,
            source_image_filename=src_path.name,
            source_image_path=str(src_path),
            target_language_code=lang_code,
            target_language_name=lang_name,
            target_image_filename=trg_path.name if trg_path else None,
            target_image_path=str(trg_path) if trg_path else None,
            image_match_method=m.method,
            image_match_confidence=m.confidence,
            image_match_status=m.status,
            alternatives=m.alternatives,
            notes=m.notes,
        )

        src_rendered: Optional[Path] = None
        trg_rendered: Optional[Path] = None
        if run_ai or write_rendered_images:
            src_rendered, src_meta = render_image_for_ocr(src_local, ghostscript_cmd, eps_dpi, render_root, write_rendered_images)
            pair.source_rendered_png = str(src_rendered) if src_rendered else ""
            if src_meta.get("error"):
                pair.notes = f"{pair.notes} | Source render: {src_meta['error']}".strip(" |")
            if trg_local:
                trg_rendered, trg_meta = render_image_for_ocr(trg_local, ghostscript_cmd, eps_dpi, render_root, write_rendered_images)
                pair.target_rendered_png = str(trg_rendered) if trg_rendered else ""
                if trg_meta.get("error"):
                    pair.notes = f"{pair.notes} | Target render: {trg_meta['error']}".strip(" |")

        src_classic = ClassicOCRResult([], [], [], "", [])
        trg_classic = ClassicOCRResult([], [], [], "", [])
        if run_classic:
            src_classic, src_err = extract_classic_ocr_evidence(
                src_local,
                tesseract_lang_source,
                tesseract_cmd,
                ghostscript_cmd,
                eps_dpi,
                min_conf=classic_ocr_min_conf,
                pair_id=pair_id,
                side="source",
            )
            if src_err:
                error_counts["classic"] += 1
                pair.classic_ocr_status = f"source_error: {src_err}"
                logging.warning("Source classic OCR error [%s]: %s", src_path.name, src_err)
            all_lines.extend(src_classic.classic_lines)
            all_tokens.extend(src_classic.classic_tokens)
            if trg_local:
                trg_classic, trg_err = extract_classic_ocr_evidence(
                    trg_local,
                    tesseract_lang_target,
                    tesseract_cmd,
                    ghostscript_cmd,
                    eps_dpi,
                    min_conf=classic_ocr_min_conf,
                    pair_id=pair_id,
                    side="target",
                )
                if trg_err:
                    error_counts["classic"] += 1
                    pair.classic_ocr_status = f"{pair.classic_ocr_status} target_error: {trg_err}".strip()
                    logging.warning("Target classic OCR error [%s]: %s", trg_path.name if trg_path else "", trg_err)
                all_lines.extend(trg_classic.classic_lines)
                all_tokens.extend(trg_classic.classic_tokens)
            if not pair.classic_ocr_status:
                pair.classic_ocr_status = "ok"

        src_llm: Optional[OCRNormalizedResult] = None
        trg_llm: Optional[OCRNormalizedResult] = None
        if run_llm:
            try:
                src_llm = normalize_ocr_with_llm(
                    src_path.name,
                    "en",
                    "English",
                    "source",
                    src_classic,
                    llm_ocr_model,
                    openai_api_key or "",
                    llm_ocr_timeout_sec,
                    llm_ocr_max_retries,
                    llm_ocr_cache_dir,
                    llm_ocr_use_cache,
                    audit=audit_log,
                    pair_id=pair_id,
                )
                cache_counts["llm_hit" if src_llm.cached else "llm_miss"] += 1
                llm_results.append(src_llm)
                if trg_local:
                    trg_llm = normalize_ocr_with_llm(
                        trg_path.name if trg_path else "",
                        lang_code,
                        lang_name,
                        "target",
                        trg_classic,
                        llm_ocr_model,
                        openai_api_key or "",
                        llm_ocr_timeout_sec,
                        llm_ocr_max_retries,
                        llm_ocr_cache_dir,
                        llm_ocr_use_cache,
                        audit=audit_log,
                        pair_id=pair_id,
                    )
                    cache_counts["llm_hit" if trg_llm.cached else "llm_miss"] += 1
                    llm_results.append(trg_llm)
                pair.llm_ocr_normalization_status = "ok"
            except Exception as exc:
                error_counts["llm"] += 1
                pair.llm_ocr_normalization_status = f"error: {exc}"
                logging.exception("LLM OCR-normalization failed for pair %s", pair_id)
                if not fallback_to_classic:
                    raise

        src_ai: Optional[OCRNormalizedResult] = None
        trg_ai: Optional[OCRNormalizedResult] = None
        if run_ai:
            try:
                if src_rendered:
                    src_ai = ai_image_ocr(
                        src_rendered,
                        src_path.name,
                        "en",
                        "English",
                        "source",
                        ai_ocr_model,
                        ai_ocr_detail,
                        openai_api_key or "",
                        ai_ocr_timeout_sec,
                        ai_ocr_max_retries,
                        ai_ocr_cache_dir,
                        ai_ocr_use_cache,
                        audit=audit_log,
                        pair_id=pair_id,
                    )
                    cache_counts["ai_hit" if src_ai.cached else "ai_miss"] += 1
                    ai_results.append(src_ai)
                if trg_rendered and trg_path:
                    trg_ai = ai_image_ocr(
                        trg_rendered,
                        trg_path.name,
                        lang_code,
                        lang_name,
                        "target",
                        ai_ocr_model,
                        ai_ocr_detail,
                        openai_api_key or "",
                        ai_ocr_timeout_sec,
                        ai_ocr_max_retries,
                        ai_ocr_cache_dir,
                        ai_ocr_use_cache,
                        audit=audit_log,
                        pair_id=pair_id,
                    )
                    cache_counts["ai_hit" if trg_ai.cached else "ai_miss"] += 1
                    ai_results.append(trg_ai)
                pair.ai_image_ocr_status = "ok"
            except Exception as exc:
                error_counts["ai"] += 1
                pair.ai_image_ocr_status = f"error: {exc}"
                logging.exception("AI image OCR failed for pair %s", pair_id)
                if not fallback_to_classic:
                    raise

        if not (src_llm or src_ai):
            src_llm = classic_fallback_result(src_path.name, "en", "English", "source", src_classic)
            llm_results.append(src_llm)
        if trg_path and not (trg_llm or trg_ai):
            trg_llm = classic_fallback_result(trg_path.name, lang_code, lang_name, "target", trg_classic)
            llm_results.append(trg_llm)

        src_tri = triangulate_side(pair_id, "source", src_path.name, src_classic, src_llm, src_ai)
        trg_tri = triangulate_side(pair_id, "target", trg_path.name if trg_path else "", trg_classic, trg_llm, trg_ai) if trg_path else []
        triangulation_rows.extend(src_tri)
        triangulation_rows.extend(trg_tri)

        if trg_path:
            if matching_mode == "llm_objects" and openai_api_key:
                try:
                    final_matches.extend(
                        match_objects_with_llm(
                            pair_id,
                            src_path.name,
                            trg_path.name,
                            lang_code,
                            lang_name,
                            src_tri,
                            trg_tri,
                            object_match_model,
                            openai_api_key,
                            object_match_timeout_sec,
                            object_match_max_retries,
                            audit=audit_log,
                        )
                    )
                except Exception as exc:
                    logging.exception("Object matching failed for pair %s; using deterministic fallback.", pair_id)
                    final_matches.extend(deterministic_object_matches(pair_id, src_path.name, trg_path.name, src_tri, trg_tri))
            elif matching_mode == "llm_text" and openai_api_key:
                try:
                    final_matches.extend(
                        legacy_llm_text_matches(
                            pair_id,
                            src_path.name,
                            trg_path.name,
                            src_tri,
                            trg_tri,
                            lang_code,
                            object_match_model,
                            openai_api_key,
                            object_match_timeout_sec,
                            object_match_max_retries,
                        )
                    )
                except Exception:
                    logging.exception("LLM text matching failed for pair %s; using deterministic fallback.", pair_id)
                    final_matches.extend(deterministic_object_matches(pair_id, src_path.name, trg_path.name, src_tri, trg_tri))
            else:
                final_matches.extend(deterministic_object_matches(pair_id, src_path.name, trg_path.name, src_tri, trg_tri))

        pair.pair_status = "needs_review" if any(r.needs_review for r in src_tri + trg_tri) else "ok"
        image_pairs.append(pair)
        if progress_every > 0 and (idx % progress_every == 0 or idx == total):
            logging.info("Progress: %d/%d image pairs processed.", idx, total)

    if temp_dir_ctx:
        temp_dir_ctx.cleanup()

    summary = {
        "run timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "script version": SCRIPT_VERSION,
        "ocr_engine": ocr_engine,
        "matching_mode": matching_mode,
        "classic OCR enabled": run_classic,
        "LLM OCR-normalization enabled": run_llm,
        "LLM OCR-normalization model": llm_ocr_model,
        "LLM OCR-normalization prompt version": LLM_OCR_NORMALIZATION_PROMPT_VERSION,
        "AI image OCR enabled": run_ai,
        "AI image OCR model": ai_ocr_model,
        "AI image OCR detail": ai_ocr_detail,
        "AI image OCR prompt version": AI_IMAGE_OCR_PROMPT_VERSION,
        "GUI match prompt version": GUI_OBJECT_MATCH_PROMPT_VERSION,
        "source image count": len(source_files),
        "matched image pair count": sum(1 for p in image_pairs if p.target_image_filename),
        "missing image pair count": sum(1 for p in image_pairs if not p.target_image_filename),
        "classic OCR error count": error_counts["classic"],
        "LLM OCR-normalization error count": error_counts["llm"],
        "AI image OCR error count": error_counts["ai"],
        "triangulation disagreement count": sum(1 for r in triangulation_rows if r.agreement_status == "disagree"),
        "final matched string count": sum(1 for m in final_matches if m.status == "matched"),
        "needs_review count": sum(1 for m in final_matches if m.status == "needs_review") + sum(1 for r in triangulation_rows if r.needs_review),
        "uncertain count": sum(1 for m in final_matches if m.status == "uncertain"),
        "unmatched count": sum(1 for m in final_matches if m.status == "unmatched"),
        "LLM normalization cache hit count": cache_counts["llm_hit"],
        "LLM normalization cache miss count": cache_counts["llm_miss"],
        "AI OCR cache hit count": cache_counts["ai_hit"],
        "AI OCR cache miss count": cache_counts["ai_miss"],
    }
    return PipelineReport(image_pairs, all_lines, all_tokens, llm_results, ai_results, triangulation_rows, final_matches, summary)


def rows_from_normalized_results(results: Sequence[OCRNormalizedResult], engine: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for result in results:
        if result.source_engine != engine:
            continue
        for obj in result.ui_objects:
            rows.append(
                {
                    "side": result.side,
                    "image_filename": result.image_filename,
                    "language_code": result.language_code,
                    "language_name": result.language_name,
                    "model": result.model,
                    "prompt_version": result.prompt_version,
                    "object_id": obj.object_id,
                    "normalized_text": obj.normalized_text,
                    "raw_evidence": "; ".join(x for x in obj.raw_evidence if x),
                    "raw_visual_evidence": obj.raw_visual_evidence,
                    "gui_role": obj.gui_role,
                    "row_group": obj.row_group,
                    "screen_area": obj.screen_area,
                    "reading_order": obj.reading_order,
                    "is_translatable_gui_string": obj.is_translatable_gui_string,
                    "correction_type": obj.correction_type,
                    "ocr_evidence_confidence": obj.ocr_evidence_confidence,
                    "normalization_confidence": obj.normalization_confidence,
                    "needs_review": obj.needs_review,
                    "review_reason": obj.review_reason,
                    "rationale": obj.rationale,
                    "cached": result.cached,
                    "warnings": "; ".join(result.warnings),
                }
            )
    return rows


def write_multi_sheet_report(report: PipelineReport, output: Path) -> None:
    def _sanitize_excel_value(value: Any) -> Any:
        if isinstance(value, str):
            return ILLEGAL_CHARACTERS_RE.sub("", value)
        return value

    def _sanitize_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        cleaned: List[Dict[str, Any]] = []
        for row in rows:
            cleaned.append({k: _sanitize_excel_value(v) for k, v in row.items()})
        return cleaned

    output.parent.mkdir(parents=True, exist_ok=True)
    sheets = {
        "Image pairs": _sanitize_rows([model_to_dict(p) for p in report.image_pairs]),
        "Raw OCR - Classic Lines": _sanitize_rows([model_to_dict(x) for x in report.classic_lines]),
        "Raw OCR - Classic Tokens": _sanitize_rows([model_to_dict(x) for x in report.classic_tokens]),
        "LLM OCR Normalized Objects": _sanitize_rows(rows_from_normalized_results(report.llm_results, "llm_ocr_normalization")),
        "AI Image OCR Objects": _sanitize_rows(rows_from_normalized_results(report.ai_results, "ai_image_ocr")),
        "OCR Triangulation": _sanitize_rows([model_to_dict(x) for x in report.triangulation_rows]),
        "Final Matches": _sanitize_rows([model_to_dict(x) for x in report.final_matches]),
        "Summary": _sanitize_rows([{"metric": k, "value": v} for k, v in report.summary.items()]),
    }
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, rows in sheets.items():
            pd.DataFrame(rows).to_excel(writer, sheet_name=name[:31], index=False)
    set_multi_sheet_excel_formatting(output)


def set_multi_sheet_excel_formatting(output_file: Path) -> None:
    from openpyxl import load_workbook

    wb = load_workbook(output_file)
    fills = {
        "matched": PatternFill(start_color="FFD9EAD3", end_color="FFD9EAD3", fill_type="solid"),
        "needs_review": PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid"),
        "uncertain": PatternFill(start_color="FFFCE5CD", end_color="FFFCE5CD", fill_type="solid"),
        "unmatched": PatternFill(start_color="FFF4CCCC", end_color="FFF4CCCC", fill_type="solid"),
        "error": PatternFill(start_color="FFD9D2E9", end_color="FFD9D2E9", fill_type="solid"),
    }
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            header = str(col[0].value or "")
            max_len = min(60, max([len(str(cell.value or "")) for cell in col[:100]] + [len(header)]))
            ws.column_dimensions[get_column_letter(col[0].column)].width = max(10, max_len + 2)
            if any(k in header.lower() for k in ("text", "evidence", "reason", "rationale", "notes", "path")):
                ws.column_dimensions[get_column_letter(col[0].column)].width = min(70, max(28, max_len + 2))
                for cell in col[1:]:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
        header_to_col = {str(cell.value or "").lower(): cell.column for cell in ws[1]}
        status_col = header_to_col.get("status") or header_to_col.get("image_match_status") or header_to_col.get("pair_status")
        review_col = header_to_col.get("needs_review")
        max_col = get_column_letter(ws.max_column)
        if status_col and ws.max_row >= 2:
            letter = get_column_letter(status_col)
            for key, fill in fills.items():
                ws.conditional_formatting.add(
                    f"A2:{max_col}{ws.max_row}",
                    FormulaRule(formula=[f'ISNUMBER(SEARCH("{key}",${letter}2))'], stopIfTrue=False, fill=fill),
                )
        if review_col and ws.max_row >= 2:
            letter = get_column_letter(review_col)
            ws.conditional_formatting.add(
                f"A2:{max_col}{ws.max_row}",
                FormulaRule(formula=[f"${letter}2=TRUE"], stopIfTrue=False, fill=fills["needs_review"]),
            )
    wb.save(output_file)


def main() -> None:
    parser = argparse.ArgumentParser(description="Match EN screenshots to target images, OCR both, and export Excel report.")
    parser.add_argument(
        "--root",
        type=Path,
        default=Path("."),
        help="Project root folder. Can be workspace root, ACROSS_* root, or image folder.",
    )
    parser.add_argument("--source-prefix", default="enis", help="Source filename prefix (default: enis).")
    parser.add_argument("--target-lang", required=True, help="Target language code matching folder suffix, e.g. hu/sk/sl.")
    parser.add_argument("--target-ocr-lang", default="eng", help="Tesseract language for target OCR (e.g. hun, slk, slv).")
    parser.add_argument("--source-ocr-lang", default="eng", help="Tesseract language for source OCR.")
    parser.add_argument("--tesseract-cmd", default=None, help="Optional full path to tesseract.exe")
    parser.add_argument("--ghostscript-cmd", default=None, help="Optional full path to gswin64c.exe")
    parser.add_argument("--eps-dpi", type=int, default=600, help="EPS rasterization DPI when using Ghostscript.")
    parser.add_argument("--progress-every", type=int, default=10, help="Log progress every N image pairs.")
    parser.add_argument("--fuzzy-threshold", type=float, default=0.72, help="Fallback fuzzy threshold (0.0-1.0).")
    parser.add_argument("--allow-fuzzy", action="store_true", help="Allow fuzzy filename matching when exact match is missing.")
    parser.add_argument(
        "--reject-ambiguous-fuzzy",
        action="store_true",
        help="When fuzzy is enabled, reject ambiguous top candidates instead of accepting low confidence.",
    )
    parser.add_argument(
        "--strict-text-only",
        action="store_true",
        help="Keep only rows where both source and target strings look like meaningful UI text.",
    )
    parser.add_argument("--ocr-engine", choices=["classic", "classic_llm", "ai", "hybrid", "triangulated"], default="classic")
    parser.add_argument("--report-format", choices=["legacy", "multi_sheet"], default="legacy")
    parser.add_argument("--matching-mode", choices=["positional", "llm_text", "llm_objects"], default="positional")
    parser.add_argument("--classic-ocr-min-conf", type=int, default=45)
    parser.add_argument("--classic-ocr-keep-all-passes", action="store_true")
    parser.add_argument("--use-llm-ocr-normalization", action="store_true")
    parser.add_argument("--llm-ocr-normalization-model", default=DEFAULT_AI_MODEL)
    parser.add_argument("--llm-ocr-normalization-timeout-sec", type=int, default=60)
    parser.add_argument("--llm-ocr-normalization-max-retries", type=int, default=2)
    parser.add_argument("--llm-ocr-normalization-cache-dir", type=Path, default=Path(".ocr_llm_norm_cache"))
    parser.add_argument("--llm-ocr-normalization-no-cache", action="store_true")
    parser.add_argument("--ai-ocr-model", default=DEFAULT_AI_MODEL)
    parser.add_argument("--ai-ocr-detail", choices=["low", "high", "auto", "original"], default="high")
    parser.add_argument("--ai-ocr-timeout-sec", type=int, default=60)
    parser.add_argument("--ai-ocr-max-retries", type=int, default=2)
    parser.add_argument("--ai-ocr-cache-dir", type=Path, default=Path(".ocr_ai_cache"))
    parser.add_argument("--ai-ocr-no-cache", action="store_true")
    parser.add_argument("--compare-ocr-engines", action="store_true")
    parser.add_argument("--fallback-to-classic", action="store_true")
    parser.add_argument("--write-rendered-images", action="store_true")
    parser.add_argument("--rendered-image-dir", type=Path, default=Path("rendered_images"))
    parser.add_argument("--ai-audit-html", type=Path, default=None, help="HTML file for readable AI request/response audit logging.")
    parser.add_argument("--disable-ai-audit-html", action="store_true", help="Disable AI request/response HTML audit generation.")
    parser.add_argument("--image-name", default=None, help="Optional source image filename filter for test runs.")
    parser.add_argument("--use-llm-matching", action="store_true", help="Use OpenAI LLM matching instead of positional line alignment.")
    parser.add_argument("--openai-model", default="gpt-4.1-mini", help="OpenAI model for LLM GUI string matching.")
    parser.add_argument("--openai-api-key", default=None, help="OpenAI API key. Falls back to OPENAI_API_KEY.")
    parser.add_argument("--llm-timeout-sec", type=int, default=45, help="Timeout for each LLM request.")
    parser.add_argument("--llm-max-retries", type=int, default=2, help="Maximum retry count for each LLM request.")
    parser.add_argument("--output", type=Path, default=Path("ocr_match_report.xlsx"), help="Excel output path.")
    parser.add_argument("--log-file", type=Path, default=Path("ocr_match_report.log"), help="Log file path.")
    parser.add_argument("--use-temp-local-copy", action="store_true", help="Copy files to local temp folder before OCR.")
    parser.add_argument("--verbose", action="store_true", help="Also print log messages to console.")
    args = parser.parse_args()

    configure_logging(args.log_file, args.verbose)
    logging.info("Starting discovery under root: %s", args.root.resolve())

    openai_api_key = args.openai_api_key or os.environ.get("OPENAI_API_KEY")
    if args.use_llm_matching and not openai_api_key:
        raise ValueError(
            "LLM matching requires an OpenAI API key. Pass --openai-api-key or set OPENAI_API_KEY."
        )
    if args.use_llm_matching:
        logging.info("LLM GUI matching enabled with model: %s", args.openai_model)
    if args.ocr_engine != "classic" and args.report_format == "legacy":
        logging.info("Non-classic OCR engine selected; switching report format to multi_sheet.")
        args.report_format = "multi_sheet"
    if args.ocr_engine in {"classic_llm", "ai", "hybrid", "triangulated"} and args.matching_mode == "positional":
        logging.info("Consider --matching-mode llm_objects for structured OCR modes.")
    logging.info("OCR engine: %s", args.ocr_engine)
    logging.info("Report format: %s", args.report_format)
    logging.info("Matching mode: %s", args.matching_mode)

    tesseract_cmd = resolve_executable(
        args.tesseract_cmd,
        ["tesseract", r"C:\Program Files\Tesseract-OCR\tesseract.exe"],
    )
    if not tesseract_cmd:
        raise FileNotFoundError(
            "Tesseract executable not found. Install Tesseract or pass --tesseract-cmd."
        )
    logging.info("Using Tesseract executable: %s", tesseract_cmd)

    ghostscript_cmd = resolve_executable(
        args.ghostscript_cmd,
        ["gswin64c", "gswin32c", r"C:\Program Files\gs\gs10.07.0\bin\gswin64c.exe"],
    )
    if ghostscript_cmd:
        logging.info("Using Ghostscript executable: %s", ghostscript_cmd)
    else:
        logging.warning("Ghostscript executable not found. EPS OCR quality may be reduced or fail.")

    image_root, lang_dirs = discover_image_layout(args.root.resolve())
    logging.info("Discovered image root: %s", image_root)
    logging.info("Discovered language folders: %s", {k: str(v) for k, v in sorted(lang_dirs.items())})

    if args.target_lang.lower() not in lang_dirs:
        raise ValueError(f"Target language '{args.target_lang}' not found. Available: {sorted(lang_dirs)}")

    source_files = infer_source_images(image_root, SUPPORTED_EXTENSIONS, args.source_prefix)
    if not source_files and "en" in lang_dirs:
        logging.info("No source images found in image root. Falling back to English(en) folder: %s", lang_dirs["en"])
        source_files = infer_source_images(lang_dirs["en"], SUPPORTED_EXTENSIONS, args.source_prefix)
    unsupported = [p for p in image_root.iterdir() if p.is_file() and p.suffix.lower() not in SUPPORTED_EXTENSIONS]
    if unsupported:
        logging.info("Unsupported/skipped files in source image root: %d", len(unsupported))

    logging.info("Inferred source matching rule: normalize first language code in filename prefix, then exact match.")
    logging.info("Variant fallback rule: accept same section/item when only the final variant suffix differs.")
    logging.info("Optional fallback rule: fuzzy filename match on language-normalized stem when --allow-fuzzy is set.")

    if args.report_format == "multi_sheet":
        lang_code = args.target_lang.lower()
        lang_name = language_name_from_dir(lang_code, lang_dirs[lang_code])
        audit_log = None if args.disable_ai_audit_html else AIAuditLog()
        report = pipeline_report(
            source_files=source_files,
            lang_code=lang_code,
            lang_name=lang_name,
            target_dir=lang_dirs[lang_code],
            tesseract_lang_target=args.target_ocr_lang,
            tesseract_lang_source=args.source_ocr_lang,
            tesseract_cmd=tesseract_cmd,
            ghostscript_cmd=ghostscript_cmd,
            eps_dpi=args.eps_dpi,
            use_temp_local=args.use_temp_local_copy,
            fuzzy_threshold=args.fuzzy_threshold,
            progress_every=args.progress_every,
            allow_fuzzy=args.allow_fuzzy,
            reject_ambiguous_fuzzy=args.reject_ambiguous_fuzzy,
            ocr_engine=args.ocr_engine,
            matching_mode=args.matching_mode,
            openai_api_key=openai_api_key,
            llm_ocr_model=args.llm_ocr_normalization_model,
            llm_ocr_timeout_sec=args.llm_ocr_normalization_timeout_sec,
            llm_ocr_max_retries=args.llm_ocr_normalization_max_retries,
            llm_ocr_cache_dir=args.llm_ocr_normalization_cache_dir,
            llm_ocr_use_cache=not args.llm_ocr_normalization_no_cache,
            ai_ocr_model=args.ai_ocr_model,
            ai_ocr_detail=args.ai_ocr_detail,
            ai_ocr_timeout_sec=args.ai_ocr_timeout_sec,
            ai_ocr_max_retries=args.ai_ocr_max_retries,
            ai_ocr_cache_dir=args.ai_ocr_cache_dir,
            ai_ocr_use_cache=not args.ai_ocr_no_cache,
            object_match_model=args.openai_model,
            object_match_timeout_sec=args.llm_timeout_sec,
            object_match_max_retries=args.llm_max_retries,
            fallback_to_classic=args.fallback_to_classic,
            use_llm_ocr_normalization=args.use_llm_ocr_normalization,
            write_rendered_images=args.write_rendered_images,
            rendered_image_dir=args.rendered_image_dir,
            classic_ocr_min_conf=args.classic_ocr_min_conf,
            image_name_filter=args.image_name,
            audit_log=audit_log,
        )
        write_multi_sheet_report(report, args.output)
        if audit_log is not None:
            audit_path = args.ai_audit_html or args.output.with_name(f"{args.output.stem}_ai_audit.html")
            audit_log.write_html(audit_path)
            logging.info("AI audit HTML written: %s", audit_path.resolve())
        logging.info("Excel report written: %s", args.output.resolve())
        return

    df = make_report(
        source_files=source_files,
        lang_code=args.target_lang.lower(),
        target_dir=lang_dirs[args.target_lang.lower()],
        tesseract_lang_target=args.target_ocr_lang,
        tesseract_lang_source=args.source_ocr_lang,
        tesseract_cmd=tesseract_cmd,
        ghostscript_cmd=ghostscript_cmd,
        eps_dpi=args.eps_dpi,
        use_temp_local=args.use_temp_local_copy,
        fuzzy_threshold=args.fuzzy_threshold,
        progress_every=args.progress_every,
        allow_fuzzy=args.allow_fuzzy,
        reject_ambiguous_fuzzy=args.reject_ambiguous_fuzzy,
        strict_text_only=args.strict_text_only,
        use_llm_matching=args.use_llm_matching,
        openai_model=args.openai_model,
        openai_api_key=openai_api_key,
        llm_timeout_sec=args.llm_timeout_sec,
        llm_max_retries=args.llm_max_retries,
    )

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    set_excel_formatting(args.output)
    logging.info("Excel report written: %s", args.output.resolve())


if __name__ == "__main__":
    main()
