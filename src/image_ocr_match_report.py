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
import sys
import tempfile
import time
import unicodedata
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
import sqlite3
import uuid
import threading
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

SCRIPT_VERSION = "2.3.0"
LLM_OCR_NORMALIZATION_PROMPT_VERSION = "llm_ocr_normalization_gui_v1"
AI_IMAGE_OCR_PROMPT_VERSION = "ai_image_ocr_gui_v1"
GUI_OBJECT_MATCH_PROMPT_VERSION = "gui_object_match_v2"
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
OBJECT_LANES = (
    "translatable_gui",
    "value_only",
    "masked_text",
    "map_background",
    "decorative_status",
    "unknown_review",
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
    source_object_lane: str = "unknown_review"
    target_object_lane: str = "unknown_review"
    raw_ocr_validation_status: Literal["confirmed", "weakly_confirmed", "not_confirmed"] = "not_confirmed"
    raw_ocr_validation_reason: str = ""
    semantic_check_status: Literal["ok", "warning", "mismatch"] = "warning"
    semantic_check_confidence: float = Field(default=0.0, ge=0.0, le=1.0)
    semantic_check_reason: str = ""
    source_guided_conflict: bool = False
    suggested_target_candidate: str = ""


class GUIObjectMatchResponse(BaseModel):
    pair_id: str
    source_image: str
    target_image: str
    matches: List[GUIMatch] = Field(default_factory=list)


class SemanticPairCheck(BaseModel):
    semantic_check_status: Literal["ok", "warning", "mismatch"] = "warning"
    semantic_check_confidence: float = Field(default=0.0, ge=0.0, le=1.0)
    semantic_check_reason: str = ""


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
    raw_ocr_full_text: List[Dict[str, Any]]
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


def _json_dumps_safe(payload: Any) -> str:
    try:
        return json.dumps(payload, ensure_ascii=False)
    except Exception:
        return json.dumps({"repr": repr(payload)}, ensure_ascii=False)


def _sdk_response_to_dict(response: Any) -> Dict[str, Any]:
    try:
        if hasattr(response, "model_dump"):
            return response.model_dump()
        if hasattr(response, "model_dump_json"):
            return json.loads(response.model_dump_json())
    except Exception:
        pass
    try:
        if hasattr(response, "to_dict"):
            return response.to_dict()
    except Exception:
        pass
    return {"repr": repr(response)}


class OpenAIAuditDB:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()
        self.conn = sqlite3.connect(str(self.path), check_same_thread=False)
        self.conn.execute("PRAGMA journal_mode=WAL;")
        self.conn.execute("PRAGMA synchronous=NORMAL;")
        self._init_schema()
        self.run_id = ""

    def _init_schema(self) -> None:
        with self._lock:
            self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS runs (
              run_id TEXT PRIMARY KEY,
              created_at TEXT NOT NULL,
              metadata_json TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS api_events (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              run_id TEXT NOT NULL,
              created_at TEXT NOT NULL,
              stage TEXT NOT NULL,
              title TEXT NOT NULL,
              model TEXT NOT NULL,
              prompt_version TEXT NOT NULL,
              cache_key TEXT,
              was_cached INTEGER NOT NULL DEFAULT 0,
              error TEXT NOT NULL DEFAULT '',
              request_json TEXT NOT NULL,
              response_json TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_api_events_stage_created ON api_events(stage, created_at);
            CREATE INDEX IF NOT EXISTS idx_api_events_cache_key ON api_events(cache_key);
            CREATE TABLE IF NOT EXISTS stage_cache (
              stage TEXT NOT NULL,
              cache_key TEXT NOT NULL,
              model_class TEXT NOT NULL,
              payload_json TEXT NOT NULL,
              metadata_json TEXT NOT NULL,
              created_at TEXT NOT NULL,
              PRIMARY KEY(stage, cache_key)
            );
            """
            )
            self.conn.commit()

    def start_run(self, metadata: Dict[str, Any]) -> str:
        run_id = f"{time.strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}"
        with self._lock:
            self.conn.execute(
                "INSERT INTO runs(run_id, created_at, metadata_json) VALUES (?, ?, ?)",
                (run_id, time.strftime("%Y-%m-%d %H:%M:%S"), _json_dumps_safe(metadata)),
            )
            self.conn.commit()
        self.run_id = run_id
        return run_id

    def log_event(
        self,
        stage: str,
        title: str,
        model: str,
        prompt_version: str,
        request: Dict[str, Any],
        response: Optional[Dict[str, Any]] = None,
        was_cached: bool = False,
        error: str = "",
        cache_key: str = "",
    ) -> None:
        with self._lock:
            self.conn.execute(
                """
                INSERT INTO api_events(
                  run_id, created_at, stage, title, model, prompt_version, cache_key, was_cached, error, request_json, response_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    self.run_id or "unknown",
                    time.strftime("%Y-%m-%d %H:%M:%S"),
                    stage,
                    title,
                    model,
                    prompt_version,
                    cache_key,
                    1 if was_cached else 0,
                    error,
                    _json_dumps_safe(request),
                    _json_dumps_safe(response or {}),
                ),
            )
            self.conn.commit()

    def cache_get_model(self, stage: str, cache_key: str, model_cls: Any) -> Optional[BaseModel]:
        with self._lock:
            row = self.conn.execute(
                "SELECT payload_json FROM stage_cache WHERE stage=? AND cache_key=?",
                (stage, cache_key),
            ).fetchone()
        if not row:
            return None
        try:
            payload = json.loads(row[0])
            return model_cls.model_validate(payload) if hasattr(model_cls, "model_validate") else model_cls.parse_obj(payload)
        except Exception as exc:
            logging.warning("Ignoring invalid DB cache entry for %s/%s: %s", stage, cache_key, exc)
            return None

    def cache_put_model(
        self,
        stage: str,
        cache_key: str,
        model_obj: BaseModel,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> None:
        with self._lock:
            self.conn.execute(
                """
                INSERT INTO stage_cache(stage, cache_key, model_class, payload_json, metadata_json, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(stage, cache_key) DO UPDATE SET
                  model_class=excluded.model_class,
                  payload_json=excluded.payload_json,
                  metadata_json=excluded.metadata_json,
                  created_at=excluded.created_at
                """,
                (
                    stage,
                    cache_key,
                    model_obj.__class__.__name__,
                    model_to_json(model_obj),
                    _json_dumps_safe(metadata or {}),
                    time.strftime("%Y-%m-%d %H:%M:%S"),
                ),
            )
            self.conn.commit()

    def close(self) -> None:
        try:
            with self._lock:
                self.conn.close()
        except Exception:
            pass


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
    logging.getLogger("httpx").setLevel(logging.WARNING)
    logging.getLogger("openai").setLevel(logging.WARNING)


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

Read the screenshot directly. Extract meaningful visible GUI text objects in the screenshot language. Do not translate and do not invent text.

Prioritize translatable UI content: labels, buttons, tabs, menu names, headers, instruction text, and short visible warning/status messages that carry user-facing meaning.
De-prioritize or ignore: standalone values (times, distances, speeds, numeric counters), masked identifiers, background map/place labels, decorative/status artifacts without user-facing meaning, and purely symbolic strings.

Keep separate GUI objects separate, join wrapped visual lines only when they form one GUI object, split merged text when needed, ignore icons/time/signal/decorative graphics, and distinguish navigation map labels from actual GUI text. Mark uncertain readings with needs_review. Return data matching the provided schema."""


def system_prompt_gui_object_match(target_language: str) -> str:
    return f"""You are a bilingual GUI string alignment specialist for vehicle infotainment screenshots.

Match English GUI objects to corresponding {target_language} GUI objects using semantic equivalence, not line position alone. Use GUI role, row group, screen area, and reading order as supporting evidence.

Consistency rules for short UI labels:
- Prefer one-to-one label alignment when plausible.
- Do not merge two English labels into one target match unless the target visibly has only one combined label.
- If target word order is reversed (for example, source labels like "Edit" and "Favourites" vs a target phrase containing equivalents in reverse order), align by lexical meaning and assign the most specific counterpart to each source label.
- If one source label cannot be isolated with confidence, return that row as uncertain/unmatched rather than forcing a merged match.
- If the target object looks merged label+description and split candidates are provided (object_id suffixes like __split_label / __split_desc), prefer matching short source labels to __split_label and longer explanatory source text to __split_desc.

Do not invent target text and do not translate the output yourself. Return unmatched or uncertain when a reliable counterpart is not present. Return data matching the provided schema."""


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
    audit_db: Optional[OpenAIAuditDB] = None,
    audit_stage: str = "",
    audit_title: str = "",
    audit_prompt_version: str = "",
    audit_request: Optional[Dict[str, Any]] = None,
    audit_cache_key: str = "",
) -> BaseModel:
    client = ensure_openai_client(api_key)
    last_error: Optional[Exception] = None
    request_for_audit = audit_request or {
        "instructions": system_prompt,
        "input": user_content,
        "schema": getattr(schema_model, "__name__", str(schema_model)),
    }
    stage_label = (audit_stage or "OpenAI structured call").strip()
    title_label = (audit_title or "").strip()
    for attempt in range(max_retries + 1):
        t0 = time.perf_counter()
        logging.info(
            "OpenAI call start | stage=%s | title=%s | model=%s | attempt=%d/%d",
            stage_label,
            title_label,
            model,
            attempt + 1,
            max_retries + 1,
        )
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
            elapsed = time.perf_counter() - t0
            logging.info(
                "OpenAI call done  | stage=%s | title=%s | model=%s | elapsed=%.2fs",
                stage_label,
                title_label,
                model,
                elapsed,
            )
            if audit is not None:
                audit.add(
                    audit_stage,
                    audit_title,
                    model,
                    audit_prompt_version,
                    request_for_audit,
                    response=model_to_dict(parsed),
                )
            if audit_db is not None:
                audit_db.log_event(
                    audit_stage,
                    audit_title,
                    model,
                    audit_prompt_version,
                    request_for_audit,
                    response={"parsed": model_to_dict(parsed), "raw_response": _sdk_response_to_dict(response)},
                    was_cached=False,
                    cache_key=audit_cache_key,
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
            elapsed = time.perf_counter() - t0
            logging.info(
                "OpenAI call done  | stage=%s | title=%s | model=%s | elapsed=%.2fs",
                stage_label,
                title_label,
                model,
                elapsed,
            )
            if audit is not None:
                audit.add(
                    audit_stage,
                    audit_title,
                    model,
                    audit_prompt_version,
                    request_for_audit,
                    response=model_to_dict(parsed),
                )
            if audit_db is not None:
                audit_db.log_event(
                    audit_stage,
                    audit_title,
                    model,
                    audit_prompt_version,
                    request_for_audit,
                    response={"parsed": model_to_dict(parsed), "raw_response": _sdk_response_to_dict(response)},
                    was_cached=False,
                    cache_key=audit_cache_key,
                )
            return parsed
        except Exception as exc:
            last_error = exc
            elapsed = time.perf_counter() - t0
            logging.warning(
                "OpenAI call error | stage=%s | title=%s | model=%s | elapsed=%.2fs | error=%s",
                stage_label,
                title_label,
                model,
                elapsed,
                exc,
            )
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
    if audit_db is not None:
        audit_db.log_event(
            audit_stage,
            audit_title,
            model,
            audit_prompt_version,
            request_for_audit,
            response=None,
            was_cached=False,
            error=str(last_error),
            cache_key=audit_cache_key,
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
    audit_db: Optional[OpenAIAuditDB] = None,
    restore_from_db_cache: bool = True,
    write_to_db_cache: bool = True,
    audit: Optional[AIAuditLog] = None,
    pair_id: str = "",
) -> OCRNormalizedResult:
    package = ocr_candidates_package(image_filename, language_code, language_name, side, evidence)
    package_json = json.dumps(package, ensure_ascii=False, sort_keys=True)
    key = cache_key_for_payload(LLM_OCR_NORMALIZATION_PROMPT_VERSION, model, package_json)
    if restore_from_db_cache and audit_db is not None:
        cached_db = audit_db.cache_get_model("llm_ocr_normalization", key, OCRNormalizedResult)
        if cached_db:
            result = cached_db  # type: ignore[assignment]
            result.cached = True
            if audit is not None:
                audit.add(
                    "LLM OCR-normalization",
                    f"{pair_id} {side} {image_filename}",
                    model,
                    LLM_OCR_NORMALIZATION_PROMPT_VERSION,
                    {"cache_key": key, "cache_backend": "sqlite", "image_filename": image_filename, "language": language_code, "side": side},
                    response=model_to_dict(result),
                    cached=True,
                )
            audit_db.log_event(
                "LLM OCR-normalization",
                f"{pair_id} {side} {image_filename}",
                model,
                LLM_OCR_NORMALIZATION_PROMPT_VERSION,
                {"cache_key": key, "cache_backend": "sqlite", "image_filename": image_filename, "language": language_code, "side": side},
                response={"parsed": model_to_dict(result)},
                was_cached=True,
                cache_key=key,
            )
            return result
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
            if audit_db is not None:
                audit_db.log_event(
                    "LLM OCR-normalization",
                    f"{pair_id} {side} {image_filename}",
                    model,
                    LLM_OCR_NORMALIZATION_PROMPT_VERSION,
                    {
                        "cache_key": key,
                        "cache_backend": "json_file",
                        "image_filename": image_filename,
                        "language": language_code,
                        "side": side,
                    },
                    response={"parsed": model_to_dict(result)},
                    was_cached=True,
                    cache_key=key,
                )
                if write_to_db_cache:
                    audit_db.cache_put_model(
                        "llm_ocr_normalization",
                        key,
                        result,
                        metadata={"model": model, "prompt_version": LLM_OCR_NORMALIZATION_PROMPT_VERSION, "side": side, "image_filename": image_filename},
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
        audit_db=audit_db,
        audit_stage="LLM OCR-normalization",
        audit_title=f"{pair_id} {side} {image_filename}",
        audit_prompt_version=LLM_OCR_NORMALIZATION_PROMPT_VERSION,
        audit_request=audit_request,
        audit_cache_key=key,
    )
    assert isinstance(result, OCRNormalizedResult)
    result.source_engine = "llm_ocr_normalization"
    result.model = model
    result.prompt_version = LLM_OCR_NORMALIZATION_PROMPT_VERSION
    result.cached = False
    if use_cache:
        write_cached_model(cache_dir, key, result)
    if write_to_db_cache and audit_db is not None:
        audit_db.cache_put_model(
            "llm_ocr_normalization",
            key,
            result,
            metadata={"model": model, "prompt_version": LLM_OCR_NORMALIZATION_PROMPT_VERSION, "side": side, "image_filename": image_filename},
        )
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
    cache_only: bool = False,
    audit_db: Optional[OpenAIAuditDB] = None,
    restore_from_db_cache: bool = True,
    write_to_db_cache: bool = True,
    audit: Optional[AIAuditLog] = None,
    pair_id: str = "",
) -> OCRNormalizedResult:
    image_hash = file_sha256(rendered_png)
    key = cache_key_for_payload(AI_IMAGE_OCR_PROMPT_VERSION, model, detail, image_hash, image_filename, side, language_code)
    if restore_from_db_cache and audit_db is not None:
        cached_db = audit_db.cache_get_model("ai_image_ocr", key, OCRNormalizedResult)
        if cached_db:
            result = cached_db  # type: ignore[assignment]
            result.cached = True
            if audit is not None:
                audit.add(
                    "AI image OCR",
                    f"{pair_id} {side} {image_filename}",
                    model,
                    AI_IMAGE_OCR_PROMPT_VERSION,
                    {"cache_key": key, "cache_backend": "sqlite", "rendered_png": str(rendered_png), "image_hash": image_hash, "detail": detail},
                    response=model_to_dict(result),
                    cached=True,
                )
            audit_db.log_event(
                "AI image OCR",
                f"{pair_id} {side} {image_filename}",
                model,
                AI_IMAGE_OCR_PROMPT_VERSION,
                {"cache_key": key, "cache_backend": "sqlite", "rendered_png": str(rendered_png), "image_hash": image_hash, "detail": detail},
                response={"parsed": model_to_dict(result)},
                was_cached=True,
                cache_key=key,
            )
            return result
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
            if audit_db is not None:
                audit_db.log_event(
                    "AI image OCR",
                    f"{pair_id} {side} {image_filename}",
                    model,
                    AI_IMAGE_OCR_PROMPT_VERSION,
                    {
                        "cache_key": key,
                        "cache_backend": "json_file",
                        "rendered_png": str(rendered_png),
                        "image_hash": image_hash,
                        "detail": detail,
                    },
                    response={"parsed": model_to_dict(result)},
                    was_cached=True,
                    cache_key=key,
                )
                if write_to_db_cache:
                    audit_db.cache_put_model(
                        "ai_image_ocr",
                        key,
                        result,
                        metadata={"model": model, "prompt_version": AI_IMAGE_OCR_PROMPT_VERSION, "side": side, "image_filename": image_filename, "image_hash": image_hash},
                    )
            return result
    if cache_only:
        raise RuntimeError(
            f"AI OCR cache miss for {image_filename} ({side}); cache-only mode will not call OpenAI."
        )
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
        audit_db=audit_db,
        audit_stage="AI image OCR",
        audit_title=f"{pair_id} {side} {image_filename}",
        audit_prompt_version=AI_IMAGE_OCR_PROMPT_VERSION,
        audit_request=audit_request,
        audit_cache_key=key,
    )
    assert isinstance(result, OCRNormalizedResult)
    result.source_engine = "ai_image_ocr"
    result.model = model
    result.prompt_version = AI_IMAGE_OCR_PROMPT_VERSION
    result.image_hash = image_hash
    result.cached = False
    if use_cache:
        write_cached_model(cache_dir, key, result)
    if write_to_db_cache and audit_db is not None:
        audit_db.cache_put_model(
            "ai_image_ocr",
            key,
            result,
            metadata={"model": model, "prompt_version": AI_IMAGE_OCR_PROMPT_VERSION, "side": side, "image_filename": image_filename, "image_hash": image_hash},
        )
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


def is_value_only_text(text: str) -> bool:
    s = clean_line(text)
    if not s:
        return True
    low = s.lower()
    if re.fullmatch(r"[0-9:./\-+% ]+", s):
        return True
    if re.search(r"\b(km|m|min|perc|am|pm)\b", low):
        letters = sum(ch.isalpha() for ch in s)
        digits = sum(ch.isdigit() for ch in s)
        if digits >= letters:
            return True
    if re.fullmatch(r"[0-9]+\s*(km|m|min|perc)", low):
        return True
    return False


def is_short_translatable_status_text(raw_text: str, obj: Optional[NormalizedGUIObject] = None) -> bool:
    text = clean_line(raw_text)
    if not text or is_value_only_text(text) or re.search(r"\d", text):
        return False
    tokens = [t for t in text.split() if t]
    if len(tokens) > 3:
        return False
    letters = [ch for ch in text if ch.isalpha()]
    if len(letters) < 2 or len(letters) > 24:
        return False
    has_warning_punctuation = bool(re.search(r"[!?]", raw_text))
    mostly_upper = not any(ch.islower() for ch in letters)
    if has_warning_punctuation:
        return True
    if obj is None:
        return False
    confidence = max(float(obj.ocr_evidence_confidence), float(obj.normalization_confidence))
    return bool(obj.is_translatable_gui_string and obj.gui_role == "status_bar" and mostly_upper and confidence >= 0.75)


def classify_object_lane(obj: NormalizedGUIObject) -> str:
    raw_text = "" if obj.normalized_text is None else str(obj.normalized_text)
    text = clean_line(raw_text)
    if not text:
        if re.search(r"\*{2,}", raw_text):
            return "masked_text"
        return "unknown_review"
    if "*" in raw_text or re.fullmatch(r"[*xX•\-. ]+", raw_text):
        return "masked_text"
    if obj.gui_role in {"map_label"}:
        return "map_background"
    if obj.gui_role in {"status_bar"}:
        if is_short_translatable_status_text(raw_text, obj):
            return "translatable_gui"
        return "decorative_status"
    if is_value_only_text(text):
        return "value_only"
    if obj.gui_role in {"title", "breadcrumb", "menu_label", "description", "button", "tab"}:
        return "translatable_gui"
    if obj.gui_role == "value":
        return "value_only"
    if is_meaningful_text(text):
        return "translatable_gui"
    return "unknown_review"


def is_label_like_value_text(text: str) -> bool:
    s = clean_line(text)
    if not s:
        return False
    if is_short_translatable_status_text(text):
        return True
    if is_value_only_text(s):
        return False
    if re.search(r"\d", s):
        return False
    tokens = [t for t in s.split() if t]
    if not tokens:
        return False
    if len(tokens) > 4:
        return False
    letters = sum(ch.isalpha() for ch in s)
    if letters < 4:
        return False
    return True


def extract_label_candidates(text: str) -> List[str]:
    s = clean_line(text)
    if not s:
        return []
    out: List[str] = [s]
    stripped = re.sub(r"^[*xX•._\- ]+", "", s).strip()
    if stripped and stripped != s:
        out.append(stripped)
    dedup: List[str] = []
    seen = set()
    for item in out:
        norm = normalize_for_llm_guard(item).lower()
        if not norm or norm in seen:
            continue
        if is_label_like_value_text(item):
            dedup.append(item)
            seen.add(norm)
    return dedup


def normalize_text_key(text: str) -> str:
    return remove_accents(clean_line(normalize_for_llm_guard(text))).lower()


def read_om_english_strings(om_xlsx: Path) -> List[str]:
    strings: List[str] = []
    if not om_xlsx.exists():
        return strings
    try:
        xls = pd.ExcelFile(om_xlsx)
    except Exception as exc:
        logging.warning("Failed to open MM strings workbook %s: %s", om_xlsx, exc)
        return strings
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(om_xlsx, sheet_name=sheet)
        except Exception as exc:
            logging.warning("Failed reading OM sheet %s: %s", sheet, exc)
            continue
        if df.empty:
            continue
        cols = [str(c).strip() for c in df.columns]
        lower_cols = [c.lower() for c in cols]
        en_col_idx = -1
        for i, c in enumerate(lower_cols):
            if c in {"en", "english", "source", "source_en", "om_en"}:
                en_col_idx = i
                break
        if en_col_idx < 0:
            for i, c in enumerate(lower_cols):
                if "en" in c or "english" in c:
                    en_col_idx = i
                    break
        if en_col_idx < 0:
            en_col_idx = 0
        col = cols[en_col_idx]
        for val in df[col].tolist():
            s = normalize_for_llm_guard("" if val is None else str(val))
            if s:
                strings.append(s)
    dedup: List[str] = []
    seen = set()
    for s in strings:
        key = normalize_text_key(s)
        if not key or key in seen:
            continue
        seen.add(key)
        dedup.append(s)
    return dedup


def build_om_strings_rows(report: "PipelineReport", om_xlsx: Path) -> List[Dict[str, Any]]:
    om_strings = read_om_english_strings(om_xlsx)
    if not om_strings:
        return []

    by_source: Dict[str, List[GUIMatch]] = {}
    for m in report.final_matches:
        key = normalize_text_key(m.source_text)
        if not key:
            continue
        by_source.setdefault(key, []).append(m)

    def _status_rank(status: str) -> int:
        # Higher is better confidence/quality.
        return {"matched": 4, "needs_review": 3, "uncertain": 2, "unmatched": 1}.get(str(status), 0)

    def _merge_reasons(matches: List[GUIMatch], field: str) -> str:
        vals: List[str] = []
        seen = set()
        for m in matches:
            v = normalize_for_llm_guard(str(getattr(m, field, "") or ""))
            if not v:
                continue
            key = v.lower()
            if key in seen:
                continue
            seen.add(key)
            vals.append(v)
        return " | ".join(vals)

    rows: List[Dict[str, Any]] = []
    for s in om_strings:
        key = normalize_text_key(s)
        hits = by_source.get(key, [])
        matched_hits = [m for m in hits if normalize_for_llm_guard(m.target_text)]
        targets: List[str] = []
        for m in matched_hits:
            t = normalize_for_llm_guard(m.target_text)
            if t and t not in targets:
                targets.append(t)
        # Row shaping rule:
        # - no hit => unmatched row with empty reasons
        # - single hit => take best-first occurrence (status-preferred, then confidence)
        # - multiple hits => force uncertain and merge rationale/reasons across candidates
        status = "unmatched"
        rationale = ""
        review_reason = ""
        semantic_check_reason = ""
        if hits:
            if len(hits) == 1:
                best = hits[0]
            else:
                best = sorted(
                    hits,
                    key=lambda m: (_status_rank(m.status), float(m.overall_confidence), float(m.semantic_confidence)),
                    reverse=True,
                )[0]
            if len(hits) == 1:
                status = str(best.status or "unmatched")
                rationale = normalize_for_llm_guard(str(best.rationale or ""))
                review_reason = normalize_for_llm_guard(str(best.review_reason or ""))
                semantic_check_reason = normalize_for_llm_guard(str(best.semantic_check_reason or ""))
            else:
                status = "uncertain"
                rationale = _merge_reasons(hits, "rationale")
                review_reason = _merge_reasons(hits, "review_reason")
                semantic_check_reason = _merge_reasons(hits, "semantic_check_reason")
        rows.append(
            {
                "om_en_string": s,
                "target_equivalent_candidates": " | ".join(targets),
                "status": status,
                "rationale": rationale,
                "review_reason": review_reason,
                "semantic_check_reason": semantic_check_reason,
                "match_count": len(hits),
            }
        )
    return rows


def corroborate_with_raw_ocr(text: str, classic: ClassicOCRResult) -> Tuple[str, str]:
    candidate = normalize_for_llm_guard(text)
    if not candidate:
        return "not_confirmed", "Empty candidate text."
    raw_blocks = [normalize_for_llm_guard(v) for v in classic.full_text_by_pass.values() if normalize_for_llm_guard(v)]
    if not raw_blocks:
        return "not_confirmed", "No raw OCR full-text evidence available."
    if any(candidate in block for block in raw_blocks):
        return "confirmed", "Exact normalized substring found in raw OCR full text."
    cand_norm = remove_accents(candidate).lower()
    cand_clean = remove_accents(clean_line(candidate)).lower()
    for block in raw_blocks:
        block_norm = remove_accents(block).lower()
        if cand_norm and cand_norm in block_norm:
            return "weakly_confirmed", "Accent-insensitive substring found in raw OCR full text."
        block_clean = remove_accents(clean_line(block)).lower()
        if cand_clean and len(cand_clean) >= 2 and cand_clean in block_clean:
            return "weakly_confirmed", "Punctuation-insensitive substring found in raw OCR full text."
        ratio = difflib.SequenceMatcher(None, cand_norm, block_norm).ratio() if cand_norm and block_norm else 0.0
        if ratio >= 0.88:
            return "weakly_confirmed", f"High fuzzy similarity against raw OCR full text ({ratio:.2f})."
    return "not_confirmed", "No supporting evidence in raw OCR full text."


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


def semantic_quality_check(
    source_text: str,
    target_text: str,
    target_language_name: str,
    model: str,
    api_key: str,
    timeout_sec: int,
    max_retries: int,
    audit: Optional[AIAuditLog] = None,
    audit_db: Optional[OpenAIAuditDB] = None,
    restore_from_db_cache: bool = True,
    write_to_db_cache: bool = True,
    pair_id: str = "",
) -> SemanticPairCheck:
    if not source_text or not target_text:
        return SemanticPairCheck(semantic_check_status="warning", semantic_check_confidence=0.0, semantic_check_reason="Missing source or target text.")
    # Fast-pass for exact/cognate-like matches to avoid unnecessary mismatch calls.
    src_norm = remove_accents(normalize_for_llm_guard(source_text)).lower()
    tgt_norm = remove_accents(normalize_for_llm_guard(target_text)).lower()
    sim = difflib.SequenceMatcher(None, src_norm, tgt_norm).ratio() if src_norm and tgt_norm else 0.0
    if src_norm == tgt_norm or sim >= 0.88:
        return SemanticPairCheck(
            semantic_check_status="ok",
            semantic_check_confidence=max(0.90, min(0.99, sim)),
            semantic_check_reason="Fast-pass: exact/near-cognate lexical match.",
        )
    cache_key = cache_key_for_payload(
        "semantic_pair_check_v1",
        model,
        target_language_name,
        normalize_for_llm_guard(source_text),
        normalize_for_llm_guard(target_text),
    )
    if restore_from_db_cache and audit_db is not None:
        cached = audit_db.cache_get_model("semantic_pair_check", cache_key, SemanticPairCheck)
        if cached:
            audit_db.log_event(
                "Semantic pair quality check",
                f"{pair_id} semantic check: {source_text[:40]} -> {target_text[:40]}",
                model,
                "semantic_pair_check_v1",
                {"pair_id": pair_id, "source_text": source_text, "target_text": target_text, "target_language_name": target_language_name},
                response={"parsed": model_to_dict(cached)},
                was_cached=True,
                cache_key=cache_key,
            )
            return cached  # type: ignore[return-value]
    schema = SemanticPairCheck
    user_text = (
        f"SOURCE_TEXT: {source_text}\n"
        f"TARGET_TEXT: {target_text}\n"
        f"TARGET_LANGUAGE: {target_language_name}\n\n"
        "Classify if target preserves source meaning in UI context.\n"
        "Return mismatch when meaning differs significantly."
    )
    parsed = call_openai_structured(
        api_key=api_key,
        model=model,
        system_prompt="You are a strict bilingual semantic validator for UI strings. Return schema only.",
        user_content=[{"type": "input_text", "text": user_text}],
        schema_model=schema,
        timeout_sec=timeout_sec,
        max_retries=max_retries,
        audit=audit,
        audit_db=audit_db,
        audit_stage="Semantic pair quality check",
        audit_title=f"{pair_id} semantic check: {source_text[:40]} -> {target_text[:40]}",
        audit_prompt_version="semantic_pair_check_v1",
        audit_request={
            "pair_id": pair_id,
            "source_text": source_text,
            "target_text": target_text,
            "target_language_name": target_language_name,
        },
        audit_cache_key=cache_key,
    )
    assert isinstance(parsed, SemanticPairCheck)
    if write_to_db_cache and audit_db is not None:
        audit_db.cache_put_model(
            "semantic_pair_check",
            cache_key,
            parsed,
            metadata={"pair_id": pair_id, "target_language_name": target_language_name, "model": model},
        )
    return parsed


def enforce_one_to_one_target_assignment(matches: List[GUIMatch]) -> None:
    used: Dict[str, List[int]] = {}
    for i, m in enumerate(matches):
        if not m.target_object_id or m.status == "unmatched":
            continue
        used.setdefault(m.target_object_id, []).append(i)
    for target_id, idxs in used.items():
        if len(idxs) <= 1:
            continue
        best = max(idxs, key=lambda i: matches[i].overall_confidence)
        for i in idxs:
            if i == best:
                continue
            if matches[i].status != "unmatched":
                matches[i].status = "uncertain"
                matches[i].review_reason = (matches[i].review_reason + " | target object reused").strip(" |")


def enforce_one_to_one_source_assignment(matches: List[GUIMatch]) -> None:
    used: Dict[str, List[int]] = {}
    for i, m in enumerate(matches):
        used.setdefault(m.source_object_id, []).append(i)
    for source_id, idxs in used.items():
        if len(idxs) <= 1:
            continue
        best = max(idxs, key=lambda i: (matches[i].overall_confidence, matches[i].semantic_confidence))
        for i in idxs:
            if i == best:
                continue
            matches[i].status = "unmatched"
            matches[i].target_object_id = None
            matches[i].target_text = ""
            matches[i].match_type = "unmatched"
            matches[i].review_reason = (matches[i].review_reason + " | duplicate source match dropped").strip(" |")


def apply_final_quality_gates(match: GUIMatch) -> None:
    if match.status == "matched" and (
        match.semantic_check_status == "mismatch" or match.raw_ocr_validation_status == "not_confirmed"
    ):
        match.status = "needs_review"
        match.review_reason = (match.review_reason + " | blocked by semantic/raw-evidence gate").strip(" |")


def _status_rank(status: str) -> int:
    return {"matched": 4, "needs_review": 3, "uncertain": 2, "unmatched": 1}.get(status, 0)


def dedupe_equivalent_matches(matches: List[GUIMatch]) -> List[GUIMatch]:
    best_by_key: Dict[Tuple[str, str], GUIMatch] = {}
    for m in matches:
        key = (
            normalize_for_llm_guard(m.source_text).lower(),
            normalize_for_llm_guard(m.target_text).lower(),
        )
        existing = best_by_key.get(key)
        if existing is None:
            best_by_key[key] = m
            continue
        lhs = (_status_rank(m.status), m.overall_confidence, m.semantic_confidence)
        rhs = (_status_rank(existing.status), existing.overall_confidence, existing.semantic_confidence)
        if lhs > rhs:
            best_by_key[key] = m
    return list(best_by_key.values())


def collapse_duplicate_source_text(matches: List[GUIMatch]) -> List[GUIMatch]:
    best_by_source: Dict[str, GUIMatch] = {}
    for m in matches:
        key = normalize_for_llm_guard(m.source_text).lower()
        if not key:
            continue
        existing = best_by_source.get(key)
        if existing is None:
            best_by_source[key] = m
            continue
        lhs = (_status_rank(m.status), m.overall_confidence, -len(normalize_for_llm_guard(m.target_text)))
        rhs = (_status_rank(existing.status), existing.overall_confidence, -len(normalize_for_llm_guard(existing.target_text)))
        if lhs > rhs:
            best_by_source[key] = m
    selected_ids = {id(v) for v in best_by_source.values()}
    passthrough = [m for m in matches if not normalize_for_llm_guard(m.source_text)]
    return [m for m in matches if id(m) in selected_ids] + passthrough


def is_icon_inference_candidate(obj: NormalizedGUIObject, classic: ClassicOCRResult) -> bool:
    text = clean_line(obj.normalized_text)
    if not text:
        return False
    corroboration, _ = corroborate_with_raw_ocr(text, classic)
    if corroboration != "not_confirmed":
        return False
    rationale_blob = " ".join(
        [
            normalize_for_llm_guard(obj.rationale).lower(),
            normalize_for_llm_guard(obj.raw_visual_evidence).lower(),
            " ".join(normalize_for_llm_guard(x).lower() for x in obj.raw_evidence),
            normalize_for_llm_guard(obj.object_id).lower(),
        ]
    )
    icon_markers = ("icon", "symbol", "arrow", "glyph", "nav_category_button")
    if not any(m in rationale_blob for m in icon_markers):
        return False
    token_count = len([t for t in text.split() if t])
    return token_count <= 4


def _ensure_unique_object_ids(objects: Sequence[NormalizedGUIObject], prefix: str) -> List[NormalizedGUIObject]:
    out: List[NormalizedGUIObject] = []
    seen: Dict[str, int] = {}
    for i, obj in enumerate(objects, start=1):
        base = normalize_for_llm_guard(obj.object_id) or f"{prefix}_{i:03d}"
        count = seen.get(base, 0) + 1
        seen[base] = count
        new_id = base if count == 1 else f"{base}__dup{count}"
        clone = obj.model_copy(deep=True)
        clone.object_id = new_id
        out.append(clone)
    return out


SPLIT_START_CUES = (
    " itt ",
    " ha ",
    " vagy ",
    " you ",
    " can ",
    " explore ",
    " connect ",
    " csatlakozz",
    " kedykolvek ",
    " whenever ",
    " when ",
)


def _split_merged_gui_object(obj: NormalizedGUIObject) -> List[NormalizedGUIObject]:
    text = normalize_for_llm_guard(obj.normalized_text)
    if not text:
        return []
    # Keep this conservative: only split likely merged label+description blocks.
    if obj.gui_role not in {"button", "menu_label", "description", "other"}:
        return []
    if len(text) < 24 or len(text.split()) < 4:
        return []
    lower = f" {remove_accents(text).lower()} "
    cut_idx = -1
    for cue in SPLIT_START_CUES:
        pos = lower.find(cue)
        if pos > 1:
            cut_idx = pos
            break
    if cut_idx < 0:
        # Secondary heuristic: split before second sentence.
        dot_pos = text.find(". ")
        if dot_pos > 8:
            cut_idx = dot_pos + 2
    if cut_idx < 0:
        return []
    left = normalize_for_llm_guard(text[:cut_idx].strip(" ;:,."))
    right = normalize_for_llm_guard(text[cut_idx:].strip(" ;:,."))
    if not left or not right:
        return []
    if len(left.split()) > 4:
        return []
    if not is_label_like_value_text(left):
        return []
    left_obj = obj.model_copy(deep=True)
    left_obj.object_id = f"{obj.object_id}__split_label"
    left_obj.normalized_text = left
    left_obj.gui_role = "menu_label" if obj.gui_role != "description" else "description"
    left_obj.rationale = (left_obj.rationale + " | synthetic split: label").strip(" |")
    left_obj.normalization_confidence = max(0.75, min(0.99, float(obj.normalization_confidence) * 0.95))

    right_obj = obj.model_copy(deep=True)
    right_obj.object_id = f"{obj.object_id}__split_desc"
    right_obj.normalized_text = right
    right_obj.gui_role = "description"
    right_obj.rationale = (right_obj.rationale + " | synthetic split: description").strip(" |")
    right_obj.normalization_confidence = max(0.70, min(0.99, float(obj.normalization_confidence) * 0.90))
    return [left_obj, right_obj]


def _augment_with_split_objects(
    objects: Sequence[NormalizedGUIObject],
    lane_map: Dict[str, str],
) -> Tuple[List[NormalizedGUIObject], Dict[str, str]]:
    out: List[NormalizedGUIObject] = []
    out_lane = dict(lane_map)
    for obj in objects:
        out.append(obj)
        split_objs = _split_merged_gui_object(obj)
        for s in split_objs:
            out.append(s)
            # Both split parts are valid match candidates.
            if s.object_id.endswith("__split_label"):
                out_lane[s.object_id] = "translatable_gui"
            elif s.object_id.endswith("__split_desc"):
                out_lane[s.object_id] = "translatable_gui"
            else:
                out_lane[s.object_id] = out_lane.get(obj.object_id, "unknown_review")
    return out, out_lane


def ai_validated_matches(
    pair_id: str,
    source_image: str,
    target_image: str,
    target_language_name: str,
    source_ai: OCRNormalizedResult,
    target_ai: OCRNormalizedResult,
    source_classic: ClassicOCRResult,
    target_classic: ClassicOCRResult,
    model: str,
    api_key: Optional[str],
    timeout_sec: int,
    max_retries: int,
    audit: Optional[AIAuditLog] = None,
    audit_db: Optional[OpenAIAuditDB] = None,
    restore_object_match_from_db_cache: bool = True,
    write_object_match_to_db_cache: bool = True,
    restore_semantic_check_from_db_cache: bool = True,
    write_semantic_check_to_db_cache: bool = True,
    parallel_workers: int = 1,
) -> List[GUIMatch]:
    source_all = _ensure_unique_object_ids(source_ai.ui_objects, "src")
    target_all = _ensure_unique_object_ids(target_ai.ui_objects, "tgt")
    source_lane = {o.object_id: classify_object_lane(o) for o in source_all}
    target_lane = {o.object_id: classify_object_lane(o) for o in target_all}
    source_all, source_lane = _augment_with_split_objects(source_all, source_lane)
    target_all, target_lane = _augment_with_split_objects(target_all, target_lane)

    def _expand_objects(
        objects: Sequence[NormalizedGUIObject],
        lane_map: Dict[str, str],
        classic: ClassicOCRResult,
    ) -> List[NormalizedGUIObject]:
        expanded: List[NormalizedGUIObject] = []
        for obj in objects:
            lane = lane_map.get(obj.object_id, "unknown_review")
            if lane != "translatable_gui":
                continue
            if is_icon_inference_candidate(obj, classic):
                continue
            # Text-only rule: keep only OCR-corroborated GUI text objects.
            corroboration, _ = corroborate_with_raw_ocr(obj.normalized_text, classic)
            if corroboration == "not_confirmed":
                continue
            expanded.append(obj)
        return expanded

    # Primary lane is translatable_gui; plus promoted label-like candidates from non-translatable lanes.
    source_objects = _expand_objects(source_all, source_lane, source_classic)
    target_objects = _expand_objects(target_all, target_lane, target_classic)

    if not source_objects:
        return []

    source_objs_dict = [model_to_dict(o) for o in source_objects]
    target_objs_dict = [model_to_dict(o) for o in target_objects]
    for item in source_objs_dict:
        item["object_lane"] = source_lane.get(item["object_id"], "unknown_review")
    for item in target_objs_dict:
        item["object_lane"] = target_lane.get(item["object_id"], "unknown_review")

    if api_key:
        object_match_cache_key = cache_key_for_payload(
            "ai_validated_object_match_v1",
            model,
            pair_id,
            json.dumps(source_objs_dict, ensure_ascii=False, sort_keys=True),
            json.dumps(target_objs_dict, ensure_ascii=False, sort_keys=True),
        )
        cached_object_match: Optional[GUIObjectMatchResponse] = None
        if restore_object_match_from_db_cache and audit_db is not None:
            cached_object_match = audit_db.cache_get_model("ai_validated_object_match", object_match_cache_key, GUIObjectMatchResponse)  # type: ignore[assignment]
        if cached_object_match is not None:
            result = cached_object_match
            audit_db.log_event(
                "AI-validated object matching",
                f"{pair_id} {source_image} -> {target_image}",
                model,
                GUI_OBJECT_MATCH_PROMPT_VERSION + "_ai_validated",
                {
                    "pair_id": pair_id,
                    "source_image": source_image,
                    "target_image": target_image,
                    "source_objects": source_objs_dict,
                    "target_objects": target_objs_dict,
                },
                response={"parsed": model_to_dict(result)},
                was_cached=True,
                cache_key=object_match_cache_key,
            )
            matches = validate_gui_matches(result, source_objects, target_objects)
        else:
            user_text = (
                f"PAIR ID: {pair_id}\n"
                f"SOURCE IMAGE: {source_image}\n"
                f"TARGET IMAGE: {target_image}\n"
                f"TARGET LANGUAGE NAME: {target_language_name}\n\n"
                f"SOURCE_OBJECTS: {json.dumps(source_objs_dict, ensure_ascii=False)}\n\n"
                f"TARGET_OBJECTS: {json.dumps(target_objs_dict, ensure_ascii=False)}\n\n"
                "Match only translatable GUI content. Do not prioritize value-only, masked, map background, or decorative/status lanes.\n"
                "For short labels, prefer one-to-one mapping and avoid merged target assignments unless unavoidable.\n"
                "If target candidates include synthetic split object ids (__split_label / __split_desc), map short label-like source text to __split_label and explanatory/long source text to __split_desc.\n"
                "If target phrase appears to combine multiple labels, map each source label to the best lexical counterpart (including reversed order) or mark uncertain/unmatched."
            )
            try:
                result = call_openai_structured(
                    api_key,
                    model,
                    system_prompt=system_prompt_gui_object_match(target_language_name),
                    user_content=[{"type": "input_text", "text": user_text}],
                    schema_model=GUIObjectMatchResponse,
                    timeout_sec=timeout_sec,
                    max_retries=max_retries,
                    audit=audit,
                    audit_db=audit_db,
                    audit_stage="AI-validated object matching",
                    audit_title=f"{pair_id} {source_image} -> {target_image}",
                    audit_prompt_version=GUI_OBJECT_MATCH_PROMPT_VERSION + "_ai_validated",
                    audit_request={
                        "pair_id": pair_id,
                        "source_image": source_image,
                        "target_image": target_image,
                        "source_objects": source_objs_dict,
                        "target_objects": target_objs_dict,
                    },
                    audit_cache_key=object_match_cache_key,
                )
                assert isinstance(result, GUIObjectMatchResponse)
                if write_object_match_to_db_cache and audit_db is not None:
                    audit_db.cache_put_model(
                        "ai_validated_object_match",
                        object_match_cache_key,
                        result,
                        metadata={"pair_id": pair_id, "model": model, "target_language_name": target_language_name},
                    )
                matches = validate_gui_matches(result, source_objects, target_objects)
            except Exception as exc:
                logging.exception(
                    "AI-validated object matching failed for pair %s (%s -> %s). Falling back to deterministic pairing. Error: %s",
                    pair_id,
                    source_image,
                    target_image,
                    exc,
                )
                matches = []
                aligned_n = max(len(source_objects), len(target_objects))
                for i in range(aligned_n):
                    s_obj = source_objects[i] if i < len(source_objects) else None
                    t_obj = target_objects[i] if i < len(target_objects) else None
                    s = s_obj.normalized_text if s_obj else ""
                    t = t_obj.normalized_text if t_obj else ""
                    matches.append(
                        GUIMatch(
                            pair_id=pair_id,
                            source_image_filename=source_image,
                            target_image_filename=target_image,
                            source_object_id=s_obj.object_id if s_obj else f"obj_{i+1:03d}",
                            target_object_id=t_obj.object_id if t_obj else None,
                            source_text=s,
                            target_text=t,
                            match_type="uncertain" if t else "unmatched",
                            semantic_confidence=0.5 if t else 0.0,
                            layout_confidence=0.6 if t else 0.0,
                            overall_confidence=0.55 if t else 0.0,
                            status="needs_review" if t else "unmatched",
                            rationale="Deterministic fallback after AI object matching timeout/error.",
                            review_reason="ai_object_match_timeout_or_error",
                        )
                    )
    else:
        # deterministic fallback
        matches = []
        aligned_n = max(len(source_objects), len(target_objects))
        for i in range(aligned_n):
            s_obj = source_objects[i] if i < len(source_objects) else None
            t_obj = target_objects[i] if i < len(target_objects) else None
            s = s_obj.normalized_text if s_obj else ""
            t = t_obj.normalized_text if t_obj else ""
            matches.append(
                GUIMatch(
                    pair_id=pair_id,
                    source_image_filename=source_image,
                    target_image_filename=target_image,
                    source_object_id=s_obj.object_id if s_obj else f"obj_{i+1:03d}",
                    target_object_id=t_obj.object_id if t_obj else None,
                    source_text=s,
                    target_text=t,
                    match_type="uncertain" if t else "unmatched",
                    semantic_confidence=0.5 if t else 0.0,
                    layout_confidence=0.6 if t else 0.0,
                    overall_confidence=0.55 if t else 0.0,
                    status="uncertain" if t else "unmatched",
                    rationale="Deterministic fallback in ai_validated mode.",
                )
            )

    # Recovery pass: if first pass omitted source objects, ask LLM to match remaining source/target objects.
    if api_key:
        matched_source_ids = {m.source_object_id for m in matches}
        matched_target_ids = {m.target_object_id for m in matches if m.target_object_id}
        remaining_source = [o for o in source_objects if o.object_id not in matched_source_ids]
        remaining_target = [o for o in target_objects if o.object_id not in matched_target_ids]
        if remaining_source and remaining_target:
            rem_source_dict = [model_to_dict(o) for o in remaining_source]
            rem_target_dict = [model_to_dict(o) for o in remaining_target]
            rem_user_text = (
                f"PAIR ID: {pair_id}\n"
                f"SOURCE IMAGE: {source_image}\n"
                f"TARGET IMAGE: {target_image}\n"
                f"TARGET LANGUAGE NAME: {target_language_name}\n\n"
                f"REMAINING_SOURCE_OBJECTS: {json.dumps(rem_source_dict, ensure_ascii=False)}\n\n"
                f"REMAINING_TARGET_OBJECTS: {json.dumps(rem_target_dict, ensure_ascii=False)}\n\n"
                "Match the remaining source GUI objects to remaining target GUI objects. "
                "Prefer one-to-one short-label mapping, even when target word order is reversed. "
                "Use synthetic split ids when available: short source labels -> __split_label, longer explanatory source text -> __split_desc. "
                "If unsure, return uncertain or unmatched. Do not invent text."
            )
            try:
                rem_cache_key = cache_key_for_payload(
                    "ai_validated_object_match_recovery_v1",
                    model,
                    pair_id,
                    json.dumps(rem_source_dict, ensure_ascii=False, sort_keys=True),
                    json.dumps(rem_target_dict, ensure_ascii=False, sort_keys=True),
                )
                cached_rem: Optional[GUIObjectMatchResponse] = None
                if restore_object_match_from_db_cache and audit_db is not None:
                    cached_rem = audit_db.cache_get_model("ai_validated_object_match_recovery", rem_cache_key, GUIObjectMatchResponse)  # type: ignore[assignment]
                if cached_rem is not None:
                    rem_result = cached_rem
                    audit_db.log_event(
                        "AI-validated object matching (recovery pass)",
                        f"{pair_id} {source_image} -> {target_image} (remaining)",
                        model,
                        GUI_OBJECT_MATCH_PROMPT_VERSION + "_ai_validated_recovery",
                        {
                            "pair_id": pair_id,
                            "source_image": source_image,
                            "target_image": target_image,
                            "remaining_source_objects": rem_source_dict,
                            "remaining_target_objects": rem_target_dict,
                        },
                        response={"parsed": model_to_dict(rem_result)},
                        was_cached=True,
                        cache_key=rem_cache_key,
                    )
                else:
                    rem_result = call_openai_structured(
                        api_key,
                        model,
                        system_prompt=system_prompt_gui_object_match(target_language_name),
                        user_content=[{"type": "input_text", "text": rem_user_text}],
                        schema_model=GUIObjectMatchResponse,
                        timeout_sec=timeout_sec,
                        max_retries=max_retries,
                        audit=audit,
                        audit_db=audit_db,
                        audit_stage="AI-validated object matching (recovery pass)",
                        audit_title=f"{pair_id} {source_image} -> {target_image} (remaining)",
                        audit_prompt_version=GUI_OBJECT_MATCH_PROMPT_VERSION + "_ai_validated_recovery",
                        audit_request={
                            "pair_id": pair_id,
                            "source_image": source_image,
                            "target_image": target_image,
                            "remaining_source_objects": rem_source_dict,
                            "remaining_target_objects": rem_target_dict,
                        },
                        audit_cache_key=rem_cache_key,
                    )
                    if write_object_match_to_db_cache and audit_db is not None:
                        audit_db.cache_put_model(
                            "ai_validated_object_match_recovery",
                            rem_cache_key,
                            rem_result,  # type: ignore[arg-type]
                            metadata={"pair_id": pair_id, "model": model, "target_language_name": target_language_name},
                        )
                assert isinstance(rem_result, GUIObjectMatchResponse)
                rem_matches = validate_gui_matches(rem_result, remaining_source, remaining_target)
                matches.extend(rem_matches)
            except Exception:
                logging.exception("AI-validated recovery matching failed for pair %s", pair_id)

    source_by_id = {o.object_id: o for o in source_all}
    source_by_id.update({o.object_id: o for o in source_objects})
    target_by_id = {o.object_id: o for o in target_all}
    target_by_id.update({o.object_id: o for o in target_objects})
    # Fallback: pair still-unmatched source and target objects by reading order to avoid silent drops.
    matched_source_ids = {m.source_object_id for m in matches if m.status != "unmatched"}
    matched_target_ids = {m.target_object_id for m in matches if m.target_object_id and m.status != "unmatched"}
    remaining_source = [o for o in source_objects if o.object_id not in matched_source_ids and is_label_like_value_text(o.normalized_text)]
    remaining_target = [o for o in target_objects if o.object_id not in matched_target_ids and is_label_like_value_text(o.normalized_text)]
    remaining_source.sort(key=lambda o: (o.row_group, o.reading_order, o.object_id))
    remaining_target.sort(key=lambda o: (o.row_group, o.reading_order, o.object_id))
    for src_obj, tgt_obj in zip(remaining_source, remaining_target):
        matches.append(
            GUIMatch(
                pair_id=pair_id,
                source_image_filename=source_image,
                target_image_filename=target_image,
                source_object_id=src_obj.object_id,
                target_object_id=tgt_obj.object_id,
                source_text=src_obj.normalized_text,
                target_text=tgt_obj.normalized_text,
                match_type="uncertain",
                semantic_confidence=0.45,
                layout_confidence=0.65,
                overall_confidence=0.55,
                status="needs_review",
                rationale="Reading-order fallback pairing for unmatched residual objects.",
                review_reason="fallback_residual_pairing",
            )
        )

    # Normalize identity fields to the current pair context regardless of model payload values.
    for m in matches:
        m.pair_id = pair_id
        m.source_image_filename = source_image
        m.target_image_filename = target_image

    semantic_jobs: List[Tuple[int, str, str]] = []
    for idx, m in enumerate(matches):
        source_obj = source_by_id.get(m.source_object_id)
        target_obj = target_by_id.get(m.target_object_id or "")
        if source_obj:
            m.source_object_lane = classify_object_lane(source_obj)
        if target_obj:
            m.target_object_lane = classify_object_lane(target_obj)
        promoted_source = "__" in (m.source_object_id or "")
        promoted_target = "__" in (m.target_object_id or "")
        if m.source_object_lane != "translatable_gui" and not promoted_source:
            if m.status != "unmatched":
                m.status = "needs_review"
                m.review_reason = (m.review_reason + f" | source lane={m.source_object_lane}").strip(" |")
        if m.target_object_lane and m.target_object_lane != "translatable_gui" and m.status == "matched" and not promoted_target:
            m.status = "needs_review"
            m.review_reason = (m.review_reason + f" | target lane={m.target_object_lane}").strip(" |")

        raw_status, raw_reason = corroborate_with_raw_ocr(m.target_text, target_classic)
        m.raw_ocr_validation_status = raw_status  # type: ignore[assignment]
        m.raw_ocr_validation_reason = raw_reason

        if api_key and m.status != "unmatched":
            semantic_jobs.append((idx, m.source_text, m.target_text))
        else:
            m.semantic_check_status = "warning"
            m.semantic_check_reason = "Semantic check unavailable (no API key or unmatched row)."

    if semantic_jobs:
        def _run_semantic(job: Tuple[int, str, str]) -> Tuple[int, Optional[SemanticPairCheck], str]:
            idx, src_text, tgt_text = job
            try:
                sem = semantic_quality_check(
                    source_text=src_text,
                    target_text=tgt_text,
                    target_language_name=target_language_name,
                    model=model,
                    api_key=api_key or "",
                    timeout_sec=timeout_sec,
                    max_retries=max_retries,
                    audit=audit,
                    audit_db=audit_db,
                    restore_from_db_cache=restore_semantic_check_from_db_cache,
                    write_to_db_cache=write_semantic_check_to_db_cache,
                    pair_id=pair_id,
                )
                return idx, sem, ""
            except Exception as exc:
                return idx, None, str(exc)

        if parallel_workers > 1 and len(semantic_jobs) > 1:
            with ThreadPoolExecutor(max_workers=min(parallel_workers, len(semantic_jobs), 6)) as ex:
                semantic_results = list(ex.map(_run_semantic, semantic_jobs))
        else:
            semantic_results = [_run_semantic(job) for job in semantic_jobs]

        for idx, sem, err in semantic_results:
            m = matches[idx]
            if sem is None:
                m.semantic_check_status = "warning"
                m.semantic_check_reason = f"Semantic check failed: {err}"
                continue
            m.semantic_check_status = sem.semantic_check_status
            m.semantic_check_confidence = sem.semantic_check_confidence
            m.semantic_check_reason = sem.semantic_check_reason
            if (
                m.semantic_check_status == "mismatch"
                and m.raw_ocr_validation_status in {"confirmed", "weakly_confirmed"}
                and m.semantic_confidence >= 0.95
                and m.layout_confidence >= 0.90
                and m.source_object_lane == "translatable_gui"
                and m.target_object_lane == "translatable_gui"
            ):
                m.semantic_check_status = "warning"
                m.semantic_check_reason = (
                    m.semantic_check_reason + " | downgraded from mismatch by high-confidence/confirmed heuristic"
                ).strip(" |")

    for m in matches:

        # Source-guided conflict (review-only, no overwrite)
        source_raw_status, _source_raw_reason = corroborate_with_raw_ocr(m.source_text, source_classic)
        if (
            m.status != "unmatched"
            and m.semantic_check_status == "mismatch"
            and m.raw_ocr_validation_status in {"weakly_confirmed", "not_confirmed"}
            and source_raw_status in {"confirmed", "weakly_confirmed"}
            and not (
                m.semantic_confidence >= 0.95
                and m.layout_confidence >= 0.90
                and m.source_object_lane == "translatable_gui"
                and m.target_object_lane == "translatable_gui"
            )
        ):
            m.source_guided_conflict = True
            m.suggested_target_candidate = ""
            m.status = "needs_review"
            m.review_reason = (m.review_reason + " | source-guided conflict").strip(" |")

        apply_final_quality_gates(m)
        if not m.source_guided_conflict:
            m.suggested_target_candidate = ""
        else:
            m.suggested_target_candidate = str(m.suggested_target_candidate or "").strip()

    # Guarantee source coverage: any source object not returned by matching is emitted as unmatched.
    existing_source_ids = {m.source_object_id for m in matches}
    for src_obj in source_objects:
        if src_obj.object_id in existing_source_ids:
            continue
        matches.append(
            GUIMatch(
                pair_id=pair_id,
                source_image_filename=source_image,
                target_image_filename=target_image,
                source_object_id=src_obj.object_id,
                target_object_id=None,
                source_text=src_obj.normalized_text,
                target_text="",
                match_type="unmatched",
                semantic_confidence=0.0,
                layout_confidence=0.0,
                overall_confidence=0.0,
                status="unmatched",
                rationale="Source object not returned by object matcher.",
                review_reason="coverage_backfill",
                source_object_lane=classify_object_lane(src_obj),
                target_object_lane="",
                raw_ocr_validation_status="not_confirmed",
                raw_ocr_validation_reason="No target candidate selected.",
                semantic_check_status="warning",
                semantic_check_confidence=0.0,
                semantic_check_reason="No target candidate selected.",
                source_guided_conflict=False,
                suggested_target_candidate="",
            )
        )

    enforce_one_to_one_source_assignment(matches)
    enforce_one_to_one_target_assignment(matches)
    return collapse_duplicate_source_text(dedupe_equivalent_matches(matches))


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
    return ocr_engine in {"classic", "classic_llm", "hybrid", "triangulated", "ai_validated"} or keep_optional_classic


def should_run_llm_normalization(ocr_engine: str, explicit: bool) -> bool:
    return ocr_engine in {"classic_llm", "triangulated"} or explicit


def should_run_ai_ocr(ocr_engine: str) -> bool:
    return ocr_engine in {"ai", "hybrid", "triangulated", "ai_validated"}


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
    openai_cache_only: bool,
    audit_db: Optional[OpenAIAuditDB],
    restore_object_match_from_db_cache: bool,
    restore_semantic_check_from_db_cache: bool,
    write_stage_cache_to_db: bool,
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
    llm_workers: int = 1,
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
    ai_required = ocr_engine in {"ai", "hybrid", "triangulated", "ai_validated"}
    llm_required = ocr_engine in {"classic_llm", "triangulated"} or (matching_mode == "llm_objects" and ocr_engine != "ai_validated")
    if openai_cache_only and ocr_engine != "ai_validated":
        raise ValueError("--openai-cache-only is currently supported for --ocr-engine ai_validated.")
    if (ai_required or llm_required) and not openai_api_key and not (openai_cache_only and ocr_engine == "ai_validated"):
        if fallback_to_classic:
            logging.warning("OpenAI API key missing; falling back to classic OCR where possible.")
            run_llm = False
            run_ai = False
            matching_mode = "positional"
        else:
            raise ValueError("AI processing requires an OpenAI API key. Pass --openai-api-key or set OPENAI_API_KEY.")
    if openai_cache_only:
        logging.info("OpenAI cache-only mode enabled; live OpenAI calls are disabled.")
        run_llm = False
    if llm_workers > 1:
        logging.info("Parallel LLM workers enabled: %d", llm_workers)

    temp_dir_ctx = tempfile.TemporaryDirectory(prefix="ocr_local_") if use_temp_local else None
    temp_root = Path(temp_dir_ctx.name) if temp_dir_ctx else None
    render_tmp_ctx = tempfile.TemporaryDirectory(prefix="ocr_render_batch_") if not write_rendered_images else None
    render_root = rendered_image_dir if write_rendered_images else Path(render_tmp_ctx.name)
    render_root.mkdir(parents=True, exist_ok=True)

    image_pairs: List[ImagePair] = []
    all_lines: List[OCRLine] = []
    all_tokens: List[OCRToken] = []
    llm_results: List[OCRNormalizedResult] = []
    ai_results: List[OCRNormalizedResult] = []
    raw_ocr_full_text: List[Dict[str, Any]] = []
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
            for pass_name, full_text in src_classic.full_text_by_pass.items():
                raw_ocr_full_text.append(
                    {
                        "pair_id": pair_id,
                        "side": "source",
                        "image_filename": src_path.name,
                        "ocr_pass": pass_name,
                        "selected_pass": any(
                            line.selected_pass and line.ocr_pass == pass_name for line in src_classic.classic_lines
                        ),
                        "full_ocr_text": full_text,
                    }
                )
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
                for pass_name, full_text in trg_classic.full_text_by_pass.items():
                    raw_ocr_full_text.append(
                        {
                            "pair_id": pair_id,
                            "side": "target",
                            "image_filename": trg_path.name if trg_path else "",
                            "ocr_pass": pass_name,
                            "selected_pass": any(
                                line.selected_pass and line.ocr_pass == pass_name for line in trg_classic.classic_lines
                            ),
                            "full_ocr_text": full_text,
                        }
                    )
            if not pair.classic_ocr_status:
                pair.classic_ocr_status = "ok"

        src_llm: Optional[OCRNormalizedResult] = None
        trg_llm: Optional[OCRNormalizedResult] = None
        if run_llm:
            try:
                src_args = (
                    src_path.name,
                    "en",
                    "English",
                    "source",
                    src_classic,
                )
                trg_args = (
                    (trg_path.name if trg_path else ""),
                    lang_code,
                    lang_name,
                    "target",
                    trg_classic,
                )
                if llm_workers > 1 and trg_local:
                    with ThreadPoolExecutor(max_workers=2) as ex:
                        fut_src = ex.submit(
                            normalize_ocr_with_llm,
                            *src_args,
                            llm_ocr_model,
                            openai_api_key or "",
                            llm_ocr_timeout_sec,
                            llm_ocr_max_retries,
                            llm_ocr_cache_dir,
                            llm_ocr_use_cache,
                            audit_db,
                            True,
                            write_stage_cache_to_db,
                            audit_log,
                            pair_id,
                        )
                        fut_trg = ex.submit(
                            normalize_ocr_with_llm,
                            *trg_args,
                            llm_ocr_model,
                            openai_api_key or "",
                            llm_ocr_timeout_sec,
                            llm_ocr_max_retries,
                            llm_ocr_cache_dir,
                            llm_ocr_use_cache,
                            audit_db,
                            True,
                            write_stage_cache_to_db,
                            audit_log,
                            pair_id,
                        )
                        src_llm = fut_src.result()
                        trg_llm = fut_trg.result()
                else:
                    src_llm = normalize_ocr_with_llm(
                        *src_args,
                        llm_ocr_model,
                        openai_api_key or "",
                        llm_ocr_timeout_sec,
                        llm_ocr_max_retries,
                        llm_ocr_cache_dir,
                        llm_ocr_use_cache,
                        audit_db=audit_db,
                        restore_from_db_cache=True,
                        write_to_db_cache=write_stage_cache_to_db,
                        audit=audit_log,
                        pair_id=pair_id,
                    )
                    if trg_local:
                        trg_llm = normalize_ocr_with_llm(
                            *trg_args,
                            llm_ocr_model,
                            openai_api_key or "",
                            llm_ocr_timeout_sec,
                            llm_ocr_max_retries,
                            llm_ocr_cache_dir,
                            llm_ocr_use_cache,
                            audit_db=audit_db,
                            restore_from_db_cache=True,
                            write_to_db_cache=write_stage_cache_to_db,
                            audit=audit_log,
                            pair_id=pair_id,
                        )
                cache_counts["llm_hit" if src_llm.cached else "llm_miss"] += 1
                llm_results.append(src_llm)
                if trg_llm:
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
                src_ai_args = (
                    src_rendered,
                    src_path.name,
                    "en",
                    "English",
                    "source",
                )
                trg_ai_args = (
                    trg_rendered,
                    (trg_path.name if trg_path else ""),
                    lang_code,
                    lang_name,
                    "target",
                )
                if llm_workers > 1 and src_rendered and trg_rendered and trg_path:
                    with ThreadPoolExecutor(max_workers=2) as ex:
                        fut_src = ex.submit(
                            ai_image_ocr,
                            *src_ai_args,
                            ai_ocr_model,
                            ai_ocr_detail,
                            openai_api_key or "",
                            ai_ocr_timeout_sec,
                            ai_ocr_max_retries,
                            ai_ocr_cache_dir,
                            ai_ocr_use_cache,
                            openai_cache_only,
                            audit_db,
                            True,
                            write_stage_cache_to_db,
                            audit_log,
                            pair_id,
                        )
                        fut_trg = ex.submit(
                            ai_image_ocr,
                            *trg_ai_args,
                            ai_ocr_model,
                            ai_ocr_detail,
                            openai_api_key or "",
                            ai_ocr_timeout_sec,
                            ai_ocr_max_retries,
                            ai_ocr_cache_dir,
                            ai_ocr_use_cache,
                            openai_cache_only,
                            audit_db,
                            True,
                            write_stage_cache_to_db,
                            audit_log,
                            pair_id,
                        )
                        src_ai = fut_src.result()
                        trg_ai = fut_trg.result()
                else:
                    if src_rendered:
                        src_ai = ai_image_ocr(
                            *src_ai_args,
                            ai_ocr_model,
                            ai_ocr_detail,
                            openai_api_key or "",
                            ai_ocr_timeout_sec,
                            ai_ocr_max_retries,
                            ai_ocr_cache_dir,
                            ai_ocr_use_cache,
                            cache_only=openai_cache_only,
                            audit_db=audit_db,
                            restore_from_db_cache=True,
                            write_to_db_cache=write_stage_cache_to_db,
                            audit=audit_log,
                            pair_id=pair_id,
                        )
                    if trg_rendered and trg_path:
                        trg_ai = ai_image_ocr(
                            *trg_ai_args,
                            ai_ocr_model,
                            ai_ocr_detail,
                            openai_api_key or "",
                            ai_ocr_timeout_sec,
                            ai_ocr_max_retries,
                            ai_ocr_cache_dir,
                            ai_ocr_use_cache,
                            cache_only=openai_cache_only,
                            audit_db=audit_db,
                            restore_from_db_cache=True,
                            write_to_db_cache=write_stage_cache_to_db,
                            audit=audit_log,
                            pair_id=pair_id,
                        )
                if src_ai:
                    cache_counts["ai_hit" if src_ai.cached else "ai_miss"] += 1
                    ai_results.append(src_ai)
                if trg_ai:
                    cache_counts["ai_hit" if trg_ai.cached else "ai_miss"] += 1
                    ai_results.append(trg_ai)
                pair.ai_image_ocr_status = "ok"
            except Exception as exc:
                error_counts["ai"] += 1
                pair.ai_image_ocr_status = f"error: {exc}"
                logging.exception("AI image OCR failed for pair %s", pair_id)
                if not fallback_to_classic or openai_cache_only:
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
            if ocr_engine == "ai_validated" and src_ai and trg_ai:
                final_matches.extend(
                    ai_validated_matches(
                        pair_id=pair_id,
                        source_image=src_path.name,
                        target_image=trg_path.name,
                        target_language_name=lang_name,
                        source_ai=src_ai,
                        target_ai=trg_ai,
                        source_classic=src_classic,
                        target_classic=trg_classic,
                        model=object_match_model,
                        api_key=None if openai_cache_only else openai_api_key,
                        timeout_sec=object_match_timeout_sec,
                        max_retries=object_match_max_retries,
                        audit=audit_log,
                        audit_db=audit_db,
                        restore_object_match_from_db_cache=restore_object_match_from_db_cache,
                        write_object_match_to_db_cache=write_stage_cache_to_db,
                        restore_semantic_check_from_db_cache=restore_semantic_check_from_db_cache,
                        write_semantic_check_to_db_cache=write_stage_cache_to_db,
                        parallel_workers=max(1, llm_workers),
                    )
                )
            elif matching_mode == "llm_objects" and openai_api_key:
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

        pair_level_review = any(r.needs_review for r in src_tri + trg_tri) or any(
            m.pair_id == pair_id and m.status in {"needs_review", "uncertain", "unmatched"} for m in final_matches
        )
        pair.pair_status = "needs_review" if pair_level_review else "ok"
        image_pairs.append(pair)
        if progress_every > 0 and (idx % progress_every == 0 or idx == total):
            logging.info("Progress: %d/%d image pairs processed.", idx, total)

    if temp_dir_ctx:
        temp_dir_ctx.cleanup()
    if render_tmp_ctx:
        render_tmp_ctx.cleanup()

    if ocr_engine == "ai_validated":
        needs_review_count = sum(1 for m in final_matches if m.status == "needs_review")
        uncertain_count = sum(1 for m in final_matches if m.status == "uncertain")
        unmatched_count = sum(1 for m in final_matches if m.status == "unmatched")
        tri_disagree_count = 0
    else:
        needs_review_count = sum(1 for m in final_matches if m.status == "needs_review") + sum(1 for r in triangulation_rows if r.needs_review)
        uncertain_count = sum(1 for m in final_matches if m.status == "uncertain")
        unmatched_count = sum(1 for m in final_matches if m.status == "unmatched")
        tri_disagree_count = sum(1 for r in triangulation_rows if r.agreement_status == "disagree")

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
        "source image count": int(len(source_files)),
        "matched image pair count": int(sum(1 for p in image_pairs if p.target_image_filename)),
        "missing image pair count": int(sum(1 for p in image_pairs if not p.target_image_filename)),
        "classic OCR error count": int(error_counts["classic"]),
        "LLM OCR-normalization error count": int(error_counts["llm"]),
        "AI image OCR error count": int(error_counts["ai"]),
        "triangulation disagreement count": int(tri_disagree_count),
        "final matched string count": int(sum(1 for m in final_matches if m.status == "matched")),
        "needs_review count": int(needs_review_count),
        "uncertain count": int(uncertain_count),
        "unmatched count": int(unmatched_count),
        "LLM normalization cache hit count": int(cache_counts["llm_hit"]),
        "LLM normalization cache miss count": int(cache_counts["llm_miss"]),
        "AI OCR cache hit count": int(cache_counts["ai_hit"]),
        "AI OCR cache miss count": int(cache_counts["ai_miss"]),
    }
    return PipelineReport(image_pairs, all_lines, all_tokens, raw_ocr_full_text, llm_results, ai_results, triangulation_rows, final_matches, summary)


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


def write_multi_sheet_report(report: PipelineReport, output: Path, om_strings_xlsx: Optional[Path] = None) -> None:
    def _sanitize_excel_value(value: Any) -> Any:
        if isinstance(value, str):
            cleaned = ILLEGAL_CHARACTERS_RE.sub("", value)
            if cleaned.startswith(("=", "+", "-", "@")):
                cleaned = "'" + cleaned
            return cleaned
        return value

    def _sanitize_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        cleaned: List[Dict[str, Any]] = []
        for row in rows:
            cleaned.append({k: _sanitize_excel_value(v) for k, v in row.items()})
        return cleaned

    def _summary_rows(summary: Dict[str, Any]) -> List[Dict[str, Any]]:
        rows: List[Dict[str, Any]] = []
        for k, v in summary.items():
            value = v
            if "count" in str(k).lower():
                try:
                    value = str(int(v))
                except Exception:
                    value = "0"
            rows.append({"metric": k, "value": value})
        return rows

    output.parent.mkdir(parents=True, exist_ok=True)
    ocr_engine = str(report.summary.get("ocr_engine", ""))
    if ocr_engine == "ai_validated":
        pair_path_map = {
            (p.source_image_filename, p.target_image_filename or ""): (p.source_image_path, p.target_image_path or "")
            for p in report.image_pairs
        }
        final_rows: List[Dict[str, Any]] = []
        for match in report.final_matches:
            row = model_to_dict(match)
            src_p, tgt_p = pair_path_map.get((match.source_image_filename, match.target_image_filename), ("", ""))
            row["source_image_path"] = src_p
            row["target_image_path"] = tgt_p
            final_rows.append(row)
        om_rows = build_om_strings_rows(report, om_strings_xlsx) if om_strings_xlsx else []
        sheets = {
            "MM strings": _sanitize_rows(om_rows),
            "Final Matches": _sanitize_rows(final_rows),
            "AI Image OCR Objects": _sanitize_rows(rows_from_normalized_results(report.ai_results, "ai_image_ocr")),
            "Image pairs": _sanitize_rows([model_to_dict(p) for p in report.image_pairs]),
            "Summary": _sanitize_rows(_summary_rows(report.summary)),
            "Raw OCR Full Text": _sanitize_rows(report.raw_ocr_full_text),
        }
    else:
        sheets = {
            "Image pairs": _sanitize_rows([model_to_dict(p) for p in report.image_pairs]),
            "Raw OCR - Classic Lines": _sanitize_rows([model_to_dict(x) for x in report.classic_lines]),
            "Raw OCR - Classic Tokens": _sanitize_rows([model_to_dict(x) for x in report.classic_tokens]),
            "LLM OCR Normalized Objects": _sanitize_rows(rows_from_normalized_results(report.llm_results, "llm_ocr_normalization")),
            "AI Image OCR Objects": _sanitize_rows(rows_from_normalized_results(report.ai_results, "ai_image_ocr")),
            "OCR Triangulation": _sanitize_rows([model_to_dict(x) for x in report.triangulation_rows]),
            "Final Matches": _sanitize_rows([model_to_dict(x) for x in report.final_matches]),
            "Summary": _sanitize_rows(_summary_rows(report.summary)),
        }
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, rows in sheets.items():
            pd.DataFrame(rows).to_excel(writer, sheet_name=name[:31], index=False)
    set_multi_sheet_excel_formatting(output)
    if ocr_engine == "ai_validated":
        from openpyxl import load_workbook
        wb = load_workbook(output)
        if "Raw OCR Full Text" in wb.sheetnames:
            wb["Raw OCR Full Text"].sheet_state = "hidden"
        sheet_order = ["MM strings", "Final Matches", "AI Image OCR Objects", "Image pairs", "Summary", "Raw OCR Full Text"]
        wb._sheets.sort(key=lambda ws: sheet_order.index(ws.title) if ws.title in sheet_order else 999)
        wb.save(output)


def set_multi_sheet_excel_formatting(output_file: Path) -> None:
    from openpyxl import load_workbook

    wb = load_workbook(output_file)
    fills = {
        "matched": PatternFill(start_color="FFD9EAD3", end_color="FFD9EAD3", fill_type="solid"),
        "needs_review": PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid"),
        "uncertain": PatternFill(start_color="FFFCE5CD", end_color="FFFCE5CD", fill_type="solid"),
        "error": PatternFill(start_color="FFD9D2E9", end_color="FFD9D2E9", fill_type="solid"),
        "warning_red": PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid"),
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
        semantic_col = header_to_col.get("semantic_check_status")
        raw_val_col = header_to_col.get("raw_ocr_validation_status")
        conflict_col = header_to_col.get("source_guided_conflict")
        max_col = get_column_letter(ws.max_column)
        if status_col and ws.max_row >= 2:
            letter = get_column_letter(status_col)
            for key, fill in fills.items():
                ws.conditional_formatting.add(
                    f"A2:{max_col}{ws.max_row}",
                    FormulaRule(formula=[f'${letter}2="{key}"'], stopIfTrue=False, fill=fill),
                )
        if review_col and ws.max_row >= 2:
            letter = get_column_letter(review_col)
            ws.conditional_formatting.add(
                f"A2:{max_col}{ws.max_row}",
                FormulaRule(formula=[f"${letter}2=TRUE"], stopIfTrue=False, fill=fills["needs_review"]),
            )
        if semantic_col and ws.max_row >= 2:
            letter = get_column_letter(semantic_col)
            ws.conditional_formatting.add(
                f"A2:{max_col}{ws.max_row}",
                FormulaRule(formula=[f'ISNUMBER(SEARCH("mismatch",${letter}2))'], stopIfTrue=False, fill=fills["warning_red"]),
            )
            ws.conditional_formatting.add(
                f"A2:{max_col}{ws.max_row}",
                FormulaRule(formula=[f'ISNUMBER(SEARCH("warning",${letter}2))'], stopIfTrue=False, fill=fills["warning_red"]),
            )
        if raw_val_col and ws.max_row >= 2:
            letter = get_column_letter(raw_val_col)
            ws.conditional_formatting.add(
                f"A2:{max_col}{ws.max_row}",
                FormulaRule(formula=[f'ISNUMBER(SEARCH("not_confirmed",${letter}2))'], stopIfTrue=False, fill=fills["warning_red"]),
            )
        if conflict_col and ws.max_row >= 2:
            letter = get_column_letter(conflict_col)
            ws.conditional_formatting.add(
                f"A2:{max_col}{ws.max_row}",
                FormulaRule(formula=[f"${letter}2=TRUE"], stopIfTrue=False, fill=fills["warning_red"]),
            )
        # Add clickable local file hyperlinks when path columns are available.
        if ws.title == "Final Matches":
            src_name_col = header_to_col.get("source_image_filename")
            tgt_name_col = header_to_col.get("target_image_filename")
            src_path_col = header_to_col.get("source_image_path")
            tgt_path_col = header_to_col.get("target_image_path")
            if src_name_col and src_path_col:
                for row_idx in range(2, ws.max_row + 1):
                    name_cell = ws.cell(row=row_idx, column=src_name_col)
                    path_cell = ws.cell(row=row_idx, column=src_path_col)
                    path_val = str(path_cell.value or "").strip()
                    if path_val:
                        try:
                            name_cell.hyperlink = Path(path_val).resolve().as_uri()
                            name_cell.style = "Hyperlink"
                        except Exception:
                            pass
            if tgt_name_col and tgt_path_col:
                for row_idx in range(2, ws.max_row + 1):
                    name_cell = ws.cell(row=row_idx, column=tgt_name_col)
                    path_cell = ws.cell(row=row_idx, column=tgt_path_col)
                    path_val = str(path_cell.value or "").strip()
                    if path_val:
                        try:
                            name_cell.hyperlink = Path(path_val).resolve().as_uri()
                            name_cell.style = "Hyperlink"
                        except Exception:
                            pass
    wb.save(output_file)


def find_review_package_script() -> Path:
    candidates: List[Path] = []
    try:
        candidates.append(Path(__file__).resolve().parents[1] / "scripts" / "build_review_package.py")
    except Exception:
        pass
    candidates.append(Path.cwd() / "scripts" / "build_review_package.py")
    for candidate in candidates:
        if candidate.exists():
            return candidate
    searched = "; ".join(str(p) for p in candidates)
    raise FileNotFoundError(f"Review package builder script not found. Checked: {searched}")


def build_review_package_after_report(
    report_xlsx: Path,
    review_root: Path,
    mm_strings_xlsx: Path,
    ghostscript_cmd: Optional[str],
    dpi: int,
    target_lang: str,
) -> Path:
    script = find_review_package_script()
    cmd = [
        sys.executable,
        str(script),
        "--workbooks",
        str(report_xlsx),
        "--review-root",
        str(review_root),
        "--mm-strings-xlsx",
        str(mm_strings_xlsx),
        "--dpi",
        str(dpi),
        "--target-lang",
        target_lang,
    ]
    if ghostscript_cmd:
        cmd.extend(["--ghostscript-cmd", ghostscript_cmd])

    proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
    for line in (proc.stdout or "").splitlines():
        logging.info("Review package: %s", line)
    for line in (proc.stderr or "").splitlines():
        logging.warning("Review package: %s", line)
    if proc.returncode != 0:
        raise RuntimeError(f"Review package generation failed with exit code {proc.returncode}.")
    return review_root / f"MM_strings_review_EN{target_lang.upper()}.xlsx"


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
    parser.add_argument("--ocr-engine", choices=["classic", "classic_llm", "ai", "hybrid", "triangulated", "ai_validated"], default="ai_validated")
    parser.add_argument("--report-format", choices=["legacy", "multi_sheet"], default="legacy")
    parser.add_argument("--matching-mode", choices=["positional", "llm_text", "llm_objects"], default="llm_objects")
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
    parser.add_argument(
        "--om-strings-xlsx",
        type=Path,
        default=Path("input/MM_strings_EN.xlsx"),
        help="Optional MM English strings workbook used to generate the 'MM strings' worksheet in ai_validated mode.",
    )
    parser.add_argument("--disable-ai-audit-html", action="store_true", help="Disable AI request/response HTML audit generation.")
    parser.add_argument("--image-name", default=None, help="Optional source image filename filter for test runs.")
    parser.add_argument("--use-llm-matching", action="store_true", help="Use OpenAI LLM matching instead of positional line alignment.")
    parser.add_argument("--openai-model", default="gpt-4.1-mini", help="OpenAI model for LLM GUI string matching.")
    parser.add_argument("--openai-api-key", default=None, help="OpenAI API key. Falls back to OPENAI_API_KEY.")
    parser.add_argument(
        "--openai-cache-only",
        action="store_true",
        help="Use cached AI OCR outputs only and never call OpenAI; ai_validated matching falls back to local rules.",
    )
    parser.add_argument("--openai-audit-db", type=Path, default=Path(".openai_audit.sqlite"), help="SQLite database for OpenAI request/response audit and stage caches.")
    parser.add_argument("--disable-openai-audit-db", action="store_true", help="Disable SQLite OpenAI audit/cache database.")
    parser.add_argument("--restore-object-match-from-db-cache", action="store_true", help="Restore ai_validated object matching from SQLite cache when available.")
    parser.add_argument("--restore-semantic-check-from-db-cache", action="store_true", help="Restore semantic quality checks from SQLite cache when available.")
    parser.add_argument("--disable-db-stage-cache-write", action="store_true", help="Do not write stage cache entries into SQLite DB.")
    parser.add_argument("--llm-timeout-sec", type=int, default=45, help="Timeout for each LLM request.")
    parser.add_argument("--llm-max-retries", type=int, default=2, help="Maximum retry count for each LLM request.")
    parser.add_argument("--llm-workers", type=int, default=1, help="Parallel workers for independent source/target LLM stages per pair (2 recommended).")
    parser.add_argument("--output", type=Path, default=Path("ocr_match_report.xlsx"), help="Excel output path.")
    parser.add_argument("--log-file", type=Path, default=Path("ocr_match_report.log"), help="Log file path.")
    parser.add_argument(
        "--build-review-package",
        action="store_true",
        help="After the multi-sheet report is written, also build to_review/MM_strings_review_ENxx.xlsx with linked PNG images.",
    )
    parser.add_argument("--review-root", type=Path, default=Path("to_review"), help="Review package output root folder.")
    parser.add_argument("--review-package-dpi", type=int, default=300, help="DPI for PNG images in the review package.")
    parser.add_argument("--use-temp-local-copy", action="store_true", help="Copy files to local temp folder before OCR.")
    parser.add_argument("--verbose", action="store_true", help="Also print log messages to console.")
    parser.add_argument("--expert-debug-mode", action="store_true", help="Allow expert/debug OCR modes like triangulated.")
    args = parser.parse_args()

    configure_logging(args.log_file, args.verbose)
    logging.info("Starting discovery under root: %s", args.root.resolve())

    openai_api_key = None if args.openai_cache_only else (args.openai_api_key or os.environ.get("OPENAI_API_KEY"))
    if args.use_llm_matching and not openai_api_key:
        raise ValueError(
            "LLM matching requires an OpenAI API key. Pass --openai-api-key or set OPENAI_API_KEY."
        )
    if args.use_llm_matching:
        logging.info("LLM GUI matching enabled with model: %s", args.openai_model)
    if args.ocr_engine == "triangulated" and not args.expert_debug_mode:
        raise ValueError("triangulated mode is expert/debug only. Re-run with --expert-debug-mode.")
    if args.ocr_engine != "classic" and args.report_format == "legacy":
        logging.info("Non-classic OCR engine selected; switching report format to multi_sheet.")
        args.report_format = "multi_sheet"
    if args.ocr_engine in {"classic_llm", "ai", "hybrid", "triangulated", "ai_validated"} and args.matching_mode == "positional":
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

    audit_db: Optional[OpenAIAuditDB] = None
    if not args.disable_openai_audit_db:
        try:
            audit_db = OpenAIAuditDB(args.openai_audit_db)
            audit_db.start_run(
                {
                    "root": str(args.root.resolve()),
                    "target_lang": args.target_lang,
                    "ocr_engine": args.ocr_engine,
                    "matching_mode": args.matching_mode,
                    "cache_only": bool(args.openai_cache_only),
                }
            )
            logging.info("OpenAI audit DB enabled: %s (run_id=%s)", args.openai_audit_db.resolve(), audit_db.run_id)
        except Exception as exc:
            logging.warning("Failed to initialize OpenAI audit DB (%s). Continuing without DB. Error: %s", args.openai_audit_db, exc)
            audit_db = None

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
            openai_cache_only=args.openai_cache_only,
            audit_db=audit_db,
            restore_object_match_from_db_cache=args.restore_object_match_from_db_cache,
            restore_semantic_check_from_db_cache=args.restore_semantic_check_from_db_cache,
            write_stage_cache_to_db=not args.disable_db_stage_cache_write,
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
            llm_workers=max(1, int(args.llm_workers)),
        )
        write_multi_sheet_report(report, args.output, om_strings_xlsx=args.om_strings_xlsx)
        if audit_log is not None:
            audit_path = args.ai_audit_html or args.output.with_name(f"{args.output.stem}_ai_audit.html")
            audit_log.write_html(audit_path)
            logging.info("AI audit HTML written: %s", audit_path.resolve())
        if args.build_review_package:
            review_out = build_review_package_after_report(
                report_xlsx=args.output,
                review_root=args.review_root,
                mm_strings_xlsx=args.om_strings_xlsx,
                ghostscript_cmd=ghostscript_cmd,
                dpi=args.review_package_dpi,
                target_lang=lang_code,
            )
            logging.info("Review package written: %s", review_out.resolve())
        if audit_db is not None:
            audit_db.close()
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
    if audit_db is not None:
        audit_db.close()
    logging.info("Excel report written: %s", args.output.resolve())


if __name__ == "__main__":
    main()
