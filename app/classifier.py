"""Slide classification.

The LLM (Qwen via remote Ollama) is used ONLY to locate content: which slide
holds the QPN, root causes, improvement actions and improvement images.
The original text is always re-read from the PPTX afterwards.

A deterministic keyword classifier runs first and is used to:
  * guarantee the QPN slide (exact "Quality Problem Notice" text wins);
  * sanity-check / repair the LLM answer;
  * act as the full fallback when Ollama is not reachable.
"""
from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass, field, asdict
from typing import Any, Dict, List, Optional

from .pptx_parser import ReportData, SlideData, norm_key, slide_text_for_llm

LOG = logging.getLogger("report_extractor.classifier")

# ----------------------------------------------------------------------------
# Heading / section vocabulary (accent-insensitive keys, see norm_key)
# ----------------------------------------------------------------------------
QPN_MARKERS = ["quality problem notice", "quality problem notification"]
QPN_WEAK_MARKERS = ["qpn"]

CAUSE_MARKERS = ["nguyen nhan", "root cause", "phan tich nguyen nhan", "cause analysis",
                 "why why", "5 why", "5why", "fishbone", "nguyen nhan trong san xuat",
                 "nguyen nhan trong kiem tra"]
TEMP_MARKERS = ["xu ly tam thoi", "phuong phap xu ly tam thoi", "bien phap tam thoi",
                "doi sach tam thoi", "temporary action", "temporary countermeasure",
                "containment action", "containment", "khac phuc tam thoi", "xu ly tuc thoi"]
IMPROVE_MARKERS = ["cai tien", "doi sach", "doi sach cai tien", "corrective action",
                   "countermeasure", "permanent action", "improvement", "khac phuc",
                   "doi sach lau dai", "cai tien trong san xuat", "cai tien trong kiem tra",
                   "bien phap khac phuc", "hanh dong khac phuc", "long term action"]
DEFECT_MARKERS = ["noi dung loi", "hien tuong", "mo ta loi", "defect", "phenomenon",
                  "problem description", "noi dung phat sinh", "tinh trang loi"]
VERIFY_MARKERS = ["kiem chung", "xac nhan hieu qua", "verification", "effect confirmation",
                  "theo doi", "ket qua kiem chung", "hieu qua doi sach", "week +", "w+1"]
STANDARD_MARKERS = ["tieu chuan hoa", "standardization", "ngan ngua tai phat",
                    "horizontal deployment", "trien khai ngang"]


def section_kind_of_heading(text: str) -> Optional[str]:
    """Return section kind for a *heading-like* line, or None.

    Priority matters: temporary > cause > verify > standard > improvement > defect.
    ("Đối sách tạm thời" contains 'doi sach' but is temporary.)
    """
    k = norm_key(text)
    if not k:
        return None
    k = re.sub(r"^(?:[ivx]+|\d+(?: \d+)*|[a-z])\s+", "", k)   # drop numbering "3." / "III"
    if any(m in k for m in QPN_MARKERS):
        return "qpn"
    # report titles ("BÁO CÁO ĐỐI SÁCH LỖI ...", "Countermeasure report") are not section headings
    if re.search(r"\b(bao cao|report)\b", k):
        return None
    if any(m in k for m in TEMP_MARKERS):
        return "temporary"
    if any(m in k for m in CAUSE_MARKERS):
        return "cause"
    if any(m in k for m in VERIFY_MARKERS):
        return "verify"
    if any(m in k for m in STANDARD_MARKERS):
        return "standard"
    if any(m in k for m in IMPROVE_MARKERS):
        return "improvement"
    if any(m in k for m in DEFECT_MARKERS):
        return "defect"
    return None


def is_heading_like(text: str, bold: bool = False, size_pt: Optional[float] = None) -> bool:
    """Short, single-line, title-ish text."""
    t = text.strip()
    if not t or "\n" in t:
        return False
    if len(t) > 90:
        return False
    if re.match(r"^[\-+•·*▪➢►→>√✓\u2022\u25cf]", t):      # bullet lines are content, not headings
        return False
    letters = [c for c in t if c.isalpha()]
    upper_ratio = (sum(1 for c in letters if c.isupper()) / len(letters)) if letters else 0
    numbered = bool(re.match(r"^\s*(?:\d+(?:\.\d+)*[.)]?|[IVX]+[.)]|[A-Za-z][.)])\s+\S", t))
    return bool(numbered or bold or upper_ratio > 0.6 or (size_pt and size_pt >= 16) or len(t) <= 40)


# ----------------------------------------------------------------------------
@dataclass
class Classification:
    management_number: str = ""
    model: str = ""
    item: str = ""
    qpn_slide: Optional[int] = None
    defect_slide: Optional[int] = None
    cause_slides: List[int] = field(default_factory=list)
    improvement_slides: List[int] = field(default_factory=list)
    improvement_image_slides: List[int] = field(default_factory=list)
    temporary_slides: List[int] = field(default_factory=list)
    verify_slides: List[int] = field(default_factory=list)
    source: str = "heuristic"          # "ollama" | "heuristic" | "ollama+heuristic"
    raw_llm: Dict[str, Any] = field(default_factory=dict)
    notes: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d.pop("raw_llm", None)
        return d


# ----------------------------------------------------------------------------
# Heuristic classifier
# ----------------------------------------------------------------------------
def _slide_kinds(slide: SlideData) -> Dict[str, int]:
    """Count heading-like hits per section kind within a slide."""
    kinds: Dict[str, int] = {}
    for b in slide.text_blocks:
        lines = b.text.split("\n")
        for i, ln in enumerate(lines):
            if not ln.strip():
                continue
            heading = is_heading_like(ln, b.bold if i == 0 else False, b.size_pt if i == 0 else None)
            kind = section_kind_of_heading(ln)
            if kind == "qpn":
                kinds["qpn"] = kinds.get("qpn", 0) + 3
            elif kind and heading:
                kinds[kind] = kinds.get(kind, 0) + 1
    return kinds


def find_qpn_slide(report: ReportData) -> Optional[int]:
    """Slide containing 'Quality Problem Notice' text (exact/close), else 'QPN' heading."""
    best: Optional[int] = None
    for s in report.slides:
        k = norm_key(s.text)
        if any(m in k for m in QPN_MARKERS):
            return s.number
    for s in report.slides:
        for b in s.text_blocks:
            first = b.text.strip().split("\n")[0]
            if norm_key(first) in ("qpn", "qpn sheet", "phieu qpn") or re.fullmatch(r"qpn.{0,20}", norm_key(first)):
                best = s.number
                break
        if best:
            return best
    return None


def heuristic_classify(report: ReportData) -> Classification:
    c = Classification(source="heuristic")
    c.qpn_slide = find_qpn_slide(report)
    current = None   # section carried across slides when a slide has no heading
    for s in report.slides:
        kinds = _slide_kinds(s)
        if s.number == c.qpn_slide:
            if "defect" in kinds or not c.defect_slide:
                c.defect_slide = c.defect_slide or s.number
            continue
        dominant = None
        if kinds:
            for k in ("temporary", "cause", "improvement", "verify", "standard", "defect"):
                if k in kinds:
                    dominant = dominant or k
            # slide with several sections -> tag each
        tagged = set(kinds.keys())
        if not tagged and current and (s.text_blocks or s.pictures):
            tagged = {current}
        if "cause" in tagged:
            c.cause_slides.append(s.number)
        if "improvement" in tagged or "standard" in tagged:
            c.improvement_slides.append(s.number)
            if s.pictures:
                c.improvement_image_slides.append(s.number)
        if "temporary" in tagged:
            c.temporary_slides.append(s.number)
        if "verify" in tagged:
            c.verify_slides.append(s.number)
        if "defect" in tagged and not c.defect_slide:
            c.defect_slide = s.number
        # the last heading kind on the slide carries over
        if kinds:
            last_kind = None
            for b in s.text_blocks:
                for ln in b.text.split("\n"):
                    k = section_kind_of_heading(ln)
                    if k and is_heading_like(ln, b.bold, b.size_pt):
                        last_kind = k
            current = last_kind if last_kind in ("cause", "improvement", "temporary") else None
    if not c.defect_slide:
        c.defect_slide = c.qpn_slide or (1 if report.slides else None)
    return c


# ----------------------------------------------------------------------------
# LLM classifier
# ----------------------------------------------------------------------------
SYSTEM_PROMPT = (
    "You are a strict document-structure classifier for Vietnamese/English quality reports "
    "(8D / countermeasure reports). You NEVER rewrite, translate or summarize content. "
    "You only answer with JSON that points to slide numbers."
)

PROMPT_TEMPLATE = """Below is the text of every slide of a PowerPoint quality report (file name: {filename}).
Slides are numbered starting at 1. Pictures are indicated as [N picture(s)].

Identify WHERE information is located. Return ONLY a JSON object with exactly these keys:
{{
  "management_number": string,   // e.g. "260918080-VOC" if present in text/file name, else ""
  "model": string,               // product model code e.g. "A185" or "SM-A185", else ""
  "item": string,                // part/item name exactly as written e.g. "Rear", "Main", "Sub", "PBA", else ""
  "qpn_slide": int|null,         // slide containing the "Quality Problem Notice" (QPN) table
  "defect_slide": int|null,      // slide describing the defect content / quantities
  "cause_slides": [int],         // slides with root cause analysis (NGUYÊN NHÂN)
  "temporary_slides": [int],     // slides with temporary handling (XỬ LÝ TẠM THỜI / containment)
  "improvement_slides": [int],   // slides with corrective/improvement actions (CẢI TIẾN, ĐỐI SÁCH, ĐỐI SÁCH LÂU DÀI)
  "improvement_image_slides": [int], // improvement slides that contain before/after evidence pictures
  "verify_slides": [int]         // slides with verification / effect confirmation results
}}
Rules: use slide numbers only; do not include temporary-handling slides in improvement_slides
unless the same slide also contains permanent improvement content; copy model/item/management
number exactly as written; use "" or null or [] when unsure. No explanations.

=== SLIDES ===
{slides}
=== END ===
"""


def build_prompt(report: ReportData, max_chars_per_slide: int = 1800) -> str:
    parts = []
    for s in report.slides:
        parts.append(f"--- Slide {s.number} ---\n{slide_text_for_llm(s, max_chars_per_slide)}")
    return PROMPT_TEMPLATE.format(filename=report.filename, slides="\n".join(parts))


def _as_int_list(v: Any, n: int) -> List[int]:
    out: List[int] = []
    if v is None:
        return out
    if isinstance(v, (int, str)):
        v = [v]
    for x in v if isinstance(v, list) else []:
        try:
            i = int(str(x).strip().lstrip("Ss").replace("lide", ""))
        except ValueError:
            continue
        if 1 <= i <= n and i not in out:
            out.append(i)
    return sorted(out)


def _as_int(v: Any, n: int) -> Optional[int]:
    lst = _as_int_list(v, n)
    return lst[0] if lst else None


def merge_llm_with_heuristic(llm: Dict[str, Any], heur: Classification, report: ReportData) -> Classification:
    n = len(report.slides)
    c = Classification(source="ollama", raw_llm=llm)
    c.management_number = str(llm.get("management_number") or "").strip()
    c.model = str(llm.get("model") or "").strip()
    c.item = str(llm.get("item") or "").strip()
    c.qpn_slide = _as_int(llm.get("qpn_slide"), n)
    c.defect_slide = _as_int(llm.get("defect_slide"), n)
    c.cause_slides = _as_int_list(llm.get("cause_slides"), n)
    c.temporary_slides = _as_int_list(llm.get("temporary_slides"), n)
    c.improvement_slides = _as_int_list(llm.get("improvement_slides"), n)
    c.improvement_image_slides = _as_int_list(llm.get("improvement_image_slides"), n)
    c.verify_slides = _as_int_list(llm.get("verify_slides"), n)

    # --- guards: the exact QPN text always wins --------------------------------
    if heur.qpn_slide and c.qpn_slide != heur.qpn_slide:
        if c.qpn_slide:
            c.notes.append(f"LLM qpn_slide={c.qpn_slide} overridden by exact text match slide {heur.qpn_slide}")
        c.qpn_slide = heur.qpn_slide
        c.source = "ollama+heuristic"
    # --- fill gaps from heuristics ----------------------------------------------
    if not c.cause_slides and heur.cause_slides:
        c.cause_slides = list(heur.cause_slides)
        c.notes.append("cause_slides from heuristic")
        c.source = "ollama+heuristic"
    if not c.improvement_slides and heur.improvement_slides:
        c.improvement_slides = list(heur.improvement_slides)
        c.notes.append("improvement_slides from heuristic")
        c.source = "ollama+heuristic"
    # heuristic-detected improvement headings that LLM missed are appended
    for s_no in heur.improvement_slides:
        if s_no not in c.improvement_slides and s_no not in c.temporary_slides and s_no != c.qpn_slide:
            c.improvement_slides.append(s_no)
            c.notes.append(f"slide {s_no} added to improvement_slides (heading match)")
    c.improvement_slides.sort()
    for s_no in heur.temporary_slides:
        if s_no not in c.temporary_slides:
            c.temporary_slides.append(s_no)
    c.temporary_slides.sort()
    if not c.improvement_image_slides:
        c.improvement_image_slides = [s for s in c.improvement_slides
                                      if report.slide(s) and report.slide(s).pictures]
    else:
        c.improvement_image_slides = [s for s in c.improvement_image_slides
                                      if report.slide(s) and report.slide(s).pictures] or \
                                     [s for s in c.improvement_slides if report.slide(s) and report.slide(s).pictures]
    if not c.defect_slide:
        c.defect_slide = heur.defect_slide
    # never let the QPN slide be treated as improvement text
    c.improvement_slides = [s for s in c.improvement_slides if s != c.qpn_slide]
    return c


def classify(report: ReportData, ollama_client=None, model: str = "") -> Classification:
    """Full classification: heuristics + (optional) remote Qwen."""
    heur = heuristic_classify(report)
    if ollama_client is None or not model:
        heur.notes.append("Ollama not used")
        return heur
    prompt = build_prompt(report)
    try:
        llm = ollama_client.generate_json(model, prompt, system=SYSTEM_PROMPT)
        LOG.debug("LLM classification for %s: %s", report.filename, json.dumps(llm, ensure_ascii=False))
    except Exception as e:  # noqa: BLE001
        LOG.warning("Ollama classification failed for %s, using heuristics: %s", report.filename, e)
        heur.notes.append(f"Ollama failed: {e}")
        return heur
    return merge_llm_with_heuristic(llm, heur, report)
