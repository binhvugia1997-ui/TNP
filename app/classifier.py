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
import time
from dataclasses import dataclass, field, asdict
from typing import Any, Dict, List, Optional, Tuple

from .pptx_parser import ReportData, SlideData, norm_key

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
DEFECT_MARKERS = ["noi dung loi", "hien tuong", "hien trang", "thuc trang", "tinh trang", "mo ta loi", "defect",
                  "phenomenon", "problem description", "problem", "current status", "noi dung phat sinh",
                  "tinh trang loi", "mo ta van de", "noi dung van de"]
VERIFY_MARKERS = ["kiem chung", "xac nhan hieu qua", "verification", "effect confirmation", "effectiveness",
                  "theo doi", "ket qua kiem chung", "hieu qua doi sach", "hieu qua sau", "hieu qua cai tien",
                  "ket qua sau cai tien", "ket qua theo doi", "week +", "w+1"]
STANDARD_MARKERS = ["tieu chuan hoa", "standardization", "ngan ngua tai phat",
                    "horizontal deployment", "trien khai ngang"]


# Improvement sub-types (PROMPT-001).  The classifier keeps treating every one of them as an *improvement
# slide* (slide selection is unchanged); the section splitter uses the sub-type to decide what may enter
# "Nội dung đối sách cải tiến": only PRODUCTION/process improvements.  Inspection/control improvements and
# follow-up / effectiveness blocks contribute ZERO text.
INSPECTION_WORDS_RE = re.compile(r"\b(?:kiem tra|kiem soat|kiem hang|inspection|inspect|control|oqc|iqc|pqc|fqc|qc)\b")
PRODUCTION_WORDS_RE = re.compile(r"\b(?:san xuat|production|lap rap|assy|assembly|gia cong|cnc|son|ep|may|"
                                 r"nhua|khuon|jig|tool|thiet bi|may moc|process)\b")
FOLLOWUP_MARKERS = ["duy tri", "theo doi hieu qua", "theo doi", "audit", "ap dung cai tien", "giam sat",
                    "monitoring", "follow up", "follow-up", "sustain", "kiem tra thuong xuyen"]
LONG_TERM_MARKERS = ["doi sach lau dai", "giai phap lau dai", "bien phap lau dai", "long term", "long-term",
                     "permanent action"]
FOLLOWUP_LINE_RE = re.compile(r"theo doi|duy tri|audit|hieu qua|ap dung cai tien|kiem tra thuong xuyen|giam sat|"
                              r"monitor|follow[- ]?up|sustain")


def is_long_term_heading(text: str) -> bool:
    k = re.sub(r"^(?:[ivx]+|\d+(?: \d+)*|[a-z])\s+", "", norm_key(text))
    return any(m in k for m in LONG_TERM_MARKERS)


def improvement_subkind(text: str) -> str:
    """'inspection' | 'followup' | 'production' for an improvement-type heading.

    * inspection: improvement heading naming inspection/control ("CẢI TIẾN TRONG KIỂM TRA",
      "Cải tiến tại công đoạn kiểm tra:") without any production/process word;
    * followup: "Duy trì và áp dụng cải tiến", "Theo dõi hiệu quả ...", audit / sustain headings;
    * production: everything else (incl. "Đối sách lâu dài" – its BODY decides later, see extractor).
    """
    k = re.sub(r"^(?:[ivx]+|\d+(?: \d+)*|[a-z])\s+", "", norm_key(text))
    if not k:
        return "production"
    if any(m in k for m in FOLLOWUP_MARKERS):
        return "followup"
    if INSPECTION_WORDS_RE.search(k) and not PRODUCTION_WORDS_RE.search(k):
        return "inspection"
    return "production"


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
    source: str = "heuristic"          # "qwen" | "heuristic" | "qwen+heuristic"
    confidence: Optional[float] = None # reported by the LLM when available (0..1)
    ambiguities: List[str] = field(default_factory=list)   # -> "Cần kiểm tra"
    raw_llm: Dict[str, Any] = field(default_factory=dict)
    notes: List[str] = field(default_factory=list)
    # diagnostics (batch_result.json)
    qpn_source: str = ""               # explicit_text | metadata | qpn_heading | defect_structure | structural | llm | ""
    qpn_override: str = ""             # why an LLM qpn_slide was not accepted
    image_slides_structural: List[int] = field(default_factory=list)   # deterministic candidates
    image_slides_llm: List[int] = field(default_factory=list)          # what the LLM proposed

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
    return find_qpn_evidence(report)[0]


def find_qpn_evidence(report: ReportData) -> Tuple[Optional[int], str]:
    """Explicit QPN evidence -> (slide, source).  Sources: explicit_text (slide text),
    metadata (alt text / OLE or picture names / notes recovered from XML), qpn_heading."""
    for s in report.slides:
        k = norm_key(s.text)
        if any(m in k for m in QPN_MARKERS):
            return s.number, "explicit_text"
    # text hidden in alt-text / OLE object names / picture names / speaker notes
    for s in report.slides:
        meta = " ".join([*s.alt_texts, *(b.shape_name for b in s.blocks), s.notes])
        k = norm_key(meta)
        if any(m in k for m in QPN_MARKERS) or re.search(r"(?<![a-z])qpn(?![a-z])", k):
            return s.number, "metadata"
    for s in report.slides:
        for b in s.text_blocks:
            first = b.text.strip().split("\n")[0]
            if norm_key(first) in ("qpn", "qpn sheet", "phieu qpn") or re.fullmatch(r"qpn.{0,20}", norm_key(first)):
                return s.number, "qpn_heading"
    return None, ""


def slide_has_qpn_evidence(report: ReportData, n: Optional[int]) -> bool:
    """True when slide ``n`` itself carries explicit QPN evidence (text / metadata / heading)."""
    if not n:
        return False
    s = report.slide(n)
    if s is None:
        return False
    one = ReportData(path=report.path, slides=[s], slide_width=report.slide_width, slide_height=report.slide_height)
    return find_qpn_evidence(one)[0] is not None


def _largest_picture_ratio(slide: SlideData) -> float:
    """Area of the largest picture relative to the slide (0 when sizes are unknown)."""
    if not slide.pictures or not slide.width or not slide.height:
        return 0.0
    area = slide.width * slide.height
    return max((p.width * p.height) for p in slide.pictures) / float(area) if area else 0.0


def structural_qpn_candidate(report: ReportData, cover: bool, first_cause: Optional[int]) -> Optional[int]:
    """QPN form present only as a picture/OLE object: the first body slide (right after the
    cover) when it carries a large picture and is not later than the root-cause section.
    Deterministic, unique by construction; None when the layout does not match."""
    first_body = 2 if cover else 1
    s = report.slide(first_body)
    if s is None or not s.pictures:
        return None
    if first_cause and first_body > first_cause:
        return None
    if first_heading_kind(s) in ("temporary", "improvement", "verify", "standard"):
        return None
    ratio = _largest_picture_ratio(s)
    if s.width and s.height and ratio < 0.12:
        return None                          # small logo / icon, not a form
    return first_body


def first_heading_kind(slide: SlideData) -> Optional[str]:
    """Kind of the first heading-like line in reading order (the slide's own section)."""
    for b in slide.text_blocks:
        for i, ln in enumerate(b.text.split("\n")):
            if not ln.strip():
                continue
            k = section_kind_of_heading(ln)
            if k and is_heading_like(ln, b.bold if i == 0 else False, b.size_pt if i == 0 else None):
                return k
    return None


def is_cover_slide(report: ReportData, slide: SlideData) -> bool:
    """Slide 1 is a cover when it has no section heading of its own and a later slide
    starts the report body (defect / QPN / cause / temporary heading)."""
    if slide.number != 1:
        return False
    if first_heading_kind(slide) in ("defect", "cause", "temporary", "qpn"):
        return False
    return any(first_heading_kind(s) in ("defect", "cause", "temporary", "qpn") or
               any(m in norm_key(s.text) for m in QPN_MARKERS) for s in report.slides[1:])


def select_improvement_image_slides(report: ReportData, improvement_slides: List[int],
                                    qpn_slide: Optional[int], defect_slide: Optional[int],
                                    temporary_slides: List[int] = ()) -> List[int]:
    """Improvement evidence = pictures on slides whose OWN section is improvement.

    A slide qualifies when it has pictures and either its first heading is an
    improvement/standardisation heading, or it has no heading at all and directly
    continues an improvement slide.  Cover, QPN, defect (HIỆN TRẠNG), temporary and
    verification-led slides never qualify, even if some improvement text appears on them.
    """
    out: List[int] = []
    imp = set(improvement_slides)
    for s in report.slides:
        n = s.number
        if n not in imp or not s.pictures:
            continue
        if n in (qpn_slide, defect_slide) or n in temporary_slides or is_cover_slide(report, s):
            continue
        fk = first_heading_kind(s)
        if fk in ("improvement", "standard"):
            out.append(n)
        elif fk is None and (n - 1) in imp and (n - 1) in out + [x for x in imp if first_heading_kind(report.slide(x)) in ("improvement", "standard")]:
            out.append(n)
    return out


def is_safe_improvement_image_candidate(report: ReportData, n: int, improvement_slides: List[int],
                                        qpn_slide: Optional[int], defect_slide: Optional[int],
                                        temporary_slides: List[int]) -> bool:
    """Structural safety check applied to LLM-proposed improvement-image slides: the slide
    must belong to the improvement section, carry pictures and not be the cover, the
    QPN/defect slide, a temporary-handling slide or a verification-led slide."""
    s = report.slide(n)
    if s is None or not s.pictures or n not in improvement_slides:
        return False
    if n in (qpn_slide, defect_slide) or n in temporary_slides or is_cover_slide(report, s):
        return False
    return first_heading_kind(s) not in ("verify", "temporary", "defect", "qpn", "cause")


def heuristic_classify(report: ReportData) -> Classification:
    c = Classification(source="heuristic")
    c.qpn_slide, c.qpn_source = find_qpn_evidence(report)
    current = None   # section carried across slides when a slide has no heading
    for s in report.slides:
        kinds = _slide_kinds(s)
        if s.number == c.qpn_slide:
            if "defect" in kinds or not c.defect_slide:
                c.defect_slide = c.defect_slide or s.number
            continue
        if is_cover_slide(report, s):
            c.notes.append("slide 1 treated as cover (no section content)")
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
        c.defect_slide = c.qpn_slide
    # QPN precedence when no explicit text: (2) defect / HIỆN TRẠNG page carrying a picture,
    # (3) structural: first body slide with a large picture before the cause section.
    if not c.qpn_slide and c.defect_slide and report.slide(c.defect_slide) and report.slide(c.defect_slide).pictures:
        c.qpn_slide = c.defect_slide
        c.qpn_source = "defect_structure"
        c.notes.append(f"QPN = trang hiện trạng/nội dung lỗi có hình (slide {c.defect_slide}); "
                       f"không có chữ 'Quality Problem Notice' trong text")
    if not c.qpn_slide:
        cover = report.slides and is_cover_slide(report, report.slides[0])
        cand = structural_qpn_candidate(report, bool(cover), c.cause_slides[0] if c.cause_slides else None)
        if cand:
            c.qpn_slide = cand
            c.qpn_source = "structural"
            if not c.defect_slide:
                c.defect_slide = cand
            c.notes.append(f"QPN = slide {cand} theo cấu trúc (trang đầu thân báo cáo có hình biểu mẫu, "
                           f"trước phần nguyên nhân)")
    c.improvement_image_slides = select_improvement_image_slides(
        report, c.improvement_slides, c.qpn_slide, c.defect_slide, c.temporary_slides)
    c.image_slides_structural = list(c.improvement_image_slides)
    return c


# ----------------------------------------------------------------------------
# LLM classifier
# ----------------------------------------------------------------------------
SYSTEM_PROMPT = (
    "You classify slides of a Vietnamese/English quality countermeasure report. "
    "Answer with one compact JSON object of slide numbers only. Never rewrite, translate or "
    "summarize content. No explanations, no reasoning text."
)

# Compact prompt: the LLM only says WHERE things are (slide numbers).  Text values
# (management number, model, item, defect, cause, improvement) are always copied from the
# PPTX by the program, so they are not requested here.
PROMPT_TEMPLATE = """Slides of a PowerPoint quality report ({n} slides). Each line: S<number> | title | pics=<pictures> | text.
Return ONLY this JSON (slide numbers as ints; null/[] when unsure):
{{"qpn_slide":int|null,"defect_slide":int|null,"cause_slides":[int],"temporary_slides":[int],"improvement_slides":[int],"improvement_image_slides":[int],"verify_slides":[int],"confidence":0..1}}
qpn_slide = "Quality Problem Notice" table; defect_slide = defect content/quantities (HIỆN TRẠNG);
cause_slides = root cause (NGUYÊN NHÂN); temporary_slides = XỬ LÝ TẠM THỜI/containment (never in improvement_slides);
improvement_slides = CẢI TIẾN/ĐỐI SÁCH/ĐỐI SÁCH LÂU DÀI; improvement_image_slides = improvement slides with before/after pictures;
verify_slides = verification/effect results.
{slides}
"""

MAX_PROMPT_CHARS_PER_SLIDE = 420
MAX_TITLE_CHARS = 90


def compact_slide_line(slide: SlideData, max_chars: int = MAX_PROMPT_CHARS_PER_SLIDE) -> str:
    """One compact line per slide: number, title, picture count, de-duplicated text.

    Repeated shape text, table pipes and blank lines are collapsed; only text is sent –
    never XML, image data or shape inventories.
    """
    lines: List[str] = []
    seen = set()
    for blk in slide.text_blocks:
        for ln in blk.text.split("\n"):
            t = re.sub(r"\s*\|\s*", " ", ln).strip()
            t = re.sub(r"\s+", " ", t)
            if len(t) < 2:
                continue
            k = norm_key(t)
            if k in seen:
                continue
            seen.add(k)
            lines.append(t)
    title = lines[0][:MAX_TITLE_CHARS] if lines else ""
    body = " / ".join(lines[1:])
    budget = max_chars - len(title)
    if len(body) > budget:
        body = body[:max(0, budget)].rstrip() + "…"
    return f"S{slide.number} | {title} | pics={len(slide.pictures)} | {body}"


def build_prompt(report: ReportData, max_chars_per_slide: int = MAX_PROMPT_CHARS_PER_SLIDE) -> str:
    slides = "\n".join(compact_slide_line(s, max_chars_per_slide) for s in report.slides)
    return PROMPT_TEMPLATE.format(n=len(report.slides), slides=slides)


def _ints_from_any(v: Any) -> List[int]:
    """Pull slide numbers out of ints, "3", "Slide 3", "3, 5-7", [..], {"slide": 3}, {"slides": [...]}."""
    out: List[int] = []
    if v is None or isinstance(v, bool):
        return out
    if isinstance(v, (int, float)):
        return [int(v)]
    if isinstance(v, str):
        txt = v.lower()
        if txt.strip() in ("", "null", "none", "n/a", "na", "-"):
            return out
        for a, b in re.findall(r"(\d+)\s*[-–]\s*(\d+)", txt):
            out.extend(range(int(a), int(b) + 1))
        txt = re.sub(r"\d+\s*[-–]\s*\d+", " ", txt)
        out.extend(int(x) for x in re.findall(r"\d+", txt))
        return out
    if isinstance(v, dict):
        for k in ("slide", "slides", "slide_number", "slide_numbers", "number", "index", "page", "pages"):
            if k in v:
                out.extend(_ints_from_any(v[k]))
        if not out:
            for val in v.values():
                out.extend(_ints_from_any(val))
        return out
    if isinstance(v, (list, tuple, set)):
        for x in v:
            out.extend(_ints_from_any(x))
    return out


def _as_int_list(v: Any, n: int) -> List[int]:
    out: List[int] = []
    for i in _ints_from_any(v):
        if 1 <= i <= n and i not in out:
            out.append(i)
    return sorted(out)


# alias groups: canonical key -> accepted key fragments (normalised: lower, non-alnum -> _)
_KEY_ALIASES: Dict[str, List[str]] = {
    "management_number": ["management_number", "management_no", "managementnumber", "mgmt_no", "management",
                          "ma_quan_ly", "so_quan_ly"],
    "model": ["model", "model_name", "model_code", "product_model"],
    "item": ["item", "item_name", "part", "part_name", "hang_muc", "linh_kien"],
    "qpn_slide": ["qpn_slide", "qpn_slides", "qpn", "quality_problem_notice", "quality_problem_notice_slide", "qpn_page"],
    "defect_slide": ["defect_slide", "defect_slides", "defect", "defect_content_slide", "noi_dung_loi", "phenomenon_slide",
                     "problem_slide"],
    "cause_slides": ["cause_slides", "cause_slide", "root_cause_slides", "root_cause_slide", "root_cause", "cause",
                     "causes", "nguyen_nhan", "analysis_slides"],
    "temporary_slides": ["temporary_slides", "temporary_slide", "temporary_handling_slides", "temporary_handling",
                         "temporary_action_slides", "temporary_actions", "containment_slides", "containment",
                         "xu_ly_tam_thoi", "temp_slides", "interim_action_slides"],
    "improvement_slides": ["improvement_slides", "improvement_slide", "improvement", "corrective_action_slides",
                           "corrective_actions", "corrective_action", "countermeasure_slides", "countermeasures",
                           "countermeasure", "permanent_action_slides", "permanent_actions", "action_slides",
                           "doi_sach", "cai_tien", "long_term_slides", "long_term_action_slides", "long_term_actions"],
    "improvement_image_slides": ["improvement_image_slides", "improvement_images", "image_slides", "picture_slides",
                                 "evidence_slides", "before_after_slides", "improvement_picture_slides", "images"],
    "verify_slides": ["verify_slides", "verification_slides", "verification", "effect_slides",
                      "effect_confirmation_slides", "kiem_chung", "result_slides"],
    "confidence": ["confidence", "confidence_score", "certainty", "score"],
}


def _norm_json_key(k: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(k).lower()).strip("_")


def normalize_llm_response(raw: Dict[str, Any]) -> Dict[str, Any]:
    """Map arbitrary-but-reasonable LLM JSON onto our canonical keys.

    Handles nested containers (``{"slides": {...}}``, ``{"result": {...}}``), key
    synonyms, and long-term action keys (merged into improvement_slides).
    """
    flat: Dict[str, Any] = {}

    def walk(obj: Any, depth: int = 0) -> None:
        if not isinstance(obj, dict) or depth > 3:
            return
        for k, v in obj.items():
            nk = _norm_json_key(k)
            if isinstance(v, dict) and not any(x in nk for x in ("slide", "qpn", "cause", "improvement", "temporary")):
                walk(v, depth + 1)      # generic container ("result", "classification", "slides")
                continue
            flat.setdefault(nk, v)
            if isinstance(v, dict):
                walk(v, depth + 1)
    walk(raw or {})

    out: Dict[str, Any] = {}
    for canon, aliases in _KEY_ALIASES.items():
        vals = [flat[a] for a in aliases if a in flat]
        if not vals:
            # fuzzy: alias contained in key (e.g. "qpn_slide_number")
            vals = [v for k, v in flat.items() if any(k.startswith(a) or k.endswith(a) for a in aliases)]
        if not vals:
            continue
        if canon in ("management_number", "model", "item"):
            v = vals[0]
            out[canon] = "" if v is None else (str(v).strip() if not isinstance(v, (list, dict)) else "")
        elif canon == "confidence":
            try:
                c = float(vals[0])
                out[canon] = c / 100.0 if c > 1 else c
            except (TypeError, ValueError):
                pass
        elif canon in ("qpn_slide", "defect_slide"):
            out[canon] = vals[0]
        else:
            merged: List[Any] = []
            for v in vals:
                merged.append(v)
            out[canon] = merged
    return out


def _as_int(v: Any, n: int) -> Optional[int]:
    lst = _as_int_list(v, n)
    return lst[0] if lst else None


def merge_llm_with_heuristic(llm_raw: Dict[str, Any], heur: Classification, report: ReportData) -> Classification:
    n = len(report.slides)
    llm = normalize_llm_response(llm_raw)
    c = Classification(source="qwen", raw_llm=llm_raw)
    c.confidence = llm.get("confidence")
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

    # --- guards: deterministic QPN evidence always wins over the LLM -------------
    llm_qpn = c.qpn_slide
    c.image_slides_llm = list(c.improvement_image_slides)
    if heur.qpn_slide:
        c.qpn_source = heur.qpn_source
        if llm_qpn and llm_qpn != heur.qpn_slide:
            c.qpn_override = (f"LLM qpn_slide={llm_qpn} rejected: deterministic evidence "
                              f"({heur.qpn_source}) = slide {heur.qpn_slide}")
            c.notes.append(c.qpn_override)
            c.source = "qwen+heuristic"
        c.qpn_slide = heur.qpn_slide
    elif llm_qpn:
        cover = report.slides and is_cover_slide(report, report.slides[0]) and llm_qpn == 1
        if cover and not slide_has_qpn_evidence(report, llm_qpn):
            c.qpn_override = f"LLM qpn_slide={llm_qpn} rejected: cover slide without QPN evidence"
            c.notes.append(c.qpn_override)
            c.qpn_slide = None
            c.source = "qwen+heuristic"
        else:
            c.qpn_source = "llm"
            c.ambiguities.append(f"QPN (slide {llm_qpn}) chỉ do AI xác định, không có bằng chứng QPN trong cấu trúc")
    # --- ambiguity: AI and keyword detection disagree completely ----------------
    if c.improvement_slides and heur.improvement_slides and not set(c.improvement_slides) & set(heur.improvement_slides):
        c.ambiguities.append(f"Slide đối sách không thống nhất: AI {c.improvement_slides} / từ khoá {heur.improvement_slides}")
    if c.cause_slides and heur.cause_slides and not set(c.cause_slides) & set(heur.cause_slides):
        c.ambiguities.append(f"Slide nguyên nhân không thống nhất: AI {c.cause_slides} / từ khoá {heur.cause_slides}")
    # temporary slides (by heading) must never be used as improvement text
    tmp_only = [s_no for s_no in heur.temporary_slides if s_no in c.improvement_slides
                and s_no not in heur.improvement_slides]
    if tmp_only:
        c.notes.append(f"slides {tmp_only} removed from improvement_slides (temporary-handling heading)")
        c.improvement_slides = [s_no for s_no in c.improvement_slides if s_no not in tmp_only]
    if c.confidence is not None and c.confidence < 0.5:
        c.ambiguities.append(f"AI báo độ tin cậy thấp ({c.confidence:.2f})")
    # --- fill gaps from heuristics ----------------------------------------------
    if not c.cause_slides and heur.cause_slides:
        c.cause_slides = list(heur.cause_slides)
        c.notes.append("cause_slides from heuristic")
        c.source = "qwen+heuristic"
    if not c.improvement_slides and heur.improvement_slides:
        c.improvement_slides = list(heur.improvement_slides)
        c.notes.append("improvement_slides from heuristic")
        c.source = "qwen+heuristic"
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
    if not c.defect_slide:
        c.defect_slide = heur.defect_slide
    structural = select_improvement_image_slides(report, c.improvement_slides, c.qpn_slide, c.defect_slide,
                                                 c.temporary_slides)
    c.image_slides_structural = list(structural)
    # final = deterministic set UNION safe LLM candidates; the LLM can never remove a valid slide
    added, rejected = [], []
    for s_no in c.image_slides_llm:
        if s_no in structural:
            continue
        if is_safe_improvement_image_candidate(report, s_no, c.improvement_slides, c.qpn_slide,
                                               c.defect_slide, c.temporary_slides):
            added.append(s_no)
        else:
            rejected.append(s_no)
    if rejected:
        c.notes.append(f"LLM improvement_image_slides {rejected} rejected (fail structural checks)")
    if added:
        c.notes.append(f"LLM improvement_image_slides {added} added (pass structural checks)")
    narrowed = [s_no for s_no in structural if s_no not in c.image_slides_llm]
    if narrowed and c.image_slides_llm:
        c.notes.append(f"LLM omitted structural improvement images {narrowed}; kept")
    c.improvement_image_slides = sorted(set(structural) | set(added))
    if not c.qpn_slide and heur.qpn_slide:
        c.qpn_slide = heur.qpn_slide
    c.notes.extend(n_ for n_ in heur.notes if n_.startswith("QPN =") and n_ not in c.notes)
    # never let the QPN slide be treated as improvement text
    c.improvement_slides = [s for s in c.improvement_slides if s != c.qpn_slide]
    return c


def classify(report: ReportData, ollama_client=None, model: str = "") -> Classification:
    """Full classification: heuristics + (optional) remote Qwen."""
    heur = heuristic_classify(report)
    if ollama_client is None or not model:
        heur.notes.append("Ollama not used (no server/model configured)")
        return heur
    prompt = build_prompt(report)
    LOG.info("Ollama request: %s slides, prompt %d chars (+system %d), model %s",
             len(report.slides), len(prompt), len(SYSTEM_PROMPT), model)
    t0 = time.perf_counter()
    try:
        llm = ollama_client.generate_json(model, prompt, system=SYSTEM_PROMPT, num_ctx=4096, num_predict=256)
        LOG.debug("LLM classification for %s: %s", report.filename, json.dumps(llm, ensure_ascii=False))
    except Exception as e:  # noqa: BLE001
        LOG.warning("Ollama classification failed for %s after %.1fs, using heuristics: %s",
                    report.filename, time.perf_counter() - t0, e)
        heur.notes.append(f"Ollama failed after {time.perf_counter() - t0:.1f}s, heuristic fallback: {e}")
        return heur
    c = merge_llm_with_heuristic(llm, heur, report)
    stats = getattr(ollama_client, "last_call", None) or {}
    c.notes.append(f"Ollama ok in {time.perf_counter() - t0:.1f}s (prompt {len(prompt)} chars, "
                   f"response {stats.get('response_chars', '?')} chars, HTTP {stats.get('status', '?')})")
    LOG.info("Classifier for %s: %s", report.filename, c.source)
    return c
