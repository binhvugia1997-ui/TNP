"""Improvement-CONTENT learning & correction mode (PROMPT-006B).

Teaches Report Extractor which PowerPoint text blocks belong to ``CẢI TIẾN TRONG SẢN XUẤT`` (the Excel
improvement/countermeasure text).  The deterministic extractor (:func:`app.extractor.split_sections` +
:mod:`app.content_region`) stays the baseline; this module only decides WHETHER a block is included – never what
its text says (verbatim, source order, no rewriting).

* every text block of the improvement-related slides becomes a :class:`ContentCandidate` with normalised PPT
  geometry/typography/keyword features (``CONTENT_FEATURE_SCHEMA = 1``, separate from the image schema);
* deterministic decision + calibrated confidence + Vietnamese evidence lines;
* user labels (``IMPROVEMENT_CONTENT`` / ``EXCLUDE_CONTENT``) → ``learning_data/content_labels.jsonl``;
* tiny local logistic regression (shared implementation) → ``learning_data/model/content_model.json``;
* final decision = hard structural exclusions > exact user correction > strong deterministic evidence > model on
  the uncertain band > deterministic fallback.
"""
from __future__ import annotations

import logging
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from .classifier import INSPECTION_WORDS_RE, is_heading_like
from .content_region import (ROLE_CAPTION, ROLE_CONTENT, ROLE_FURNITURE, ROLE_SIDEBAR, ROLE_TITLE, classify_blocks,
                             inline_anchor_kind, title_text)
from .image_learning import (CONFIDENCE_VI, ImageModel, LabelStore, TrainingError, confidence_band, load_model,
                             save_model, train_model)
from .improvement_pictures import is_blue, is_decorative_picture
from .pptx_parser import Block, ReportData, clean_text, norm_key

LOG = logging.getLogger("report_extractor.content_learning")

CONTENT_FEATURE_SCHEMA = 1
CONTENT_LABELS_FILE = "content_labels.jsonl"
CONTENT_MODEL_FILE = "content_model.json"
CONTENT_LABELS = ("IMPROVEMENT_CONTENT", "EXCLUDE_CONTENT")
CONTENT_UNLABELED = "UNLABELED"
CONTENT_LABEL_VI = {"IMPROVEMENT_CONTENT": "Nội dung cải tiến", "EXCLUDE_CONTENT": "Không lấy",
                    CONTENT_UNLABELED: "Chưa xác nhận"}
# optional detailed internal classification (diagnostics only – the learned target is INCLUDE vs EXCLUDE)
DETAIL_KINDS = ("IMPROVEMENT_HEADING", "IMPROVEMENT_BODY", "BEFORE_TEXT", "AFTER_TEXT", "CONTROL_CONTENT",
                "TEMPORARY_CONTENT", "CAUSE_CONTENT", "DECORATION", "OTHER")

THRESHOLD_INCLUDE = 0.75
THRESHOLD_EXCLUDE = 0.35
# text blocks are far more numerous than pictures -> stronger minimum than the image classifier
MIN_EXAMPLES = 20
MIN_PER_CLASS = 5
MSG_NOT_ENOUGH = "Chưa đủ dữ liệu học nội dung cải tiến — đang dùng quy tắc hiện tại."
MSG_MODEL_UNAVAILABLE = "Mô hình nội dung không khả dụng — sử dụng quy tắc mặc định."

CONTENT_FEATURE_NAMES: Tuple[str, ...] = (
    "relative_x", "relative_y", "relative_width", "relative_height",
    "text_length", "paragraph_count", "reading_order", "is_bold", "font_size", "has_blue_text", "is_heading_like",
    "contains_cai_tien", "contains_truoc", "contains_sau", "contains_tam_thoi", "contains_inspection",
    "under_improvement_title", "under_temporary_title", "under_inspection_title", "in_improvement_item",
    "near_picture", "near_arrow",
    "role_title", "role_sidebar", "role_caption", "role_furniture",
    "det_improvement", "det_excluded_section",
)

_CAI_TIEN_RE = re.compile(r"cai tien|doi sach|improve")
_TRUOC_RE = re.compile(r"\btruoc\b|before")
_SAU_RE = re.compile(r"\bsau\b|after")
_TAM_THOI_RE = re.compile(r"xu ly tam thoi|tam thoi|temporary")
_SECTION_VI = {"improvement": "CẢI TIẾN TRONG SẢN XUẤT", "standard": "tiêu chuẩn / SOP", "inspection":
               "cải tiến kiểm tra / kiểm soát", "followup": "theo dõi / duy trì", "verify": "kiểm chứng",
               "temporary": "XỬ LÝ TẠM THỜI", "cause": "NGUYÊN NHÂN", "other": "chưa phân loại"}
INCLUDED_KINDS = ("improvement", "standard")
HARD_EXCLUDED_KINDS = ("temporary",)                       # known excluded temporary-action region
SOFT_EXCLUDED_KINDS = ("inspection", "followup", "verify", "cause")


@dataclass
class ContentCandidate:
    candidate_id: str                 # ManagementNumber|S<slide>|SH<shape>|<reading-order>
    management_number: str
    source_name: str
    source_file: str
    slide: int
    shape_id: int
    order: int
    shape_kind: str
    bounds: Tuple[int, int, int, int]
    slide_size: Tuple[int, int]
    text: str
    features: Dict[str, float]
    nearest_title: str = ""
    nearest_heading: str = ""
    section_kind: str = ""
    role: str = ROLE_CONTENT
    detail_kind: str = "OTHER"
    deterministic_kind: str = "other"         # improvement | excluded_section | furniture | other
    deterministic_how: str = ""
    hard_excluded: bool = False
    hard_reason: str = ""
    confidence: float = 0.5                   # deterministic confidence that the block is improvement content
    learned_probability: Optional[float] = None
    evidence: List[str] = field(default_factory=list)
    decision: str = "review"                  # include | exclude | review
    decision_source: str = "rules"
    user_label: str = CONTENT_UNLABELED
    baseline_included: bool = False           # what the deterministic extractor did with this block

    def to_dict(self) -> dict:
        return asdict(self)

    @property
    def label_text(self) -> str:
        return f"S{self.slide}#SH{self.shape_id}"

    @property
    def confidence_band(self) -> str:
        return confidence_band(self.confidence if self.learned_probability is None else self.learned_probability)


def content_candidate_id(report_key: str, slide: int, shape_id: int, order: int) -> str:
    return f"{report_key}|S{slide}|SH{shape_id}|{order}"


def _rel(v: float, total: float) -> float:
    return round(v / total, 4) if total else 0.0


def _near(b: Block, others: Sequence[Block], W: int, H: int) -> bool:
    """A picture/arrow is 'near' when it overlaps or sits within ~6 % of the slide of the block rectangle."""
    gx, gy = 0.06 * W, 0.06 * H
    for o in others:
        if o.left - gx <= b.right and o.right + gx >= b.left and o.top - gy <= b.bottom and o.bottom + gy >= b.top:
            return True
    return False


def _kinds_by_block(sections) -> Dict[Tuple[int, int], List[str]]:
    """{(slide, shape_id): [section kinds that hold at least one line of the block]}."""
    out: Dict[Tuple[int, int], List[str]] = {}
    for s in sections:
        for sid in s.sources:
            if sid < 0:
                continue
            lst = out.setdefault((s.slide, sid), [])
            if s.kind not in lst:
                lst.append(s.kind)
    return out


def _heading_by_block(sections) -> Dict[Tuple[int, int], str]:
    out: Dict[Tuple[int, int], str] = {}
    for s in sections:
        for sid in s.sources:
            if sid >= 0:
                out.setdefault((s.slide, sid), s.heading or (s.lines[0] if s.lines else ""))
    return out


def _deterministic(role: str, kinds: List[str], f: Dict[str, float], heading: str
                   ) -> Tuple[str, str, bool, str, float, List[str], str]:
    """-> (kind, how, hard, hard_reason, confidence, evidence, detail_kind)."""
    ev: List[str] = []
    if role == ROLE_TITLE:
        return "furniture", "slide title / header", True, "tiêu đề slide", 0.0, ["tiêu đề slide – loại trừ cứng"], "DECORATION"
    if role == ROLE_SIDEBAR:
        return "furniture", "sidebar label", True, "nhãn tròn / sidebar", 0.0, ["nhãn sidebar – loại trừ cứng"], "DECORATION"
    if role == ROLE_CAPTION:
        return ("furniture", "caption button", True, "nút chú thích ảnh", 0.0,
                ["nút chú thích ảnh (Trước/Sau cải tiến) – loại trừ cứng"], "DECORATION")
    if role == ROLE_FURNITURE:
        return ("furniture", "page furniture", True, "footer / logo / trang trí", 0.0,
                ["footer / logo / chữ trang trí – loại trừ cứng"], "DECORATION")
    if f["under_improvement_title"]:
        ev.append("nằm trong khu vực CẢI TIẾN TRONG SẢN XUẤT")
    if heading and f["in_improvement_item"]:
        ev.append(f"thuộc mục {heading.strip()[:60]}")
    if f["contains_sau"] and f["in_improvement_item"]:
        ev.append("có dòng + Sau:")
    if f["contains_truoc"] and f["in_improvement_item"]:
        ev.append("có dòng + Trước:")
    if f["has_blue_text"]:
        ev.append("có chữ xanh (mô tả Sau)")
    if any(k in HARD_EXCLUDED_KINDS for k in kinds) and not any(k in INCLUDED_KINDS for k in kinds):
        return ("excluded_section", "temporary section", True, "khu vực XỬ LÝ TẠM THỜI", 0.02,
                ev + ["thuộc khu vực XỬ LÝ TẠM THỜI – loại trừ theo quy tắc"], "TEMPORARY_CONTENT")
    if any(k in INCLUDED_KINDS for k in kinds):
        ev.append("thứ tự đọc phù hợp")
        if f["is_heading_like"] and f["paragraph_count"] <= 1:
            detail = "IMPROVEMENT_HEADING"
        elif f["contains_sau"]:
            detail = "AFTER_TEXT"
        elif f["contains_truoc"]:
            detail = "BEFORE_TEXT"
        else:
            detail = "IMPROVEMENT_BODY"
        strong = f["contains_sau"] or f["contains_truoc"] or f["has_blue_text"] or f["in_improvement_item"]
        c = 0.95 if strong else (0.88 if f["under_improvement_title"] else 0.8)
        if len(kinds) > 1:
            c = min(c, 0.8)
            ev.append("khối chứa cả dòng thuộc mục bị loại – kiểm tra")
        return "improvement", "improvement section", False, "", c, ev, detail
    if kinds:
        k = kinds[0]
        vi = _SECTION_VI.get(k, k)
        detail = {"inspection": "CONTROL_CONTENT", "followup": "CONTROL_CONTENT", "verify": "CONTROL_CONTENT",
                  "cause": "CAUSE_CONTENT"}.get(k, "OTHER")
        if k in SOFT_EXCLUDED_KINDS:
            return ("excluded_section", f"{k} section", False, "", 0.1,
                    ev + [f"thuộc khu vực {vi} – không phải nội dung cải tiến sản xuất"], detail)
        # "other": content on an improvement slide that no heading claimed -> uncertain
        c = 0.45
        if f["under_inspection_title"] or f["contains_inspection"]:
            c = 0.3
            ev.append("có từ ngữ kiểm tra / kiểm soát")
        if f["under_improvement_title"] and (f["contains_cai_tien"] or f["has_blue_text"]):
            c = 0.6
        ev.append("không có tiêu đề mục xác định → cần người xác nhận")
        return "other", "unclassified content", False, "", c, ev, "OTHER"
    ev.append("không thuộc khu vực nội dung nào đã nhận diện")
    return "other", "outside recognised sections", False, "", 0.4, ev, "OTHER"


def build_content_candidates(report: ReportData, sections, slide_numbers: Sequence[int], management_number: str = "",
                             source_file: str = "") -> List[ContentCandidate]:
    """Structured candidates for every TEXT block of the given slides (deterministic features only)."""
    W, H = report.slide_width or 1, report.slide_height or 1
    src = Path(source_file) if source_file else None
    report_key = management_number or (src.stem if src else "report")
    kinds_map = _kinds_by_block(sections)
    heading_map = _heading_by_block(sections)
    out: List[ContentCandidate] = []
    for n in sorted(set(int(x) for x in slide_numbers)):
        s = report.slide(n)
        if s is None:
            continue
        roles = classify_blocks(s)
        if not roles:
            continue
        title = title_text(s)
        tkey = norm_key(title)
        from .extractor import _semantic_kind                      # lazy: extractor imports learning modules
        tkind = _semantic_kind(title.splitlines()[0]) if title else None
        pics = [p for p in s.pictures if not is_decorative_picture(p, W, H)]
        arrows = list(getattr(s, "arrows", []))
        n_blocks = max(1, len(roles))
        for idx, r in enumerate(roles):
            b = r.block
            txt = b.text.strip()
            key = norm_key(txt)
            lines = [ln for ln in txt.splitlines() if ln.strip()]
            kinds = kinds_map.get((n, b.shape_id), [])
            heading = heading_map.get((n, b.shape_id), "")
            first = lines[0] if lines else ""
            in_item = bool(heading) and any(k in INCLUDED_KINDS for k in kinds) and norm_key(heading) != norm_key(first)
            inline = any(inline_anchor_kind(ln) for ln in lines)
            f: Dict[str, float] = {
                "relative_x": _rel(b.left, W), "relative_y": _rel(b.top, H),
                "relative_width": _rel(b.width, W), "relative_height": _rel(b.height, H),
                "text_length": round(min(len(txt), 600) / 600.0, 4),
                "paragraph_count": float(min(len(lines), 20)),
                "reading_order": round(idx / n_blocks, 4),
                "is_bold": 1.0 if b.bold else 0.0,
                "font_size": round(float(b.size_pt or 0.0) / 40.0, 4),
                "has_blue_text": 1.0 if any(is_blue(c) for c in (b.line_colors or [])) else 0.0,
                "is_heading_like": 1.0 if (lines and is_heading_like(first, b.bold, b.size_pt)) else 0.0,
                "contains_cai_tien": 1.0 if _CAI_TIEN_RE.search(key) else 0.0,
                "contains_truoc": 1.0 if (_TRUOC_RE.search(key) or inline) else 0.0,
                "contains_sau": 1.0 if (_SAU_RE.search(key) or inline) else 0.0,
                "contains_tam_thoi": 1.0 if _TAM_THOI_RE.search(key) else 0.0,
                "contains_inspection": 1.0 if INSPECTION_WORDS_RE.search(key) else 0.0,
                "under_improvement_title": 1.0 if tkind in INCLUDED_KINDS else 0.0,
                "under_temporary_title": 1.0 if tkind == "temporary" else 0.0,
                "under_inspection_title": 1.0 if tkind in ("inspection", "followup", "verify") else 0.0,
                "in_improvement_item": 1.0 if in_item else 0.0,
                "near_picture": 1.0 if _near(b, pics, W, H) else 0.0,
                "near_arrow": 1.0 if _near(b, arrows, W, H) else 0.0,
                "role_title": 1.0 if r.role == ROLE_TITLE else 0.0,
                "role_sidebar": 1.0 if r.role == ROLE_SIDEBAR else 0.0,
                "role_caption": 1.0 if r.role == ROLE_CAPTION else 0.0,
                "role_furniture": 1.0 if r.role == ROLE_FURNITURE else 0.0,
                "det_improvement": 1.0 if any(k in INCLUDED_KINDS for k in kinds) else 0.0,
                "det_excluded_section": 1.0 if (kinds and not any(k in INCLUDED_KINDS for k in kinds)) else 0.0,
            }
            kind, how, hard, hard_reason, conf, ev, detail = _deterministic(r.role, kinds, f, heading)
            out.append(ContentCandidate(
                candidate_id=content_candidate_id(report_key, n, b.shape_id, b.order),
                management_number=management_number, source_name=src.name if src else "",
                source_file=str(src) if src else "", slide=n, shape_id=b.shape_id, order=b.order, shape_kind=b.kind,
                bounds=(b.left, b.top, b.width, b.height), slide_size=(W, H), text=txt, features=f,
                nearest_title=title[:80], nearest_heading=heading[:80], section_kind=kinds[0] if kinds else "",
                role=r.role, detail_kind=detail, deterministic_kind=kind, deterministic_how=how, hard_excluded=hard,
                hard_reason=hard_reason, confidence=conf, evidence=ev,
                baseline_included=any(k in INCLUDED_KINDS for k in kinds)))
        _ = tkey
    return out


# ============================================================================ decision / explanation
def decide_content(cands: Sequence[ContentCandidate], model: Optional[ImageModel] = None,
                   overrides: Optional[Dict[str, str]] = None) -> List[ContentCandidate]:
    overrides = overrides or {}
    for c in cands:
        c.evidence = [e for e in c.evidence if not e.startswith("mô hình học") and not e.startswith("người dùng")]
        lbl = overrides.get(c.candidate_id, CONTENT_UNLABELED)
        c.user_label = lbl if lbl in CONTENT_LABELS else CONTENT_UNLABELED
        c.learned_probability = None
        if c.hard_excluded:                                              # 1. hard structural exclusions
            c.decision, c.decision_source = "exclude", "rules"
            continue
        if c.user_label != CONTENT_UNLABELED:                            # 2. exact user correction
            c.decision = "include" if c.user_label == "IMPROVEMENT_CONTENT" else "exclude"
            c.decision_source = "user"
            c.evidence.append(f"người dùng đã xác nhận: {CONTENT_LABEL_VI[c.user_label]}")
            continue
        if c.confidence >= THRESHOLD_INCLUDE:                            # 3. strong deterministic evidence
            c.decision, c.decision_source = "include", "rules"
        elif c.confidence <= THRESHOLD_EXCLUDE:
            c.decision, c.decision_source = "exclude", "rules"
        else:                                                            # 4./5. model on the uncertain band
            c.decision, c.decision_source = "review", "rules"
            if model is not None:
                p = model.probability(c.features)
                c.learned_probability = round(p, 3)
                if p >= THRESHOLD_INCLUDE:
                    c.decision, c.decision_source = "include", "model"
                    c.evidence.append("mô hình học: giống các khối Nội dung cải tiến đã xác nhận")
                elif p <= THRESHOLD_EXCLUDE:
                    c.decision, c.decision_source = "exclude", "model"
                    c.evidence.append("mô hình học: giống các khối Không lấy đã xác nhận")
                else:
                    c.evidence.append("mô hình học: chưa chắc chắn → giữ kết quả quy tắc hiện tại")
    return list(cands)


def explain_content(c: ContentCandidate) -> str:
    result = {"include": "Nội dung cải tiến", "exclude": "Không lấy", "review": "Chưa chắc chắn (giữ quy tắc hiện tại)"}[c.decision]
    if c.user_label != CONTENT_UNLABELED:
        result = CONTENT_LABEL_VI[c.user_label]
    lines = [f"Slide {c.slide}", f"Kết quả: {result}",
             f"Độ tin cậy: {CONFIDENCE_VI[c.confidence_band]} ({c.confidence:.2f}"
             + (f", mô hình {c.learned_probability:.2f}" if c.learned_probability is not None else "") + ")",
             "Bằng chứng:"]
    lines += [f"- {e}" for e in c.evidence] or ["- (không có)"]
    return "\n".join(lines)


def effective_included(c: ContentCandidate) -> bool:
    """Block included in the rebuilt text: review band falls back to the deterministic extractor's own choice."""
    if c.decision == "include":
        return True
    if c.decision == "exclude":
        return False
    return c.baseline_included


def rebuild_improvement_text(report: ReportData, sections, cands: Sequence[ContentCandidate]) -> Optional[str]:
    """Verbatim rebuild of the improvement text from the included blocks, in slide / reading order.

    Returns None when no learned/user decision changes the deterministic selection (baseline text kept
    byte-for-byte).  Lines of a block that the deterministic extractor already placed in a section are reused
    (same heading handling); a newly included block contributes its own non-empty lines."""
    changed = any(effective_included(c) != c.baseline_included for c in cands)
    if not changed:
        return None
    by_block: Dict[Tuple[int, int], List[str]] = {}
    for s in sections:
        for ln, sid in zip(s.lines, s.sources):
            if sid >= 0:
                by_block.setdefault((s.slide, sid), []).append(ln)
    parts: List[str] = []
    for c in sorted(cands, key=lambda c: (c.slide, c.order)):
        if not effective_included(c):
            continue
        lines = by_block.get((c.slide, c.shape_id))
        if not lines:
            lines = [ln.rstrip() for ln in c.text.splitlines() if ln.strip()]
        t = clean_text("\n".join(lines))
        if t and t not in parts:
            parts.append(t)
    return "\n\n".join(parts)


# ============================================================================ runtime facade
class ContentLearning:
    def __init__(self, directory: Path, enabled: bool = True):
        self.dir = Path(directory)
        self.enabled = enabled
        self.store = LabelStore(self.dir, CONTENT_LABELS_FILE, CONTENT_LABELS, CONTENT_FEATURE_NAMES,
                                CONTENT_FEATURE_SCHEMA, positive="IMPROVEMENT_CONTENT")
        self.model: Optional[ImageModel] = None
        self.model_status = "missing"
        if enabled:
            self.reload_model()

    def reload_model(self) -> str:
        self.model, self.model_status = load_model(self.dir, CONTENT_MODEL_FILE, CONTENT_FEATURE_SCHEMA,
                                                   CONTENT_FEATURE_NAMES)
        return self.model_status

    def counts(self) -> Dict[str, int]:
        c = self.store.counts()
        return {"total": c["total"], "include": c["positive"], "exclude": c["negative"]}

    def status_text(self) -> str:
        c = self.counts()
        base = f"Nội dung cải tiến: {c['include']}   Không lấy: {c['exclude']}"
        if self.model is not None:
            return f"{base}   Mô hình: đã huấn luyện ({self.model.n_examples} mẫu)"
        if self.model_status in ("corrupt", "incompatible"):
            return f"{base}   {MSG_MODEL_UNAVAILABLE}"
        return f"{base}   {MSG_NOT_ENOUGH}"

    def overrides(self) -> Dict[str, str]:
        try:
            return {k: v["label"] for k, v in self.store.latest().items()}
        except OSError:
            return {}

    def apply(self, cands: Sequence[ContentCandidate]) -> List[ContentCandidate]:
        if not self.enabled:
            return decide_content(cands, None, {})
        return decide_content(cands, self.model, self.overrides())

    def train(self) -> Tuple[bool, str]:
        try:
            model = train_model(self.store, min_examples=MIN_EXAMPLES, min_per_class=MIN_PER_CLASS,
                                msg=MSG_NOT_ENOUGH, pos_name="Nội dung cải tiến", neg_name="Không lấy")
        except TrainingError as e:
            return False, str(e)
        save_model(model, self.dir, CONTENT_MODEL_FILE)
        self.model, self.model_status = model, "ok"
        return True, (f"Đã cập nhật mô hình nội dung từ {model.n_examples} mẫu (Nội dung cải tiến {model.n_after}, "
                      f"Không lấy {model.n_non_after}; khớp dữ liệu huấn luyện {model.train_accuracy:.0%}).")


def select_content_with_learning(report: ReportData, sections, slide_numbers: Sequence[int],
                                 learning: Optional[ContentLearning], management_number: str = "",
                                 source_file: str = "") -> Tuple[Optional[str], List[ContentCandidate], List[str]]:
    """(rebuilt text or None, candidates, extra review reasons)."""
    cands = build_content_candidates(report, sections, slide_numbers, management_number, source_file)
    try:
        cands = learning.apply(cands) if learning else decide_content(cands)
    except Exception as e:  # noqa: BLE001 – learning must never stop processing
        LOG.warning("content learning skipped: %s", e)
        cands = decide_content(cands)
    reasons: List[str] = []
    for c in cands:
        if c.decision_source == "model" and effective_included(c) != c.baseline_included:
            what = "đưa vào" if c.decision == "include" else "bỏ"
            reasons.append(f"Khối chữ slide {c.slide} (#{c.shape_id}) được {what} theo mô hình học – kiểm tra nếu cần")
    text = rebuild_improvement_text(report, sections, cands)
    return text, cands, reasons
