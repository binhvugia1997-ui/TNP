"""Improvement-image learning & correction mode (PROMPT-006).

Sits ON TOP of the deterministic After-picture detector (:mod:`app.improvement_pictures`, PROMPT-004C) and is fully
independent from Qwen/Ollama:

* every content picture of an improvement slide becomes an :class:`ImageCandidate` with deterministic, normalised
  PPT-layout features (no OCR, no screenshot pixels);
* the deterministic decision gets a calibrated confidence 0–1 and short Vietnamese evidence lines;
* the user labels candidates (AFTER / BEFORE / CONTROL / IGNORE) in the review window; labels are appended to
  ``learning_data/image_labels.jsonl`` (user data, preserved by the updater, never shipped in a release);
* a small local logistic-regression model (pure Python, structured features only) is trained on confirmed labels
  (AFTER vs NON_AFTER) once enough examples of BOTH classes exist and is stored in ``learning_data/model/``;
* final decision = hard exclusions + deterministic evidence + learned probability (uncertain cases only)
  + user-confirmed overrides (authoritative).  Missing/corrupt/incompatible model → deterministic rules only.
"""
from __future__ import annotations

import hashlib
import json
import logging
import math
import re
from dataclasses import asdict, dataclass, field
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from .content_region import ROLE_CAPTION, ROLE_CONTENT, classify_blocks, title_text
from .improvement_pictures import (PictureRef, PictureSelection, _arrow_usable, _arrow_votes, _blue_anchors,
                                   _caption_anchors, _caption_distance, _claims, _INSPECTION_RE, _inline_anchors,
                                   is_decorative_picture)
from .pptx_parser import Block, ReportData, SlideData, norm_key

LOG = logging.getLogger("report_extractor.image_learning")

IMAGE_FEATURE_SCHEMA = 1
LEARNING_DIR_NAME = "learning_data"
LABELS_FILE = "image_labels.jsonl"
THUMBS_DIR = "thumbnails"
MODEL_DIR = "model"
MODEL_FILE = "image_model.json"

LABELS = ("AFTER", "BEFORE", "CONTROL", "IGNORE")
UNLABELED = "UNLABELED"
LABEL_VI = {"AFTER": "Sau cải tiến", "BEFORE": "Trước cải tiến", "CONTROL": "Kiểm tra / Kiểm soát",
            "IGNORE": "Không lấy", UNLABELED: "Chưa xác nhận"}
NON_AFTER = "NON_AFTER"

# review thresholds (confidence that the picture is an After picture)
THRESHOLD_INCLUDE = 0.75
THRESHOLD_EXCLUDE = 0.35
# learning requirements – never train from a handful of examples or a single class
MIN_EXAMPLES = 10
MIN_PER_CLASS = 3

MSG_NOT_ENOUGH = "Chưa đủ dữ liệu học — đang dùng quy tắc hiện tại."
MSG_MODEL_UNAVAILABLE = "Mô hình ảnh không khả dụng — sử dụng quy tắc mặc định."

# ordered numeric feature vector consumed by the model (schema 1)
FEATURE_NAMES: Tuple[str, ...] = (
    "relative_x", "relative_y", "relative_width", "relative_height", "relative_area",
    "block_rel_x", "block_rel_y",
    "explicit_after", "explicit_before", "blue_after_text", "caption_distance",
    "has_arrow", "arrow_destination", "arrow_source", "arrow_horizontal",
    "picture_count", "picture_order", "is_last_in_block", "is_first_in_block",
    "production_block", "inspection_control", "logo_decorative", "deterministic_after", "deterministic_before",
)

CONFIDENCE_VI = {"high": "Cao", "medium": "Trung bình", "low": "Thấp"}


# ============================================================================ candidates
@dataclass
class ImageCandidate:
    candidate_id: str                 # stable identity (report key + slide + shape id) – no absolute path inside
    management_number: str
    source_name: str                  # file name only (diagnostics); classifier never depends on it
    source_file: str                  # full path for the current run (opening thumbnails / re-apply)
    slide: int
    picture_id: int                   # shape id in the slide
    order: int
    bounds: Tuple[int, int, int, int] # EMU left, top, width, height
    slide_size: Tuple[int, int]
    block_bounds: Tuple[int, int, int, int]
    features: Dict[str, float]
    nearby_text: str = ""
    nearest_heading: str = ""
    nearest_caption: str = ""
    nearest_text_color: str = ""
    deterministic_kind: str = "ambiguous"     # after | before | ambiguous | excluded
    deterministic_how: str = ""
    hard_excluded: bool = False
    hard_reason: str = ""
    confidence: float = 0.5                   # deterministic confidence that this is an After picture
    learned_probability: Optional[float] = None
    evidence: List[str] = field(default_factory=list)
    decision: str = "review"                  # include | exclude | review
    decision_source: str = "rules"            # rules | model | user
    user_label: str = UNLABELED
    thumbnail: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @property
    def label_text(self) -> str:
        return f"S{self.slide}#{self.picture_id}"

    @property
    def confidence_band(self) -> str:
        return confidence_band(self.confidence if self.learned_probability is None else self.learned_probability)


def confidence_band(c: float) -> str:
    if c >= THRESHOLD_INCLUDE or c <= 1 - THRESHOLD_INCLUDE:
        return "high"
    if c >= 0.6 or c <= 0.4:
        return "medium"
    return "low"


def candidate_id_for(report_key: str, slide: int, shape_id: int, order: int) -> str:
    return f"{report_key}|S{slide}|#{shape_id}|{order}"


def _rel(v: float, total: float) -> float:
    return round(v / total, 4) if total else 0.0


def _block_bounds(slide: SlideData, roles, W: int, H: int) -> Tuple[int, int, int, int]:
    """Bounding box of the business block = content text + content pictures (titles/sidebars/logos excluded)."""
    xs, ys, xe, ye = [], [], [], []
    for r in roles:
        if r.role == ROLE_CONTENT:
            b = r.block
            xs.append(b.left); ys.append(b.top); xe.append(b.right); ye.append(b.bottom)
    for p in slide.pictures:
        if not is_decorative_picture(p, W, H):
            xs.append(p.left); ys.append(p.top); xe.append(p.right); ye.append(p.bottom)
    if not xs:
        return 0, 0, W, H
    return min(xs), min(ys), max(xe) - min(xs), max(ye) - min(ys)


def _nearby_text(p: Block, roles, W: int, H: int) -> Tuple[str, str, str]:
    """(nearby content text, nearest caption text, nearest text colour) by geometric proximity."""
    best_txt, best_d = "", None
    cap_txt, cap_d = "", None
    color = ""
    for r in roles:
        b = r.block
        if not b.text.strip():
            continue
        d = _caption_distance(p, _as_anchor(b), W, H)
        if d is None:
            continue
        if r.role == ROLE_CAPTION:
            if cap_d is None or d < cap_d:
                cap_d, cap_txt = d, b.text.strip()
        elif r.role == ROLE_CONTENT:
            if best_d is None or d < best_d:
                best_d, best_txt = d, b.text.strip()
                color = next((c for c in (b.line_colors or []) if c), "")
    return best_txt[:160], cap_txt[:80], color


def _as_anchor(b: Block):
    from .improvement_pictures import Anchor
    return Anchor("x", b.left, b.top, b.width, b.height, "text")


def _deterministic_confidence(kind: str, how: str, f: Dict[str, float]) -> Tuple[float, List[str]]:
    """Calibrated confidence that the picture is an After picture + explainable factors (no hidden reasoning)."""
    ev: List[str] = []
    if f["logo_decorative"]:
        return 0.0, ["ảnh trang trí / logo / mũi tên – loại trừ cứng"]
    if f["inspection_control"]:
        ev.append("thuộc mục kiểm tra / kiểm soát")
    if f["production_block"]:
        ev.append("thuộc block cải tiến sản xuất")
    if f["explicit_after"]:
        ev.append("có chú thích 'Sau cải tiến' đi kèm")
    if f["explicit_before"]:
        ev.append("có chú thích 'Trước cải tiến' đi kèm")
    if f["blue_after_text"]:
        ev.append("gần nội dung chữ xanh (mô tả Sau)")
    if f["arrow_destination"]:
        ev.append("nằm phía đích mũi tên")
    if f["arrow_source"]:
        ev.append("nằm phía nguồn mũi tên (Trước)")
    how_l = (how or "").lower()
    if kind == "after":
        if how_l.startswith("caption") or how_l.startswith("inline"):
            c = 0.95
        elif "blue after text +" in how_l:
            c = 0.92
        elif how_l.startswith("row neighbour"):
            c = 0.85
        elif "blue" in how_l:
            c = 0.82
        else:                                   # arrow destination only
            c = 0.78
    elif kind == "before":
        if how_l.startswith("caption") or how_l.startswith("inline"):
            c = 0.05
        elif how_l.startswith("row neighbour"):
            c = 0.12
        else:
            c = 0.20
    elif kind == "excluded":
        c = 0.05 if f["inspection_control"] else 0.10
        ev.append(f"loại trừ theo quy tắc: {how}" if how else "loại trừ theo quy tắc")
    else:                                       # ambiguous – genuinely uncertain, stays in the review band
        c = 0.5
        if f["blue_after_text"] and f["arrow_source"]:
            c = 0.45
        ev.append("không có chú thích / chữ xanh / mũi tên quyết định → cần người xác nhận")
    return round(c, 2), ev


def build_candidates(report: ReportData, slide_numbers: Sequence[int], sel: PictureSelection,
                     management_number: str = "", source_file: str = "") -> List[ImageCandidate]:
    """Structured candidate records for every picture of the improvement slides (deterministic features only)."""
    W, H = report.slide_width or 1, report.slide_height or 1
    src = Path(source_file) if source_file else None
    report_key = management_number or (src.stem if src else "report")
    decided: Dict[Tuple[int, int], PictureRef] = {}
    for r in list(sel.after) + list(sel.rejected):
        decided[(r.slide, id(r.block))] = r
    out: List[ImageCandidate] = []
    for n in sorted(set(int(x) for x in slide_numbers)):
        s = report.slide(n)
        if s is None or not s.pictures:
            continue
        roles = classify_blocks(s)
        head = norm_key(title_text(s))
        insp_slide = bool(head and _INSPECTION_RE.search(head)
                          and not re.search(r"san xuat|production|process|cong doan", head))
        captions = _caption_anchors(roles)
        inlines, inspection_ranges = _inline_anchors(roles)
        from .extractor import excluded_bands
        bands = excluded_bands(s, H)
        content = [p for p in s.pictures if not is_decorative_picture(p, W, H)]
        claims = _claims(content, captions, W, H)
        blue = [a for a in _blue_anchors(roles, inspection_ranges) if not any(y0 <= a.cy < y1 for y0, y1, _k in bands)]
        arrows = [a for a in getattr(s, "arrows", []) if _arrow_usable(a, W, H, bands, inspection_ranges, roles)]
        votes = _arrow_votes(content, arrows, W, H)
        bx, by, bw, bh = _block_bounds(s, roles, W, H)
        pics = sorted(s.pictures, key=lambda b: b.order)
        n_content = len(content)
        content_ids = [id(p) for p in sorted(content, key=lambda b: b.order)]
        for p in pics:
            deco = is_decorative_picture(p, W, H)
            cy = p.top + p.height / 2
            in_insp = insp_slide or any(y0 <= cy < y1 for y0, y1 in inspection_ranges)
            mine = claims.get(id(p), [])
            near_cap = None
            for a in captions:
                d = _caption_distance(p, a, W, H)
                if d is not None and (near_cap is None or d < near_cap[0]):
                    near_cap = (d, a)
            cap_kind = mine[0].kind if len({a.kind for a in mine}) == 1 else (near_cap[1].kind if near_cap else "")
            inline_kind = ""
            if inlines:
                ordered = sorted(inlines, key=lambda a: a.top)
                for i, a in enumerate(ordered):
                    nxt = ordered[i + 1].top if i + 1 < len(ordered) else None
                    if a.top - 0.02 * H <= cy and (nxt is None or cy < nxt):
                        inline_kind = a.kind
            blue_d = None
            for a in blue:
                d = _caption_distance(p, a, W, H)
                if d is not None and (blue_d is None or d < blue_d):
                    blue_d = d
            vote = votes.get(id(p))
            arrow_dir = ""
            if vote:
                m = re.search(r"\b(right|left|up|down)\b", vote[1] or "")
                arrow_dir = m.group(1) if m else ""
            order_idx = content_ids.index(id(p)) if id(p) in content_ids else -1
            nearby, cap_txt, color = _nearby_text(p, roles, W, H)
            ref = decided.get((n, id(p)))
            kind = ref.kind if ref else ("excluded" if deco else "ambiguous")
            how = (ref.anchor or ref.reason) if ref else (deco or "")
            feats: Dict[str, float] = {
                "relative_x": _rel(p.left, W), "relative_y": _rel(p.top, H),
                "relative_width": _rel(p.width, W), "relative_height": _rel(p.height, H),
                "relative_area": _rel(p.width * p.height, W * H),
                "block_rel_x": _rel(p.left - bx, bw), "block_rel_y": _rel(p.top - by, bh),
                "explicit_after": 1.0 if (cap_kind == "after" or inline_kind == "after") else 0.0,
                "explicit_before": 1.0 if (cap_kind == "before" or inline_kind == "before") else 0.0,
                "blue_after_text": 1.0 if blue_d is not None else 0.0,
                "caption_distance": round(near_cap[0], 4) if near_cap else 1.0,
                "has_arrow": 1.0 if vote else 0.0,
                "arrow_destination": 1.0 if (vote and vote[0] == "after") else 0.0,
                "arrow_source": 1.0 if (vote and vote[0] == "before") else 0.0,
                "arrow_horizontal": 1.0 if arrow_dir in ("right", "left") else 0.0,
                "picture_count": float(n_content), "picture_order": float(order_idx + 1),
                "is_last_in_block": 1.0 if (order_idx >= 0 and order_idx == n_content - 1) else 0.0,
                "is_first_in_block": 1.0 if order_idx == 0 else 0.0,
                "production_block": 0.0 if in_insp else 1.0,
                "inspection_control": 1.0 if in_insp else 0.0,
                "logo_decorative": 1.0 if deco else 0.0,
                "deterministic_after": 1.0 if kind == "after" else 0.0,
                "deterministic_before": 1.0 if kind == "before" else 0.0,
            }
            conf, ev = _deterministic_confidence(kind, how, feats)
            c = ImageCandidate(
                candidate_id=candidate_id_for(report_key, n, p.shape_id, p.order),
                management_number=management_number, source_name=src.name if src else "",
                source_file=str(src) if src else "", slide=n, picture_id=p.shape_id, order=p.order,
                bounds=(p.left, p.top, p.width, p.height), slide_size=(W, H), block_bounds=(bx, by, bw, bh),
                features=feats, nearby_text=nearby, nearest_heading=title_text(s)[:80], nearest_caption=cap_txt,
                nearest_text_color=color, deterministic_kind=kind, deterministic_how=how,
                hard_excluded=bool(deco), hard_reason=deco or "", confidence=conf, evidence=ev)
            out.append(c)
    return out


def feature_vector(features: Dict[str, float]) -> List[float]:
    return [float(features.get(k, 0.0)) for k in FEATURE_NAMES]


# ============================================================================ label store (JSONL, user data)
def learning_dir(root: Optional[Path] = None, create: bool = True) -> Path:
    """``learning_data`` folder.  ``root=None`` -> beside the portable executable (runtime_paths.learning_dir)."""
    if root is None:
        from . import runtime_paths
        p = runtime_paths.learning_dir(create=create)
    else:
        p = Path(root) / LEARNING_DIR_NAME
    if create:
        try:
            for sub in ("", THUMBS_DIR, MODEL_DIR):
                (p / sub).mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
    return p


class LabelStore:
    """Append-only ``image_labels.jsonl``; the LAST record per candidate_id is the current label (audit kept)."""

    def __init__(self, directory: Path):
        self.dir = Path(directory)
        self.path = self.dir / LABELS_FILE

    def records(self) -> List[dict]:
        if not self.path.exists():
            return []
        out = []
        for line in self.path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except Exception:  # noqa: BLE001 – one damaged line never hides the others
                continue
            if isinstance(rec, dict) and rec.get("candidate_id") and rec.get("label") in LABELS:
                out.append(rec)
        return out

    def latest(self) -> Dict[str, dict]:
        cur: Dict[str, dict] = {}
        for rec in self.records():
            cur[rec["candidate_id"]] = rec
        return cur

    def current_label(self, candidate_id: str) -> str:
        rec = self.latest().get(candidate_id)
        return rec["label"] if rec else UNLABELED

    def label(self, cand: ImageCandidate, label: str, note: str = "") -> Optional[dict]:
        """Confirm ``label`` for ``cand``.  Returns the written record, or None when it is a duplicate of the
        current label (duplicate-label protection).  A changed label supersedes the old one (previous_label kept)."""
        if label not in LABELS:
            raise ValueError(f"Nhãn không hợp lệ: {label}")
        prev = self.latest().get(cand.candidate_id)
        if prev and prev.get("label") == label and prev.get("schema_version") == IMAGE_FEATURE_SCHEMA:
            return None
        rec = {
            "schema_version": IMAGE_FEATURE_SCHEMA, "candidate_id": cand.candidate_id, "label": label,
            "previous_label": prev.get("label") if prev else None,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "management_number": cand.management_number, "source_name": cand.source_name,
            "slide": cand.slide, "picture_id": cand.picture_id,
            "features": {k: cand.features.get(k, 0.0) for k in FEATURE_NAMES},
            "deterministic_kind": cand.deterministic_kind, "deterministic_confidence": cand.confidence,
            "note": note,
        }
        self.dir.mkdir(parents=True, exist_ok=True)
        with open(self.path, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
        LOG.info("IMAGE_LABEL candidate=%s label=%s previous=%s", cand.candidate_id, label, rec["previous_label"])
        return rec

    def counts(self) -> Dict[str, int]:
        cur = self.latest()
        after = sum(1 for r in cur.values() if r["label"] == "AFTER")
        return {"total": len(cur), "after": after, "non_after": len(cur) - after,
                "before": sum(1 for r in cur.values() if r["label"] == "BEFORE"),
                "control": sum(1 for r in cur.values() if r["label"] == "CONTROL"),
                "ignore": sum(1 for r in cur.values() if r["label"] == "IGNORE")}

    def export(self, target: Path) -> Path:
        target = Path(target)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps(self.records(), ensure_ascii=False, indent=1), encoding="utf-8")
        return target


def save_thumbnail(cand: ImageCandidate, blob: Optional[bytes], directory: Path, max_side: int = 240) -> str:
    """GUI-review thumbnail (never a learning input).  Returns the path or '' when not practical."""
    if not blob:
        return ""
    try:
        from PIL import Image
        im = Image.open(BytesIO(blob))
        im.load()
        im = im.convert("RGB")
        im.thumbnail((max_side, max_side))
        name = hashlib.sha1(cand.candidate_id.encode("utf-8")).hexdigest()[:16] + ".png"
        d = Path(directory) / THUMBS_DIR
        d.mkdir(parents=True, exist_ok=True)
        target = d / name
        im.save(target, "PNG", optimize=True)
        cand.thumbnail = str(target)
        return str(target)
    except Exception as e:  # noqa: BLE001
        LOG.debug("thumbnail skipped for %s: %s", cand.candidate_id, e)
        return ""


# ============================================================================ tiny local model (logistic regression)
@dataclass
class ImageModel:
    weights: List[float]
    bias: float
    means: List[float]
    scales: List[float]
    feature_names: List[str]
    schema_version: int
    trained_at: str
    n_examples: int
    n_after: int
    n_non_after: int
    train_accuracy: float = 0.0

    def to_dict(self) -> dict:
        return asdict(self)

    def probability(self, features: Dict[str, float]) -> float:
        x = feature_vector(features)
        z = self.bias
        for w, v, m, s in zip(self.weights, x, self.means, self.scales):
            z += w * ((v - m) / s)
        z = max(-30.0, min(30.0, z))
        return 1.0 / (1.0 + math.exp(-z))


class TrainingError(ValueError):
    """Vietnamese, user-facing: not enough data / single class / incompatible schema."""


def training_examples(records: Iterable[dict]) -> Tuple[List[List[float]], List[int]]:
    xs, ys = [], []
    for rec in records:
        if rec.get("schema_version") != IMAGE_FEATURE_SCHEMA or rec.get("label") not in LABELS:
            continue
        xs.append(feature_vector(rec.get("features") or {}))
        ys.append(1 if rec["label"] == "AFTER" else 0)                   # BEFORE/CONTROL/IGNORE -> NON_AFTER
    return xs, ys


def check_training_data(ys: Sequence[int]) -> Optional[str]:
    n_after = sum(ys)
    n_non = len(ys) - n_after
    if len(ys) < MIN_EXAMPLES or n_after < MIN_PER_CLASS or n_non < MIN_PER_CLASS:
        return (f"{MSG_NOT_ENOUGH} (cần ≥ {MIN_EXAMPLES} mẫu và ≥ {MIN_PER_CLASS} mẫu mỗi lớp; hiện có "
                f"{len(ys)} mẫu: Sau cải tiến {n_after}, Không phải Sau {n_non})")
    return None


def train_model(store_or_records, epochs: int = 400, lr: float = 0.1, l2: float = 0.01) -> ImageModel:
    """Batch gradient-descent logistic regression on standardised features.  Deterministic, dependency-free."""
    records = store_or_records.latest().values() if isinstance(store_or_records, LabelStore) else store_or_records
    xs, ys = training_examples(records)
    problem = check_training_data(ys)
    if problem:
        raise TrainingError(problem)
    n, d = len(xs), len(FEATURE_NAMES)
    means = [sum(x[j] for x in xs) / n for j in range(d)]
    scales = []
    for j in range(d):
        var = sum((x[j] - means[j]) ** 2 for x in xs) / n
        scales.append(math.sqrt(var) if var > 1e-12 else 1.0)
    xn = [[(x[j] - means[j]) / scales[j] for j in range(d)] for x in xs]
    w = [0.0] * d
    b = 0.0
    for _ in range(epochs):
        gw = [0.0] * d
        gb = 0.0
        for x, y in zip(xn, ys):
            z = b + sum(wi * xi for wi, xi in zip(w, x))
            p = 1.0 / (1.0 + math.exp(-max(-30.0, min(30.0, z))))
            err = p - y
            gb += err
            for j in range(d):
                gw[j] += err * x[j]
        b -= lr * gb / n
        for j in range(d):
            w[j] -= lr * (gw[j] / n + l2 * w[j])
    model = ImageModel(weights=[round(v, 6) for v in w], bias=round(b, 6), means=[round(v, 6) for v in means],
                       scales=[round(v, 6) for v in scales], feature_names=list(FEATURE_NAMES),
                       schema_version=IMAGE_FEATURE_SCHEMA, trained_at=datetime.now().isoformat(timespec="seconds"),
                       n_examples=n, n_after=sum(ys), n_non_after=n - sum(ys))
    correct = sum(1 for x, y in zip(xs, ys)
                  if (model.probability(dict(zip(FEATURE_NAMES, x))) >= 0.5) == bool(y))
    model.train_accuracy = round(correct / n, 3)
    return model


def model_path(directory: Path) -> Path:
    return Path(directory) / MODEL_DIR / MODEL_FILE


def save_model(model: ImageModel, directory: Path) -> Path:
    p = model_path(directory)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(model.to_dict(), ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(p)
    LOG.info("IMAGE_MODEL saved %s examples=%s after=%s non_after=%s", p, model.n_examples, model.n_after,
             model.n_non_after)
    return p


def load_model(directory: Path) -> Tuple[Optional[ImageModel], str]:
    """(model, status).  status: 'ok' | 'missing' | 'corrupt' | 'incompatible' – never raises."""
    p = model_path(directory)
    if not p.exists():
        return None, "missing"
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        model = ImageModel(**{k: data[k] for k in ImageModel.__dataclass_fields__ if k in data})
    except Exception as e:  # noqa: BLE001
        LOG.warning("IMAGE_MODEL corrupt %s: %s", p, e)
        return None, "corrupt"
    if model.schema_version != IMAGE_FEATURE_SCHEMA or list(model.feature_names) != list(FEATURE_NAMES) \
            or len(model.weights) != len(FEATURE_NAMES) or len(model.means) != len(FEATURE_NAMES):
        LOG.warning("IMAGE_MODEL incompatible schema=%s features=%s", model.schema_version, len(model.feature_names))
        return None, "incompatible"
    return model, "ok"


# ============================================================================ combination
def decide(cands: Sequence[ImageCandidate], model: Optional[ImageModel] = None,
           overrides: Optional[Dict[str, str]] = None) -> List[ImageCandidate]:
    """hard exclusions + deterministic evidence + learned probability (uncertain band only) + user overrides."""
    overrides = overrides or {}
    for c in cands:
        c.evidence = [e for e in c.evidence if not e.startswith("mô hình học") and not e.startswith("người dùng")]
        lbl = overrides.get(c.candidate_id, UNLABELED)
        c.user_label = lbl if lbl in LABELS else UNLABELED
        if c.hard_excluded:                                              # logos / arrows / icons: never inserted
            c.decision, c.decision_source = "exclude", "rules"
            continue
        if c.user_label != UNLABELED:                                    # authoritative over rules and model
            c.decision = "include" if c.user_label == "AFTER" else "exclude"
            c.decision_source = "user"
            c.evidence.append(f"người dùng đã xác nhận: {LABEL_VI[c.user_label]}")
            continue
        if c.confidence >= THRESHOLD_INCLUDE:
            c.decision, c.decision_source = "include", "rules"
        elif c.confidence <= THRESHOLD_EXCLUDE:
            c.decision, c.decision_source = "exclude", "rules"
        else:
            c.decision, c.decision_source = "review", "rules"
            if model is not None:
                p = model.probability(c.features)
                c.learned_probability = round(p, 3)
                if p >= THRESHOLD_INCLUDE:
                    c.decision, c.decision_source = "include", "model"
                    c.evidence.append("mô hình học: phù hợp với các mẫu Sau cải tiến đã xác nhận")
                elif p <= THRESHOLD_EXCLUDE:
                    c.decision, c.decision_source = "exclude", "model"
                    c.evidence.append("mô hình học: giống các mẫu Không phải Sau đã xác nhận")
                else:
                    c.evidence.append("mô hình học: chưa chắc chắn → cần người xác nhận")
    return list(cands)


def explain(c: ImageCandidate) -> str:
    result = {"include": LABEL_VI["AFTER"], "exclude": "Không lấy", "review": "Cần kiểm tra"}[c.decision]
    if c.decision == "exclude" and c.deterministic_kind == "before":
        result = LABEL_VI["BEFORE"]
    if c.user_label != UNLABELED:
        result = LABEL_VI[c.user_label]
    lines = [f"Ảnh slide {c.slide} (#{c.picture_id})", f"Kết quả: {result}",
             f"Độ tin cậy: {CONFIDENCE_VI[c.confidence_band]} ({c.confidence:.2f}"
             + (f", mô hình {c.learned_probability:.2f}" if c.learned_probability is not None else "") + ")",
             "Bằng chứng:"]
    lines += [f"- {e}" for e in c.evidence] or ["- (không có)"]
    if c.nearby_text:
        lines.append(f"Chữ gần ảnh: {c.nearby_text[:80]}")
    return "\n".join(lines)


# ============================================================================ runtime facade
class ImageLearning:
    """Everything the pipeline/GUI needs; safe when the learning folder is unavailable."""

    def __init__(self, directory: Optional[Path] = None, enabled: bool = True):
        self.dir = Path(directory) if directory else learning_dir()
        self.enabled = enabled
        self.store = LabelStore(self.dir)
        self.model: Optional[ImageModel] = None
        self.model_status = "missing"
        if enabled:
            self.reload_model()

    def reload_model(self) -> str:
        self.model, self.model_status = load_model(self.dir)
        return self.model_status

    def status_text(self) -> str:
        c = self.store.counts()
        base = f"Mẫu đã xác nhận: {c['total']}   Sau cải tiến: {c['after']}   Không phải Sau: {c['non_after']}"
        if self.model is not None:
            return (f"{base}   Mô hình: đã huấn luyện {self.model.trained_at[:16].replace('T', ' ')} "
                    f"({self.model.n_examples} mẫu)")
        if self.model_status in ("corrupt", "incompatible"):
            return f"{base}   {MSG_MODEL_UNAVAILABLE}"
        return f"{base}   {MSG_NOT_ENOUGH}"

    def overrides(self) -> Dict[str, str]:
        try:
            return {k: v["label"] for k, v in self.store.latest().items()}
        except OSError:
            return {}

    def apply(self, cands: Sequence[ImageCandidate]) -> List[ImageCandidate]:
        if not self.enabled:
            return decide(cands, None, {})
        return decide(cands, self.model, self.overrides())

    def train(self) -> Tuple[bool, str]:
        try:
            model = train_model(self.store)
        except TrainingError as e:
            return False, str(e)
        save_model(model, self.dir)
        self.model, self.model_status = model, "ok"
        return True, (f"Đã cập nhật mô hình ảnh từ {model.n_examples} mẫu (Sau cải tiến {model.n_after}, "
                      f"Không phải Sau {model.n_non_after}; khớp dữ liệu huấn luyện {model.train_accuracy:.0%}).")


def select_with_learning(report: ReportData, slide_numbers: Sequence[int], sel: PictureSelection,
                         learning: Optional[ImageLearning], management_number: str = "", source_file: str = ""
                         ) -> Tuple[List[PictureRef], List[ImageCandidate], List[str]]:
    """Final After refs after learning/overrides + candidates + extra review reasons."""
    cands = build_candidates(report, slide_numbers, sel, management_number, source_file)
    try:
        cands = learning.apply(cands) if learning else decide(cands)
    except Exception as e:  # noqa: BLE001 – learning must never stop processing
        LOG.warning("image learning skipped: %s", e)
        cands = decide(cands)
    by_key: Dict[Tuple[int, int], PictureRef] = {}
    for r in list(sel.after) + list(sel.rejected):
        by_key[(r.slide, r.block.shape_id)] = r
    refs: List[PictureRef] = []
    reasons: List[str] = []
    for c in cands:
        r = by_key.get((c.slide, c.picture_id))
        if r is None:
            continue
        if c.decision == "include":
            how = r.anchor if c.decision_source == "rules" else (
                "mô hình học" if c.decision_source == "model" else "người dùng xác nhận")
            refs.append(PictureRef(r.slide, r.block, "after", "", how))
        elif c.decision_source == "model" and c.decision == "exclude":
            reasons.append(f"Ảnh slide {c.slide} (#{c.picture_id}) bị loại theo mô hình học – kiểm tra nếu cần")
    refs.sort(key=lambda x: (x.slide, x.block.order))
    return refs, cands, reasons
