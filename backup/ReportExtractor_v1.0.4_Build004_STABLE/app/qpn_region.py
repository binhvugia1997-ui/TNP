"""Object/region based QPN isolation (PROMPT-001, rules 2-5).

The QPN Excel cell must show ONLY the "Quality Problem Notice" panel of the classified QPN slide –
never the slide title ("1. HIỆN TRẠNG"), the sidebar circle, dotted separators, footer/logo or slide
whitespace.  The panel is located deterministically from the PPTX object structure and geometry:

    1. a direct picture/object carrying QPN evidence (alt text / shape name "QPN", "Quality Problem Notice");
    2. a text/table shape whose text contains the QPN heading -> that shape seeds the panel and every
       shape that belongs to the same coherent region (horizontally aligned with the seed, vertically
       contiguous, not title/sidebar/footer furniture) is added -> union of QPN MEMBER shapes only;
    3. the single dominant content picture of a QPN-classified slide (a screenshot of the notice) when
       the slide holds no QPN text at all and no other picture is comparable.

Anything else fails CLOSED: ``locate_qpn_region`` returns ``None`` with an explicit reason, the caller
omits the QPN image and adds "Cần kiểm tra" – it never falls back to the whole slide.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

from .classifier import QPN_MARKERS, QPN_WEAK_MARKERS
from .content_region import ROLE_CONTENT, classify_blocks
from .pptx_parser import Block, SlideData, norm_key

# relative-geometry tolerances (fractions of slide width / height)
H_ALIGN_TOL = 0.03          # a member may stick out of the seed's horizontal extent by this much
H_OVERLAP_MIN = 0.60        # ... or must overlap the panel's horizontal extent by >= 60 % of its own width
V_GAP_MAX = 0.04            # vertical gap between the panel and a member must be <= 4 % of slide height
MIN_PANEL_AREA = 0.12       # a credible QPN panel covers >= 12 % of the slide
MAX_PANEL_AREA = 0.95       # ... and is not the whole slide
DOMINANT_PICTURE_AREA = 0.20
DOMINANT_RATIO = 2.5        # dominant picture must be >= 2.5x the area of the next biggest picture
CROP_PAD = 0.004            # padding around the panel when cropping the rendered slide

_QPN_NAME_RE = re.compile(r"\b(?:qpn|quality problem notice)\b")


@dataclass
class QpnRegion:
    left: int                       # EMU, absolute slide coordinates
    top: int
    width: int
    height: int
    method: str                     # "picture_evidence" | "text_panel" | "table_panel" | "dominant_picture"
    members: List[str] = field(default_factory=list)      # shape names / first words (diagnostics)
    excluded: List[str] = field(default_factory=list)     # what was deliberately left out (diagnostics)
    picture: Optional[Block] = None                       # when the QPN is one picture object (direct export)

    @property
    def right(self) -> int:
        return self.left + self.width

    @property
    def bottom(self) -> int:
        return self.top + self.height

    def bbox(self) -> Tuple[int, int, int, int]:
        return self.left, self.top, self.right, self.bottom

    def contains(self, b: Block, tol: int = 0) -> bool:
        return (b.left >= self.left - tol and b.top >= self.top - tol and
                b.right <= self.right + tol and b.bottom <= self.bottom + tol)

    def describe(self, W: int, H: int) -> str:
        if not W or not H:
            return self.method
        return (f"{self.method}: x {self.left / W:.0%}-{self.right / W:.0%}, y {self.top / H:.0%}-{self.bottom / H:.0%}, "
                f"{len(self.members)} đối tượng")


@dataclass
class QpnLocate:
    region: Optional[QpnRegion]
    reason: str = ""                # why isolation failed (fail-closed) – empty on success


def _has_qpn_text(text: str) -> bool:
    k = norm_key(text or "")
    return any(m in k for m in QPN_MARKERS)


def _has_qpn_evidence(b: Block) -> bool:
    return bool(_QPN_NAME_RE.search(norm_key(f"{b.shape_name} {b.alt_text}")) or
                any(m in norm_key(b.alt_text or "") for m in QPN_WEAK_MARKERS))


def _union(a: Tuple[int, int, int, int], b: Block) -> Tuple[int, int, int, int]:
    return min(a[0], b.left), min(a[1], b.top), max(a[2], b.right), max(a[3], b.bottom)


def _belongs(bbox: Tuple[int, int, int, int], seed_x: Tuple[int, int], b: Block, W: int, H: int) -> bool:
    """Coherent-panel membership: horizontally aligned with the seed and vertically contiguous with the panel."""
    l, t, r, btm = bbox
    tol_x = W * H_ALIGN_TOL
    inside_x = b.left >= seed_x[0] - tol_x and b.right <= seed_x[1] + tol_x
    ov = max(0, min(b.right, r) - max(b.left, l))
    overlap_x = b.width > 0 and ov / b.width >= H_OVERLAP_MIN
    if not (inside_x or overlap_x):
        return False
    gap = max(b.top - btm, t - b.bottom, 0)
    return gap <= H * V_GAP_MAX


def locate_qpn_region(slide: SlideData, W: int, H: int) -> QpnLocate:
    """Deterministic QPN panel of ``slide`` (slide size ``W`` x ``H`` in EMU) or an explicit failure reason."""
    if not slide or not W or not H:
        return QpnLocate(None, "slide không có dữ liệu hình học")
    roles = classify_blocks(slide)
    furniture = {id(r.block): r.role for r in roles if r.role != ROLE_CONTENT}
    excluded = [f"{r.role}: {r.text.strip().splitlines()[0][:40]}" for r in roles if r.role != ROLE_CONTENT and r.text.strip()]
    pictures = list(slide.pictures)
    slide_area = float(W * H)

    # 1. direct picture/object with QPN evidence --------------------------------------------------------
    evidenced = [p for p in pictures if _has_qpn_evidence(p) and p.width * p.height / slide_area >= MIN_PANEL_AREA * 0.5]
    if evidenced:
        p = max(evidenced, key=lambda b: b.width * b.height)
        return QpnLocate(QpnRegion(p.left, p.top, p.width, p.height, "picture_evidence",
                                   [p.shape_name or "picture"], excluded, picture=p))

    # 2. text / table shape carrying the QPN heading seeds a coherent panel ---------------------------------
    seeds = [b for b in slide.text_blocks if _has_qpn_text(b.text) and id(b) not in furniture]
    if not seeds:
        # the notice's own heading may sit in the header band of a dedicated QPN page ("Quality Problem
        # Notice" as the page title); it is then the top of the panel – other furniture is still excluded
        seeds = [b for b in slide.text_blocks if _has_qpn_text(b.text) and furniture.get(id(b)) == "title"]
        furniture = {k: v for k, v in furniture.items() if not any(k == id(b) for b in seeds)}
    if seeds:
        seed = sorted(seeds, key=lambda b: (0 if b.kind == "table" else 1, -(b.width * b.height)))[0]
        bbox = (seed.left, seed.top, seed.right, seed.bottom)
        seed_x = (seed.left, seed.right)
        members: List[Block] = [seed]
        candidates = [b for b in slide.blocks if b is not seed and id(b) not in furniture and
                      (b.is_text or b.kind == "picture") and b.width > 0 and b.height > 0]
        changed = True
        while changed:
            changed = False
            for b in candidates:
                if b in members:
                    continue
                if b.kind == "picture":
                    cx, cy = b.left + b.width / 2, b.top + b.height / 2
                    tol = W * H_ALIGN_TOL
                    inside = bbox[0] - tol <= cx <= bbox[2] + tol and bbox[1] - tol <= cy <= bbox[3] + tol
                    if not inside and not _belongs(bbox, seed_x, b, W, H):
                        continue
                elif not _belongs(bbox, seed_x, b, W, H):
                    continue
                members.append(b)
                bbox = _union(bbox, b)
                changed = True
        l, t, r, btm = bbox
        area = (r - l) * (btm - t) / slide_area
        method = "table_panel" if seed.kind == "table" else "text_panel"
        if area < MIN_PANEL_AREA:
            return QpnLocate(None, f"vùng QPN quá nhỏ ({area:.0%} slide) – không đủ tin cậy")
        if area > MAX_PANEL_AREA:
            return QpnLocate(None, f"vùng QPN gần bằng cả slide ({area:.0%}) – không tách được khỏi slide")
        names = [(m.shape_name or m.kind) + (f" '{m.text.strip().splitlines()[0][:25]}'" if m.is_text else "")
                 for m in members]
        left_out = excluded + [f"{b.kind}: {(b.text or b.shape_name).strip().splitlines()[0][:40]}"
                               for b in candidates if b not in members]
        return QpnLocate(QpnRegion(l, t, r - l, btm - t, method, names, left_out))

    # 3. dominant content picture of a QPN slide (screenshot of the notice) -----------------------------
    from .improvement_pictures import is_decorative_picture
    content_pics = sorted((p for p in pictures if not is_decorative_picture(p, W, H)),
                          key=lambda b: b.width * b.height, reverse=True)
    if content_pics:
        p = content_pics[0]
        a0 = p.width * p.height / slide_area
        a1 = (content_pics[1].width * content_pics[1].height / slide_area) if len(content_pics) > 1 else 0.0
        if a0 >= DOMINANT_PICTURE_AREA and (a1 == 0.0 or a0 / a1 >= DOMINANT_RATIO):
            return QpnLocate(QpnRegion(p.left, p.top, p.width, p.height, "dominant_picture",
                                       [p.shape_name or "picture"], excluded, picture=p))
        return QpnLocate(None, f"không có ảnh QPN nổi trội (ảnh lớn nhất {a0:.0%} slide, ảnh kế {a1:.0%})")
    return QpnLocate(None, "không tìm thấy đối tượng 'Quality Problem Notice' hay ảnh QPN trên slide")


def crop_box_px(region: QpnRegion, W: int, H: int, img_w: int, img_h: int) -> Tuple[int, int, int, int]:
    """Pixel crop box of ``region`` on a rendering of the whole slide (``img_w`` x ``img_h``)."""
    sx, sy = img_w / float(W), img_h / float(H)
    pad_x, pad_y = int(W * CROP_PAD * sx), int(H * CROP_PAD * sy)
    l = max(0, int(region.left * sx) - pad_x)
    t = max(0, int(region.top * sy) - pad_y)
    r = min(img_w, int(round(region.right * sx)) + pad_x)
    b = min(img_h, int(round(region.bottom * sy)) + pad_y)
    return l, t, max(l + 1, r), max(t + 1, b)
