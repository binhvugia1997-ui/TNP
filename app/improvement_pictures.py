"""After-only selection of improvement pictures (deterministic geometry, no LLM).

"Hình ảnh cải tiến" receives ONLY the pictures that belong to "Sau cải tiến" of a
production/process improvement item:

  * the improvement block is located first (slide classified as improvement);
  * caption buttons ("Trước cải tiến" / "Sau cải tiến" / "Before" / "After") and the inline
    business lines "+ Trước:" / "+ Sau:" are the anchors;
  * every picture is assigned to the nearest anchor using geometry relative to the slide;
  * when the deck has no explicit Before/After wording, two further deterministic signals are used
    (PROMPT-004C): BLUE native text of the same production block that is spatially associated with the
    picture (After convention) and a directional arrow whose DESTINATION side holds the picture
    ("[Before] → [After]"); both come from PPTX structure only – no OCR, no pixel analysis, no LLM;
  * Before pictures, arrows, logos, decorative objects and pictures of inspection/control
    improvements are excluded;
  * a picture that cannot be attributed safely is NOT copied and the slide is flagged:
    "Cần kiểm tra: Không xác định chắc chắn ảnh Sau cải tiến tại slide X".
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from .classifier import improvement_subkind, is_heading_like, section_kind_of_heading
from .content_region import (ROLE_CAPTION, ROLE_CONTENT, ROLE_FURNITURE, ROLE_SIDEBAR, ROLE_TITLE, BlockRole,
                             classify_blocks, inline_anchor_kind, title_text)
from .pptx_parser import Block, ReportData, SlideData, norm_key
from .report_identity import report_scope_key

LOG = logging.getLogger("report_extractor.improvement_pictures")

AMBIGUOUS_REASON = "Cần kiểm tra: Không xác định chắc chắn ảnh Sau cải tiến tại slide {n}"

# relative thresholds (fractions of slide width / height)
MIN_PICTURE_AREA = 0.010        # smaller pictures are logos / arrows / icons
MIN_PICTURE_SIDE = 0.05
TITLE_BAND = 0.14
CAPTION_GAP_MAX = 0.12          # caption directly above/below the picture
CAPTION_SIDE_GAP_MAX = 0.06     # caption beside the picture
COLUMN_HEADER_GAP_MAX = 0.55    # caption used as a column header further above
TIE_TOLERANCE = 0.008
ANCHOR_TOLERANCE = 0.02

_INSPECTION_RE = re.compile(r"kiem tra|kiem soat|inspection|control|quan ly|tieu chuan kiem")
_TEMPORARY_RE = re.compile(r"xu ly tam thoi|temporary action|temporary countermeasure|containment|tam thoi")
_VERIFICATION_RE = re.compile(r"kiem chung|xac nhan hieu qua|hieu qua (?:cai tien|doi sach|sau)|verification|effectiveness|"
                              r"theo doi|monitoring|audit|sustain|duy tri|giam sat|follow[- ]?up|"
                              r"kiem tra thuong xuyen")
_CONTROL_ONLY_RE = re.compile(r"kiem soat|doi sach kiem soat|phuong phap kiem tra|tieu chuan kiem tra|"
                              r"kiem tra|inspection|control|verification|kiem chung")

SEMANTIC_PRODUCTION = "PRODUCTION_IMPROVEMENT"
SEMANTIC_INSPECTION = "INSPECTION_CONTROL"
SEMANTIC_TEMPORARY = "TEMPORARY_ACTION"
SEMANTIC_VERIFICATION = "VERIFICATION"
SEMANTIC_OTHER = "OTHER"

# PROMPT-004C evidence thresholds (fractions of slide size)
ARROW_CORRIDOR = 0.12           # picture must overlap the arrow's axis band (± this much of the slide) to be on a side
ARROW_REACH = 0.45              # farthest a picture may lie from the arrow along its axis
ARROW_MAX_SIDE = 0.35           # bigger shapes are not separators but content (e.g. a big chevron banner)
ARROW_GROUP_GAP = 0.12          # a gap wider than this (fraction of slide) ends the picture group on one side
BLUE_HUE = (190.0, 260.0)       # HSV hue range accepted as "blue"
BLUE_MIN_SAT = 0.35
BLUE_MIN_VAL = 0.30


@dataclass
class Anchor:
    kind: str                   # "before" | "after"
    left: int
    top: int
    width: int
    height: int
    source: str                 # "caption" | "inline"

    @property
    def right(self) -> int:
        return self.left + self.width

    @property
    def bottom(self) -> int:
        return self.top + self.height

    @property
    def cy(self) -> float:
        return self.top + self.height / 2


@dataclass
class ItemRegion:
    """A structural production item or semantic section on a slide."""
    slide: int
    owner_id: str
    heading: str
    bounds: Tuple[int, int, int, int]
    source_order: int
    semantic_role: str
    confident_defect_ownership: bool
    region_kind: str = "item"   # item | section

    @property
    def left(self) -> int:
        return self.bounds[0]

    @property
    def top(self) -> int:
        return self.bounds[1]

    @property
    def width(self) -> int:
        return self.bounds[2]

    @property
    def height(self) -> int:
        return self.bounds[3]

    @property
    def right(self) -> int:
        return self.left + self.width

    @property
    def bottom(self) -> int:
        return self.top + self.height


@dataclass
class OwnerResult:
    region: Optional[ItemRegion]
    confident: bool
    reason: str
    semantic_role: str = SEMANTIC_OTHER

    @property
    def owner_id(self) -> str:
        return self.region.owner_id if self.region else ""

    @property
    def heading(self) -> str:
        return self.region.heading if self.region else ""


@dataclass
class PictureRef:
    slide: int
    block: Block
    kind: str                   # "after" | "before" | "ambiguous" | "excluded"
    reason: str = ""
    anchor: str = ""            # how it was decided
    owner_id: str = ""
    owner_heading: str = ""
    source_order: int = 0
    temporal_role: str = "UNKNOWN"       # BEFORE | AFTER | UNKNOWN
    semantic_role: str = SEMANTIC_OTHER
    confident_owner: bool = False
    excel_output_eligible: bool = False
    exclusion_reason: str = ""
    report_scope_id: str = ""       # transient source-report scope; never used to generalize learning
    confidence: Optional[float] = None  # deterministic/model candidate confidence; absent when not measured

    @property
    def label(self) -> str:
        return f"S{self.slide}#{self.block.shape_id}"


@dataclass
class PictureSelection:
    after: List[PictureRef] = field(default_factory=list)
    rejected: List[PictureRef] = field(default_factory=list)
    reasons: List[str] = field(default_factory=list)       # -> Cần kiểm tra
    notes: List[str] = field(default_factory=list)         # diagnostics only
    slides_with_after: List[int] = field(default_factory=list)


def _overlap(a1: int, a2: int, b1: int, b2: int) -> int:
    return max(0, min(a2, b2) - max(a1, b1))


def _gap(a1: int, a2: int, b1: int, b2: int) -> int:
    """Distance between two 1-D intervals (0 when overlapping)."""
    return max(0, max(a1, b1) - min(a2, b2))


def semantic_role_for_heading(text: str) -> str:
    """Structural meaning of a section/item heading, independent of Before/After wording."""
    key = norm_key(text)
    if not key:
        return SEMANTIC_OTHER
    if _TEMPORARY_RE.search(key):
        return SEMANTIC_TEMPORARY
    kind = section_kind_of_heading(text)
    if kind in ("verify", "followup") or _VERIFICATION_RE.search(key):
        return SEMANTIC_VERIFICATION
    if kind == "improvement":
        subkind = improvement_subkind(text)
        if subkind == "inspection":
            return SEMANTIC_INSPECTION
        if subkind == "followup":
            return SEMANTIC_VERIFICATION
        return SEMANTIC_PRODUCTION
    if _CONTROL_ONLY_RE.search(key):
        return SEMANTIC_INSPECTION
    if kind == "standard":
        return SEMANTIC_PRODUCTION
    return SEMANTIC_OTHER


def _generic_production_heading(text: str) -> bool:
    key = re.sub(r"^(?:\d+\s+)+", "", norm_key(text)).strip(" .:;-_")
    return key in {
        "cai tien", "cai tien san xuat", "cai tien trong san xuat", "production improvement",
        "improvement", "process improvement", "doi sach", "doi sach cai tien", "countermeasure",
        "corrective action", "san xuat", "production", "process",
    }


# PROMPT-025: a logical improvement-item anchor may also open with a defect name
# ("Lỗi …", "Mẻ …", "Xước …", "Hiện trạng …") – the real report family names items after defects.
_DEFECT_OPENER_RE = re.compile(
    r"^(?:loi|hien trang|hien tuong|me|xuoc|bong|lech|cong venh|venh|bien dang|ro|nut|gay|dinh|tray|dom|"
    r"khuyet tat|phong rop|diem loi)\b")


def _is_item_anchor_text(text: str) -> bool:
    """PROMPT-025 item-anchor wording: a colon-terminated production heading ("Cải tiến lỗi X (...):")
    or a line opening with a defect name ("Lỗi lệch ATN sau ép nhỰA").  Plain body sentences never
    match.  The colon test uses the raw text (``norm_key`` strips trailing punctuation)."""
    key = norm_key(text)
    if not key:
        return False
    if text.rstrip().endswith(":") and semantic_role_for_heading(text) == SEMANTIC_PRODUCTION:
        return True
    return bool(_DEFECT_OPENER_RE.match(key))


def _visual_row_below(line_top: int, line_bottom: int, visual_tops: Sequence[int], H: int) -> bool:
    """PROMPT-025 co-signal: an item anchor is authored together with its visual evidence – a caption
    button or content picture must start within half a slide height below the anchor line."""
    window_top = line_top - int(0.02 * H)
    window_bottom = line_bottom + int(0.5 * H)
    return any(window_top <= top <= window_bottom for top in visual_tops)


def _line_box(block: Block, index: int, n_lines: int) -> Tuple[int, int, int, int]:
    line_h = max(1, int(block.height / max(1, n_lines)))
    return block.left, int(block.top + index * line_h), block.width, line_h


def slide_items(slide: SlideData) -> List[ItemRegion]:
    """Find item/section headings and estimate their line-level regions in absolute slide geometry.

    Section headings establish local semantics. A specific production heading such as "Cải tiến lỗi A"
    or a short item heading below a production section becomes a confident logical owner. Generic slide
    headings alone never count as a defect owner.

    PROMPT-025: a slide may carry several logical improvement items. Besides the recognized section
    headings, an item anchor is a short, non-bullet line under a production context that either is
    colon-terminated with production semantics ("Cải tiến lỗi X (...):") or opens with a defect name
    ("Lỗi lệch ATN sau ép nhựa"). Real decks repeat the item-heading style inside body frames (mid-block,
    not bold), so the first-line/bold requirement of the defect-heading branch is intentionally relaxed;
    an authored visual row (caption button or content picture) must follow within half the slide height.
    """
    roles = sorted(classify_blocks(slide), key=lambda r: (r.block.order, r.block.top, r.block.left))
    items: List[ItemRegion] = []
    current_semantic = SEMANTIC_OTHER
    duplicates: Dict[str, int] = {}
    W = slide.width or 1
    H = slide.height or 1
    visual_tops: List[int] = [role.block.top for role in roles
                              if role.role == ROLE_CAPTION and role.block.text.strip()]
    for picture in slide.pictures:
        if not is_decorative_picture(picture, W, H):
            visual_tops.append(picture.top)
    visual_tops.sort()
    for role in roles:
        if role.role not in (ROLE_TITLE, ROLE_CONTENT, ROLE_SIDEBAR):
            continue
        block = role.block
        lines = [line for line in block.text.splitlines()]
        n_lines = max(1, len(lines))
        for i, raw in enumerate(lines):
            text = raw.strip()
            if not text:
                continue
            key = norm_key(text)
            if re.match(r"^[\-+•·*▪➢►→>]+", text):
                continue
            heading_like = is_heading_like(text, block.bold if i == 0 else False,
                                           block.size_pt if i == 0 else None)
            kind = section_kind_of_heading(text) if heading_like else None
            semantic = semantic_role_for_heading(text) if (heading_like or i == 0) else SEMANTIC_OTHER

            # Recognized section headings (including inspection/control, temporary, verification) update
            # the context before subsequent item/body headings are inspected.
            if semantic != SEMANTIC_OTHER:
                if role.role == ROLE_SIDEBAR and semantic == SEMANTIC_PRODUCTION:
                    current_semantic = SEMANTIC_PRODUCTION
                    continue                       # a generic production sidebar is context, not an image owner
                logical = semantic == SEMANTIC_PRODUCTION and not _generic_production_heading(text)
                region_kind = "item" if logical else "section"
                if semantic == SEMANTIC_PRODUCTION:
                    current_semantic = SEMANTIC_PRODUCTION
                elif semantic in (SEMANTIC_INSPECTION, SEMANTIC_TEMPORARY, SEMANTIC_VERIFICATION):
                    current_semantic = semantic
                line_key = norm_key(text).strip(" .:;-_")
                occurrence = duplicates.get(line_key, 0) + 1
                duplicates[line_key] = occurrence
                owner = f"item:{line_key}" if logical else f"section:S{slide.number}:{block.order}:{i}"
                if logical and occurrence > 1:
                    owner += f":{occurrence}"
                bounds = ((block.left, block.top, block.width, block.height)
                          if logical and i == 0 else _line_box(block, i, n_lines))
                items.append(ItemRegion(slide.number, owner, text, bounds,
                                        block.order * 100 + i, semantic, logical, region_kind))
                continue

            # A defect heading may omit the words "cải tiến" while remaining under an explicit
            # production section. Use block starts/defect headings only; never turn body bullets into owners.
            if (current_semantic == SEMANTIC_PRODUCTION and heading_like and
                    (i == 0 or kind == "defect") and not re.search(
                        r"\b(hien trang|truoc|sau|before|after|nguyen nhan|temporary|tam thoi)\b", key)):
                if kind in ("cause", "temporary", "verify", "standard", "improvement", "qpn"):
                    continue
                line_key = key.strip(" .:;-_")
                if not line_key or _generic_production_heading(text):
                    continue
                occurrence = duplicates.get(line_key, 0) + 1
                duplicates[line_key] = occurrence
                owner = f"item:{line_key}" + (f":{occurrence}" if occurrence > 1 else "")
                bounds = (block.left, block.top, block.width, block.height) if i == 0 else _line_box(block, i, n_lines)
                items.append(ItemRegion(slide.number, owner, text, bounds,
                                        block.order * 100 + i, SEMANTIC_PRODUCTION, True, "item"))
                continue

            # PROMPT-025 multi-item anchor: colon-terminated production item heading or defect-name
            # opener under a production context, with authored visual evidence below (see above).
            if (current_semantic == SEMANTIC_PRODUCTION and len(text) <= 120
                    and inline_anchor_kind(text) is None
                    and not _generic_production_heading(text)
                    and _is_item_anchor_text(text)):
                line_box = _line_box(block, i, n_lines)
                if _visual_row_below(line_box[1], line_box[1] + line_box[3], visual_tops, H):
                    line_key = key.strip(" .:;-_")
                    occurrence = duplicates.get(line_key, 0) + 1
                    duplicates[line_key] = occurrence
                    owner = f"item:{line_key}" + (f":{occurrence}" if occurrence > 1 else "")
                    bounds = ((block.left, block.top, block.width, block.height)
                              if i == 0 else line_box)
                    items.append(ItemRegion(slide.number, owner, text, bounds,
                                            block.order * 100 + i, SEMANTIC_PRODUCTION, True, "item"))
                    continue
    items.sort(key=lambda item: (item.source_order, item.bounds[1], item.bounds[0]))
    # An item's content often continues well below its title (inline +Sau rows, grouped paragraphs, or
    # picture tables). Extend only confident item regions to the next logical item or non-production
    # section; a broad production slide title/sidebar is context, not an ownership boundary.
    slide_bottom = int(slide.height or max((item.bottom for item in items), default=1))
    footer_tops = [role.block.top for role in roles
                   if role.role == ROLE_FURNITURE and role.block.top / max(1, slide_bottom) >= 0.75]
    if footer_tops:
        slide_bottom = min(slide_bottom, min(footer_tops))
    boundaries = [item for item in items
                  if item.region_kind == "item" or item.semantic_role != SEMANTIC_PRODUCTION]
    for item in items:
        if not item.confident_defect_ownership:
            continue
        x, y, width, _height = item.bounds
        # PROMPT-025: a boundary clips this item only when it is horizontally relevant — it overlaps the
        # item's own band or spans nearly the full slide width. A narrow sidebar label on the opposite
        # side (e.g. an inspection oval in the left margin) must not cut the item's authored visual row
        # (captions + Before/After pictures below the text) out of its own span.
        following = []
        for other in boundaries:
            if other.source_order <= item.source_order or other.top <= item.top:
                continue
            o_left, _o_top, o_width, _o_height = other.bounds
            full_width = o_width >= 0.55 * W
            overlaps = not (o_left + o_width <= x or o_left >= x + width)
            if full_width or overlaps:
                following.append(other.top)
        if following:
            # A later heading inside the same text box divides the original box; clip at that line
            # even when the first-line block geometry spans the rest of the paragraph group.
            bottom = max(y + 1, min(slide_bottom, min(following)))
        else:
            bottom = max(y + 1, slide_bottom)
        item.bounds = (x, y, width, bottom - y)
    return items


def _region_distance(point_block: Block, region: ItemRegion, W: int, H: int) -> float:
    cx, cy = point_block.left + point_block.width / 2, point_block.top + point_block.height / 2
    dx = _gap(int(cx), int(cx), region.left, region.right) / max(1, W)
    dy = _gap(int(cy), int(cy), region.top, region.bottom) / max(1, H)
    return dx + 1.25 * dy


def owner_of(picture: Block, items: Sequence[ItemRegion], W: Optional[int] = None,
             H: Optional[int] = None) -> OwnerResult:
    """Resolve a picture's nearest structural section/item owner; fail closed on ties or weak evidence."""
    if not items:
        return OwnerResult(None, False, "no item/section heading", SEMANTIC_OTHER)
    W = int(W or max([picture.right, *(i.right for i in items), 1]))
    H = int(H or max([picture.bottom, *(i.bottom for i in items), 1]))
    ranked = sorted(((_region_distance(picture, region, W, H), region) for region in items),
                    key=lambda pair: (pair[0], pair[1].source_order))
    score, region = ranked[0]
    if score > 0.42:
        return OwnerResult(None, False, f"nearest heading is too far ({score:.3f})", SEMANTIC_OTHER)
    if region.semantic_role == SEMANTIC_PRODUCTION and not region.confident_defect_ownership:
        # A generic production slide heading is context, not an item owner. Prefer a nearby specific child
        # heading when it is spatially comparable to that broad section title.
        specific = [(d, item) for d, item in ranked[1:]
                    if item.semantic_role == SEMANTIC_PRODUCTION and item.confident_defect_ownership
                    and d - score <= 0.12]
        if specific:
            score, region = specific[0]
    competing = [(d, item) for d, item in ranked[1:]
                 if item.owner_id != region.owner_id and d - score <= 0.025
                 and (item.confident_defect_ownership or item.semantic_role != region.semantic_role)]
    if competing:
        other = competing[0][1]
        return OwnerResult(None, False, f"ambiguous ownership: {region.heading!r} vs {other.heading!r}", SEMANTIC_OTHER)
    if region.semantic_role != SEMANTIC_PRODUCTION:
        return OwnerResult(region, False, f"section semantic role is {region.semantic_role}", region.semantic_role)
    if not region.confident_defect_ownership:
        return OwnerResult(region, False, "only a generic production section is known", region.semantic_role)
    # A second production item at nearly the same distance is not a confident owner.
    other_items = [(d, item) for d, item in ranked[1:]
                   if item.confident_defect_ownership and item.owner_id != region.owner_id]
    if other_items and other_items[0][0] - score <= 0.04:
        return OwnerResult(None, False, f"ambiguous logical item owner near {region.heading!r}", SEMANTIC_OTHER)
    return OwnerResult(region, True, "unique nearby production defect/improvement heading", region.semantic_role)


def group_refs(refs: Sequence["PictureRef"]) -> Dict[Tuple[str, str], List["PictureRef"]]:
    """Eligible pictures grouped by (source report, logical defect owner) in deterministic source order.

    ``owner_id`` intentionally contains only the logical item identity within a report. Keeping the report scope in
    the grouping key means identical headings/shape IDs in separately parsed reports cannot merge if a caller passes
    a multi-report collection. Empty scopes remain compatible for manually constructed, report-local refs.
    """
    groups: Dict[Tuple[str, str], List[PictureRef]] = {}
    for ref in sorted(refs, key=lambda x: (x.source_order, x.slide, x.block.order, x.report_scope_id)):
        if (not ref.excel_output_eligible or ref.temporal_role != "AFTER"
                or ref.semantic_role != SEMANTIC_PRODUCTION or not ref.confident_owner or not ref.owner_id):
            continue
        groups.setdefault((ref.report_scope_id, ref.owner_id), []).append(ref)
    return groups


def is_decorative_picture(p: Block, W: int, H: int) -> Optional[str]:
    if W <= 0 or H <= 0:
        return None
    area = (p.width * p.height) / float(W * H)
    if area < MIN_PICTURE_AREA or p.width / W < MIN_PICTURE_SIDE or p.height / H < MIN_PICTURE_SIDE:
        return "small picture (logo / arrow / icon)"
    if (p.top + p.height) / H <= TITLE_BAND and p.width / W < 0.3:
        return "picture in header band (logo)"
    if p.width / W < 0.10 and p.height / H < 0.10:
        return "small picture"
    ratio = (p.width / p.height) if p.height else 0
    if ratio > 6 or (ratio and ratio < 1 / 6):
        return "thin picture (arrow / divider)"
    n = norm_key(p.shape_name or "")
    if re.search(r"\b(arrow|logo|icon)\b", n) or re.search(r"\b(arrow|logo|icon)\b", norm_key(p.alt_text or "")):
        return f"decorative by name ({p.shape_name})"
    return None


def is_blue(color: str) -> bool:
    """'#rrggbb' → True when the colour is unmistakably blue (HSV; theme colours already resolved by the parser)."""
    c = (color or "").lstrip("#")
    if len(c) != 6:
        return False
    try:
        r, g, b = (int(c[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    except ValueError:
        return False
    mx, mn = max(r, g, b), min(r, g, b)
    v, d = mx, mx - mn
    if mx <= 0 or d <= 0:
        return False
    sat = d / mx
    if mx == r:
        h = (60 * ((g - b) / d)) % 360
    elif mx == g:
        h = 60 * ((b - r) / d) + 120
    else:
        h = 60 * ((r - g) / d) + 240
    return BLUE_HUE[0] <= h <= BLUE_HUE[1] and sat >= BLUE_MIN_SAT and v >= BLUE_MIN_VAL


def _blue_anchors(roles: Sequence[BlockRole], inspection_ranges: Sequence[Tuple[int, int]]) -> List[Anchor]:
    """After anchors from BLUE native text lines of CONTENT blocks (production convention: After descriptions are
    written in blue).  Titles, sidebars, captions, logos and footer text never qualify (role filter) and lines inside
    an inspection/control item are ignored.  Consecutive blue lines form one anchor rectangle."""
    out: List[Anchor] = []
    for r in roles:
        if r.role != ROLE_CONTENT:
            continue
        b = r.block
        lines = b.text.split("\n")
        colors = list(b.line_colors or [])
        if not colors or not any(is_blue(c) for c in colors):
            continue
        n = max(1, len(lines))
        line_h = b.height / n if b.height else 0
        start = None
        for i in range(n + 1):
            blue = i < n and i < len(colors) and is_blue(colors[i]) and lines[i].strip() != ""
            if blue and start is None:
                start = i
            elif not blue and start is not None:
                y0, y1 = int(b.top + start * line_h), int(b.top + i * line_h)
                cy = (y0 + y1) / 2
                if not any(a0 <= cy < a1 for a0, a1 in inspection_ranges):
                    out.append(Anchor("after", b.left, y0, b.width, max(1, y1 - y0), "blue"))
                start = None
    return out


def _arrow_usable(a: Block, W: int, H: int, bands, inspection_ranges, roles: Sequence[BlockRole]) -> bool:
    """Only arrows that live inside the business area of a production block count as Before→After separators."""
    if not a.direction or W <= 0 or H <= 0:
        return False
    if a.width / W > ARROW_MAX_SIDE and a.height / H > ARROW_MAX_SIDE:
        return False
    cy = a.top + a.height / 2
    cx = a.left + a.width / 2
    if cy / H <= 0.12 or cy / H >= 0.92:                       # title band / footer decoration
        return False
    if cx / W <= 0.10 and a.width / W < 0.12:                  # left sidebar ornaments
        return False
    if any(y0 <= cy < y1 for y0, y1, _k in bands):
        return False
    if any(y0 <= cy < y1 for y0, y1 in inspection_ranges):
        return False
    return True


def _arrow_side(p: Block, a: Block, W: int, H: int) -> Optional[str]:
    """'source' | 'destination' | None – where picture ``p`` sits relative to directional arrow ``a``."""
    horizontal = a.direction in ("left", "right")
    if horizontal:
        axis = a.top + a.height / 2
        if _gap(p.top, p.bottom, int(axis - ARROW_CORRIDOR * H), int(axis + ARROW_CORRIDOR * H)) > 0:
            return None
        if p.right <= a.left + a.width * 0.25:
            side, dist = "left", (a.left - p.right) / W
        elif p.left >= a.right - a.width * 0.25:
            side, dist = "right", (p.left - a.right) / W
        else:
            return None
        if dist > ARROW_REACH:
            return None
        return "destination" if side == a.direction else "source"
    axis = a.left + a.width / 2
    if _gap(p.left, p.right, int(axis - ARROW_CORRIDOR * W), int(axis + ARROW_CORRIDOR * W)) > 0:
        return None
    if p.bottom <= a.top + a.height * 0.25:
        side, dist = "up", (a.top - p.bottom) / H
    elif p.top >= a.bottom - a.height * 0.25:
        side, dist = "down", (p.top - a.bottom) / H
    else:
        return None
    if dist > ARROW_REACH:
        return None
    return "destination" if side == a.direction else "source"


def _contiguous_groups(pictures: Sequence[Block], a: Block, sides: Dict[int, str], W: int, H: int) -> Dict[int, str]:
    """Keep, on each side of the arrow, only the pictures that form one contiguous group starting at the arrow
    (a large gap means another block / unrelated picture further along the row or column)."""
    horizontal = a.direction in ("left", "right")
    out: Dict[int, str] = {}
    for side in ("source", "destination"):
        group = [p for p in pictures if sides.get(id(p)) == side]
        if not group:
            continue
        if horizontal:
            towards_right = (side == "destination") == (a.direction == "right")
            group.sort(key=lambda p: p.left if towards_right else -p.right)
            edge = a.right if towards_right else a.left
            for p in group:
                gap = (p.left - edge) if towards_right else (edge - p.right)
                if gap / W > ARROW_GROUP_GAP:
                    break
                out[id(p)] = side
                edge = p.right if towards_right else p.left
        else:
            downwards = (side == "destination") == (a.direction == "down")
            group.sort(key=lambda p: p.top if downwards else -p.bottom)
            edge = a.bottom if downwards else a.top
            for p in group:
                gap = (p.top - edge) if downwards else (edge - p.bottom)
                if gap / H > ARROW_GROUP_GAP:
                    break
                out[id(p)] = side
                edge = p.bottom if downwards else p.top
    return out


def _arrow_votes(pictures: Sequence[Block], arrows: Sequence[Block], W: int, H: int) -> Dict[int, Tuple[str, str]]:
    """{id(picture): ('after'|'before', how)} from arrows that really separate two picture groups
    (at least one picture on the source side AND one on the destination side)."""
    votes: Dict[int, Tuple[str, str]] = {}
    conflicts: set = set()
    for a in arrows:
        sides = {p_id: side for p_id, side in ((id(p), _arrow_side(p, a, W, H)) for p in pictures) if side}
        sides = _contiguous_groups(pictures, a, sides, W, H)
        if "source" not in sides.values() or "destination" not in sides.values():
            continue
        for p in pictures:
            side = sides.get(id(p))
            if not side:
                continue
            kind = "after" if side == "destination" else "before"
            prev = votes.get(id(p))
            if prev and prev[0] != kind:
                conflicts.add(id(p))
            votes[id(p)] = (kind, f"arrow '{a.direction}' {side} side")
    for pid in conflicts:
        votes.pop(pid, None)
    return votes


def _inline_anchors(roles: Sequence[BlockRole]) -> Tuple[List[Anchor], List[Tuple[int, int]]]:
    """Anchors from '+ Trước:' / '+ Sau:' lines and vertical ranges of inspection/control items."""
    anchors: List[Anchor] = []
    inspection_ranges: List[Tuple[int, int]] = []
    for r in roles:
        if r.role != ROLE_CONTENT:
            continue
        b = r.block
        lines = b.text.split("\n")
        n = max(1, len(lines))
        line_h = b.height / n if b.height else 0
        item_heads: List[Tuple[int, bool]] = []      # (y, is_inspection)
        for i, ln in enumerate(lines):
            if not ln.strip():
                continue
            y = int(b.top + i * line_h)
            k = inline_anchor_kind(ln)
            if k:
                anchors.append(Anchor(k, b.left, y, b.width, max(1, int(line_h)), "inline"))
                continue
            stripped = ln.strip()
            is_bullet = bool(re.match(r"^[\-\+\u2022\u00b7\*\u25aa\u27a2\u25ba>]", stripped))
            if not is_bullet:
                item_heads.append((y, bool(_INSPECTION_RE.search(norm_key(stripped)))))
        for j, (y, insp) in enumerate(item_heads):
            if insp:
                y_end = item_heads[j + 1][0] if j + 1 < len(item_heads) else b.bottom
                inspection_ranges.append((y, y_end))
    return anchors, inspection_ranges


def _caption_anchors(roles: Sequence[BlockRole]) -> List[Anchor]:
    out = []
    for r in roles:
        if r.role == ROLE_CAPTION and r.caption_kind:
            b = r.block
            out.append(Anchor(r.caption_kind, b.left, b.top, b.width, b.height, "caption"))
    return out


def _caption_distance(p: Block, a: Anchor, W: int, H: int) -> Optional[float]:
    """Relative distance between picture and caption when they are geometrically related."""
    h_ov = _overlap(p.left, p.right, a.left, a.right)
    v_ov = _overlap(p.top, p.bottom, a.top, a.bottom)
    if h_ov and v_ov:
        return 0.0
    min_w = max(1, min(p.width, a.width))
    min_h = max(1, min(p.height, a.height))
    v_gap = _gap(p.top, p.bottom, a.top, a.bottom) / H
    h_gap = _gap(p.left, p.right, a.left, a.right) / W
    if h_ov >= 0.5 * min_w:                              # caption above / below
        if v_gap <= CAPTION_GAP_MAX:
            return v_gap
        if a.bottom <= p.top and v_gap <= COLUMN_HEADER_GAP_MAX:     # column header
            return v_gap + 0.10
    if v_ov >= 0.5 * min_h and h_gap <= CAPTION_SIDE_GAP_MAX:        # caption beside
        return h_gap + 0.02
    return None


def _claims(pictures: Sequence[Block], captions: Sequence[Anchor], W: int, H: int) -> Dict[int, List[Anchor]]:
    """Every caption claims the picture it is closest to (a caption labels exactly one picture)."""
    out: Dict[int, List[Anchor]] = {}
    for a in captions:
        best = None
        for p in pictures:
            d = _caption_distance(p, a, W, H)
            if d is None:
                continue
            below = 0 if a.bottom <= p.top + p.height * 0.25 else 1      # a caption labels the picture under it
            key = (round(d, 3), below)
            if best is None or key < best[0]:
                best = (key, p)
        if best is not None:
            out.setdefault(id(best[1]), []).append(a)
    return out


def classify_picture(p: Block, captions: Sequence[Anchor], inlines: Sequence[Anchor], W: int, H: int,
                     claims: Optional[Dict[int, List[Anchor]]] = None, blue: Optional[Sequence[Anchor]] = None,
                     arrow_votes: Optional[Dict[int, Tuple[str, str]]] = None) -> Tuple[str, str]:
    """('after'|'before'|'ambiguous', how).  Evidence priority: caption button → inline '+ Trước/+ Sau' line →
    blue After text (+ arrow agreement) → arrow destination side → ambiguous (never 'the only picture')."""
    # 1. caption buttons (strongest evidence)
    mine = (claims or {}).get(id(p), [])
    kinds = {a.kind for a in mine}
    if len(kinds) == 1:
        return mine[0].kind, f"caption '{mine[0].kind}'"
    if len(kinds) > 1:
        return "ambiguous", "captions of both kinds label this picture"
    scored = []
    for a in captions:
        d = _caption_distance(p, a, W, H)
        if d is not None:
            above = 0 if a.bottom <= p.top + p.height * 0.25 else 1        # captions usually sit above
            scored.append((d, above, a))
    if scored:
        scored.sort(key=lambda t: (t[0], t[1]))
        best_d, _, best = scored[0]
        others = [a for d, ab, a in scored[1:] if a.kind != best.kind and d - best_d <= TIE_TOLERANCE]
        if others:
            return "ambiguous", "two captions at the same distance"
        return best.kind, f"caption '{best.kind}'"
    # 2. inline "+ Trước:" / "+ Sau:" lines own the vertical range down to the next anchor
    inline_note = ""
    if inlines:
        ordered = sorted(inlines, key=lambda a: a.top)
        cy = p.top + p.height / 2
        owner = None
        for i, a in enumerate(ordered):
            nxt = ordered[i + 1].top if i + 1 < len(ordered) else None
            if a.top - ANCHOR_TOLERANCE * H <= cy and (nxt is None or cy < nxt):
                owner = a
        if owner is not None:
            return owner.kind, f"inline '{owner.kind}' line"
        inline_note = "picture above the first Trước/Sau line"
    # 3. BLUE After text of the same block, spatially associated (PROMPT-004C evidence 2)
    blue_hit = None
    for a in blue or ():
        d = _caption_distance(p, a, W, H)
        if d is not None and (blue_hit is None or d < blue_hit):
            blue_hit = d
    # 4. directional arrow: picture on the DESTINATION side (PROMPT-004C evidence 3)
    vote = (arrow_votes or {}).get(id(p))
    if blue_hit is not None and vote and vote[0] == "after":
        return "after", f"blue After text + {vote[1]}"
    if blue_hit is not None and vote and vote[0] == "before":
        return "ambiguous", "blue text beside the picture but the arrow puts it on the source side"
    if blue_hit is not None:
        return "after", "blue After text associated with the picture"
    if vote:
        return vote[0], vote[1]
    if inline_note:
        return "ambiguous", inline_note
    return "ambiguous", "no Trước/Sau anchor, blue After text or Before→After arrow for this picture"


def _caption_owner_for_picture(picture: Block, after_pictures: Sequence[Block],
                               claims: Dict[int, List[Anchor]], items: Sequence[ItemRegion],
                               W: int, H: int) -> Optional[OwnerResult]:
    """Propagate an explicit After-caption owner across its adjacent same-row picture group.

    The caption detector gives each caption one closest picture. Remaining pictures in that authored row still belong
    to the same visual item, even when a later item heading is geometrically closer. If a row component is anchored by
    multiple item captions, ownership stays ambiguous rather than merging defects.
    """
    direct_anchors = [anchor for anchor in claims.get(id(picture), []) if anchor.kind == "after"]
    anchored_owners: List[OwnerResult] = [owner_of(anchor, items, W, H) for anchor in direct_anchors]
    if not anchored_owners:
        component = {id(picture)}
        changed = True
        while changed:
            changed = False
            members = [candidate for candidate in after_pictures if id(candidate) in component]
            for candidate in after_pictures:
                if id(candidate) in component:
                    continue
                if any(_same_row_adjacent(candidate, member, W, H) for member in members):
                    component.add(id(candidate))
                    changed = True
        for anchor_picture in after_pictures:
            if id(anchor_picture) not in component:
                continue
            anchors = [anchor for anchor in claims.get(id(anchor_picture), []) if anchor.kind == "after"]
            anchored_owners.extend(owner_of(anchor, items, W, H) for anchor in anchors)
    if not anchored_owners:
        return None
    owner_keys = {owner.owner_id or f"unresolved:{owner.reason}" for owner in anchored_owners}
    if len(owner_keys) > 1:
        return OwnerResult(None, False, "ambiguous ownership between nearby After-caption groups", SEMANTIC_OTHER)
    return anchored_owners[0]


def _attach_owner(ref: PictureRef, slide: SlideData, items: Sequence[ItemRegion],
                  slide_is_inspection: bool, W: int, H: int, after_pictures: Sequence[Block] = (),
                  claims: Optional[Dict[int, List[Anchor]]] = None) -> None:
    result = owner_of(ref.block, items, W, H)
    if ref.temporal_role == "AFTER":
        caption_owner = _caption_owner_for_picture(ref.block, after_pictures, claims or {}, items, W, H)
        if caption_owner is not None:
            caption_is_ambiguous = not caption_owner.confident and "ambiguous" in caption_owner.reason
            if not (caption_is_ambiguous and result.confident and result.region is not None):
                result = caption_owner
    ref.owner_id = result.owner_id
    ref.owner_heading = result.heading
    ref.confident_owner = result.confident and not slide_is_inspection
    ref.semantic_role = (SEMANTIC_INSPECTION if slide_is_inspection else
                         result.semantic_role if result.region is not None else SEMANTIC_OTHER)
    ref.source_order = int(ref.slide) * 1_000_000 + int(ref.block.order)
    if ref.temporal_role == "UNKNOWN":
        ref.temporal_role = {"after": "AFTER", "before": "BEFORE"}.get(ref.kind, "UNKNOWN")
    if slide_is_inspection:
        ref.exclusion_reason = "slide title identifies inspection/control content"
    elif result.region is None:
        ref.exclusion_reason = result.reason
    elif not result.confident:
        ref.exclusion_reason = result.reason


def _finalize_picture(ref: PictureRef, slide: SlideData, items: Sequence[ItemRegion],
                      slide_is_inspection: bool, W: int, H: int, sel: PictureSelection,
                      after_pictures: Sequence[Block] = (), claims: Optional[Dict[int, List[Anchor]]] = None) -> None:
    _attach_owner(ref, slide, items, slide_is_inspection, W, H, after_pictures, claims)
    if ref.kind == "excluded":
        ref.exclusion_reason = ref.reason or ref.exclusion_reason
        sel.rejected.append(ref)
        return
    if ref.temporal_role == "BEFORE":
        ref.kind = "before"
        ref.exclusion_reason = ref.reason or "Before image"
        sel.rejected.append(ref)
        return

    eligible = (ref.temporal_role == "AFTER" and ref.semantic_role == SEMANTIC_PRODUCTION
                and ref.confident_owner and bool(ref.owner_id))
    ref.excel_output_eligible = bool(eligible)
    if eligible:
        ref.kind = "after"
        ref.exclusion_reason = ""
        sel.after.append(ref)
        return

    # Known non-production semantics are definite exclusions. Weak ownership or semantics remain
    # ambiguous for review, but neither can enter the final workbook.
    if ref.semantic_role in (SEMANTIC_INSPECTION, SEMANTIC_TEMPORARY, SEMANTIC_VERIFICATION):
        ref.kind = "excluded"
        default_reason = {
            SEMANTIC_INSPECTION: "inspection/control content",
            SEMANTIC_TEMPORARY: "temporary-action content",
            SEMANTIC_VERIFICATION: "verification/follow-up content",
        }[ref.semantic_role]
        ref.exclusion_reason = ref.exclusion_reason or default_reason
        ref.reason = ref.reason or ref.exclusion_reason
    else:
        ref.kind = "ambiguous"
        reasons = []
        if ref.temporal_role != "AFTER":
            reasons.append("temporal role is unknown")
        if ref.semantic_role != SEMANTIC_PRODUCTION:
            reasons.append(f"semantic role is {ref.semantic_role}")
        if not ref.confident_owner:
            reasons.append(ref.exclusion_reason or "no confident logical item owner")
        ref.exclusion_reason = "; ".join(dict.fromkeys(reasons)) or "not confidently eligible for Excel"
        generic = AMBIGUOUS_REASON.format(n=ref.slide)
        if generic not in sel.reasons:
            sel.reasons.append(generic)
    sel.rejected.append(ref)


def _anchor_in_span(anchor, item: ItemRegion, H: int, pad: float = 0.02) -> bool:
    """PROMPT-025: an anchor (caption button / inline '+ Trước:/+ Sau:' line / blue text / arrow) belongs
    to the logical item whose vertical span contains it."""
    cy = getattr(anchor, "cy", None)
    if cy is None:
        cy = anchor.top + anchor.height / 2
    return item.top - pad * H <= cy <= item.bottom + pad * H


def _item_for_picture(p: Block, prod_items: Sequence[ItemRegion], W: int, H: int) -> Optional[ItemRegion]:
    """The logical item a picture belongs to: span containment by centre, nearest item as fallback."""
    if not prod_items:
        return None
    cy = p.top + p.height / 2
    for item in prod_items:
        if item.top <= cy < item.bottom:
            return item
    return min(prod_items, key=lambda item: _region_distance(p, item, W, H))


def _item_anchor_bundles(prod_items: Sequence[ItemRegion], captions, inlines, blue, arrows, claims,
                         content_pics, W: int, H: int) -> Dict[str, Dict[str, object]]:
    """PROMPT-025: per-item temporal anchors. Captions, inline '+ Trước:/+ Sau:' lines, blue After text and
    transition arrows are scoped to the item span, so one item's anchors never classify another item's
    pictures. An item without in-span anchors of a kind falls back to the slide-wide list (prior behavior)."""
    bundles: Dict[str, Dict[str, object]] = {}
    for item in prod_items:
        caps = [a for a in captions if _anchor_in_span(a, item, H)]
        ins = [a for a in inlines if _anchor_in_span(a, item, H)]
        ble = [a for a in blue if _anchor_in_span(a, item, H)]
        arr = [a for a in arrows if _anchor_in_span(a, item, H)]
        votes_i = _arrow_votes(content_pics, arr, W, H) if arr else {}
        if caps:
            cap_ids = {id(c) for c in caps}
            claims_i = {pid: [a for a in anchors if id(a) in cap_ids]
                        for pid, anchors in claims.items()}
        else:
            claims_i = claims
        bundles[item.owner_id] = {"captions": caps, "inlines": ins, "blue": ble,
                                  "votes": votes_i, "claims": claims_i}
    return bundles


def select_after_pictures(report: ReportData, slide_numbers: Sequence[int]) -> PictureSelection:
    """Only confidently owned After + production-improvement pictures are Excel eligible."""
    sel = PictureSelection()
    W, H = report.slide_width or 1, report.slide_height or 1
    for n in sorted(set(int(x) for x in slide_numbers)):
        slide: Optional[SlideData] = report.slide(n)
        if slide is None or not slide.pictures:
            continue
        roles = classify_blocks(slide)
        items = slide_items(slide)
        prod_items = [it for it in items
                      if it.region_kind == "item" and it.semantic_role == SEMANTIC_PRODUCTION]
        LOG.info("ITEM_SEGMENTATION report=%s scope=%s slide=%s items=%d headings=%s",
                 report.filename, report_scope_key(report.path), n, len(prod_items),
                 [it.heading for it in prod_items])
        head = norm_key(title_text(slide))
        slide_is_inspection = bool(head and _INSPECTION_RE.search(head) and
                                   not re.search(r"san xuat|production|process|cong doan", head))
        captions = _caption_anchors(roles)
        inlines, inspection_ranges = _inline_anchors(roles)
        from .extractor import excluded_bands                      # lazy: extractor imports this module
        bands = excluded_bands(slide, H)
        content_pics = [p for p in slide.pictures if not is_decorative_picture(p, W, H)]
        claims = _claims(content_pics, captions, W, H)
        blue = _blue_anchors(roles, inspection_ranges)
        blue = [a for a in blue if not any(y0 <= a.cy < y1 for y0, y1, _k in bands)]
        arrows = [a for a in getattr(slide, "arrows", [])
                  if _arrow_usable(a, W, H, bands, inspection_ranges, roles)]
        votes = _arrow_votes(content_pics, arrows, W, H)
        bundles = _item_anchor_bundles(prod_items, captions, inlines, blue, arrows, claims,
                                       content_pics, W, H)

        # Keep temporal classification independent, then apply semantic/item eligibility after row
        # neighbours have had a chance to inherit a caption's Before/After label.
        working = PictureSelection()
        ambiguous = False
        for picture in sorted(slide.pictures, key=lambda b: b.order):
            why = is_decorative_picture(picture, W, H)
            if why:
                working.rejected.append(PictureRef(n, picture, "excluded", why,
                                                   temporal_role="UNKNOWN"))
                continue
            owner_item = _item_for_picture(picture, prod_items, W, H)
            bundle = bundles.get(owner_item.owner_id) if owner_item is not None else None
            if bundle is not None:
                caps_i = bundle["captions"] or captions
                ins_i = bundle["inlines"] or inlines
                ble_i = bundle["blue"] or blue
                claims_i = bundle["claims"]
                votes_i = bundle["votes"]
            else:
                caps_i, ins_i, ble_i, claims_i, votes_i = captions, inlines, blue, claims, votes
            cy = picture.top + picture.height / 2
            band = next((k for y0, y1, k in bands if y0 <= cy < y1), None)
            if band:
                band_label = "inspection/control" if band == "inspection" else band
                band_kind, band_anchor = classify_picture(picture, caps_i, ins_i, W, H, claims_i, ble_i, votes_i)
                band_temporal = {"after": "AFTER", "before": "BEFORE"}.get(band_kind, "UNKNOWN")
                working.rejected.append(PictureRef(
                    n, picture, "excluded", f"{band_label} section block (mid-slide heading)", band_anchor,
                    temporal_role=band_temporal))
                continue
            if any(y0 <= cy < y1 for y0, y1 in inspection_ranges) and not captions:
                working.rejected.append(PictureRef(n, picture, "excluded", "inspection/control item"))
                continue
            kind, how = classify_picture(picture, caps_i, ins_i, W, H, claims_i, ble_i, votes_i)
            temporal = {"after": "AFTER", "before": "BEFORE"}.get(kind, "UNKNOWN")
            if kind == "after":
                working.after.append(PictureRef(n, picture, "after", "", how, temporal_role=temporal))
            elif kind == "before":
                working.rejected.append(PictureRef(n, picture, "before", "Before picture", how,
                                                   temporal_role=temporal))
            else:
                ambiguous = True
                working.rejected.append(PictureRef(n, picture, "ambiguous", how, temporal_role=temporal))

        ambiguous = _propagate_row_neighbours(working, n, W, H) if ambiguous else ambiguous
        after_blocks = [candidate.block for candidate in working.after]
        for candidate in sorted(working.after, key=lambda r: r.block.order):
            _finalize_picture(candidate, slide, items, slide_is_inspection, W, H, sel, after_blocks, claims)
        for rejected in sorted(working.rejected, key=lambda r: r.block.order):
            # Prior exclusions (decorative/section/inspection) are still given explicit item semantics.
            _finalize_picture(rejected, slide, items, slide_is_inspection, W, H, sel, after_blocks, claims)
        generic = AMBIGUOUS_REASON.format(n=n)
        if ambiguous and generic not in sel.reasons:
            sel.reasons.append(generic)
        if any(r.slide == n and r.excel_output_eligible for r in sel.after):
            sel.slides_with_after.append(n)

    scope_id = report_scope_key(report.path)
    for ref in sel.after + sel.rejected:
        ref.report_scope_id = scope_id
    for ref in sorted(sel.after + sel.rejected, key=lambda r: (r.slide, r.block.order)):
        status = "eligible" if ref.excel_output_eligible else "excluded"
        detail = ref.exclusion_reason or ref.anchor or ""
        sel.notes.append(
            f"{ref.label}: {status}; temporal={ref.temporal_role}; semantic={ref.semantic_role}; "
            f"owner={ref.owner_id or 'unknown'} ({ref.owner_heading!r}); {detail}")
    return sel


def _same_row_adjacent(a: Block, b: Block, W: int, H: int) -> bool:
    v_ov = _overlap(a.top, a.bottom, b.top, b.bottom)
    if v_ov < 0.5 * max(1, min(a.height, b.height)):
        return False
    return _gap(a.left, a.right, b.left, b.right) / W <= 0.08


def _propagate_row_neighbours(sel: PictureSelection, n: int, W: int, H: int) -> bool:
    """Resolve ambiguous pictures of slide ``n`` from caption-decided neighbours in the same row.
    Returns True when at least one picture of the slide is still ambiguous."""
    changed = True
    while changed:
        changed = False
        decided = [r for r in sel.after + sel.rejected
                   if r.slide == n and r.kind in ("after", "before") and r.anchor.startswith(("caption", "row neighbour"))]
        for r in list(sel.rejected):
            if r.slide != n or r.kind != "ambiguous":
                continue
            kinds = {d.kind for d in decided if _same_row_adjacent(r.block, d.block, W, H)}
            if len(kinds) != 1:
                continue
            k = kinds.pop()
            sel.rejected.remove(r)
            how = f"row neighbour of caption '{k}'"
            if k == "after":
                sel.after.append(PictureRef(n, r.block, "after", "", how))
                sel.after.sort(key=lambda x: (x.slide, x.block.order))
            else:
                sel.rejected.append(PictureRef(n, r.block, "before", "Before picture", how))
            changed = True
    return any(r.slide == n and r.kind == "ambiguous" for r in sel.rejected)


def selection_summary(sel: PictureSelection) -> Dict[str, object]:
    return {"after": [r.label for r in sel.after], "before": [r.label for r in sel.rejected if r.kind == "before"],
            "ambiguous": [r.label for r in sel.rejected if r.kind == "ambiguous"],
            "excluded": [r.label for r in sel.rejected if r.kind == "excluded"]}
