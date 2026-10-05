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

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from .content_region import (ROLE_CAPTION, ROLE_CONTENT, BlockRole, classify_blocks, inline_anchor_kind,
                             title_text)
from .pptx_parser import Block, ReportData, SlideData, norm_key

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
class PictureRef:
    slide: int
    block: Block
    kind: str                   # "after" | "before" | "ambiguous" | "excluded"
    reason: str = ""
    anchor: str = ""            # how it was decided

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


def select_after_pictures(report: ReportData, slide_numbers: Sequence[int]) -> PictureSelection:
    """After-only pictures of the given improvement slides, in slide/reading order."""
    sel = PictureSelection()
    W, H = report.slide_width or 1, report.slide_height or 1
    for n in sorted(set(int(x) for x in slide_numbers)):
        s: Optional[SlideData] = report.slide(n)
        if s is None or not s.pictures:
            continue
        roles = classify_blocks(s)
        head = norm_key(title_text(s))
        if head and _INSPECTION_RE.search(head) and not re.search(r"san xuat|production|process|cong doan", head):
            for p in s.pictures:
                sel.rejected.append(PictureRef(n, p, "excluded", "inspection/control improvement slide"))
            sel.notes.append(f"S{n}: ảnh thuộc cải tiến kiểm tra/kiểm soát – không chèn")
            continue
        captions = _caption_anchors(roles)
        inlines, inspection_ranges = _inline_anchors(roles)
        from .extractor import excluded_bands                      # lazy: extractor imports this module
        bands = excluded_bands(s, H)
        content_pics = [p for p in s.pictures if not is_decorative_picture(p, W, H)]
        claims = _claims(content_pics, captions, W, H)
        blue = _blue_anchors(roles, inspection_ranges)
        blue = [a for a in blue if not any(y0 <= a.cy < y1 for y0, y1, _k in bands)]
        arrows = [a for a in getattr(s, "arrows", []) if _arrow_usable(a, W, H, bands, inspection_ranges, roles)]
        votes = _arrow_votes(content_pics, arrows, W, H)
        ambiguous = False
        for p in sorted(s.pictures, key=lambda b: b.order):
            why = is_decorative_picture(p, W, H)
            if why:
                sel.rejected.append(PictureRef(n, p, "excluded", why))
                continue
            cy = p.top + p.height / 2
            band = next((k for y0, y1, k in bands if y0 <= cy < y1), None)
            if band:                                                # PROMPT-001: excluded section block -> zero pictures
                sel.rejected.append(PictureRef(n, p, "excluded", f"{band} section block (mid-slide heading)"))
                continue
            if any(y0 <= cy < y1 for y0, y1 in inspection_ranges) and not captions:
                sel.rejected.append(PictureRef(n, p, "excluded", "inspection/control item"))
                continue
            kind, how = classify_picture(p, captions, inlines, W, H, claims, blue, votes)
            if kind == "after":
                sel.after.append(PictureRef(n, p, "after", "", how))
            elif kind == "before":
                sel.rejected.append(PictureRef(n, p, "before", "Before picture", how))
            else:
                ambiguous = True
                sel.rejected.append(PictureRef(n, p, "ambiguous", how))
        # pictures grouped in the same row next to a caption-labelled picture share its caption
        # (a caption button usually labels the whole group of pictures beside it)
        ambiguous = _propagate_row_neighbours(sel, n, W, H) if ambiguous else ambiguous
        if ambiguous:
            sel.reasons.append(AMBIGUOUS_REASON.format(n=n))
        if any(r.slide == n for r in sel.after):
            sel.slides_with_after.append(n)
    for r in sel.rejected:
        sel.notes.append(f"{r.label}: {r.kind} – {r.reason or r.anchor}")
    for r in sel.after:
        sel.notes.append(f"{r.label}: after – {r.anchor}")
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
