"""Content-region analysis of a slide (deterministic, geometry + structure based).

The production report layout contains, besides the main business content:
slide title/header, left-side circular section labels, decorative shapes, image
caption buttons ("Trước cải tiến" / "Sau cải tiến"), arrows and logos.
Only the main content region may end up in the Excel text fields.

Every *text* block of a slide gets a role:

    title      slide title/header ("1. NGUYÊN NHÂN")            -> excluded, sets section kind
    sidebar    left circular / rotated section label            -> excluded
    caption    image caption button ("Trước cải tiến", "Sau")   -> excluded, used as picture anchor
    furniture  logos, page numbers, footer text                  -> excluded
    content    main business content                             -> copied verbatim

All thresholds are relative to the slide size (no absolute pixel positions).
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import List, Optional, Tuple

from .classifier import is_heading_like, section_kind_of_heading
from .pptx_parser import Block, SlideData, norm_key, strip_accents

ROLE_TITLE = "title"
ROLE_SIDEBAR = "sidebar"
ROLE_CAPTION = "caption"
ROLE_FURNITURE = "furniture"
ROLE_CONTENT = "content"

# relative geometry thresholds
TITLE_BAND = 0.17          # blocks whose bottom edge is above 17 % of the slide height are header candidates
TITLE_TOP_MAX = 0.12       # ... and whose top edge is above 12 %
SIDEBAR_RIGHT_MAX = 0.22   # left strip: right edge <= 22 % of the slide width
SIDEBAR_WIDTH_MAX = 0.22
FOOTER_BAND = 0.92         # bottom strip
CORNER_LEFT_MIN = 0.78     # top-right corner (logos)
SIDEBAR_MAX_CHARS = 45
CAPTION_MAX_CHARS = 40

_CAPTION_RE = re.compile(
    r"^(?:hinh anh |hinh |anh |image |picture |photo )?"
    r"(?:(?P<before>truoc|before)|(?P<after>sau|after))"
    r"(?: khi)?(?: cai tien| improvement| khac phuc)?(?: :)?$")
_CAPTION_NEUTRAL_RE = re.compile(r"^(?:hinh anh|hinh anh minh hoa|anh minh hoa|hinh minh hoa|image|picture|photo|"
                                 r"minh hoa|truoc / sau|truoc sau|before / after|before after)$")
_SYMBOL_ONLY_RE = re.compile(r"^[\W_\d]+$")
_LABEL_PRST = {"ellipse", "roundRect", "flowChartConnector", "donut", "chevron", "homePlate", "pentagon",
               "hexagon", "octagon", "diamond", "wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout"}
_ARROW_PRST_RE = re.compile(r"arrow|chevron|homePlate", re.IGNORECASE)

# inline "+ Trước:" / "+ Sau:" business lines (these stay in the text, and anchor pictures)
_INLINE_ANCHOR_RE = re.compile(r"^[\+\-\u2022\u00b7\*\u25aa\u27a2\u25ba>\u2192]?\s*"
                               r"(?:hinh anh |anh )?(?P<kind>truoc|sau|before|after)"
                               r"(?: khi)?(?: cai tien| improvement)?\s*(?::|-|\u2013)")


@dataclass
class BlockRole:
    block: Block
    role: str
    caption_kind: Optional[str] = None      # "before" | "after" | None (neutral caption)
    note: str = ""

    @property
    def text(self) -> str:
        return self.block.text


@dataclass
class StructuralCauseRegion:
    """Body text structurally owned by a separate left-side cause marker."""
    marker: BlockRole
    body: List[BlockRole]
    excluded: List[Tuple[BlockRole, str]]


def _rel(v: int, total: int) -> float:
    return (v / total) if total else 0.0


def caption_kind(text: str) -> Optional[str]:
    """'before' / 'after' for caption buttons such as 'Trước cải tiến', 'Sau', 'After'; None otherwise."""
    k = norm_key(text)
    k = re.sub(r"\s*:\s*$", "", k).strip()
    m = _CAPTION_RE.match(k)
    if not m:
        return None
    return "before" if m.group("before") else "after"


def is_neutral_caption(text: str) -> bool:
    return bool(_CAPTION_NEUTRAL_RE.match(norm_key(text)))


def _fold_line(line: str) -> str:
    """Accent/case fold that KEEPS punctuation (':' matters for '+ Trước:')."""
    return re.sub(r"\s+", " ", strip_accents(line or "").lower()).strip()


def inline_anchor_kind(line: str) -> Optional[str]:
    """'before' / 'after' when the business line starts with '+ Trước:' / '+ Sau:' (text keeps the line)."""
    m = _INLINE_ANCHOR_RE.match(_fold_line(line))
    if not m:
        return None
    return "before" if m.group("kind") in ("truoc", "before") else "after"


def is_slide_level_heading(line: str) -> bool:
    """Numbered ('3. CẢI TIẾN …') or fully upper-case section heading = page structure, not content."""
    t = line.strip()
    if not t or section_kind_of_heading(t) is None:
        return False
    if re.match(r"^\s*(?:\d+(?:\.\d+)*[.)]?|[IVX]+[.)])\s+\S", t):
        return True
    letters = [c for c in t if c.isalpha()]
    return bool(letters) and all(c.isupper() for c in letters)


def _first_line(b: Block) -> str:
    return b.text.strip().splitlines()[0] if b.text.strip() else ""


def classify_blocks(slide: SlideData) -> List[BlockRole]:
    """Assign a role to every text block of the slide (pictures are handled by improvement_pictures)."""
    W, H = slide.width or 1, slide.height or 1
    roles: List[BlockRole] = []
    text_blocks = slide.text_blocks
    for b in text_blocks:
        txt = b.text.strip()
        lines = [ln for ln in txt.splitlines() if ln.strip()]
        first = _first_line(b)
        n_chars = len(txt)
        top_r, bottom_r = _rel(b.top, H), _rel(b.bottom, H)
        left_r, right_r, width_r = _rel(b.left, W), _rel(b.right, W), _rel(b.width, W)
        rotated = abs(((b.rotation or 0.0) + 180) % 360 - 180) >= 45

        # --- symbols only (arrows drawn as text, numbering bubbles) -----------------
        if _SYMBOL_ONLY_RE.match(txt):
            roles.append(BlockRole(b, ROLE_FURNITURE, note="symbols only"))
            continue

        # --- caption buttons ("Trước cải tiến", "Sau cải tiến", "Before", "After") ---
        ck = caption_kind(txt) if n_chars <= CAPTION_MAX_CHARS and len(lines) == 1 else None
        if ck:
            roles.append(BlockRole(b, ROLE_CAPTION, caption_kind=ck, note="caption"))
            continue
        if len(lines) == 1 and n_chars <= CAPTION_MAX_CHARS and is_neutral_caption(txt):
            roles.append(BlockRole(b, ROLE_CAPTION, caption_kind=None, note="neutral caption"))
            continue

        # --- slide title / header --------------------------------------------------
        is_title_placeholder = b.kind == "title"
        in_title_band = top_r <= TITLE_TOP_MAX and bottom_r <= TITLE_BAND + 0.08
        heading = len(lines) <= 2 and is_heading_like(first, b.bold, b.size_pt)
        if is_title_placeholder and len(lines) <= 2:
            roles.append(BlockRole(b, ROLE_TITLE, note="title placeholder"))
            continue
        if in_title_band and heading and (section_kind_of_heading(first) is not None or is_slide_level_heading(first)
                                          or width_r >= 0.3):
            roles.append(BlockRole(b, ROLE_TITLE, note="header band"))
            continue

        # --- left circular / rotated section labels ---------------------------------
        short = n_chars <= SIDEBAR_MAX_CHARS and len(lines) <= 3
        in_left_strip = right_r <= SIDEBAR_RIGHT_MAX and width_r <= SIDEBAR_WIDTH_MAX
        if short and (in_left_strip or rotated or (b.prst in _LABEL_PRST and width_r <= 0.3 and
                                                   section_kind_of_heading(first) is not None)):
            roles.append(BlockRole(b, ROLE_SIDEBAR, note="sidebar label"))
            continue

        # --- page furniture: footer, top-right logos, tiny text ---------------------
        if len(lines) == 1 and n_chars <= 60 and top_r >= FOOTER_BAND:
            roles.append(BlockRole(b, ROLE_FURNITURE, note="footer"))
            continue
        if len(lines) <= 2 and n_chars <= 30 and left_r >= CORNER_LEFT_MIN and top_r <= TITLE_TOP_MAX:
            roles.append(BlockRole(b, ROLE_FURNITURE, note="corner logo text"))
            continue
        if b.size_pt is not None and b.size_pt <= 7 and n_chars <= 40:
            roles.append(BlockRole(b, ROLE_FURNITURE, note="tiny text"))
            continue
        if b.prst and _ARROW_PRST_RE.search(b.prst) and n_chars <= 12:
            roles.append(BlockRole(b, ROLE_FURNITURE, note="arrow shape"))
            continue

        roles.append(BlockRole(b, ROLE_CONTENT))
    return roles


def _is_separate_cause_marker(role: BlockRole, W: int) -> bool:
    text = norm_key(role.text).strip(" .:;,-")
    if text not in {"nguyen nhan", "phan tich nguyen nhan", "root cause", "cause analysis"}:
        return False
    b = role.block
    # A separate marker is a narrow left label/sidebar, not a sentence in the body.
    return role.role == ROLE_SIDEBAR or (b.left / max(1, W) <= 0.25 and b.width / max(1, W) <= 0.25)


def _cause_candidate_exclusion(role: BlockRole) -> Optional[str]:
    """Return why an adjacent block cannot be a root-cause body."""
    b = role.block
    first = _first_line(b)
    key = norm_key(first)
    if not key:
        return "empty body"
    if key.strip(" .:;,-") in {"nguyen nhan", "phan tich nguyen nhan", "root cause", "cause analysis"}:
        return "marker label, not body"
    if (b.kind == "title" or (role.role == ROLE_TITLE and is_slide_level_heading(first))
            or re.search(r"\b(bao cao|report)\b", key)):
        return "slide/report title"
    kind = section_kind_of_heading(first) if is_heading_like(first, b.bold, b.size_pt) else None
    if kind and kind != "cause":
        return f"non-cause section heading ({kind})"
    if re.search(r"\b(hien trang|current status|temporary action|xu ly tam thoi)\b", key):
        return "current-state or temporary-action text"
    return None


def cause_sidebar_regions(slide: SlideData) -> List[StructuralCauseRegion]:
    """Detect separate left-side ``Nguyên nhân`` markers and assign adjacent body text by geometry.

    Ownership uses relative horizontal adjacency and vertical affinity, then excludes titles,
    furniture, current-state, temporary, and improvement section blocks. Grouped shapes and
    table blocks already carry absolute slide geometry from :mod:`pptx_parser`.
    """
    W, H = slide.width or 1, slide.height or 1
    roles = classify_blocks(slide)
    markers = [r for r in roles if _is_separate_cause_marker(r, W)]
    regions: List[StructuralCauseRegion] = []
    for marker in markers:
        mb = marker.block
        body: List[BlockRole] = []
        excluded: List[Tuple[BlockRole, str]] = []
        for candidate in roles:
            b = candidate.block
            if b is mb or not b.text.strip():
                continue
            if candidate.role not in (ROLE_CONTENT, ROLE_TITLE):
                continue
            # Body must be to the marker's right (small overlaps are tolerated for rounded sidebars).
            if b.right <= mb.left or b.left < mb.left - 0.02 * W:
                continue
            xgap = max(0, b.left - mb.right) / W
            if xgap > 0.16 or b.width / W < 0.20:
                continue
            ygap = max(0, max(mb.top, b.top) - min(mb.bottom, b.bottom)) / H
            if ygap > 0.40:
                continue
            reason = _cause_candidate_exclusion(candidate)
            if reason:
                excluded.append((candidate, reason))
            else:
                body.append(candidate)
        body.sort(key=lambda r: (r.block.order, r.block.top, r.block.left))
        regions.append(StructuralCauseRegion(marker=marker, body=body, excluded=excluded))
    return regions


def content_blocks(slide: SlideData) -> List[Block]:
    return [r.block for r in classify_blocks(slide) if r.role == ROLE_CONTENT]


def caption_blocks(slide: SlideData) -> List[BlockRole]:
    return [r for r in classify_blocks(slide) if r.role == ROLE_CAPTION]


def title_text(slide: SlideData) -> str:
    """Header text of the slide (title role blocks and a slide-level first line of the first content block)."""
    roles = classify_blocks(slide)
    parts = [r.text.strip() for r in roles if r.role == ROLE_TITLE]
    if not parts:
        for r in roles:
            if r.role == ROLE_CONTENT:
                first = _first_line(r.block)
                if is_slide_level_heading(first):
                    parts.append(first)
                break
    return "\n".join(parts)


def excluded_summary(slide: SlideData) -> List[str]:
    """Diagnostics: what was left out of the text fields and why."""
    out = []
    for r in classify_blocks(slide):
        if r.role != ROLE_CONTENT:
            t = r.text.strip().replace("\n", " / ")
            out.append(f"S{slide.number} {r.role}: {t[:60]}")
    return out
