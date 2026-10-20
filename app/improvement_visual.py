"""Owner-scoped After visual regions and their rendered-slide crop geometry.

Semantic picture selection remains in :mod:`improvement_pictures` / image learning. This module builds the final
visual evidence from refs that have already passed the strict Excel eligibility gate.

PROMPT-027R: a picture is only a SEED.  The region of one authored After block is the union of its member
pictures, positively associated captions and visual annotations, bounded by its item span and by the next
improvement heading (a hard structural obstacle), plus bounded deterministic padding.  Nothing here is derived
from pixels, from literal words or from a candidate's own bounds.
"""
from __future__ import annotations

import logging
import math
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Sequence, Set, Tuple

from .cancellation import check_cancelled
from .content_region import ROLE_CAPTION, ROLE_FURNITURE, ROLE_SIDEBAR, ROLE_TITLE, classify_blocks
from .improvement_pictures import (CAPTION_GAP_MAX, CAPTION_SIDE_GAP_MAX, COLUMN_HEADER_GAP_MAX,
                                   SEMANTIC_PRODUCTION, ItemRegion, PictureRef, gap_is_clear, group_refs,
                                   is_decorative_picture, slide_items, structural_obstacles)
from .pptx_parser import Block, ReportData
from .report_identity import report_scope_key

LOG = logging.getLogger("report_extractor.images")
BBox = Tuple[int, int, int, int]  # left, top, right, bottom; PPTX EMU coordinates

# Padding and object association are relative to the slide, not report-specific coordinates.
DEFAULT_SAFE_PADDING_FRACTION = 0.008
DEFAULT_ASSOCIATION_GAP_FRACTION = 0.006
# PROMPT-025: absolute authored adjacency tolerance (unchanged).
CLUSTER_ROW_GAP_FRACTION = 0.08
CLUSTER_COL_GAP_FRACTION = 0.08
# PROMPT-027R §7: adjacency is also relative to the authored PICTURE size (a local, structural spacing scale)
# and requires that no structural object (another picture, caption or item heading) sits in the gap.
PICTURE_GAP_RATIO = 1.0
# Pictures that share an After caption or an authored group are related more loosely, still gap-checked.
RELATED_GAP_RATIO = 2.0
# PROMPT-027R §6/§8: which visual objects belong to the After block.
OVERLAY_FRACTION = 0.04       # a shape covering this share of a member picture is an overlay (annotation)
CONTAINMENT_MIN = 0.5         # a visual shape must lie mostly inside the After footprint …
TEXT_CONTAINMENT_MIN = 0.6    # … and text must lie mostly inside it (body paragraphs never qualify)
FOOTER_BAND = 0.75            # furniture below this fraction of the slide height is a footer boundary


@dataclass
class ImprovementVisualRegion:
    """One logically owned After group, rendered and cropped from one source slide.

    PROMPT-025: the region is ITEM-scoped (``improvement_item_id`` + ``item_index``), never slide-scoped.
    ``after_block_index`` distinguishes multiple authored After groups of one item (§19).
    PROMPT-027R: ``span_*`` / ``boundary_*`` record the structural item span and the boundary that actually
    limited the crop; ``raw_bbox`` / ``padding`` expose the geometry before and after bounded padding.
    """

    report_scope_id: str
    management_number: str
    slide_index: int
    improvement_item_id: str
    improvement_heading: str
    semantic_role: str
    temporal_role: str
    bbox: BBox
    source_order: int
    confidence: Optional[float]
    excel_output_eligible: bool
    anchor_text: str = ""
    pictures: List[PictureRef] = field(default_factory=list)
    included_objects: List[Block] = field(default_factory=list)
    excluded_objects: List[str] = field(default_factory=list)
    item_index: int = -1
    after_block_index: int = 0
    span_top: int = 0
    span_bottom: int = 0
    boundary_source: str = ""
    boundary_item_id: str = ""
    boundary_top: Optional[int] = None
    raw_bbox: Optional[BBox] = None
    padding: Tuple[int, int, int, int] = (0, 0, 0, 0)

    @property
    def width(self) -> int:
        return max(0, self.bbox[2] - self.bbox[0])

    @property
    def height(self) -> int:
        return max(0, self.bbox[3] - self.bbox[1])

    @property
    def region_id(self) -> str:
        return f"{self.report_scope_id}:S{self.slide_index}:{self.improvement_item_id}"

    @property
    def block_id(self) -> str:
        """Unique region identity; identical to ``region_id`` for the item's first After block."""
        if self.after_block_index:
            return f"{self.region_id}#A{self.after_block_index}"
        return self.region_id

    @property
    def picture_shape_ids(self) -> List[int]:
        return [int(ref.block.shape_id) for ref in self.pictures]


@dataclass(frozen=True)
class _Span:
    """Vertical span of one improvement item and the structural boundary that ends it (PROMPT-027R §10)."""
    top: int
    bottom: int
    source: str                  # next_item_heading | next_section_heading | footer | slide_bottom
    next_item_id: str = ""
    next_heading_top: Optional[int] = None


@dataclass(frozen=True)
class _Boundary:
    """A heading or footer below an item that a crop must never reach (an obstacle, never a member)."""
    left: int
    top: int
    right: int
    bottom: int
    source: str
    item_id: str = ""


def _eligible(ref: PictureRef) -> bool:
    """Reassert PROMPT-015 at the visual-evidence boundary; labels never bypass this gate."""
    return bool(ref.excel_output_eligible and ref.temporal_role == "AFTER"
                and ref.semantic_role == SEMANTIC_PRODUCTION and ref.confident_owner and ref.owner_id)


def _same_group(a: Block, b: Block) -> bool:
    """Group ancestry as association evidence (PROMPT-027R §8): same innermost authored group."""
    pa = getattr(a, "group_path", ()) or ()
    pb = getattr(b, "group_path", ()) or ()
    return bool(pa) and bool(pb) and pa[-1] == pb[-1]


def _cluster_with_reasons(pictures: Sequence[PictureRef], captions: Sequence[Block], slide_width: int,
                          slide_height: int, obstacles: Sequence[BBox] = (),
                          claim_obstacles: Optional[Sequence[BBox]] = None
                          ) -> Tuple[List[List[PictureRef]], List[str]]:
    """Clusters + human-readable join decisions (PROMPT-027R §7).

    Two same-owner After pictures are ONE block when they are authored in the same band (row: vertical overlap,
    or column: horizontal overlap), are not further apart than the picture scale allows, and no structural object
    sits in the gap between them.  Transitivity keeps a chain A-B-C together.  Two pictures claimed by DIFFERENT
    After captions are separate blocks.  Shared caption / shared authored group relax only the distance.
    """
    W, H = max(1, slide_width), max(1, slide_height)
    n = len(pictures)
    gap_source = list(obstacles)
    claim_source = list(obstacles if claim_obstacles is None else claim_obstacles)
    rects = [_rotated_rect(ref.block) for ref in pictures]
    claimed: List[Set[int]] = [
        {id(caption) for caption in captions if _caption_distance(ref.block, caption, W, H, claim_source) is not None}
        for ref in pictures]
    caption_rects = {id(caption): _rotated_rect(caption) for caption in captions}

    def side(caption_id: int, index: int) -> str:
        return "above" if caption_rects[caption_id][3] <= rects[index][1] else "below"

    parent = list(range(n))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i: int, j: int) -> None:
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[max(ri, rj)] = min(ri, rj)

    def _starts_block(i: int, j: int, axis: str) -> bool:
        """An After caption ABOVE the later picture (in reading order) opens a new block there, unless the earlier
        picture is labelled by that same caption.  Unlabelled pictures before it keep joining the previous block."""
        if axis == "row":
            earlier, later = (i, j) if rects[i][0] <= rects[j][0] else (j, i)
        else:
            earlier, later = (i, j) if rects[i][1] <= rects[j][1] else (j, i)
        above_later = {cid for cid in claimed[later] if side(cid, later) == "above"}
        return bool(above_later) and not (claimed[earlier] & above_later)

    reasons: List[str] = []
    for i in range(n):
        for j in range(i + 1, n):
            a, b = rects[i], rects[j]
            shared = claimed[i] & claimed[j]
            # Two pictures labelled by DIFFERENT After captions from the SAME side (each caption above its own block,
            # or each beneath its own block) are two blocks.  A caption above one picture and another beneath its
            # neighbour is the normal authoring of ONE block with two labels, so it does not split.
            sides_i = {side(cid, i) for cid in claimed[i]}
            sides_j = {side(cid, j) for cid in claimed[j]}
            if claimed[i] and claimed[j] and not shared and (sides_i & sides_j):
                reasons.append(f"{i}-{j}:separate-after-captions")
                continue
            same_group = _same_group(pictures[i].block, pictures[j].block)
            ratio = RELATED_GAP_RATIO if (shared or same_group) else PICTURE_GAP_RATIO
            exempt = {caption_rects[cid] for cid in shared if cid in caption_rects}
            gap_obstacles = [o for o in gap_source if o not in exempt]
            joined = ""
            if _overlap_y(a, b) >= 0.5 * max(1, min(a[3] - a[1], b[3] - b[1])):
                min_w = max(1, min(a[2] - a[0], b[2] - b[0]))
                if _gap_x(a, b) <= max(CLUSTER_ROW_GAP_FRACTION * W, ratio * min_w):
                    if gap_is_clear(a, b, gap_obstacles):
                        joined = "row"
                    else:
                        reasons.append(f"{i}-{j}:row-gap-holds-structure")
            if not joined and _overlap_x(a, b) >= 0.5 * max(1, min(a[2] - a[0], b[2] - b[0])):
                min_h = max(1, min(a[3] - a[1], b[3] - b[1]))
                if _gap_y(a, b) <= max(CLUSTER_COL_GAP_FRACTION * H, ratio * min_h):
                    if gap_is_clear(a, b, gap_obstacles):
                        joined = "column"
                    else:
                        reasons.append(f"{i}-{j}:column-gap-holds-structure")
            if joined and _starts_block(i, j, joined):
                reasons.append(f"{i}-{j}:new-block-at-caption")
                continue
            if joined:
                union(i, j)
                suffix = ("+shared-caption" if shared else "") + ("+group" if same_group else "")
                reasons.append(f"{i}-{j}:{joined}{suffix}")
    clusters: Dict[int, List[PictureRef]] = {}
    for i, ref in enumerate(pictures):
        clusters.setdefault(find(i), []).append(ref)
    ordered = [sorted(members, key=lambda r: (r.source_order, r.block.order))
               for _, members in sorted(clusters.items(), key=lambda kv: min(r.source_order for r in kv[1]))]
    return ordered, reasons


def cluster_after_pictures(pictures: Sequence[PictureRef], captions: Sequence[Block],
                           slide_width: int, slide_height: int, obstacles: Sequence[BBox] = (),
                           claim_obstacles: Optional[Sequence[BBox]] = None) -> List[List[PictureRef]]:
    """PROMPT-025 §18 / PROMPT-027R §7: cluster one item's eligible After pictures into logical After blocks.

    ``obstacles`` are structural rectangles (other pictures, captions, item headings) that may separate members;
    ``claim_obstacles`` is the separator set used when deciding which caption labels which picture.
    """
    clusters, _reasons = _cluster_with_reasons(pictures, captions, slide_width, slide_height, obstacles,
                                               claim_obstacles)
    return clusters


def _rotated_rect(block: Block) -> BBox:
    left, top, right, bottom = block.left, block.top, block.right, block.bottom
    if not block.rotation or block.width <= 0 or block.height <= 0:
        return left, top, right, bottom
    cx, cy = left + block.width / 2.0, top + block.height / 2.0
    theta = math.radians(block.rotation)
    cosine, sine = math.cos(theta), math.sin(theta)
    points = []
    for x, y in ((left, top), (right, top), (right, bottom), (left, bottom)):
        dx, dy = x - cx, y - cy
        points.append((cx + dx * cosine - dy * sine, cy + dx * sine + dy * cosine))
    return (int(math.floor(min(p[0] for p in points))), int(math.floor(min(p[1] for p in points))),
            int(math.ceil(max(p[0] for p in points))), int(math.ceil(max(p[1] for p in points))))


def _union(boxes: Sequence[BBox]) -> BBox:
    return (min(box[0] for box in boxes), min(box[1] for box in boxes),
            max(box[2] for box in boxes), max(box[3] for box in boxes))


def _expand(box: BBox, dx: int, dy: int) -> BBox:
    return box[0] - dx, box[1] - dy, box[2] + dx, box[3] + dy


def _safe_padding(box: BBox, obstacles: Sequence[BBox], pad_x: int, pad_y: int,
                  slide_width: int, slide_height: int) -> Tuple[int, int, int, int]:
    """Keep safe padding where possible, but do not expand across adjacent unowned visual objects.

    Leave roughly one pixel of separation at the built-in renderer's default 1920x1080 slide scale so the outward
    crop rounding does not pull a neighboring arrow/logo/Before picture or annotation back into the region.
    """
    guard_x = max(1, int(math.ceil(slide_width / 1920)))
    guard_y = max(1, int(math.ceil(slide_height / 1080)))
    left, top, right, bottom = pad_x, pad_y, pad_x, pad_y
    for obstacle in obstacles:
        overlaps_x = _overlap_x(box, obstacle) > 0
        overlaps_y = _overlap_y(box, obstacle) > 0
        if overlaps_x:
            if obstacle[3] <= box[1]:
                top = min(top, max(0, box[1] - obstacle[3] - guard_y))
            elif obstacle[1] >= box[3]:
                bottom = min(bottom, max(0, obstacle[1] - box[3] - guard_y))
        if overlaps_y:
            if obstacle[2] <= box[0]:
                left = min(left, max(0, box[0] - obstacle[2] - guard_x))
            elif obstacle[0] >= box[2]:
                right = min(right, max(0, obstacle[0] - box[2] - guard_x))
    return left, top, right, bottom


def _overlap_x(a: BBox, b: BBox) -> int:
    return max(0, min(a[2], b[2]) - max(a[0], b[0]))


def _overlap_y(a: BBox, b: BBox) -> int:
    return max(0, min(a[3], b[3]) - max(a[1], b[1]))


def _intersection_area(a: BBox, b: BBox) -> int:
    return _overlap_x(a, b) * _overlap_y(a, b)


def _area(box: BBox) -> int:
    return max(0, box[2] - box[0]) * max(0, box[3] - box[1])


def _gap_x(a: BBox, b: BBox) -> int:
    return max(0, max(a[0], b[0]) - min(a[2], b[2]))


def _gap_y(a: BBox, b: BBox) -> int:
    return max(0, max(a[1], b[1]) - min(a[3], b[3]))


def _caption_distance(picture: Block, caption: Block, slide_width: int, slide_height: int,
                      obstacles: Optional[Sequence[BBox]] = None) -> Optional[float]:
    """Same relative above/below/side association used by deterministic Before/After selection.

    PROMPT-027R §5: with ``obstacles`` a caption cannot label a picture across another caption, picture or heading.
    """
    p, c = _rotated_rect(picture), _rotated_rect(caption)
    h_overlap, v_overlap = _overlap_x(p, c), _overlap_y(p, c)
    if h_overlap and v_overlap:
        return 0.0
    min_width = max(1, min(p[2] - p[0], c[2] - c[0]))
    min_height = max(1, min(p[3] - p[1], c[3] - c[1]))
    v_gap = _gap_y(p, c) / max(1, slide_height)
    h_gap = _gap_x(p, c) / max(1, slide_width)
    distance: Optional[float] = None
    if h_overlap >= 0.5 * min_width:                              # caption above / below
        if v_gap <= CAPTION_GAP_MAX:
            distance = v_gap
        elif c[3] <= p[1] and v_gap <= COLUMN_HEADER_GAP_MAX:     # column header
            distance = v_gap + 0.10
    if distance is None and v_overlap >= 0.5 * min_height and h_gap <= CAPTION_SIDE_GAP_MAX:  # caption beside
        distance = h_gap + 0.02
    if distance is None or not obstacles:
        return distance
    if not gap_is_clear(p, c, list(obstacles)):
        return None
    return distance


def _segment_fraction_in_rect(points: Tuple[int, int, int, int], box: BBox) -> float:
    """Fraction of a straight segment inside a rectangle (Liang–Barsky clipping)."""
    x0, y0, x1, y1 = points
    dx, dy = x1 - x0, y1 - y0
    t0, t1 = 0.0, 1.0
    for p, q in ((-dx, x0 - box[0]), (dx, box[2] - x0),
                 (-dy, y0 - box[1]), (dy, box[3] - y0)):
        if abs(p) < 1e-12:
            if q < 0:
                return 0.0
            continue
        ratio = q / p
        if p < 0:
            t0 = max(t0, ratio)
        else:
            t1 = min(t1, ratio)
        if t0 > t1:
            return 0.0
    return max(0.0, min(1.0, t1 - t0))


def _centres_inside(box: BBox, pictures: Sequence[Block]) -> List[Block]:
    return [picture for picture in pictures
            if box[0] <= picture.left + picture.width / 2 <= box[2]
            and box[1] <= picture.top + picture.height / 2 <= box[3]]


def _is_arrow(block: Block, arrow_z: set) -> bool:
    return (block.z_order in arrow_z or bool(block.direction) or "arrow" in (block.prst or "").lower()
            or "chevron" in (block.prst or "").lower())


def _object_intersects_core(block: Block, core: BBox, expanded_core: BBox, arrow_z: set,
                            picture_boxes: Sequence[BBox]) -> bool:
    is_arrow = _is_arrow(block, arrow_z)
    if block.line_endpoints:
        core_fraction = _segment_fraction_in_rect(block.line_endpoints, core)
        picture_fraction = max((_segment_fraction_in_rect(block.line_endpoints, box)
                                for box in picture_boxes), default=0.0)
        if is_arrow:
            # A transition arrow may point toward the After column without belonging to it. Keep arrows
            # that lie mostly inside the After group or visibly overlay one of its pictures.
            return picture_fraction >= 0.04 or core_fraction >= 0.55
        length = math.hypot(block.line_endpoints[2] - block.line_endpoints[0],
                            block.line_endpoints[3] - block.line_endpoints[1])
        group_span = max(core[2] - core[0], core[3] - core[1], 1)
        return picture_fraction >= 0.05 or core_fraction >= 0.30 or (
            core_fraction > 0 and length <= group_span * 1.5)
    rect = _rotated_rect(block)
    if is_arrow:
        rect_area = max(1, _area(rect))
        overlays_picture = any(_intersection_area(rect, picture) / rect_area >= 0.04
                               for picture in picture_boxes)
        centre_inside = (core[0] <= (rect[0] + rect[2]) / 2 <= core[2]
                         and core[1] <= (rect[1] + rect[3]) / 2 <= core[3])
        return overlays_picture or centre_inside
    return _intersection_area(rect, expanded_core) > 0


def _role_for_block(block: Block, roles_by_z: Dict[int, list]) -> Optional[object]:
    matches = roles_by_z.get(block.z_order, []) if block.z_order >= 0 else []
    if not matches:
        return None
    # A shape may contain several paragraphs. Keep the role whose bounds best match this visual/text block.
    return min(matches, key=lambda role: abs(role.block.left - block.left) + abs(role.block.top - block.top))


def _describe(block: Block) -> str:
    preview = " ".join((block.text or "").split())[:64]
    return (f"S{block.z_order} SH{block.shape_id} {block.shape_name or block.shape_type or block.kind} "
            f"bbox={_rotated_rect(block)}" + (f" text={preview!r}" if preview else ""))


def _group_label(block: Block) -> str:
    path = getattr(block, "group_path", ()) or ()
    return "/".join(str(part) for part in path) if path else "-"


def _item_span(owner: Optional[ItemRegion], items: Sequence[ItemRegion], W: int, H: int) -> Optional[_Span]:
    """The item span and the boundary that ends it, using the SAME criterion as :func:`slide_items`.

    A later heading ends the span only when it is horizontally related to the owner's heading (or spans nearly the
    whole slide).  The boundary source is recorded so Windows acceptance can see which rule limited the item.
    """
    if owner is None:
        return None
    o_left, o_top, o_right, o_bottom = owner.heading_box
    following: List[ItemRegion] = []
    for other in items:
        if other is owner or other.top <= owner.top or other.source_order <= owner.source_order:
            continue
        if not (other.region_kind == "item" or other.semantic_role != SEMANTIC_PRODUCTION):
            continue                                       # generic production headings are context, not ends
        x0, _y0, x1, _y1 = other.heading_box
        full_width = (x1 - x0) >= 0.55 * W
        overlaps = not (x1 <= o_left or x0 >= o_right)
        if full_width or overlaps:
            following.append(other)
    nxt = min(following, key=lambda it: it.top) if following else None
    if nxt is not None and nxt.top <= owner.bottom:
        source = "next_item_heading" if nxt.region_kind == "item" else "next_section_heading"
        return _Span(owner.top, owner.bottom, source, nxt.owner_id, nxt.top)
    source = "slide_bottom" if owner.bottom >= H else "footer"
    return _Span(owner.top, owner.bottom, source)


def _obstacles_below(owner: Optional[ItemRegion], items: Sequence[ItemRegion], roles, W: int, H: int
                     ) -> List[_Boundary]:
    """Every heading (item or section) and footer below the owner: hard obstacles for the final crop."""
    out: List[_Boundary] = []
    for other in items:
        if other is owner or (owner is not None and other.top <= owner.top):
            continue
        left, top, right, bottom = other.heading_box
        if right <= left or bottom <= top:
            continue
        source = "next_item_heading" if other.region_kind == "item" else "next_section_heading"
        out.append(_Boundary(left, top, right, bottom, source, other.owner_id))
    for role in roles:
        block = role.block
        if role.role == ROLE_FURNITURE and H and block.text.strip() and block.top / H >= FOOTER_BAND:
            out.append(_Boundary(block.left, block.top, block.right, block.bottom, "footer"))
    return out


def _clamp_bottom(bbox: BBox, members_bottom: int, boundaries: Sequence[_Boundary], guard_y: int
                  ) -> Tuple[int, Optional[_Boundary], List[_Boundary]]:
    """Lower the crop bottom above every boundary it horizontally reaches (PROMPT-027R §10/§11).

    The test is made against the FINAL crop's horizontal extent, not against the picture extent: a heading that
    shares a single column with the crop leaks just as much as one that sits under the pictures.  A boundary that
    would cut through a member picture is reported as a conflict instead of dropping the member.
    """
    bottom = bbox[3]
    applied: Optional[_Boundary] = None
    conflicts: List[_Boundary] = []
    for boundary in sorted(boundaries, key=lambda b: (b.top, b.left)):
        if boundary.top <= bbox[1]:
            continue                                     # above the crop top: not in the way
        if min(bbox[2], boundary.right) - max(bbox[0], boundary.left) <= 0:
            continue                                     # horizontally disjoint from the crop
        if boundary.top >= bottom:
            break                                        # sorted: every later boundary is lower still
        limit = boundary.top - guard_y
        if limit < members_bottom:
            conflicts.append(boundary)
            continue
        bottom = limit
        applied = boundary
    return bottom, applied, conflicts


def _separators(slide, roles, items: Sequence[ItemRegion], members: Sequence[PictureRef], W: int, H: int
                ) -> List[BBox]:
    """Structural objects that may separate two members of one block (PROMPT-027R §7).

    The members themselves never separate each other.  Captions and headings that overlap a member are that
    member's own label/annotation and are not separators.  Free text is never a separator.
    """
    member_ids = {id(ref.block) for ref in members}
    member_rects = [_rotated_rect(ref.block) for ref in members]

    def clear_of_members(rect: BBox) -> bool:
        return all(_intersection_area(rect, m) == 0 for m in member_rects)

    out: List[BBox] = []
    for picture in slide.pictures:
        if id(picture) in member_ids or is_decorative_picture(picture, W, H):
            continue
        out.append(_rotated_rect(picture))
    for role in roles:
        if role.role == ROLE_CAPTION and role.block.text.strip():
            rect = _rotated_rect(role.block)
            if clear_of_members(rect):
                out.append(rect)
    for item in items:
        left, top, right, bottom = item.heading_box
        if right > left and bottom > top and clear_of_members((left, top, right, bottom)):
            out.append((left, top, right, bottom))
    return out


def _assign_captions_to_clusters(clusters: Sequence[Sequence[PictureRef]], owner_captions: Sequence[Block],
                                 W: int, H: int, report_name: str, scope_id: str, slide_number: int,
                                 owner_id: str, management_number: str,
                                 obstacles: Optional[Sequence[BBox]] = None) -> List[List[Block]]:
    """PROMPT-025 §14: each of the item's After captions joins the single closest After block
    (caption association is item-scoped; a tie between blocks keeps the caption out)."""
    assignment: List[List[Block]] = [[] for _ in clusters]
    for caption in owner_captions:
        scored: List[Tuple[Tuple[float, int], int]] = []
        for cluster_index, cluster in enumerate(clusters):
            for ref in cluster:
                distance = _caption_distance(ref.block, caption, W, H, obstacles)
                if distance is None:
                    continue
                above_penalty = 0 if caption.bottom <= ref.block.top + ref.block.height * 0.25 else 1
                scored.append(((distance, above_penalty), cluster_index))
        if not scored:
            continue
        scored.sort(key=lambda item: (item[0], item[1]))
        best_score, best_cluster = scored[0]
        if best_score[0] > COLUMN_HEADER_GAP_MAX + 0.10:
            continue
        if (len(scored) > 1 and scored[1][0][1] == best_score[1]
                and scored[1][0][0] - best_score[0] <= 0.02):
            LOG.info("REGION_EXCLUDE MN=%s report=%s scope=%s slide=%s item=%s caption=%r reason=ambiguous-after-block",
                     management_number or "-", report_name, scope_id, slide_number, owner_id,
                     " ".join(caption.text.split())[:80])
            continue
        assignment[best_cluster].append(caption)
    return assignment


def _trace_member(report: ReportData, ref: PictureRef, scope_id: str, owner_id: str, item_index: int,
                  management_number: str) -> None:
    """PROMPT-027R §5: one diagnostic line per member picture (ownership, temporal/semantic state, group)."""
    LOG.info("REGION_MEMBER MN=%s report=%s scope=%s slide=%s item=%s item_index=%d shape=%s source_order=%s "
             "bbox=%s temporal=%s semantic=%s owner=%s confident=%s eligible=%s group=%s",
             management_number or "-", report.filename, scope_id, ref.slide, owner_id, item_index,
             ref.block.shape_id, ref.source_order, _rotated_rect(ref.block), ref.temporal_role, ref.semantic_role,
             ref.owner_id or "-", ref.confident_owner, ref.excel_output_eligible, _group_label(ref.block))


def build_improvement_visual_regions(report: ReportData, refs: Sequence[PictureRef], management_number: str = "",
                                     safe_padding_fraction: float = DEFAULT_SAFE_PADDING_FRACTION,
                                     association_gap_fraction: float = DEFAULT_ASSOCIATION_GAP_FRACTION,
                                     should_cancel: Optional[Callable[[], bool]] = None
                                     ) -> List[ImprovementVisualRegion]:
    """Build one slide-crop region for each eligible (report, logical item, slide, After block) group.

    Pictures are the ownership seeds. Only structurally associated After captions and visual objects are added.
    Before captions, unselected pictures, next-item headings and free body text are excluded. All geometry is PPTX
    EMU and all thresholds are normalized to the source slide dimensions.  ``should_cancel`` is honoured between
    improvement-item groups (PROMPT-024R): each item is an independent unit of work.
    """
    check_cancelled(should_cancel)
    expected_scope = report_scope_key(report.path)
    for ref in refs:
        scope_matches = not ref.report_scope_id or ref.report_scope_id == expected_scope
        if not _eligible(ref) or not scope_matches:
            reason = ("report-scope-mismatch" if not scope_matches else "final-semantic-eligibility-gate")
            LOG.info("REGION_EXCLUDE MN=%s report=%s scope=%s slide=%s item=%s temporal=%s semantic=%s "
                     "confident_owner=%s excel_eligible=%s reason=%s",
                     management_number or "-", report.filename, ref.report_scope_id or "local", ref.slide,
                     ref.owner_id or "unknown", ref.temporal_role, ref.semantic_role, ref.confident_owner,
                     ref.excel_output_eligible, reason)
    scoped_refs = [ref for ref in refs if _eligible(ref)
                   and (not ref.report_scope_id or ref.report_scope_id == expected_scope)]
    grouped: Dict[Tuple[str, str, int], List[PictureRef]] = {}
    for (_scope, owner_id), owner_refs in group_refs(scoped_refs).items():
        for ref in owner_refs:
            scope = ref.report_scope_id or expected_scope
            grouped.setdefault((scope, owner_id, ref.slide), []).append(ref)

    regions: List[ImprovementVisualRegion] = []
    for (scope_id, owner_id, slide_number), pictures in sorted(
            grouped.items(), key=lambda item: (min((r.source_order, r.block.order) for r in item[1]),
                                               item[0][2], item[0][1])):
        # safe boundary: between improvement-item groups (PROMPT-024R §45)
        check_cancelled(should_cancel)
        slide = report.slide(slide_number)
        if slide is None or not pictures:
            continue
        slide_width = int(report.slide_width or slide.width or 1)
        slide_height = int(report.slide_height or slide.height or 1)
        W, H = max(1, slide_width), max(1, slide_height)

        roles = classify_blocks(slide)
        roles_by_z: Dict[int, list] = {}
        for role in roles:
            if role.block.z_order >= 0:
                roles_by_z.setdefault(role.block.z_order, []).append(role)

        # PROMPT-025/027R: the item span (anchor .. next heading) and every other heading bound the crop.
        item_regions = slide_items(slide)
        production = [it for it in item_regions if it.region_kind == "item" and it.semantic_role == SEMANTIC_PRODUCTION]
        owner_item = next((it for it in item_regions if it.owner_id == owner_id), None)
        item_index = next((index for index, it in enumerate(production) if it.owner_id == owner_id), -1)
        span = _item_span(owner_item, item_regions, W, H)
        if owner_item is not None and span is not None:
            LOG.info("ITEM_SPAN MN=%s report=%s scope=%s slide=%s item_index=%d item_id=%s start_y=%d end_y=%d "
                     "start_in=%.3f end_in=%.3f next_item_id=%s next_heading_y=%s boundary_source=%s",
                     management_number or "-", report.filename, scope_id, slide_number, item_index, owner_id,
                     span.top, span.bottom, span.top / 914400.0, span.bottom / 914400.0,
                     span.next_item_id or "-",
                     span.next_heading_top if span.next_heading_top is not None else "-", span.source)
        obstacles_below = _obstacles_below(owner_item, item_regions, roles, W, H)

        for ref in sorted(pictures, key=lambda r: (r.source_order, r.block.order)):
            _trace_member(report, ref, scope_id, owner_id, item_index, management_number)

        # An After button/caption is assigned to the closest final-eligible owner on this slide, not
        # merely copied because it happens to be near a picture. This prevents a neighbouring Before label
        # or another item's After label from being absorbed into the crop.
        claim_obstacles = structural_obstacles(slide, roles, item_regions)
        eligible_on_slide: Dict[str, List[PictureRef]] = {}
        for (other_scope, other_owner, other_slide), owner_refs in grouped.items():
            if other_scope == scope_id and other_slide == slide_number:
                eligible_on_slide.setdefault(other_owner, []).extend(owner_refs)
        caption_owner: Dict[int, str] = {}
        for role in roles:
            if role.role != ROLE_CAPTION or role.caption_kind != "after":
                continue
            owner_distances: Dict[str, Tuple[float, int]] = {}
            for candidate_owner, candidate_refs in eligible_on_slide.items():
                for ref in candidate_refs:
                    distance = _caption_distance(ref.block, role.block, W, H, claim_obstacles)
                    if distance is not None:
                        # Prefer a label just above its photo row over an equally close caption below the
                        # preceding row (common when one defect ends where the next begins).
                        above_penalty = 0 if role.block.bottom <= ref.block.top + ref.block.height * 0.25 else 1
                        score = (distance, above_penalty)
                        owner_distances[candidate_owner] = min(owner_distances.get(candidate_owner, (float("inf"), 1)),
                                                               score)
            if not owner_distances:
                continue
            distances = sorted((score, candidate_owner) for candidate_owner, score in owner_distances.items())
            best_score, best_owner = distances[0]
            if best_score[0] > COLUMN_HEADER_GAP_MAX + 0.10:
                continue
            if (len(distances) > 1 and distances[1][0][1] == best_score[1]
                    and distances[1][0][0] - best_score[0] <= 0.02):
                LOG.info("REGION_EXCLUDE MN=%s report=%s scope=%s slide=%s caption=%r reason=ambiguous-owner",
                         management_number or "-", report.filename, scope_id, slide_number, role.text[:80])
                continue
            caption_owner[id(role.block)] = best_owner

        owner_captions = [role.block for role in roles
                          if role.role == ROLE_CAPTION and role.caption_kind == "after"
                          and caption_owner.get(id(role.block)) == owner_id]

        # PROMPT-025 §18/§19 + PROMPT-027R §7: cluster this item's After pictures into authored After blocks.
        separators = _separators(slide, roles, item_regions, pictures, W, H)
        clusters, join_reasons = _cluster_with_reasons(pictures, owner_captions, W, H, separators, claim_obstacles)
        cluster_captions = _assign_captions_to_clusters(clusters, owner_captions, W, H, report.filename,
                                                        scope_id, slide_number, owner_id, management_number,
                                                        claim_obstacles)
        LOG.info("REGION_CLUSTERS MN=%s report=%s scope=%s slide=%s item=%s item_index=%d clusters=%d "
                 "pictures=%s joins=%s", management_number or "-", report.filename, scope_id, slide_number,
                 owner_id, item_index, len(clusters), [[ref.label for ref in cluster] for cluster in clusters],
                 join_reasons or "-")
        for after_block_index, cluster in enumerate(clusters):
            region = _region_for_cluster(
                report=report, slide=slide, slide_width=slide_width, slide_height=slide_height,
                scope_id=scope_id, owner_id=owner_id, slide_number=slide_number, pictures=cluster,
                captions=cluster_captions[after_block_index], roles=roles, roles_by_z=roles_by_z,
                owner_item=owner_item, span=span, obstacles_below=obstacles_below, items=item_regions,
                item_index=item_index, after_block_index=after_block_index, management_number=management_number,
                safe_padding_fraction=safe_padding_fraction, association_gap_fraction=association_gap_fraction)
            regions.append(region)
    regions.sort(key=lambda r: (r.slide_index, r.item_index if r.item_index >= 0 else 10 ** 6,
                                r.after_block_index, r.source_order))
    return regions


def _region_for_cluster(report: ReportData, slide, slide_width: int, slide_height: int,
                        scope_id: str, owner_id: str, slide_number: int,
                        pictures: Sequence[PictureRef], captions: Sequence[Block],
                        roles, roles_by_z: Dict[int, list], owner_item: Optional[ItemRegion],
                        span: Optional[_Span], obstacles_below: Sequence[_Boundary],
                        items: Sequence[ItemRegion], item_index: int, after_block_index: int,
                        management_number: str, safe_padding_fraction: float, association_gap_fraction: float
                        ) -> ImprovementVisualRegion:
    """Build ONE item-scoped After region for one clustered After block of one improvement item."""
    W, H = max(1, slide_width), max(1, slide_height)
    picture_ids = {id(ref.block) for ref in pictures}
    picture_boxes = [_rotated_rect(ref.block) for ref in pictures]
    other_pictures = [pic for pic in slide.pictures if id(pic) not in picture_ids]

    caption_boxes = [_rotated_rect(caption) for caption in captions]
    core = _union(picture_boxes + caption_boxes)
    core_area = max(1, _area(core))
    slide_area = W * H
    association_x = max(1, int(W * max(0.0, association_gap_fraction)))
    association_y = max(1, int(H * max(0.0, association_gap_fraction)))
    expanded_core = _expand(core, association_x, association_y)

    before_caption_z = {role.block.z_order for role in roles
                        if role.role == ROLE_CAPTION and role.caption_kind == "before"}
    selected_caption_z = {caption.z_order for caption in captions if caption.z_order >= 0}
    selected_caption_ids = {id(caption) for caption in captions}
    arrow_z = {arrow.z_order for arrow in slide.arrows if arrow.z_order >= 0 and arrow.direction}
    own_span = (owner_item.top, owner_item.bottom) if owner_item is not None else None
    foreign_spans = [(it.top, it.bottom) for it in items
                     if it is not owner_item and it.region_kind == "item" and it.semantic_role == SEMANTIC_PRODUCTION]
    foreign_headings = [it.heading_box for it in items if it is not owner_item]

    included: List[Block] = []
    excluded: List[str] = [
        f"picture SH{pic.shape_id} bbox={_rotated_rect(pic)} reason=not-owned-by-this-item"
        for pic in other_pictures
    ]

    def decide(obj: Block, is_text: bool) -> Tuple[bool, str]:
        """Membership of one non-picture object (PROMPT-027R §6/§8/§11).

        Ownership first: an object whose centre lies in another item's span, or inside another item's heading, is
        that item's.  Then visual containment: a shape/text must lie mostly within this After footprint (or overlay
        one of its pictures).  A body paragraph that merely grazes the footprint is excluded, never absorbed.
        """
        box = _rotated_rect(obj)
        cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
        in_own = own_span is not None and own_span[0] <= cy < own_span[1]
        in_foreign = any(lo <= cy < hi for lo, hi in foreign_spans)
        if in_foreign and not in_own:
            return False, "other-item-span"
        if any(left <= cx < right and top <= cy < bottom for left, top, right, bottom in foreign_headings):
            return False, "next-item-heading"
        if obj.line_endpoints or _is_arrow(obj, arrow_z):
            if _object_intersects_core(obj, core, expanded_core, arrow_z, picture_boxes):
                return True, ""
            return False, ("outside-after-core-transition-or-decoration" if _is_arrow(obj, arrow_z)
                           else "outside-after-core")
        area = _area(box)
        containment = _intersection_area(box, expanded_core) / max(1, area)
        if is_text:
            if containment < TEXT_CONTAINMENT_MIN:
                return False, "text-outside-visual-footprint"
        else:
            overlays = any(_intersection_area(box, picture) / max(1, area) >= OVERLAY_FRACTION
                           for picture in picture_boxes)
            if containment < CONTAINMENT_MIN and not overlays:
                return False, "shape-outside-visual-footprint"
        other_centres = _centres_inside(box, other_pictures)
        if other_centres:
            return False, ("text-over-nonselected-picture" if is_text else "spans-nonselected-picture")
        if area > max(core_area * 6, int(slide_area * 0.40)):
            return False, ("oversized-neighbouring-text" if is_text else "oversized-neighbouring-background")
        return True, ""

    # Every child yielded from a PPTX group retains absolute geometry and z-order in annotations.
    for obj in slide.annotations:
        obj_role = _role_for_block(obj, roles_by_z)
        if obj.z_order in before_caption_z or (obj_role is not None and obj_role.role == ROLE_CAPTION
                                               and obj_role.caption_kind == "before"):
            excluded.append(_describe(obj) + " reason=before-caption")
            continue
        if obj_role is not None and obj_role.role == ROLE_CAPTION and obj_role.caption_kind == "after":
            if obj_role.block.z_order in selected_caption_z:
                included.append(obj)
            else:
                excluded.append(_describe(obj) + " reason=other-after-owner")
            continue
        visible = bool(obj.fill_visible or obj.line_visible or (obj.text or "").strip())
        if not visible:
            continue
        if obj_role is not None and obj_role.role in (ROLE_TITLE, ROLE_SIDEBAR, ROLE_FURNITURE):
            excluded.append(_describe(obj) + f" reason={obj_role.role}")
            continue
        ok, why = decide(obj, is_text=bool((obj.text or "").strip()))
        if ok:
            included.append(obj)
        elif _is_arrow(obj, arrow_z) or _intersection_area(_rotated_rect(obj), expanded_core) > 0:
            excluded.append(_describe(obj) + f" reason={why}")     # transitions are always reported (§16)

    # Text boxes are semantic blocks, so include them separately from vector styling records. Captions other than
    # the selected After caption, headings, page furniture and side labels never trigger a crop.
    for block in slide.blocks:
        if block.kind in ("picture", "visual") or not block.is_text:
            continue
        role = _role_for_block(block, roles_by_z)
        if role is not None and role.role == ROLE_CAPTION:
            if role.caption_kind == "after" and id(block) in selected_caption_ids:
                included.append(block)
            elif role.caption_kind == "before":
                excluded.append(_describe(block) + " reason=before-caption")
            continue
        if role is not None and role.role in (ROLE_TITLE, ROLE_SIDEBAR, ROLE_FURNITURE):
            continue
        ok, why = decide(block, is_text=True)
        if ok:
            included.append(block)
        elif _intersection_area(_rotated_rect(block), expanded_core) > 0:
            excluded.append(_describe(block) + f" reason={why}")

    # De-duplicate the style record and its text block for diagnostics/bounds while keeping both when
    # they provide different geometry (e.g. a caption inside a rounded button).
    unique: List[Block] = []
    seen = set()
    for obj in included:
        key = (obj.kind, obj.z_order, obj.shape_id, obj.left, obj.top, obj.width, obj.height, obj.text)
        if key not in seen:
            seen.add(key)
            unique.append(obj)
    included = unique

    bounds = picture_boxes + caption_boxes + [_rotated_rect(obj) for obj in included]
    unpadded = _union(bounds)
    max_stroke = max((int(obj.line_width or 0) for obj in included), default=0)
    pad_x = max(int(W * max(0.0, safe_padding_fraction)), max_stroke // 2)
    pad_y = max(int(H * max(0.0, safe_padding_fraction)), max_stroke // 2)
    included_ids = {id(obj) for obj in included}
    obstacle_boxes = [_rotated_rect(pic) for pic in other_pictures]
    for obj in [*slide.annotations, *slide.blocks]:
        if obj.kind == "picture" or id(obj) in included_ids:
            continue
        if obj.fill_visible or obj.line_visible or (obj.text or "").strip():
            obstacle_boxes.append(_rotated_rect(obj))
    padding = _safe_padding(unpadded, obstacle_boxes, pad_x, pad_y, W, H)
    padded = (unpadded[0] - padding[0], unpadded[1] - padding[1],
              unpadded[2] + padding[2], unpadded[3] + padding[3])
    bbox = (max(0, padded[0]), max(0, padded[1]), min(slide_width, padded[2]), min(slide_height, padded[3]))

    # PROMPT-027R §10/§11: the final crop is clamped by every heading below it that it horizontally reaches.
    # Members (pictures and their captions) always stay inside: a boundary that would cut one is a reported
    # conflict, never a silent loss of evidence.
    members = picture_boxes + caption_boxes
    members_top = min(box[1] for box in members)
    members_bottom = max(box[3] for box in members)
    guard_y = max(1, int(math.ceil(H / 1080)))
    bottom, applied, conflicts = _clamp_bottom(bbox, members_bottom, obstacles_below, guard_y)
    top = bbox[1]
    if span is not None:
        top = max(top, span.top)                         # never above this item's own heading
    if top > members_top:
        top = members_top
    clamped = (bbox[0], top, bbox[2], bottom)
    boundary_source = applied.source if applied is not None else (span.source if span is not None else "")
    boundary_item = applied.item_id if applied is not None else (span.next_item_id if span is not None else "")
    boundary_top = applied.top if applied is not None else (span.next_heading_top if span is not None else None)
    if clamped != bbox:
        LOG.info("REGION_CLAMP MN=%s report=%s scope=%s slide=%s item=%s block=%d raw_bbox=%s bbox=%s "
                 "raw_bottom=%d final_bottom=%d boundary_source=%s next_heading_y=%s",
                 management_number or "-", report.filename, scope_id, slide_number, owner_id, after_block_index,
                 bbox, clamped, bbox[3], bottom, boundary_source, boundary_top if boundary_top is not None else "-")
    for conflict in conflicts:
        LOG.warning("REGION_BOUNDARY_CONFLICT MN=%s report=%s scope=%s slide=%s item=%s block=%d "
                    "boundary_source=%s boundary_y=%d members_bottom=%d (members kept, crop not clamped there)",
                    management_number or "-", report.filename, scope_id, slide_number, owner_id,
                    after_block_index, conflict.source, conflict.top, members_bottom)
    bbox = clamped

    heading = next((ref.owner_heading for ref in pictures if ref.owner_heading), owner_id)
    confidences = [ref.confidence for ref in pictures if ref.confidence is not None]
    confidence = min(confidences) if confidences else None
    anchor_text = " | ".join(dict.fromkeys(" ".join(c.text.split()) for c in captions if c.text.strip()))
    region = ImprovementVisualRegion(
        report_scope_id=scope_id, management_number=management_number, slide_index=slide_number,
        improvement_item_id=owner_id, improvement_heading=heading,
        semantic_role=SEMANTIC_PRODUCTION, temporal_role="AFTER", bbox=bbox,
        source_order=min((ref.source_order for ref in pictures), default=0), confidence=confidence,
        excel_output_eligible=True, anchor_text=anchor_text, pictures=list(pictures),
        included_objects=included, excluded_objects=excluded,
        item_index=item_index, after_block_index=after_block_index,
        span_top=span.top if span is not None else 0, span_bottom=span.bottom if span is not None else 0,
        boundary_source=boundary_source, boundary_item_id=boundary_item, boundary_top=boundary_top,
        raw_bbox=unpadded, padding=tuple(int(v) for v in padding),
    )
    LOG.info("REGION_ITEM MN=%s report=%s scope=%s slide=%s item=%s item_index=%d block=%d heading=%r "
             "semantic=%s temporal=%s eligible=%s confidence=%s pictures=%s",
             management_number or "-", report.filename, scope_id, slide_number, owner_id, item_index,
             after_block_index, heading, region.semantic_role, region.temporal_role,
             region.excel_output_eligible, region.confidence, [ref.label for ref in pictures])
    LOG.info("REGION_ANCHOR MN=%s report=%s scope=%s slide=%s item=%s anchor=%r",
             management_number or "-", report.filename, scope_id, slide_number, owner_id,
             anchor_text or "(none)")
    LOG.info("REGION_EVIDENCE MN=%s report=%s scope=%s slide=%s item=%s block=%d picture_ids=%s "
             "region_bbox=%s region_bottom_y=%d boundary_source=%s next_heading_y=%s",
             management_number or "-", report.filename, scope_id, slide_number, owner_id, after_block_index,
             region.picture_shape_ids, bbox, bbox[3], boundary_source,
             boundary_top if boundary_top is not None else "-")
    for obj in included:
        LOG.info("REGION_OBJECT MN=%s report=%s scope=%s slide=%s item=%s included=%s",
                 management_number or "-", report.filename, scope_id, slide_number, owner_id, _describe(obj))
    for detail in excluded:
        LOG.info("REGION_EXCLUDE MN=%s report=%s scope=%s slide=%s item=%s object=%s",
                 management_number or "-", report.filename, scope_id, slide_number, owner_id, detail)
    LOG.info("REGION_BBOX MN=%s report=%s scope=%s slide=%s item=%s bbox=%s unpadded=%s "
             "padding=(left=%s,top=%s,right=%s,bottom=%s) objects=%s",
             management_number or "-", report.filename, scope_id, slide_number, owner_id, bbox, unpadded,
             *padding, len(included))
    return region


def crop_box_px(bbox: BBox, slide_width: int, slide_height: int,
                image_width: int, image_height: int) -> Tuple[int, int, int, int]:
    """Convert an EMU bbox to an outward-rounded, slide-clamped rendered-pixel rectangle."""
    if min(slide_width, slide_height, image_width, image_height) <= 0:
        raise ValueError("slide and rendered image dimensions must be positive")
    sx, sy = image_width / slide_width, image_height / slide_height
    left = max(0, min(image_width, int(math.floor(bbox[0] * sx))))
    top = max(0, min(image_height, int(math.floor(bbox[1] * sy))))
    right = max(0, min(image_width, int(math.ceil(bbox[2] * sx))))
    bottom = max(0, min(image_height, int(math.ceil(bbox[3] * sy))))
    if right <= left or bottom <= top:
        raise ValueError(f"visual region has empty pixel crop {bbox}")
    return left, top, right, bottom
