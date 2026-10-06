"""Owner-scoped After visual regions and their rendered-slide crop geometry.

Semantic picture selection remains in :mod:`improvement_pictures` / image learning. This module only
builds final visual evidence from refs that have already passed the strict Excel eligibility gate.
"""
from __future__ import annotations

import logging
import math
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from .content_region import (ROLE_CAPTION, ROLE_FURNITURE, ROLE_SIDEBAR, ROLE_TITLE, classify_blocks)
from .improvement_pictures import (CAPTION_GAP_MAX, CAPTION_SIDE_GAP_MAX, COLUMN_HEADER_GAP_MAX,
                                   SEMANTIC_PRODUCTION, PictureRef, group_refs)
from .pptx_parser import Block, ReportData
from .report_identity import report_scope_key

LOG = logging.getLogger("report_extractor.images")
BBox = Tuple[int, int, int, int]  # left, top, right, bottom; PPTX EMU coordinates

# Padding and object association are relative to the slide, not report-specific coordinates.
DEFAULT_SAFE_PADDING_FRACTION = 0.008
DEFAULT_ASSOCIATION_GAP_FRACTION = 0.006


@dataclass
class ImprovementVisualRegion:
    """One logically owned After group, rendered and cropped from one source slide."""

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

    @property
    def width(self) -> int:
        return max(0, self.bbox[2] - self.bbox[0])

    @property
    def height(self) -> int:
        return max(0, self.bbox[3] - self.bbox[1])

    @property
    def region_id(self) -> str:
        return f"{self.report_scope_id}:S{self.slide_index}:{self.improvement_item_id}"


def _eligible(ref: PictureRef) -> bool:
    """Reassert PROMPT-015 at the visual-evidence boundary; labels never bypass this gate."""
    return bool(ref.excel_output_eligible and ref.temporal_role == "AFTER"
                and ref.semantic_role == SEMANTIC_PRODUCTION and ref.confident_owner and ref.owner_id)


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


def _caption_distance(picture: Block, caption: Block, slide_width: int, slide_height: int) -> Optional[float]:
    """Same relative above/below/side association used by deterministic Before/After selection."""
    p, c = _rotated_rect(picture), _rotated_rect(caption)
    h_overlap, v_overlap = _overlap_x(p, c), _overlap_y(p, c)
    if h_overlap and v_overlap:
        return 0.0
    min_width = max(1, min(p[2] - p[0], c[2] - c[0]))
    min_height = max(1, min(p[3] - p[1], c[3] - c[1]))
    v_gap = _gap_y(p, c) / max(1, slide_height)
    h_gap = _gap_x(p, c) / max(1, slide_width)
    if h_overlap >= 0.5 * min_width:
        if v_gap <= CAPTION_GAP_MAX:
            return v_gap
        if c[3] <= p[1] and v_gap <= COLUMN_HEADER_GAP_MAX:
            return v_gap + 0.10
    if v_overlap >= 0.5 * min_height and h_gap <= CAPTION_SIDE_GAP_MAX:
        return h_gap + 0.02
    return None


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


def build_improvement_visual_regions(report: ReportData, refs: Sequence[PictureRef], management_number: str = "",
                                     safe_padding_fraction: float = DEFAULT_SAFE_PADDING_FRACTION,
                                     association_gap_fraction: float = DEFAULT_ASSOCIATION_GAP_FRACTION
                                     ) -> List[ImprovementVisualRegion]:
    """Build one slide-crop region for each eligible (report, logical item, slide) group.

    Pictures are the ownership seeds. Only geometrically associated After captions and visual objects are
    added. Before captions, unselected pictures, broad background shapes, and center transition arrows are
    excluded. All geometry is PPTX EMU and all thresholds are normalized to the source slide dimensions.
    """
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
        slide = report.slide(slide_number)
        if slide is None or not pictures:
            continue
        slide_width = int(report.slide_width or slide.width or 1)
        slide_height = int(report.slide_height or slide.height or 1)
        W, H = max(1, slide_width), max(1, slide_height)
        picture_ids = {id(ref.block) for ref in pictures}
        picture_boxes = [_rotated_rect(ref.block) for ref in pictures]
        other_pictures = [pic for pic in slide.pictures if id(pic) not in picture_ids]

        roles = classify_blocks(slide)
        roles_by_z: Dict[int, list] = {}
        for role in roles:
            if role.block.z_order >= 0:
                roles_by_z.setdefault(role.block.z_order, []).append(role)

        # An After button/caption is assigned to the closest final-eligible owner on this slide, not
        # merely copied because it happens to be near a picture. This prevents a neighbouring Before label
        # or another item's After label from being absorbed into the crop.
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
                    distance = _caption_distance(ref.block, role.block, W, H)
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

        captions = [role.block for role in roles
                    if role.role == ROLE_CAPTION and role.caption_kind == "after"
                    and caption_owner.get(id(role.block)) == owner_id]
        core_boxes = picture_boxes + [_rotated_rect(caption) for caption in captions]
        core = _union(core_boxes)
        association_x = max(1, int(W * max(0.0, association_gap_fraction)))
        association_y = max(1, int(H * max(0.0, association_gap_fraction)))
        expanded_core = _expand(core, association_x, association_y)

        before_caption_z = {role.block.z_order for role in roles
                            if role.role == ROLE_CAPTION and role.caption_kind == "before"}
        selected_caption_z = {caption.z_order for caption in captions if caption.z_order >= 0}
        arrow_z = {arrow.z_order for arrow in slide.arrows if arrow.z_order >= 0 and arrow.direction}
        included: List[Block] = []
        excluded: List[str] = [
            f"picture SH{pic.shape_id} bbox={_rotated_rect(pic)} reason=not-owned-by-this-item"
            for pic in other_pictures
        ]

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
            if not _object_intersects_core(obj, core, expanded_core, arrow_z, picture_boxes):
                if _is_arrow(obj, arrow_z):
                    excluded.append(_describe(obj) + " reason=outside-after-core-transition-or-decoration")
                continue
            obj_box = _rotated_rect(obj)
            obj_area = _area(obj_box)
            core_area = max(1, _area(core))
            slide_area = W * H
            other_centres = _centres_inside(obj_box, other_pictures)
            if other_centres and (obj_area > core_area * 1.5 or len(other_centres) > 0):
                excluded.append(_describe(obj) + " reason=spans-nonselected-picture")
                continue
            if obj_area > max(core_area * 6, int(slide_area * 0.40)) and not obj.line_endpoints:
                excluded.append(_describe(obj) + " reason=oversized-neighbouring-background")
                continue
            included.append(obj)

        # Text boxes are semantic blocks, so include them separately from vector styling records. Captions
        # other than the selected After caption, headings, page furniture and side labels never trigger a crop.
        selected_caption_ids = {id(caption) for caption in captions}
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
            if not _object_intersects_core(block, core, expanded_core, arrow_z, picture_boxes):
                continue
            block_box = _rotated_rect(block)
            block_area = _area(block_box)
            if any(_centres_inside(block_box, other_pictures)):
                excluded.append(_describe(block) + " reason=text-over-nonselected-picture")
                continue
            if block_area > max(_area(core) * 6, int(slide_area * 0.40)):
                excluded.append(_describe(block) + " reason=oversized-neighbouring-text")
                continue
            included.append(block)

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

        bounds = picture_boxes + [_rotated_rect(obj) for obj in included]
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
        )
        regions.append(region)
        LOG.info("REGION_ITEM MN=%s report=%s scope=%s slide=%s item=%s heading=%r semantic=%s temporal=%s "
                 "eligible=%s confidence=%s pictures=%s",
                 management_number or "-", report.filename, scope_id, slide_number, owner_id, heading,
                 region.semantic_role, region.temporal_role, region.excel_output_eligible, region.confidence,
                 [ref.label for ref in pictures])
        LOG.info("REGION_ANCHOR MN=%s report=%s scope=%s slide=%s item=%s anchor=%r",
                 management_number or "-", report.filename, scope_id, slide_number, owner_id,
                 anchor_text or "(none)")
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
    return regions


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
