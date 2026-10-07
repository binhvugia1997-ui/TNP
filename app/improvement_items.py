"""PROMPT-025 — explicit logical ImprovementItem segmentation for multi-item slides.

A production-improvement slide may carry 0, 1, 2, 3 … logical improvement items
(Report -> Slide[] -> ImprovementItem[] -> Before/After regions -> evidence).
The historical extractor derived item identity only from detected heading-like
lines; when a later item's heading escaped that detection the earlier item
absorbed the whole slide and the After evidence was cropped at slide scope.

This module turns the anchor/spans produced by :func:`app.improvement_pictures.slide_items`
(whose heading detection was extended for the real multi-item heading styles) into an
explicit, item-scoped model:

    ImprovementItem
        report identity / slide / stable item index + id
        item bbox / vertical span
        heading (anchor text)
        improvement text lines (item-scoped, never the next item's text)
        Before / After visual candidates
        associated captions
        semantic classification + diagnostics

Segmentation happens BEFORE final evidence cropping; every later stage
(temporal classification, ownership, region construction, rendering, Excel
mapping, learning identity) consumes these items.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from .content_region import ROLE_CAPTION, ROLE_CONTENT, classify_blocks
from .improvement_pictures import (SEMANTIC_PRODUCTION, ItemRegion, PictureRef, PictureSelection,
                                   is_decorative_picture, slide_items)
from .pptx_parser import Block, ReportData, SlideData, clean_text
from .report_identity import report_scope_key

LOG = logging.getLogger("report_extractor.items")

BBox = Tuple[int, int, int, int]  # left, top, width, height (PPTX EMU)


@dataclass
class ImprovementItem:
    """One logical production-improvement item on one slide (PROMPT-025 §5)."""

    report_scope_id: str
    slide: int
    index: int                          # stable 0-based item index within the slide
    item_id: str                        # report-local logical identity (== PictureRef.owner_id)
    heading: str                        # anchor/heading text
    bounds: BBox                        # item vertical span: anchor .. next item/section boundary
    region: Optional[ItemRegion] = field(default=None, repr=False)   # source anchor/span (single source of truth)
    text_lines: List[str] = field(default_factory=list)    # item-scoped improvement text (heading first)
    text_sources: List[int] = field(default_factory=list)  # shape_id per text line
    picture_candidates: List[Block] = field(default_factory=list)   # content pictures in the span
    before_pictures: List[PictureRef] = field(default_factory=list)  # filled from the selection
    after_pictures: List[PictureRef] = field(default_factory=list)   # filled from the selection
    captions: List[Block] = field(default_factory=list)             # caption buttons in the span
    after_blocks: int = 0               # clustered After visual groups (filled by the region builder)
    semantic_role: str = SEMANTIC_PRODUCTION
    confident: bool = True
    source_order: int = 0
    diagnostics: List[str] = field(default_factory=list)

    # ---------------------------------------------------------------- geometry
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

    @property
    def scoped_item_id(self) -> str:
        """Full report-scoped identity: scope | slide | item (PROMPT-020 isolation)."""
        return f"{self.report_scope_id}|S{self.slide}|{self.item_id}"

    @property
    def text(self) -> str:
        return clean_text("\n".join(self.text_lines))

    def contains_y(self, y: float) -> bool:
        return self.top <= y < self.bottom

    # ---------------------------------------------------------------- filling
    def attach_selection(self, sel: PictureSelection) -> None:
        """Attach the deterministic picture decisions that belong to this item."""
        self.after_pictures = [ref for ref in sel.after
                               if ref.slide == self.slide and ref.owner_id == self.item_id]
        self.before_pictures = [ref for ref in sel.rejected
                                if ref.slide == self.slide and ref.owner_id == self.item_id
                                and ref.kind == "before"]
        self.diagnostics.append(
            f"after_pictures={len(self.after_pictures)} before_pictures={len(self.before_pictures)}")

    def summary(self) -> dict:
        """JSON-serializable item summary (diagnostics / bridge DTO / batch_result.json)."""
        return {
            "slide": self.slide,
            "index": self.index,
            "itemId": self.item_id,
            "scopedItemId": self.scoped_item_id,
            "heading": self.heading,
            "text": self.text,
            "bounds": {"x": self.left, "y": self.top, "w": self.width, "h": self.height},
            "beforePictures": [ref.label for ref in self.before_pictures],
            "afterPictures": [ref.label for ref in self.after_pictures],
            "captions": [" ".join(c.text.split()) for c in self.captions if c.text.strip()],
            "afterBlocks": self.after_blocks,
            "semantic": self.semantic_role,
            "confident": self.confident,
        }


def production_item_regions(items: Sequence[ItemRegion]) -> List[ItemRegion]:
    """Logical production-improvement items of a slide, in stable reading order."""
    return sorted((it for it in items
                   if it.region_kind == "item" and it.semantic_role == SEMANTIC_PRODUCTION),
                  key=lambda it: (it.source_order, it.top, it.left))


def item_region_for_owner(items: Sequence[ItemRegion], owner_id: str) -> Optional[ItemRegion]:
    return next((it for it in items if it.owner_id == owner_id), None)


def item_index_for_owner(items: Sequence[ItemRegion], owner_id: str) -> int:
    for index, it in enumerate(production_item_regions(items)):
        if it.owner_id == owner_id:
            return index
    return -1


def segment_improvement_items(report: ReportData, slide: SlideData,
                              items: Optional[Sequence[ItemRegion]] = None,
                              roles=None) -> List[ImprovementItem]:
    """Segment one slide into logical ImprovementItems (PROMPT-025 §6).

    1. anchors/spans come from :func:`slide_items` (extended heading detection);
    2. every content line is assigned to the item whose vertical span contains it,
       so an item never absorbs the next item's heading/body text;
    3. content pictures and caption buttons are assigned by span containment.

    No OCR, no pixel analysis, no screenshot input — PPTX structure only.
    """
    if items is None:
        items = slide_items(slide)
    if roles is None:
        roles = classify_blocks(slide)
    W, H = slide.width or 1, slide.height or 1
    scope_id = report_scope_key(report.path)
    prod = production_item_regions(items)

    out: List[ImprovementItem] = []
    for index, region in enumerate(prod):
        item = ImprovementItem(
            report_scope_id=scope_id, slide=slide.number, index=index, item_id=region.owner_id,
            heading=region.heading, bounds=region.bounds, region=region,
            semantic_role=region.semantic_role, confident=region.confident_defect_ownership,
            source_order=region.source_order)

        # --- item-scoped text: content lines whose line box lies inside the span ---
        for role in roles:
            if role.role != ROLE_CONTENT:
                continue
            block = role.block
            lines = block.text.split("\n")
            n_lines = max(1, len(lines))
            for i, raw in enumerate(lines):
                text = raw.strip()
                if not text:
                    continue
                line_top = int(block.top + i * (block.height / n_lines)) if block.height else block.top
                if not item.contains_y(line_top):
                    continue
                item.text_lines.append(raw.rstrip())
                item.text_sources.append(block.shape_id)

        # --- visual candidates + captions by span containment ---
        for picture in slide.pictures:
            if is_decorative_picture(picture, W, H):
                continue
            if item.contains_y(picture.top + picture.height / 2):
                item.picture_candidates.append(picture)
        for role in roles:
            if role.role == ROLE_CAPTION and role.block.text.strip() \
                    and item.contains_y(role.block.top + role.block.height / 2):
                item.captions.append(role.block)

        item.diagnostics.append(
            f"span=({item.left},{item.top},{item.width},{item.height}) "
            f"text_lines={len(item.text_lines)} picture_candidates={len(item.picture_candidates)} "
            f"captions={len(item.captions)}")
        out.append(item)
    return out


def segment_report_items(report: ReportData, slide_numbers: Sequence[int],
                         log: bool = True) -> Dict[int, List[ImprovementItem]]:
    """Segment every given slide of a report; logs the per-slide item count (PROMPT-025 §27)."""
    out: Dict[int, List[ImprovementItem]] = {}
    for n in sorted(set(int(x) for x in slide_numbers)):
        slide = report.slide(n)
        if slide is None:
            continue
        items = segment_improvement_items(report, slide)
        out[n] = items
        if log:
            LOG.info("ITEM_SEGMENTATION report=%s scope=%s slide=%s items=%d",
                     report.filename, report_scope_key(report.path), n, len(items))
            for item in items:
                LOG.info("ITEM_DETAIL report=%s scope=%s slide=%s index=%d item=%s heading=%r "
                         "span=(%d,%d,%d,%d) %s",
                         report.filename, item.report_scope_id, n, item.index, item.item_id,
                         item.heading, item.left, item.top, item.width, item.height,
                         "; ".join(item.diagnostics))
    return out


def attach_selection(items_by_slide: Dict[int, List[ImprovementItem]], sel: PictureSelection) -> None:
    """Fill the Before/After picture refs of every item from a finished selection."""
    for slide_number, items in items_by_slide.items():
        for item in items:
            item.attach_selection(sel)
