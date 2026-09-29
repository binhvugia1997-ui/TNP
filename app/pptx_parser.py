"""Read PowerPoint structure with python-pptx.

Everything extracted here is the *original* content of the report.  The parser
never rewrites text; it only orders it (top-to-bottom, left-to-right) so the
reading order matches the slide layout.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterator, List, Optional, Tuple

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

EMU_PER_INCH = 914400


# ----------------------------------------------------------------------------
# Data structures
# ----------------------------------------------------------------------------
@dataclass
class TextRunInfo:
    text: str
    bold: bool = False
    size_pt: Optional[float] = None


@dataclass
class Block:
    """One logical piece of slide content (paragraph, table row, picture)."""
    kind: str                          # "paragraph" | "table" | "picture" | "title"
    text: str = ""                     # original text ("" for pictures)
    left: int = 0                      # EMU
    top: int = 0
    width: int = 0
    height: int = 0
    shape_id: int = 0
    shape_name: str = ""
    order: int = 0                     # reading order inside slide
    bold: bool = False
    size_pt: Optional[float] = None
    rows: List[List[str]] = field(default_factory=list)   # for tables
    image_blob: Optional[bytes] = None                    # for pictures
    image_ext: str = ""
    level: int = 0                                        # paragraph indent level

    @property
    def is_text(self) -> bool:
        return self.kind in ("paragraph", "table", "title") and bool(self.text.strip())

    @property
    def bottom(self) -> int:
        return self.top + self.height

    @property
    def right(self) -> int:
        return self.left + self.width


@dataclass
class SlideData:
    number: int                        # 1-based
    blocks: List[Block] = field(default_factory=list)
    width: int = 0
    height: int = 0
    notes: str = ""

    @property
    def text_blocks(self) -> List[Block]:
        return [b for b in self.blocks if b.is_text]

    @property
    def pictures(self) -> List[Block]:
        return [b for b in self.blocks if b.kind == "picture"]

    @property
    def text(self) -> str:
        """All text of the slide in reading order (original wording)."""
        return "\n".join(b.text for b in self.text_blocks)

    @property
    def title(self) -> str:
        for b in self.blocks:
            if b.kind == "title" and b.text.strip():
                return b.text.strip()
        tb = self.text_blocks
        return tb[0].text.strip().splitlines()[0] if tb else ""


@dataclass
class ReportData:
    path: Path
    slides: List[SlideData] = field(default_factory=list)
    slide_width: int = 0
    slide_height: int = 0

    def slide(self, number: int) -> Optional[SlideData]:
        if 1 <= number <= len(self.slides):
            return self.slides[number - 1]
        return None

    @property
    def filename(self) -> str:
        return self.path.name

    def all_text(self) -> str:
        return "\n".join(s.text for s in self.slides)


# ----------------------------------------------------------------------------
# Text helpers
# ----------------------------------------------------------------------------
def strip_accents(s: str) -> str:
    nfkd = unicodedata.normalize("NFD", s)
    out = "".join(ch for ch in nfkd if unicodedata.category(ch) != "Mn")
    return out.replace("đ", "d").replace("Đ", "D")


def norm_key(s: str) -> str:
    """Accent-insensitive, case-insensitive, whitespace-collapsed key."""
    s = strip_accents(s or "").lower()
    s = re.sub(r"[\s_\-–—:.,;()\[\]/\\]+", " ", s)
    return s.strip()


def clean_text(s: str) -> str:
    """Normalize only whitespace/line breaks; keep every word untouched."""
    if not s:
        return ""
    s = s.replace("\x0b", "\n").replace("\r", "\n")
    s = unicodedata.normalize("NFC", s)
    lines = [re.sub(r"[ \t\u00a0]+", " ", ln).rstrip() for ln in s.split("\n")]
    # collapse 3+ blank lines but keep intentional single blank lines
    out: List[str] = []
    blank = 0
    for ln in lines:
        if ln.strip():
            blank = 0
            out.append(ln)
        else:
            blank += 1
            if blank <= 1:
                out.append("")
    return "\n".join(out).strip("\n")


# ----------------------------------------------------------------------------
# Parsing
# ----------------------------------------------------------------------------
def _paragraph_text(paragraph) -> str:
    # python-pptx paragraph.text turns vertical tabs into "\v" for line breaks
    return "".join(r.text for r in paragraph.runs) if paragraph.runs else paragraph.text


def _iter_shapes(shapes, offset=(0, 0)) -> Iterator[Tuple[Any, Tuple[int, int]]]:
    """Yield (shape, absolute_offset) flattening groups."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            # child coordinates are relative to group's child offset; python-pptx
            # already stores absolute-like coordinates for most decks, so we keep
            # a simple approach: use child coordinates directly.
            yield from _iter_shapes(shape.shapes, offset)
        else:
            yield shape, offset


def _frame_blocks(shape, offset, is_title: bool) -> List[Block]:
    blocks: List[Block] = []
    tf = shape.text_frame
    left = int(shape.left or 0) + offset[0]
    top = int(shape.top or 0) + offset[1]
    width = int(shape.width or 0)
    height = int(shape.height or 0)
    lines: List[Tuple[str, int, bool, Optional[float]]] = []
    for i, p in enumerate(tf.paragraphs):
        txt = clean_text(_paragraph_text(p))
        bold = any(bool(r.font.bold) for r in p.runs) if p.runs else False
        size = None
        for r in p.runs:
            if r.font.size is not None:
                size = r.font.size.pt
                break
        lines.append((txt, p.level or 0, bold, size))
    # Keep a text frame as ONE block (so multi-line content stays together) but
    # remember paragraph metadata for heading detection of the first line.
    text = "\n".join(t for t, *_ in lines)
    if not text.strip():
        return blocks
    b = Block(
        kind="title" if is_title else "paragraph",
        text=text,
        left=left,
        top=top,
        width=width,
        height=height,
        shape_id=shape.shape_id,
        shape_name=shape.name,
        bold=lines[0][2] if lines else False,
        size_pt=lines[0][3] if lines else None,
        level=lines[0][1] if lines else 0,
    )
    blocks.append(b)
    return blocks


def _table_block(shape, offset) -> Optional[Block]:
    tbl = shape.table
    rows: List[List[str]] = []
    for r in tbl.rows:
        cells: List[str] = []
        prev = None
        for c in r.cells:
            t = clean_text(c.text_frame.text if c.text_frame is not None else c.text)
            # merged cells repeat their text in python-pptx spans; avoid duplicates
            if c.is_spanned or (prev is not None and t == prev and t):
                t = "" if c.is_spanned else t
            cells.append(t)
            prev = t
        rows.append(cells)
    lines = []
    for cells in rows:
        non_empty = [c for c in cells if c]
        if non_empty:
            lines.append(" | ".join(non_empty))
    if not lines:
        return None
    return Block(
        kind="table",
        text="\n".join(lines),
        rows=rows,
        left=int(shape.left or 0) + offset[0],
        top=int(shape.top or 0) + offset[1],
        width=int(shape.width or 0),
        height=int(shape.height or 0),
        shape_id=shape.shape_id,
        shape_name=shape.name,
    )


def _picture_block(shape, offset) -> Optional[Block]:
    try:
        image = shape.image
        blob = image.blob
        ext = image.ext
    except Exception:
        return None
    return Block(
        kind="picture",
        left=int(shape.left or 0) + offset[0],
        top=int(shape.top or 0) + offset[1],
        width=int(shape.width or 0),
        height=int(shape.height or 0),
        shape_id=shape.shape_id,
        shape_name=shape.name,
        image_blob=blob,
        image_ext=ext,
    )


def reading_order(blocks: List[Block], slide_height: int) -> List[Block]:
    """Sort blocks top-to-bottom, then left-to-right, using row bands.

    Blocks whose vertical centre falls in the same band (~4% of slide height)
    are considered the same row and ordered by ``left``.
    """
    band = max(1, int(slide_height * 0.04)) if slide_height else 200000
    def key(b: Block):
        return (round(b.top / band), b.left, b.top)
    ordered = sorted(blocks, key=key)
    for i, b in enumerate(ordered):
        b.order = i
    return ordered


def parse_pptx(path: str | Path) -> ReportData:
    """Parse a .pptx into :class:`ReportData` (read-only, never modifies file)."""
    path = Path(path)
    prs = Presentation(str(path))
    report = ReportData(path=path, slide_width=int(prs.slide_width or 0),
                        slide_height=int(prs.slide_height or 0))
    for idx, slide in enumerate(prs.slides, start=1):
        sd = SlideData(number=idx, width=report.slide_width, height=report.slide_height)
        blocks: List[Block] = []
        title_shape_id = None
        try:
            if slide.shapes.title is not None:
                title_shape_id = slide.shapes.title.shape_id
        except Exception:
            title_shape_id = None
        for shape, offset in _iter_shapes(slide.shapes):
            try:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or shape.shape_type == MSO_SHAPE_TYPE.LINKED_PICTURE:
                    pb = _picture_block(shape, offset)
                    if pb:
                        blocks.append(pb)
                    continue
                if getattr(shape, "has_table", False) and shape.has_table:
                    tb = _table_block(shape, offset)
                    if tb:
                        blocks.append(tb)
                    continue
                if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                    blocks.extend(_frame_blocks(shape, offset, shape.shape_id == title_shape_id))
                    continue
                # placeholder pictures / OLE objects with an image
                if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER and hasattr(shape, "image"):
                    pb = _picture_block(shape, offset)
                    if pb:
                        blocks.append(pb)
            except Exception:
                # a single broken shape must not break the slide
                continue
        sd.blocks = reading_order(blocks, report.slide_height)
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
                sd.notes = clean_text(slide.notes_slide.notes_text_frame.text)
        except Exception:
            pass
        report.slides.append(sd)
    return report


def slide_text_for_llm(slide: SlideData, max_chars: int = 1800) -> str:
    """Compact but faithful text for classification prompts (truncated only for size)."""
    parts = []
    for b in slide.text_blocks:
        t = b.text.strip()
        if b.kind == "table":
            t = "[TABLE]\n" + t
        parts.append(t)
    txt = "\n".join(parts)
    if slide.pictures:
        txt += f"\n[{len(slide.pictures)} picture(s)]"
    if len(txt) > max_chars:
        txt = txt[:max_chars] + " …[truncated]"
    return txt
