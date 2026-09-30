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
    alt_text: str = ""                                    # cNvPr/@descr or @title (pictures, OLE objects)
    origin: str = "shape"                                 # "shape" | "ole" | "alternate_content"
    prst: str = ""                                        # preset geometry (rect, ellipse, roundRect, rightArrow…)
    rotation: float = 0.0                                 # degrees
    n_lines: int = 0                                      # paragraphs in the text frame

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
    xml_stats: dict = field(default_factory=dict)          # raw-XML inventory (diagnostics)

    @property
    def alt_texts(self) -> List[str]:
        return [b.alt_text for b in self.blocks if b.alt_text]

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


def _shape_geometry(shape) -> Tuple[str, float]:
    """(preset geometry name, rotation in degrees) – best effort."""
    prst = ""
    try:
        g = shape._element.xpath(".//a:prstGeom")
        if g:
            prst = g[0].get("prst", "") or ""
    except Exception:
        pass
    rot = 0.0
    try:
        rot = float(shape.rotation or 0.0)
    except Exception:
        pass
    return prst, rot


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
        n_lines=len(lines),
    )
    b.prst, b.rotation = _shape_geometry(shape)
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
    descr = ""
    try:
        cnv = shape._element.xpath(".//p:cNvPr")
        if cnv:
            descr = " ".join(filter(None, [cnv[0].get("descr", ""), cnv[0].get("title", "")]))
    except Exception:
        pass
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
        alt_text=descr.strip(),
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


# ----------------------------------------------------------------------------
# Raw-XML sweep: content python-pptx does not expose
# ----------------------------------------------------------------------------
_NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
_R_EMBED = "{%s}embed" % _NS["r"]


def _xfrm(elm) -> Tuple[int, int, int, int]:
    """(left, top, width, height) in EMU from the nearest a:xfrm / p:xfrm."""
    for x in elm.iter():
        if x.tag in ("{%s}xfrm" % _NS["p"], "{%s}xfrm" % _NS["a"]):
            off = x.find("a:off", _NS)
            ext = x.find("a:ext", _NS)
            if off is not None and ext is not None:
                return (int(off.get("x", 0)), int(off.get("y", 0)), int(ext.get("cx", 0)), int(ext.get("cy", 0)))
    return (0, 0, 0, 0)


def _cnvpr(elm):
    for x in elm.iter("{%s}cNvPr" % _NS["p"]):
        return x
    return None


def _blob(slide_part, rid: str):
    try:
        part = slide_part.related_part(rid)
        return part.blob, (part.partname.ext or "").lstrip(".").lower()
    except Exception:
        return None, ""


def xml_inventory(slide) -> dict:
    """Counts of the XML constructs that matter for diagnosis."""
    root = slide._element
    def count(tag_ns, tag):
        return sum(1 for _ in root.iter("{%s}%s" % (_NS[tag_ns], tag)))
    return {
        "sp": count("p", "sp"), "pic": count("p", "pic"), "grpSp": count("p", "grpSp"),
        "graphicFrame": count("p", "graphicFrame"), "oleObj": count("p", "oleObj"),
        "AlternateContent": count("mc", "AlternateContent"), "tbl": count("a", "tbl"),
        "a_t_runs": count("a", "t"),
    }


def raw_xml_text(slide) -> str:
    """Every a:t run in the slide XML (document order) – used to detect text the shape API missed."""
    return clean_text("\n".join((t.text or "") for t in slide._element.iter("{%s}t" % _NS["a"])))


def sweep_unexposed_shapes(slide, seen_ids: set) -> List[Block]:
    """Blocks for content python-pptx does not yield:

    * OLE objects (embedded Excel/Word forms such as a Quality Problem Notice):
      their preview picture (p:oleObj/p:pic) becomes a picture block;
    * shapes inside mc:AlternateContent (Choice or Fallback) – text and pictures;
    * alt text (descr/title) of any picture/OLE object.
    """
    blocks: List[Block] = []
    root = slide._element
    part = slide.part

    # --- OLE objects -----------------------------------------------------------
    for gf in root.iter("{%s}graphicFrame" % _NS["p"]):
        ole = next(iter(gf.iter("{%s}oleObj" % _NS["p"])), None)
        if ole is None:
            continue
        cnv = _cnvpr(gf)
        sid = int(cnv.get("id", 0)) if cnv is not None else 0
        key = ("ole", sid)
        if key in seen_ids:
            continue
        seen_ids.add(key)
        left, top, width, height = _xfrm(gf)
        name = (cnv.get("name", "") if cnv is not None else "") or ""
        descr = " ".join(filter(None, [cnv.get("descr", "") if cnv is not None else "",
                                       cnv.get("title", "") if cnv is not None else "",
                                       ole.get("name", ""), ole.get("progId", "")]))
        blip = next(iter(ole.iter("{%s}blip" % _NS["a"])), None)
        blob, ext = (None, "")
        if blip is not None and blip.get(_R_EMBED):
            blob, ext = _blob(part, blip.get(_R_EMBED))
        blocks.append(Block(kind="picture", left=left, top=top, width=width, height=height, shape_id=sid,
                            shape_name=f"OLE:{name}", image_blob=blob, image_ext=ext, alt_text=descr.strip(),
                            origin="ole"))
        # text inside the OLE fallback (rare) – capture as paragraph
        txt = clean_text("\n".join((t.text or "") for t in gf.iter("{%s}t" % _NS["a"])))
        if txt:
            blocks.append(Block(kind="paragraph", text=txt, left=left, top=top, width=width, height=height,
                                shape_id=sid, shape_name=f"OLE-text:{name}", origin="ole"))

    # --- shapes hidden inside mc:AlternateContent ---------------------------------
    for ac in root.iter("{%s}AlternateContent" % _NS["mc"]):
        for elm in ac.iter():
            tag = elm.tag
            if tag == "{%s}sp" % _NS["p"]:
                cnv = _cnvpr(elm)
                sid = int(cnv.get("id", 0)) if cnv is not None else 0
                if ("sp", sid) in seen_ids:
                    continue
                seen_ids.add(("sp", sid))
                txt = clean_text("\n".join((t.text or "") for t in elm.iter("{%s}t" % _NS["a"])))
                if txt:
                    left, top, width, height = _xfrm(elm)
                    blocks.append(Block(kind="paragraph", text=txt, left=left, top=top, width=width, height=height,
                                        shape_id=sid, shape_name=(cnv.get("name", "") if cnv is not None else ""),
                                        origin="alternate_content"))
            elif tag == "{%s}pic" % _NS["p"] and not any(a.tag == "{%s}oleObj" % _NS["p"] for a in elm.iterancestors()):
                cnv = _cnvpr(elm)
                sid = int(cnv.get("id", 0)) if cnv is not None else 0
                if ("pic", sid) in seen_ids:
                    continue
                seen_ids.add(("pic", sid))
                blip = next(iter(elm.iter("{%s}blip" % _NS["a"])), None)
                blob, ext = (None, "")
                if blip is not None and blip.get(_R_EMBED):
                    blob, ext = _blob(part, blip.get(_R_EMBED))
                left, top, width, height = _xfrm(elm)
                blocks.append(Block(kind="picture", left=left, top=top, width=width, height=height, shape_id=sid,
                                    shape_name=(cnv.get("name", "") if cnv is not None else ""), image_blob=blob,
                                    image_ext=ext, alt_text=(cnv.get("descr", "") if cnv is not None else ""),
                                    origin="alternate_content"))
            elif tag == "{%s}graphicFrame" % _NS["p"] and any(True for _ in elm.iter("{%s}tbl" % _NS["a"])):
                cnv = _cnvpr(elm)
                sid = int(cnv.get("id", 0)) if cnv is not None else 0
                if ("tbl", sid) in seen_ids:
                    continue
                seen_ids.add(("tbl", sid))
                rows: List[List[str]] = []
                for tr in elm.iter("{%s}tr" % _NS["a"]):
                    rows.append([clean_text("\n".join((t.text or "") for t in tc.iter("{%s}t" % _NS["a"])))
                                 for tc in tr.iter("{%s}tc" % _NS["a"])])
                lines = [" | ".join(c for c in r if c) for r in rows if any(r)]
                if lines:
                    left, top, width, height = _xfrm(elm)
                    blocks.append(Block(kind="table", text="\n".join(lines), rows=rows, left=left, top=top,
                                        width=width, height=height, shape_id=sid, origin="alternate_content"))
    return blocks


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
        # content python-pptx does not expose (OLE objects, mc:AlternateContent)
        seen = set()
        for b in blocks:
            seen.add(("pic" if b.kind == "picture" else ("tbl" if b.kind == "table" else "sp"), b.shape_id))
        try:
            blocks.extend(sweep_unexposed_shapes(slide, seen))
        except Exception:
            pass
        try:
            sd.xml_stats = xml_inventory(slide)
        except Exception:
            sd.xml_stats = {}
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
