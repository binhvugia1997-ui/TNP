"""Read PowerPoint structure with python-pptx.

Everything extracted here is the *original* content of the report.  The parser
never rewrites text; it only orders it (top-to-bottom, left-to-right) so the
reading order matches the slide layout.
"""
from __future__ import annotations

import logging
import math
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterator, List, Optional, Tuple

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

LOG = logging.getLogger("report_extractor.pptx_parser")
EMU_PER_INCH = 914400
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


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
    """One logical piece of slide content or render-only geometry (paragraph, table row, picture, visual shape)."""
    kind: str                          # paragraph | table | picture | title | visual
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
    line_colors: List[str] = field(default_factory=list)  # '#rrggbb' per paragraph ('' = inherited/unknown)
    direction: str = ""                                   # arrows only: "right" | "left" | "up" | "down" | ""
    z_order: int = -1                                      # original drawing order, independent of reading order
    fill_color: str = ""
    line_color: str = ""
    fill_visible: bool = False
    line_visible: bool = False
    line_width: int = 0                                    # EMU
    line_endpoints: Optional[Tuple[int, int, int, int]] = None  # absolute EMU for straight connectors
    shape_type: str = ""                                  # python-pptx shape kind for visual rendering/diagnostics

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
    arrows: List[Block] = field(default_factory=list)      # directional arrow shapes/connectors (structure only,
                                                           # never content; kept out of ``blocks`` on purpose)
    annotations: List[Block] = field(default_factory=list)  # non-picture visual shapes and text with z-order/geometry

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


class GroupXform:
    """Affine map from a group's CHILD coordinate space to absolute slide EMU.

    DrawingML stores every child of ``p:grpSp`` in the group's child space (``a:chOff`` / ``a:chExt``) which is
    mapped onto the group's own ``a:off`` / ``a:ext``.  python-pptx returns the raw child values, so once a group
    has been moved or resized (``chOff != off`` – the normal case in real decks) the raw numbers point at the
    wrong place on the slide.  Nested groups compose their transforms.
    """
    __slots__ = ("ox", "oy", "sx", "sy", "chx", "chy")

    def __init__(self, ox=0, oy=0, sx=1.0, sy=1.0, chx=0, chy=0):
        self.ox, self.oy, self.sx, self.sy, self.chx, self.chy = ox, oy, sx, sy, chx, chy

    IDENTITY: "GroupXform"

    def child(self, group) -> "GroupXform":
        """Transform for the children of ``group`` (itself positioned through *self*)."""
        try:
            xfrm = group._element.find("{%s}grpSpPr/{%s}xfrm" % (_NS_P, _NS_A))
            off = xfrm.find("{%s}off" % _NS_A)
            ext = xfrm.find("{%s}ext" % _NS_A)
            ch_off = xfrm.find("{%s}chOff" % _NS_A)
            ch_ext = xfrm.find("{%s}chExt" % _NS_A)
            gx, gy = int(off.get("x", 0)), int(off.get("y", 0))
            gw, gh = int(ext.get("cx", 0)), int(ext.get("cy", 0))
            chx = int(ch_off.get("x", 0)) if ch_off is not None else gx
            chy = int(ch_off.get("y", 0)) if ch_off is not None else gy
            chw = int(ch_ext.get("cx", 0)) if ch_ext is not None else gw
            chh = int(ch_ext.get("cy", 0)) if ch_ext is not None else gh
        except Exception:
            return self
        sx = (gw / chw) if (chw and gw) else 1.0
        sy = (gh / chh) if (chh and gh) else 1.0
        # the group's own box is expressed in the parent's space -> map it first
        ax, ay, aw, ah = self.apply(gx, gy, gw, gh)
        sx *= (aw / gw) if gw else 1.0
        sy *= (ah / gh) if gh else 1.0
        return GroupXform(ax, ay, sx, sy, chx, chy)

    def apply(self, left, top, width, height) -> Tuple[int, int, int, int]:
        left, top, width, height = int(left or 0), int(top or 0), int(width or 0), int(height or 0)
        return (int(round(self.ox + (left - self.chx) * self.sx)), int(round(self.oy + (top - self.chy) * self.sy)),
                int(round(width * self.sx)), int(round(height * self.sy)))


GroupXform.IDENTITY = GroupXform()


def shape_geometry(shape, offset) -> Tuple[int, int, int, int]:
    """Absolute (left, top, width, height) of ``shape``; ``offset`` is a :class:`GroupXform` or a legacy (dx, dy)."""
    if isinstance(offset, GroupXform):
        return offset.apply(shape.left, shape.top, shape.width, shape.height)
    dx, dy = offset
    return int(shape.left or 0) + dx, int(shape.top or 0) + dy, int(shape.width or 0), int(shape.height or 0)


def _iter_shapes(shapes, offset=None) -> Iterator[Tuple[Any, GroupXform]]:
    """Yield (shape, GroupXform) flattening groups recursively (nested groups compose their transforms)."""
    xf = offset if isinstance(offset, GroupXform) else GroupXform.IDENTITY
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(shape.shapes, xf.child(shape))
        else:
            yield shape, xf


_THEME_SLOTS = {"TEXT_1": "dk1", "DARK_1": "dk1", "BACKGROUND_1": "lt1", "LIGHT_1": "lt1", "TEXT_2": "dk2",
                "DARK_2": "dk2", "BACKGROUND_2": "lt2", "LIGHT_2": "lt2", "ACCENT_1": "accent1", "ACCENT_2": "accent2",
                "ACCENT_3": "accent3", "ACCENT_4": "accent4", "ACCENT_5": "accent5", "ACCENT_6": "accent6",
                "HYPERLINK": "hlink", "FOLLOWED_HYPERLINK": "folHlink"}


def theme_color_map(prs) -> dict:
    """{'accent1': '#4f81bd', 'dk1': '#000000', …} from the first slide master's theme (best effort)."""
    out: dict = {}
    try:
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        theme_part = prs.slide_masters[0].part.part_related_by(RT.THEME)
        from lxml import etree
        root = etree.fromstring(theme_part.blob)
        scheme = root.find(".//{%s}clrScheme" % _NS_A)
        for child in (scheme if scheme is not None else []):
            slot = etree.QName(child).localname
            srgb = child.find("{%s}srgbClr" % _NS_A)
            sysc = child.find("{%s}sysClr" % _NS_A)
            val = srgb.get("val") if srgb is not None else (sysc.get("lastClr") if sysc is not None else None)
            if val:
                out[slot] = "#" + val.lower()
    except Exception:
        pass
    return out


def _run_color(run, theme: dict) -> str:
    """'#rrggbb' of a run's font colour (RGB or theme-resolved), '' when inherited/unknown."""
    try:
        color = run.font.color
        if color is None or color.type is None:
            return ""
        from pptx.enum.dml import MSO_COLOR_TYPE
        if color.type == MSO_COLOR_TYPE.RGB:
            return "#" + str(color.rgb).lower()
        if color.type == MSO_COLOR_TYPE.SCHEME:
            name = str(color.theme_color).split(".")[-1].split(" ")[0]
            return theme.get(_THEME_SLOTS.get(name, ""), "")
    except Exception:
        return ""
    return ""


def _paragraph_color(paragraph, theme: dict) -> str:
    """Colour covering the majority of the paragraph's characters ('' when inherited)."""
    weights: dict = {}
    total = 0
    for r in paragraph.runs:
        n = len((r.text or "").strip())
        if not n:
            continue
        total += n
        weights[_run_color(r, theme)] = weights.get(_run_color(r, theme), 0) + n
    if not total:
        return ""
    color, n = max(weights.items(), key=lambda kv: kv[1])
    return color if n * 2 >= total else ""


_ARROW_BASE = {"rightArrow": "right", "leftArrow": "left", "upArrow": "up", "downArrow": "down",
               "notchedRightArrow": "right", "stripedRightArrow": "right", "homePlate": "right", "chevron": "right",
               "bentUpArrow": "up", "uturnArrow": "", "leftRightArrow": "", "upDownArrow": "", "quadArrow": "",
               "leftRightUpArrow": "", "curvedRightArrow": "right", "curvedLeftArrow": "left",
               "curvedUpArrow": "up", "curvedDownArrow": "down", "swooshArrow": "right", "circularArrow": ""}
_DIRS = ("right", "down", "left", "up")           # clockwise rotation order


def _rotate_dir(d: str, rotation: float, flip_h: bool, flip_v: bool) -> str:
    if not d:
        return ""
    if flip_h and d in ("left", "right"):
        d = "left" if d == "right" else "right"
    if flip_v and d in ("up", "down"):
        d = "up" if d == "down" else "down"
    steps = int(round((rotation % 360) / 90.0)) % 4
    return _DIRS[(_DIRS.index(d) + steps) % 4]


def arrow_block(shape, offset, prst: str, rotation: float) -> Optional[Block]:
    """A directional arrow (auto-shape or connector with an arrow head) as a structural Block(kind='arrow')."""
    try:
        elm = shape._element
        xfrm = elm.find(".//{%s}xfrm" % _NS_A)
        flip_h = bool(xfrm is not None and xfrm.get("flipH") in ("1", "true"))
        flip_v = bool(xfrm is not None and xfrm.get("flipV") in ("1", "true"))
        direction = ""
        if prst in _ARROW_BASE:
            direction = _rotate_dir(_ARROW_BASE[prst], rotation, flip_h, flip_v)
        elif "Connector" in prst or prst == "line" or etree_localname(elm) == "cxnSp":
            ln = elm.find(".//{%s}ln" % _NS_A)
            head = ln.find("{%s}headEnd" % _NS_A) if ln is not None else None
            tail = ln.find("{%s}tailEnd" % _NS_A) if ln is not None else None
            head_arrow = head is not None and head.get("type", "none") not in ("none", None)
            tail_arrow = tail is not None and tail.get("type", "none") not in ("none", None)
            if head_arrow == tail_arrow:
                return None                                   # no head or double-headed: not directional
            w, h = int(shape.width or 0), int(shape.height or 0)
            if w >= h:
                d = "left" if flip_h else "right"             # start→end runs left→right unless flipped
            else:
                d = "up" if flip_v else "down"
            if head_arrow:                                    # arrow head at the START point
                d = {"right": "left", "left": "right", "up": "down", "down": "up"}[d]
            direction = _rotate_dir(d, rotation, False, False)
        else:
            return None
        a_left, a_top, a_w, a_h = shape_geometry(shape, offset)
        b = Block(kind="arrow", left=a_left, top=a_top, width=a_w, height=a_h, shape_id=shape.shape_id,
                  shape_name=shape.name, prst=prst, rotation=rotation, direction=direction)
        return b
    except Exception:
        return None


def etree_localname(elm) -> str:
    tag = elm.tag
    return tag.split("}", 1)[1] if "}" in tag else tag


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


def _color_hex(color, theme: dict) -> str:
    """Resolve an Office RGB/theme color to '#rrggbb' when possible."""
    try:
        from pptx.enum.dml import MSO_COLOR_TYPE
        if color.type == MSO_COLOR_TYPE.RGB:
            return "#" + str(color.rgb).lower()
        if color.type == MSO_COLOR_TYPE.SCHEME:
            name = str(color.theme_color).split(".")[-1].split(" ")[0]
            return theme.get(_THEME_SLOTS.get(name, ""), "")
    except Exception:
        pass
    return ""


def _line_endpoints(shape, offset, left: int, top: int, width: int, height: int,
                    rotation: float, prst: str) -> Optional[Tuple[int, int, int, int]]:
    """Absolute endpoints of a straight PPTX line/connector, or None for area annotations."""
    try:
        is_line = prst == "line" or etree_localname(shape._element) == "cxnSp" \
            or shape.shape_type == MSO_SHAPE_TYPE.LINE
        if not is_line:
            return None
        xfrm = shape._element.find(".//{%s}xfrm" % _NS_A)
        flip_h = bool(xfrm is not None and xfrm.get("flipH") in ("1", "true"))
        flip_v = bool(xfrm is not None and xfrm.get("flipV") in ("1", "true"))
        diagonal = flip_h ^ flip_v
        if diagonal:
            p1, p2 = (left, top + height), (left + width, top)
        else:
            p1, p2 = (left, top), (left + width, top + height)
        cx, cy = left + width / 2.0, top + height / 2.0
        theta = math.radians(rotation)
        cosine, sine = math.cos(theta), math.sin(theta)
        def rotate(point):
            dx, dy = point[0] - cx, point[1] - cy
            return (int(round(cx + dx * cosine - dy * sine)),
                    int(round(cy + dx * sine + dy * cosine)))
        p1, p2 = rotate(p1), rotate(p2)
        return p1[0], p1[1], p2[0], p2[1]
    except Exception:
        return None


def _style_reference_color(shape, ref_name: str, theme: dict) -> Tuple[bool, str]:
    """Best-effort resolution of an inherited Office theme fill/line reference."""
    try:
        ref = shape._element.find(".//{%s}%s" % (_NS_A, ref_name))
        if ref is None or int(ref.get("idx", "0")) <= 0:
            return False, ""
        srgb = ref.find("{%s}srgbClr" % _NS_A)
        scheme = ref.find("{%s}schemeClr" % _NS_A)
        if srgb is not None and srgb.get("val"):
            return True, "#" + srgb.get("val").lower()
        if scheme is not None and scheme.get("val"):
            name = scheme.get("val")
            return True, theme.get(name, theme.get(_THEME_SLOTS.get(name.upper(), ""), ""))
    except Exception:
        pass
    return False, ""


def _annotation_block(shape, offset, prst: str, rotation: float, z_order: int, theme: dict) -> Block:
    """Geometry/style record for one non-picture visual object; excluded from semantic text classification."""
    left, top, width, height = shape_geometry(shape, offset)
    fill_color, fill_visible = "", False
    line_color, line_visible, line_width = "", False, 0
    visual_text = ""
    try:
        visual_text = clean_text(shape.text_frame.text) if shape.has_text_frame else ""
    except Exception:
        pass
    try:
        from pptx.enum.dml import MSO_FILL
        fill = shape.fill
        fill_visible = fill.type is not None and fill.type != MSO_FILL.BACKGROUND
        if fill_visible:
            fill_color = _color_hex(fill.fore_color, theme)
        if not fill_visible:
            sp_pr = shape._element.find("{%s}spPr" % _NS_P)
            explicit_fill = (sp_pr is not None and any(sp_pr.find("{%s}%s" % (_NS_A, tag)) is not None
                                                       for tag in ("solidFill", "noFill", "gradFill", "blipFill",
                                                                   "pattFill", "grpFill")))
            if not explicit_fill:
                fill_visible, fill_color = _style_reference_color(shape, "fillRef", theme)
    except Exception:
        pass
    try:
        line = shape.line
        line_visible = line.fill.type is not None
        line_color = _color_hex(line.color, theme) if line_visible else ""
        line_width = int(line.width or 0) if line_visible else 0
        sp_pr = shape._element.find("{%s}spPr" % _NS_P)
        line_element = sp_pr.find("{%s}ln" % _NS_A) if sp_pr is not None else None
        explicit_no_line = (line_element is not None and line_element.find("{%s}noFill" % _NS_A) is not None)
        if not line_visible and not explicit_no_line:
            line_visible, line_color = _style_reference_color(shape, "lnRef", theme)
            line_width = 12700 if line_visible else 0
    except Exception:
        pass
    return Block(
        kind="visual", text=visual_text, left=left, top=top, width=width, height=height,
        shape_id=getattr(shape, "shape_id", 0), shape_name=getattr(shape, "name", ""),
        prst=prst, rotation=rotation, z_order=z_order, fill_color=fill_color, line_color=line_color,
        fill_visible=fill_visible, line_visible=line_visible, line_width=line_width,
        line_endpoints=_line_endpoints(shape, offset, left, top, width, height, rotation, prst),
        shape_type=str(getattr(shape, "shape_type", "")),
    )


def _frame_blocks(shape, offset, is_title: bool, theme: Optional[dict] = None) -> List[Block]:
    blocks: List[Block] = []
    theme = theme or {}
    tf = shape.text_frame
    colors: List[str] = []
    left, top, width, height = shape_geometry(shape, offset)
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
        colors.append(_paragraph_color(p, theme))
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
        line_colors=colors,
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
    left, top, width, height = shape_geometry(shape, offset)
    return Block(
        kind="table",
        text="\n".join(lines),
        rows=rows,
        left=left,
        top=top,
        width=width,
        height=height,
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
    left, top, width, height = shape_geometry(shape, offset)
    return Block(
        kind="picture",
        left=left,
        top=top,
        width=width,
        height=height,
        shape_id=shape.shape_id,
        shape_name=shape.name,
        image_blob=blob,
        image_ext=ext,
        alt_text=descr.strip(),
    )


def _dedupe_shape_ids(blocks: List[Block], slide_no: int = 0) -> None:
    """Make ``shape_id`` unique per slide for text/table/picture blocks.

    Copy-pasted (grouped) shapes may carry the same ``cNvPr id``; downstream, sections, candidates and learning
    labels are keyed by ``(slide, shape_id)``, so a collision silently attaches one shape's evidence to another
    shape's text.  The first occurrence keeps its id; later ones get fresh ids above the slide maximum.
    """
    used = set()
    next_id = max([b.shape_id for b in blocks] + [0]) + 1
    for b in blocks:
        if b.kind == "arrow":
            continue
        if b.shape_id in used:
            LOG.debug("slide %s: duplicate shape id %s (%r) -> %s", slide_no, b.shape_id, b.text[:30], next_id)
            b.shape_id = next_id
            next_id += 1
        used.add(b.shape_id)


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
    theme = theme_color_map(prs)
    for idx, slide in enumerate(prs.slides, start=1):
        sd = SlideData(number=idx, width=report.slide_width, height=report.slide_height)
        blocks: List[Block] = []
        arrows: List[Block] = []
        annotations: List[Block] = []
        title_shape_id = None
        try:
            if slide.shapes.title is not None:
                title_shape_id = slide.shapes.title.shape_id
        except Exception:
            title_shape_id = None
        for z_order, (shape, offset) in enumerate(_iter_shapes(slide.shapes)):
            try:
                is_picture = (shape.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE)
                              or (shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER and hasattr(shape, "image")))
                if is_picture:
                    pb = _picture_block(shape, offset)
                    if pb:
                        pb.z_order = z_order
                        blocks.append(pb)
                    continue
                prst, rot = _shape_geometry(shape)
                is_table = getattr(shape, "has_table", False) and shape.has_table
                visual = _annotation_block(shape, offset, prst, rot, z_order, theme)
                annotations.append(visual)
                if is_table:
                    tb = _table_block(shape, offset)
                    if tb:
                        tb.z_order = z_order
                        visual.text = tb.text
                        blocks.append(tb)
                    continue
                if prst in _ARROW_BASE or "Connector" in prst or prst == "line" or \
                        etree_localname(shape._element) == "cxnSp":
                    ab = arrow_block(shape, offset, prst, rot)
                    if ab is not None:
                        ab.z_order = z_order
                        visual.direction = ab.direction
                        arrows.append(ab)
                if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                    text_blocks = _frame_blocks(shape, offset, shape.shape_id == title_shape_id, theme)
                    for text_block in text_blocks:
                        text_block.z_order = z_order
                    blocks.extend(text_blocks)
                    continue
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
        _dedupe_shape_ids(blocks, idx)
        sd.blocks = reading_order(blocks, report.slide_height)
        sd.arrows = sorted(arrows, key=lambda a: (a.top, a.left))
        sd.annotations = annotations
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
