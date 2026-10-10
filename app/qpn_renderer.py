"""Render a complete PPTX slide (e.g. the whole Quality Problem Notice page) to PNG.

Renderer chain (first one that works wins):
  1. Microsoft PowerPoint COM automation (Windows, PowerPoint installed)
  2. LibreOffice headless  (pptx -> pdf -> png via PyMuPDF if available)
  3. Built-in Pillow renderer (draws pictures, common vector shapes, tables and text boxes
     from PPTX geometry). Always available; effects, SmartArt, complex freeforms, theme transforms,
     advanced typography and some grouped-shape details remain lower fidelity than Office renderers.

Nothing here is sent to the cloud.  Source PPTX files are opened read-only.
"""
from __future__ import annotations

import logging
import math
import os
import shutil
import subprocess
import sys
import tempfile
from io import BytesIO
from pathlib import Path
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from PIL import Image, ImageDraw, ImageFont

from .cancellation import CancellationRequested, check_cancelled
from .qpn_region import QpnRegion, crop_box_px, locate_qpn_region
from .pptx_parser import Block, ReportData, SlideData

LOG = logging.getLogger("report_extractor.renderer")
EMU_PER_INCH = 914400
EMU_PER_PT = 12700

# PROMPT-027: bump whenever the rendered full-slide bitmap's GEOMETRY or paint semantics change, so a
# cached preview produced by older (wrong) rendering can never be reused (§29).
RENDER_SCHEMA_VERSION = 3

# Fallback paint used ONLY for a slot that is genuinely visible but whose colour could not be resolved
# from the PPTX/theme.  Never applied to a slot the author hid (No Fill / No Line) — §7/§9.
DEFAULT_SHAPE_FILL = "#4f81bd"
DEFAULT_SHAPE_LINE = "#000000"
DEFAULT_LINE_WIDTH_EMU = 12700
DEFAULT_TEXT_COLOR = "black"

# Text-frame layout defaults (ECMA-376).  Calibrated so 1pt of authored text maps to the same visual
# size PowerPoint uses at the rendered scale; MIN_TEXT_PX bounds the shrink-to-fit fallback.
DEFAULT_TEXT_PT = 14.0
DEFAULT_LINE_SPACING = 1.2
TEXT_SIZE_CALIBRATION = 1.05
MIN_TEXT_PX = 10


# ----------------------------------------------------------------------------
# Availability checks
# ----------------------------------------------------------------------------
POWERPOINT_PROGID = "PowerPoint.Application"


def _powerpoint_progid_registered(pythoncom_module) -> Tuple[bool, str]:
    """Is PowerPoint's COM class registered?  COM's own lookup first (the same registry COM itself reads, no process
    is started), then the explicit HKCR probe in the default, 64-bit and 32-bit registry views (§15: a single brittle
    probe must never mark an installed PowerPoint as missing)."""
    try:
        pythoncom_module.CLSIDFromProgID(POWERPOINT_PROGID)
        return True, "progid-com"
    except Exception as exc:  # noqa: BLE001 – a missing ProgID is the normal negative answer here
        com_reason = type(exc).__name__
    try:
        import winreg  # type: ignore
        for access in (0, getattr(winreg, "KEY_WOW64_64KEY", 0), getattr(winreg, "KEY_WOW64_32KEY", 0)):
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, POWERPOINT_PROGID, 0, winreg.KEY_READ | access):
                    return True, "progid-registry"
            except OSError:
                continue
    except Exception:  # noqa: BLE001 – registry access denied / unavailable: fall back to the COM answer
        pass
    return False, f"progid-not-registered ({com_reason})"


def powerpoint_probe() -> Tuple[bool, str, str]:
    """(available, stage, detail).  Side-effect free: it never starts PowerPoint.  ``stage`` names the first check that
    failed (``availability`` / ``import``); ``detail`` is a safe, path-free explanation for diagnostics."""
    if sys.platform != "win32":
        return False, "availability", "not-windows"
    try:
        import pythoncom  # type: ignore  # pywin32
    except Exception as exc:  # noqa: BLE001
        return False, "import", f"pythoncom {type(exc).__name__}"
    try:
        import win32com.client  # type: ignore  # noqa: F401
    except Exception as exc:  # noqa: BLE001
        return False, "import", f"win32com.client {type(exc).__name__}"
    registered, how = _powerpoint_progid_registered(pythoncom)
    if not registered:
        return False, "availability", how
    return True, "", how


def powerpoint_available() -> bool:
    """True when the PowerPoint COM backend can be attempted on this machine (no cached answer: a later install or a
    repaired COM registration is seen immediately, PROMPT-027R §15)."""
    return bool(powerpoint_probe()[0])


def powerpoint_unavailable_detail() -> str:
    return powerpoint_probe()[2]


def find_soffice() -> Optional[str]:
    for name in ("soffice", "soffice.exe", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    if sys.platform == "win32":
        for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                     os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
            cand = Path(base) / "LibreOffice" / "program" / "soffice.exe"
            if cand.exists():
                return str(cand)
    return None


def pymupdf_available() -> bool:
    try:
        import fitz  # noqa: F401
        return True
    except Exception:
        return False


def renderer_status() -> Dict[str, str]:
    return {
        "powerpoint": "OK" if powerpoint_available() else "Unavailable",
        "libreoffice": "OK" if (find_soffice() and pymupdf_available()) else "Unavailable",
        "builtin": "OK",
    }


# PROMPT-027 §5/§6: only Microsoft PowerPoint COM exports the COMPLETE authored slide pixel-faithfully
# (no trimming, no repositioning, no whitespace normalization).  LibreOffice and the built-in renderer
# are honest fallbacks: they must be reported as reduced fidelity, never presented as PowerPoint output.
FAITHFUL_BACKENDS = frozenset({"powerpoint"})


def is_faithful_backend(backend: str) -> bool:
    """True when ``backend`` reproduces the authored slide faithfully (drives the UI fidelity notice)."""
    return str(backend or "") in FAITHFUL_BACKENDS


def _safe_reason(report: ReportData, exc: BaseException) -> str:
    """Concise, log-safe failure reason: exception class + message with local paths removed (§3/§42).

    Renderer errors routinely embed absolute Windows paths from COM/LibreOffice.  Diagnostics must stay
    useful without leaking the tester's filesystem into logs that can reach the UI layer.
    """
    text = str(exc).strip()
    for secret in (str(getattr(report, "path", "") or ""), getattr(report, "filename", "") or ""):
        if secret:
            text = text.replace(secret, "<report>")
    text = " ".join(text.split())
    if len(text) > 180:
        text = text[:177] + "..."
    return f"{type(exc).__name__}: {text}" if text else type(exc).__name__


# ----------------------------------------------------------------------------
# Fonts (Unicode Vietnamese capable)
# ----------------------------------------------------------------------------
_FONT_CACHE: Dict[Tuple[str, int], ImageFont.FreeTypeFont] = {}


def _font_candidates() -> List[str]:
    c: List[str] = []
    from .runtime_paths import resource_root
    c.append(str(resource_root() / "assets" / "DejaVuSans.ttf"))
    if sys.platform == "win32":
        win = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
        c += [str(win / n) for n in ("arial.ttf", "segoeui.ttf", "tahoma.ttf", "calibri.ttf", "times.ttf")]
    c += ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
          "/usr/share/fonts/dejavu/DejaVuSans.ttf",
          "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
          "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
          "/System/Library/Fonts/Supplemental/Arial.ttf"]
    return c


def get_font(size: int, bold: bool = False) -> ImageFont.ImageFont:
    key = ("b" if bold else "r", size)
    if key in _FONT_CACHE:
        return _FONT_CACHE[key]
    font = None
    for path in _font_candidates():
        p = path
        if bold:
            for b in (path.replace(".ttf", "bd.ttf"), path.replace("Sans.ttf", "Sans-Bold.ttf"),
                      path.replace("-Regular", "-Bold")):
                if Path(b).exists():
                    p = b
                    break
        if Path(p).exists():
            try:
                font = ImageFont.truetype(p, size)
                break
            except Exception:
                continue
    if font is None:
        font = ImageFont.load_default()
    _FONT_CACHE[key] = font
    return font


# ----------------------------------------------------------------------------
# 1) PowerPoint COM
# ----------------------------------------------------------------------------
class PowerPointStageError(RuntimeError):
    """A PowerPoint COM failure tagged with the stage that failed (PROMPT-027R §14).

    ``stage`` is one of import / com-init / dispatch / presentation-open / slide-export; ``reason`` is path-free and
    safe for logs; ``hresult`` is the COM error code when the exception carries one.
    """

    def __init__(self, stage: str, exc: BaseException, pptx: Optional[Path] = None):
        self.stage = stage
        self.exception_type = type(exc).__name__
        self.hresult = _hresult_text(exc)
        self.reason = _com_reason(exc, pptx)
        super().__init__(f"{stage}: {self.reason}")


def _hresult_text(exc: BaseException) -> str:
    code = getattr(exc, "hresult", None)
    if not isinstance(code, int):
        args = getattr(exc, "args", ()) or ()
        code = args[0] if args and isinstance(args[0], int) else None
    return f"0x{code & 0xFFFFFFFF:08X}" if isinstance(code, int) else "-"


def _com_reason(exc: BaseException, pptx: Optional[Path]) -> str:
    text = str(exc).strip()
    if pptx is not None:
        for secret in (str(pptx), Path(pptx).name):
            if secret:
                text = text.replace(secret, "<report>")
    text = " ".join(text.split())
    if len(text) > 180:
        text = text[:177] + "..."
    return text or type(exc).__name__


def render_with_powerpoint(pptx: Path, slide_numbers: Sequence[int], out_dir: Path,
                           width_px: int = 1920) -> Dict[int, Path]:
    """Export COMPLETE authored slides through an APP-OWNED PowerPoint instance (PROMPT-027R §16).

    The user's interactive PowerPoint is never attached to, never closed and never quit: the instance is created with
    ``DispatchEx``, the source is opened read-only without a window, only that presentation is closed and only the
    instance this function created is quit.  CoInitialize / CoUninitialize are balanced.  Every failure raises
    :class:`PowerPointStageError` naming the stage that failed, so diagnostics say WHERE COM broke.
    """
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception as exc:  # noqa: BLE001
        raise PowerPointStageError("import", exc, pptx)
    out: Dict[int, Path] = {}
    try:
        pythoncom.CoInitialize()
    except Exception as exc:  # noqa: BLE001 – RPC_E_CHANGED_MODE etc.: never CoUninitialize what we did not start
        raise PowerPointStageError("com-init", exc, pptx)
    app = None
    pres = None
    created_app = False
    cleanup_failures: List[Tuple[str, BaseException]] = []
    try:
        try:
            app = win32com.client.DispatchEx(POWERPOINT_PROGID)           # always a NEW, application-owned instance
            created_app = True
        except Exception as exc:  # noqa: BLE001
            raise PowerPointStageError("dispatch", exc, pptx)
        try:
            # WithWindow=False keeps it invisible; ReadOnly=True never modifies the source
            pres = app.Presentations.Open(str(pptx.resolve()), True, False, False)
        except Exception as exc:  # noqa: BLE001
            raise PowerPointStageError("presentation-open", exc, pptx)
        try:
            ratio = pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth
            h = int(width_px * ratio)
            for n in slide_numbers:
                target = out_dir / f"slide_{n:03d}.png"
                pres.Slides(n).Export(str(target.resolve()), "PNG", width_px, h)
                out[n] = target
        except Exception as exc:  # noqa: BLE001
            raise PowerPointStageError("slide-export", exc, pptx)
    finally:
        if pres is not None:
            try:
                pres.Close()                                              # only the presentation WE opened
            except Exception as exc:  # noqa: BLE001
                cleanup_failures.append(("presentation-close", exc))
        if created_app and app is not None:
            try:
                app.Quit()                                                # only the instance WE created
            except Exception as exc:  # noqa: BLE001
                cleanup_failures.append(("app-quit", exc))
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass
        for step, exc in cleanup_failures:
            LOG.info("SLIDE_RENDER_CLEANUP backend=powerpoint stage=cleanup step=%s exception_type=%s reason=%s",
                     step, type(exc).__name__, _com_reason(exc, pptx))
    return out


# ----------------------------------------------------------------------------
# 2) LibreOffice
# ----------------------------------------------------------------------------
def render_with_libreoffice(pptx: Path, slide_numbers: Sequence[int], out_dir: Path,
                            dpi: int = 150) -> Dict[int, Path]:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice not found")
    import fitz  # PyMuPDF
    with tempfile.TemporaryDirectory(prefix="re_lo_") as tmp:
        cmd = [soffice, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", tmp, str(pptx)]
        subprocess.run(cmd, check=True, timeout=300, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pdfs = list(Path(tmp).glob("*.pdf"))
        if not pdfs:
            raise RuntimeError("LibreOffice produced no PDF")
        doc = fitz.open(str(pdfs[0]))
        out: Dict[int, Path] = {}
        for n in slide_numbers:
            if 1 <= n <= doc.page_count:
                page = doc[n - 1]
                pix = page.get_pixmap(dpi=dpi)
                target = out_dir / f"slide_{n:03d}.png"
                pix.save(str(target))
                out[n] = target
        doc.close()
    return out


# ----------------------------------------------------------------------------
# 3) Built-in renderer
# ----------------------------------------------------------------------------
def _wrap(text: str, font: ImageFont.ImageFont, max_w: int, draw: ImageDraw.ImageDraw) -> List[str]:
    lines: List[str] = []
    for para in text.split("\n"):
        if not para.strip():
            lines.append("")
            continue
        words = para.split(" ")
        cur = ""
        for w in words:
            trial = (cur + " " + w).strip()
            if draw.textlength(trial, font=font) <= max_w or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
    return lines


def render_builtin(report: ReportData, slide: SlideData, width_px: int = 1920) -> Image.Image:
    sw = report.slide_width or 12192000
    sh = report.slide_height or 6858000
    scale = width_px / sw
    height_px = max(1, int(sh * scale))
    img = Image.new("RGB", (width_px, height_px), "white")
    draw = ImageDraw.Draw(img)
    px = lambda emu: int(emu * scale)  # noqa: E731

    # Preserve PPTX drawing order. Shape fills/strokes are drawn before the text frame belonging to that shape.
    events = []
    serial = 0
    for block in slide.annotations:
        events.append((block.z_order if block.z_order >= 0 else 10**9, 0, serial, "visual", block))
        serial += 1
    for block in slide.blocks:
        if block.kind == "picture":
            priority, kind = 1, "picture"
        elif block.kind == "table":
            priority, kind = 2, "table"
        elif block.kind in ("paragraph", "title"):
            priority, kind = 3, "text"
        else:
            continue
        events.append((block.z_order if block.z_order >= 0 else 10**9, priority, serial, kind, block))
        serial += 1
    arrow_by_z = {arrow.z_order: arrow for arrow in slide.arrows if arrow.z_order >= 0}
    for _z, _priority, _serial, kind, block in sorted(events, key=lambda event: event[:3]):
        if kind == "visual":
            _draw_visual_shape(draw, block, px, scale, arrow_by_z.get(block.z_order))
        elif kind == "picture":
            _draw_picture(img, draw, block, px)
        elif kind == "table" and block.rows:
            _draw_table(draw, block, px)
        elif kind == "text" and block.text.strip():
            _draw_text(draw, block, px, scale)
    return img


def _draw_picture(canvas: Image.Image, draw: ImageDraw.ImageDraw, block: Block, px) -> None:
    if not block.image_blob:
        return
    try:
        with Image.open(BytesIO(block.image_blob)) as source:
            pic = source.convert("RGBA") if source.mode in ("P", "LA", "RGBA") else source.convert("RGB")
        width, height = max(1, px(block.width)), max(1, px(block.height))
        pic = pic.resize((width, height), Image.LANCZOS)
        x, y = px(block.left), px(block.top)
        if pic.mode == "RGBA":
            canvas.paste(pic, (x, y), pic)
        else:
            canvas.paste(pic, (x, y))
    except Exception as exc:  # noqa: BLE001
        LOG.debug("picture render failed: %s", exc)
        draw.rectangle([px(block.left), px(block.top), px(block.right), px(block.bottom)], outline="gray")


def _rotate_points(points, cx: float, cy: float, rotation: float):
    if not rotation:
        return points
    theta = math.radians(rotation)
    cosine, sine = math.cos(theta), math.sin(theta)
    return [(cx + (x - cx) * cosine - (y - cy) * sine,
             cy + (x - cx) * sine + (y - cy) * cosine) for x, y in points]


def _visual_polygon(block: Block, px):
    x0, y0, x1, y1 = px(block.left), px(block.top), px(block.right), px(block.bottom)
    width, height = max(1, x1 - x0), max(1, y1 - y0)
    prst = (block.prst or "rect").lower()
    if prst in ("ellipse", "oval", "donut"):
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        points = [(cx + width / 2 * math.cos(2 * math.pi * i / 64),
                   cy + height / 2 * math.sin(2 * math.pi * i / 64)) for i in range(64)]
    elif "arrow" in prst:
        if "left" in prst:
            raw = [(x1, y0 + height * .25), (x0 + width * .4, y0 + height * .25),
                   (x0 + width * .4, y0), (x0, cy := (y0 + y1) / 2),
                   (x0 + width * .4, y1), (x0 + width * .4, y0 + height * .75),
                   (x1, y0 + height * .75)]
        elif "up" in prst:
            raw = [(x0 + width * .25, y1), (x0 + width * .25, y0 + height * .4),
                   (x0, y0 + height * .4), ((x0 + x1) / 2, y0),
                   (x1, y0 + height * .4), (x0 + width * .75, y0 + height * .4),
                   (x0 + width * .75, y1)]
        elif "down" in prst:
            raw = [(x0 + width * .25, y0), (x0 + width * .75, y0),
                   (x0 + width * .75, y0 + height * .6), (x1, y0 + height * .6),
                   ((x0 + x1) / 2, y1), (x0, y0 + height * .6),
                   (x0 + width * .25, y0 + height * .6)]
        else:
            raw = [(x0, y0 + height * .25), (x0 + width * .6, y0 + height * .25),
                   (x0 + width * .6, y0), (x1, (y0 + y1) / 2),
                   (x0 + width * .6, y1), (x0 + width * .6, y0 + height * .75),
                   (x0, y0 + height * .75)]
        points = raw
    elif "triangle" in prst:
        points = [(x0, y1), ((x0 + x1) / 2, y0), (x1, y1)]
    elif prst in ("diamond",):
        points = [((x0 + x1) / 2, y0), (x1, (y0 + y1) / 2),
                  ((x0 + x1) / 2, y1), (x0, (y0 + y1) / 2)]
    elif "chevron" in prst:
        points = [(x0, y0), (x0 + width * .62, y0), (x1, (y0 + y1) / 2),
                  (x0 + width * .62, y1), (x0, y1), (x0 + width * .38, (y0 + y1) / 2)]
    else:
        points = [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    return _rotate_points(points, cx, cy, block.rotation)


def _draw_line_arrowhead(draw: ImageDraw.ImageDraw, tip: Tuple[float, float], direction: str,
                         size: int, color) -> None:
    vectors = {"right": (1, 0), "left": (-1, 0), "down": (0, 1), "up": (0, -1)}
    dx, dy = vectors.get(direction or "", (1, 0))
    bx, by = tip[0] - dx * size, tip[1] - dy * size
    half = max(2, size * .55)
    points = [tip, (bx - dy * half, by + dx * half), (bx + dy * half, by - dx * half)]
    draw.polygon(points, fill=color)


def _draw_visual_shape(draw: ImageDraw.ImageDraw, block: Block, px, scale: float, arrow=None) -> None:
    """Paint one vector shape strictly according to its resolved style (PROMPT-027 §9).

    ``fill_visible``/``line_visible`` are authoritative: when a slot is invisible NO pixels are painted
    for it, and no colour is invented.  A shape authored ``No Fill`` + ``No Line`` therefore contributes
    nothing at all — the previous code turned such a shape into a black rectangle because an unresolved
    line colour fell back to ``#000000`` while the line was still (wrongly) considered visible.
    """
    fill_color = (block.fill_color or DEFAULT_SHAPE_FILL) if block.fill_visible else None
    line_color = (block.line_color or DEFAULT_SHAPE_LINE) if block.line_visible else None
    stroke = max(1, int(round((block.line_width or DEFAULT_LINE_WIDTH_EMU) * scale))) if block.line_visible else 0
    if block.line_endpoints:
        # A straight connector/line has no fill concept: its stroke IS the shape.  An explicit No Line
        # hides it; otherwise it stays visible (an invisible connector would carry no meaning at all).
        if block.line_explicit_none or (not block.line_visible and not fill_color):
            return
        paint = line_color or fill_color or DEFAULT_SHAPE_LINE
        x0, y0, x1, y1 = block.line_endpoints
        start, end = (px(x0), px(y0)), (px(x1), px(y1))
        draw.line((start, end), fill=paint, width=max(1, stroke))
        if arrow is not None and arrow.direction:
            direction = arrow.direction
            tip = max((start, end), key=lambda p: p[0]) if direction == "right" else \
                min((start, end), key=lambda p: p[0]) if direction == "left" else \
                max((start, end), key=lambda p: p[1]) if direction == "down" else \
                min((start, end), key=lambda p: p[1])
            _draw_line_arrowhead(draw, tip, direction, max(5, stroke * 5), paint)
        return
    if fill_color is None and line_color is None:
        return                                                 # No Fill + No Line -> nothing to paint
    points = _visual_polygon(block, px)
    draw.polygon(points, fill=fill_color, outline=line_color, width=max(1, stroke))


def _text_layout_lines(b: Block, font, box_w: int, draw: ImageDraw.ImageDraw,
                       wrap: bool) -> List[Tuple[str, int]]:
    """Wrapped display lines of ``b.text`` tagged with their source paragraph index.

    The paragraph index is what makes per-paragraph spacing and per-paragraph colour correct: indexing
    those lists by the wrapped-line number (as before) desynchronizes them as soon as any paragraph
    wraps onto more than one line.
    """
    out: List[Tuple[str, int]] = []
    for index, paragraph in enumerate(b.text.split("\n")):
        if not paragraph.strip():
            out.append(("", index))
            continue
        if not wrap:                                           # <a:bodyPr wrap="none">
            out.append((paragraph, index))
            continue
        current = ""
        for word in paragraph.split(" "):
            trial = (current + " " + word).strip()
            if draw.textlength(trial, font=font) <= box_w or not current:
                current = trial
            else:
                out.append((current, index))
                current = word
        if current:
            out.append((current, index))
    return out


def _spacing(b: Block, values: Sequence[float], index: int) -> float:
    return float(values[index]) if index < len(values) else 0.0


def _draw_text(draw: ImageDraw.ImageDraw, b: Block, px, scale: float) -> None:
    """Draw a text frame at its AUTHORED position inside its box (PROMPT-027 §11).

    PowerPoint places text using the body insets, the vertical anchor (``t``/``ctr``/``b``), the
    ``normAutofit`` font scale and per-paragraph line/paragraph spacing.  Ignoring those pinned every
    run to the top-left corner, which is exactly why real slides with authored whitespace and centred
    labels rendered as if their content had been shifted upward.  All geometry here derives from the
    PPTX values — there is no fixed offset, no per-slide and no per-report correction.
    """
    x0, y0, x1, y1 = px(b.left), px(b.top), px(b.right), px(b.bottom)
    ins_l, ins_t = px(b.inset_left), px(b.inset_top)
    ins_r, ins_b = px(b.inset_right), px(b.inset_bottom)
    area_x = x0 + ins_l
    area_y = y0 + ins_t
    area_w = max(8, x1 - ins_r - area_x)
    area_h = max(8, y1 - ins_b - area_y)
    wrap = b.wrap != "none"

    def measure(size: int) -> Tuple[List[Tuple[str, int]], List[float], float]:
        """Return (lines, per-line advance, total height) for one candidate font size."""
        font = get_font(size, b.bold)
        lines = _text_layout_lines(b, font, area_w, draw, wrap)
        advances: List[float] = []
        total = 0.0
        previous_paragraph = None
        for _text, paragraph in lines:
            multiple = b.line_spacing[paragraph] if paragraph < len(b.line_spacing) else None
            line_h = size * (multiple if isinstance(multiple, float) and multiple > 0 else DEFAULT_LINE_SPACING)
            advance = line_h
            if paragraph != previous_paragraph:
                advance += _spacing(b, b.space_before, paragraph) * 12700 * scale
                if previous_paragraph is not None:
                    advance += _spacing(b, b.space_after, previous_paragraph) * 12700 * scale
            previous_paragraph = paragraph
            advances.append(advance)
            total += advance
        return lines, advances, total

    size_pt = b.size_pt or DEFAULT_TEXT_PT
    size = max(MIN_TEXT_PX, int(size_pt * EMU_PER_PT * scale * b.autofit_scale * TEXT_SIZE_CALIBRATION))
    lines, advances, total = measure(size)
    # Bounded shrink-to-fit: only when the authored text genuinely overflows the authored box, and only
    # down to MIN_TEXT_PX.  This is a fallback-renderer approximation of PowerPoint's autofit, never a
    # positional correction.
    while total > area_h and size > MIN_TEXT_PX:
        size -= 1
        lines, advances, total = measure(size)

    anchor = b.vertical_anchor
    if anchor == "ctr":
        cursor = area_y + max(0.0, (area_h - total) / 2.0)
    elif anchor == "b":
        cursor = area_y + max(0.0, area_h - total)
    else:
        cursor = float(area_y)
    for (line, paragraph), advance in zip(lines, advances):
        if line:
            color = b.line_colors[paragraph] if paragraph < len(b.line_colors) and b.line_colors[paragraph] \
                else DEFAULT_TEXT_COLOR
            draw.text((area_x, cursor), line, fill=color, font=get_font(size, b.bold))
        cursor += advance


def _draw_table(draw: ImageDraw.ImageDraw, b: Block, px) -> None:
    rows = b.rows
    n_rows = len(rows)
    n_cols = max(len(r) for r in rows) if rows else 0
    if not n_rows or not n_cols:
        return
    x0, y0, x1, y1 = px(b.left), px(b.top), px(b.right), px(b.bottom)
    cw = max(1, (x1 - x0) // n_cols)
    rh = max(1, (y1 - y0) // n_rows)
    font = get_font(max(10, min(22, int(rh * 0.28))))
    for ri, row in enumerate(rows):
        for ci in range(n_cols):
            cx0, cy0 = x0 + ci * cw, y0 + ri * rh
            draw.rectangle([cx0, cy0, cx0 + cw, cy0 + rh], outline="black", width=1)
            txt = row[ci] if ci < len(row) else ""
            if txt:
                lines = _wrap(txt, font, cw - 6, draw)
                y = cy0 + 2
                for ln in lines:
                    if y > cy0 + rh - 4:
                        break
                    draw.text((cx0 + 3, y), ln, fill="black", font=font)
                    y += int(font.size * 1.2) if hasattr(font, "size") else 12


# ----------------------------------------------------------------------------
# Public API
# ----------------------------------------------------------------------------
class SlideRenderer:
    """Renders slides with the best available backend; caches per file.

    PROMPT-027 §3 / PROMPT-027R §14: every backend attempt is logged with a ``purpose`` (learning_preview /
    after_evidence / qpn_panel / improvement_sheet) and, for failures, the STAGE that broke (availability, import,
    com-init, dispatch, presentation-open, slide-export, render), so Windows acceptance can tell which renderer actually
    produced the preview and why a higher-fidelity backend was not used.  Only ``powerpoint`` is pixel-faithful to
    PowerPoint (:data:`FAITHFUL_BACKENDS`); LibreOffice and the built-in renderer are honest fallbacks and are reported
    as reduced fidelity rather than pretending to be PowerPoint output (§6).  No failure is silently skipped.
    """

    def __init__(self, prefer: Sequence[str] = ("powerpoint", "libreoffice", "builtin"),
                 width_px: int = 1920, dpi: int = 150):
        self.prefer = list(prefer)
        self.width_px = width_px
        self.dpi = dpi
        self.last_backend = ""
        self.last_purpose = ""
        self.backend_tried: List[str] = []
        self.last_failures: List[Dict[str, str]] = []

    def availability(self) -> Dict[str, bool]:
        """Per-backend availability for cache identity (PROMPT-027R §30).  Cheap; never starts a renderer."""
        out: Dict[str, bool] = {}
        for backend in self.prefer:
            if backend == "powerpoint":
                out[backend] = bool(powerpoint_available())
            elif backend == "libreoffice":
                out[backend] = bool(find_soffice() and pymupdf_available())
            elif backend == "builtin":
                out[backend] = True
            else:
                out[backend] = False
        return out

    def render(self, report: ReportData, slide_numbers: Sequence[int], out_dir: Path,
               should_cancel: Optional[Callable[[], bool]] = None,
               purpose: str = "") -> Dict[int, Path]:
        """Render slides. ``should_cancel`` is only observed BEFORE a backend starts and BETWEEN
        individual builtin slide renders; a running PowerPoint/LibreOffice/builtin render of one slide is
        never interrupted mid-operation (PROMPT-024R §15). Cancellation therefore raises at the nearest
        safe boundary and leaves no partially written render behind.

        ``purpose`` labels the consumer (``learning_preview`` / ``after_evidence`` / ``qpn_panel``) in
        the structured ``SLIDE_RENDER`` diagnostics; it never changes rendering behaviour.
        """
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        wanted = [n for n in slide_numbers if report.slide(n)]
        if not wanted:
            return {}
        check_cancelled(should_cancel)
        errors: List[str] = []
        self.backend_tried = []
        self.last_failures = []
        slide_width = int(report.slide_width or 0)
        slide_height = int(report.slide_height or 0)
        for backend in self.prefer:
            check_cancelled(should_cancel)
            self.backend_tried.append(backend)
            try:
                if backend == "powerpoint":
                    if not powerpoint_available():
                        self._fail(purpose, backend, report, stage="availability", reason="unavailable",
                                   detail=powerpoint_unavailable_detail())
                        continue
                    res = render_with_powerpoint(report.path, wanted, out_dir, self.width_px)
                elif backend == "libreoffice":
                    if not (find_soffice() and pymupdf_available()):
                        self._fail(purpose, backend, report, stage="availability", reason="unavailable")
                        continue
                    res = render_with_libreoffice(report.path, wanted, out_dir, self.dpi)
                elif backend == "builtin":
                    res = {}
                    for n in wanted:
                        check_cancelled(should_cancel)
                        img = render_builtin(report, report.slide(n), self.width_px)
                        target = out_dir / f"slide_{n:03d}.png"
                        img.save(target, "PNG")
                        res[n] = target
                else:
                    self._fail(purpose, backend, report, stage="availability", reason="unknown-backend")
                    continue
                if res and all(n in res for n in wanted):
                    self.last_backend = backend
                    self.last_purpose = purpose
                    self._log_render(purpose, backend, report, res, slide_width, slide_height,
                                     attempted=self.backend_tried)
                    return res
                self._fail(purpose, backend, report, stage="render", reason="incomplete-result")
                errors.append(f"{backend}: incomplete result")
            except CancellationRequested:
                raise                                    # cooperative cancel is never a renderer failure
            except PowerPointStageError as e:
                self._fail(purpose, backend, report, stage=e.stage, reason=e.reason,
                           exception_type=e.exception_type, hresult=e.hresult)
                errors.append(f"{backend}[{e.stage}]: {e.reason}")
            except Exception as e:  # noqa: BLE001
                reason = _safe_reason(report, e)
                LOG.warning("Renderer %s failed for %s: %s", backend, report.filename, reason)
                self._fail(purpose, backend, report, stage="render", reason=reason,
                           exception_type=type(e).__name__)
                errors.append(f"{backend}: {reason}")
        LOG.error("SLIDE_RENDER_FAILED purpose=%s report=%s slides=%s tried=%s errors=%s",
                  purpose or "-", report.filename, wanted, ",".join(self.backend_tried) or "-",
                  "; ".join(errors) or "-")
        raise RuntimeError("Không render được slide: " + "; ".join(errors))

    @property
    def degraded(self) -> bool:
        """True when a backend that SHOULD have been tried actually failed during the last render (not merely absent).

        A degraded preview is reused only briefly, so a transient COM failure cannot pin the fallback (§30)."""
        return any(failure.get("stage") not in ("availability", "") for failure in self.last_failures)

    # ------------------------------------------------------------------ diagnostics (PROMPT-027 §3 / 027R §14)
    def _fail(self, purpose: str, backend: str, report: ReportData, stage: str, reason: str,
              exception_type: str = "", hresult: str = "", detail: str = "") -> None:
        self.last_failures.append({"backend": backend, "stage": stage, "reason": reason})
        LOG.warning("SLIDE_RENDER_BACKEND_FAILED purpose=%s backend=%s stage=%s exception_type=%s hresult=%s "
                    "reason=%s detail=%s report=%s",
                    purpose or "-", backend, stage, exception_type or "-", hresult or "-", reason,
                    detail or "-", report.filename)

    @staticmethod
    def _log_render(purpose: str, backend: str, report: ReportData, res: Dict[int, Path],
                    slide_width: int, slide_height: int, attempted: Sequence[str] = ()) -> None:
        for number in sorted(res):
            width = height = 0
            try:
                with Image.open(res[number]) as image:     # header only; pixels are never loaded here
                    width, height = image.size
            except Exception:  # noqa: BLE001
                pass
            faithful = backend in FAITHFUL_BACKENDS
            LOG.info("SLIDE_RENDER purpose=%s backend=%s slide=%s attempted_backend=%s selected_backend=%s "
                     "faithful=%s report=%s rendered=%sx%s slide_emu=%sx%s schema=v%s",
                     purpose or "-", backend, number, ",".join(attempted) or backend, backend,
                     str(faithful).lower(), report.filename, width, height, slide_width, slide_height,
                     RENDER_SCHEMA_VERSION)

    @staticmethod
    def _log_backend_failed(purpose: str, backend: str, report: ReportData, reason: str) -> None:
        LOG.warning("SLIDE_RENDER_BACKEND_FAILED purpose=%s backend=%s stage=render exception_type=- hresult=- "
                    "reason=%s detail=- report=%s", purpose or "-", backend, reason, report.filename)


@dataclass
class QpnRender:
    path: Optional[Path]            # None when the QPN panel could not be isolated (fail-closed)
    backend: str = ""
    region: Optional[QpnRegion] = None
    reason: str = ""                # why no image was produced
    crop_px: Tuple[int, int, int, int] = (0, 0, 0, 0)

    @property
    def ok(self) -> bool:
        return self.path is not None


def render_qpn_panel(report: ReportData, qpn_slide: int, target_png: Path,
                     renderer: Optional[SlideRenderer] = None) -> QpnRender:
    """PRODUCTION QPN image = the QPN panel ONLY (PROMPT-001).

    The panel is located from the PPTX objects (:func:`locate_qpn_region`).  A direct picture object is
    exported as-is; otherwise the slide is rendered and CROPPED to the panel's bounds.  Whole-slide output
    (even white-trimmed) is never produced: without a confident panel the result is ``path=None`` + reason.
    """
    slide = report.slide(qpn_slide)
    if slide is None:
        return QpnRender(None, reason=f"slide {qpn_slide} không tồn tại")
    W, H = report.slide_width or slide.width, report.slide_height or slide.height
    loc = locate_qpn_region(slide, W, H)
    if loc.region is None:
        return QpnRender(None, reason=loc.reason)
    region = loc.region
    target_png.parent.mkdir(parents=True, exist_ok=True)
    if region.picture is not None and region.picture.image_blob:
        try:
            with Image.open(BytesIO(region.picture.image_blob)) as im:
                im.convert("RGB").save(target_png, "PNG", optimize=True)
            return QpnRender(target_png, "picture", region)
        except Exception as e:  # noqa: BLE001 – EMF/WMF previews: fall through to render + crop
            LOG.debug("QPN picture blob not decodable (%s) – cropping rendered slide instead", e)
    renderer = renderer or SlideRenderer()
    with tempfile.TemporaryDirectory(prefix="re_qpn_") as tmp:
        res = renderer.render(report, [qpn_slide], Path(tmp), purpose="qpn_panel")
        with Image.open(res[qpn_slide]) as im:
            im = im.convert("RGB")
            box = crop_box_px(region, W, H, im.width, im.height)
            im.crop(box).save(target_png, "PNG", optimize=True)
    return QpnRender(target_png, renderer.last_backend, region, crop_px=box)


def render_qpn(report: ReportData, qpn_slide: int, target_png: Path,
               renderer: Optional[SlideRenderer] = None) -> Tuple[Path, str]:
    """Compatibility wrapper: QPN panel image or ``RuntimeError`` (never a whole-slide fallback)."""
    res = render_qpn_panel(report, qpn_slide, target_png, renderer)
    if not res.ok:
        raise RuntimeError(f"Không tách được QPN khỏi slide {qpn_slide}: {res.reason}")
    return res.path, res.backend


def trim_white_margins(im: "Image.Image", threshold: int = 245, pad: int = 8) -> "Image.Image":
    """Crop uniform (near-)white borders around the content; keeps a small padding, never crops content."""
    try:
        from PIL import ImageChops, ImageOps
        gray = ImageOps.grayscale(im)
        mask = gray.point(lambda v: 255 if v < threshold else 0)
        bbox = mask.getbbox()
        if not bbox:
            return im
        l, t, r, b = bbox
        l, t = max(0, l - pad), max(0, t - pad)
        r, b = min(im.width, r + pad), min(im.height, b + pad)
        if (r - l) < im.width * 0.2 or (b - t) < im.height * 0.2:
            return im
        return im.crop((l, t, r, b))
    except Exception:  # noqa: BLE001
        return im
