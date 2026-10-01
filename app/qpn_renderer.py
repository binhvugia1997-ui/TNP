"""Render a complete PPTX slide (e.g. the whole Quality Problem Notice page) to PNG.

Renderer chain (first one that works wins):
  1. Microsoft PowerPoint COM automation (Windows, PowerPoint installed)
  2. LibreOffice headless  (pptx -> pdf -> png via PyMuPDF if available)
  3. Built-in Pillow renderer (draws pictures, tables and text boxes at their
     original positions).  Always available, lower fidelity but never fails.

Nothing here is sent to the cloud.  Source PPTX files are opened read-only.
"""
from __future__ import annotations

import logging
import os
import shutil
import subprocess
import sys
import tempfile
from io import BytesIO
from pathlib import Path
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple

from PIL import Image, ImageDraw, ImageFont

from .qpn_region import QpnRegion, crop_box_px, locate_qpn_region
from .pptx_parser import Block, ReportData, SlideData

LOG = logging.getLogger("report_extractor.renderer")
EMU_PER_INCH = 914400


# ----------------------------------------------------------------------------
# Availability checks
# ----------------------------------------------------------------------------
def powerpoint_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import winreg  # type: ignore
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application"):
            pass
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


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


# ----------------------------------------------------------------------------
# Fonts (Unicode Vietnamese capable)
# ----------------------------------------------------------------------------
_FONT_CACHE: Dict[Tuple[str, int], ImageFont.FreeTypeFont] = {}


def _font_candidates() -> List[str]:
    c: List[str] = []
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    c.append(str(base / "assets" / "DejaVuSans.ttf"))
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
def render_with_powerpoint(pptx: Path, slide_numbers: Sequence[int], out_dir: Path,
                           width_px: int = 1920) -> Dict[int, Path]:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
    out: Dict[int, Path] = {}
    pythoncom.CoInitialize()
    app = None
    pres = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        # WithWindow=False keeps it invisible; ReadOnly=True never modifies the source
        pres = app.Presentations.Open(str(pptx.resolve()), True, False, False)
        ratio = pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth
        h = int(width_px * ratio)
        for n in slide_numbers:
            target = out_dir / f"slide_{n:03d}.png"
            pres.Slides(n).Export(str(target.resolve()), "PNG", width_px, h)
            out[n] = target
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if app is not None and app.Presentations.Count == 0:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
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
    height_px = int(sh * scale)
    img = Image.new("RGB", (width_px, height_px), "white")
    draw = ImageDraw.Draw(img)
    px = lambda emu: int(emu * scale)  # noqa: E731

    # pictures first (background), then tables & text
    for b in slide.blocks:
        if b.kind != "picture" or not b.image_blob:
            continue
        try:
            pic = Image.open(BytesIO(b.image_blob))
            pic = pic.convert("RGBA") if pic.mode in ("P", "LA", "RGBA") else pic.convert("RGB")
            w, h = max(1, px(b.width)), max(1, px(b.height))
            pic = pic.resize((w, h), Image.LANCZOS)
            if pic.mode == "RGBA":
                img.paste(pic, (px(b.left), px(b.top)), pic)
            else:
                img.paste(pic, (px(b.left), px(b.top)))
        except Exception as e:  # noqa: BLE001
            LOG.debug("picture render failed: %s", e)
            draw.rectangle([px(b.left), px(b.top), px(b.right), px(b.bottom)], outline="gray")

    for b in slide.blocks:
        if b.kind == "table" and b.rows:
            _draw_table(draw, b, px)
        elif b.kind in ("paragraph", "title") and b.text.strip():
            _draw_text(draw, b, px, scale)
    return img


def _draw_text(draw: ImageDraw.ImageDraw, b: Block, px, scale: float) -> None:
    x0, y0, x1, y1 = px(b.left), px(b.top), px(b.right), px(b.bottom)
    box_w = max(20, x1 - x0 - 8)
    size_pt = b.size_pt or 14
    size = max(10, int(size_pt * 12700 * scale * 1.05))
    font = get_font(size, b.bold)
    lines = _wrap(b.text, font, box_w, draw)
    line_h = int(size * 1.25)
    # shrink to fit if text overflows badly
    while line_h * len(lines) > (y1 - y0) * 1.6 and size > 10:
        size -= 1
        font = get_font(size, b.bold)
        lines = _wrap(b.text, font, box_w, draw)
        line_h = int(size * 1.25)
    y = y0 + 3
    for ln in lines:
        draw.text((x0 + 4, y), ln, fill="black", font=font)
        y += line_h


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
    """Renders slides with the best available backend; caches per file."""

    def __init__(self, prefer: Sequence[str] = ("powerpoint", "libreoffice", "builtin"),
                 width_px: int = 1920, dpi: int = 150):
        self.prefer = list(prefer)
        self.width_px = width_px
        self.dpi = dpi
        self.last_backend = ""

    def render(self, report: ReportData, slide_numbers: Sequence[int], out_dir: Path) -> Dict[int, Path]:
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        wanted = [n for n in slide_numbers if report.slide(n)]
        if not wanted:
            return {}
        errors: List[str] = []
        for backend in self.prefer:
            try:
                if backend == "powerpoint":
                    if not powerpoint_available():
                        continue
                    res = render_with_powerpoint(report.path, wanted, out_dir, self.width_px)
                elif backend == "libreoffice":
                    if not (find_soffice() and pymupdf_available()):
                        continue
                    res = render_with_libreoffice(report.path, wanted, out_dir, self.dpi)
                elif backend == "builtin":
                    res = {}
                    for n in wanted:
                        img = render_builtin(report, report.slide(n), self.width_px)
                        target = out_dir / f"slide_{n:03d}.png"
                        img.save(target, "PNG")
                        res[n] = target
                else:
                    continue
                if res and all(n in res for n in wanted):
                    self.last_backend = backend
                    return res
                errors.append(f"{backend}: incomplete result")
            except Exception as e:  # noqa: BLE001
                LOG.warning("Renderer %s failed for %s: %s", backend, report.filename, e)
                errors.append(f"{backend}: {e}")
        raise RuntimeError("Không render được slide: " + "; ".join(errors))


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
        res = renderer.render(report, [qpn_slide], Path(tmp))
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
