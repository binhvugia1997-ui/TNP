"""Stand-in for the Windows-only PowerPoint COM backend, for Linux runs.

Production rendering is PowerPoint-only (fail-closed): on a machine without PowerPoint
the batch pipeline raises instead of degrading to the built-in renderer.  Linux tests
that exercise the production batch/learning paths therefore drive the real
``SlideRenderer.render`` loop with this stubbed COM call, so they prove the pipeline
works when PowerPoint renders, which is the Windows reality.

The stub raster is a white slide with every PICTURE shape drawn as a filled rectangle
of the picture's average colour at its authored geometry.  That is enough for the
pipeline's geometry and colour-presence assertions (region crops must contain the
After evidence colour and exclude Before/arrow colours); it is not a faithful
render and must never be presented as one.  Tests that assert fallback/degraded
behavior patch the same attributes themselves (their patches apply after this one
and win); this helper must never be autouse suite-wide.
"""
import io
from contextlib import contextmanager
from pathlib import Path

import pytest
from PIL import Image, ImageDraw


def _average_color(blob: bytes):
    try:
        with Image.open(io.BytesIO(blob)) as im:
            return im.convert("RGB").resize((1, 1)).getpixel((0, 0))
    except Exception:
        return None


def _fake_export(pptx, slide_numbers, out_dir, width_px=1920):
    out = {}
    pictures = {}
    try:
        from pptx import Presentation
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        prs = Presentation(str(pptx))
        slide_w, slide_h = prs.slide_width, prs.slide_height
        height_px = max(1, int(width_px * slide_h / slide_w))
        for n in slide_numbers:
            rects = []
            for shape in prs.slides[n - 1].shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    color = _average_color(shape.image.blob)
                    if color is None:
                        continue
                    rects.append((
                        shape.left / slide_w * width_px,
                        shape.top / slide_h * height_px,
                        (shape.left + shape.width) / slide_w * width_px,
                        (shape.top + shape.height) / slide_h * height_px,
                        color,
                    ))
            pictures[n] = (width_px, height_px, rects)
    except Exception:
        pictures = {}
    for n in slide_numbers:
        width, height, rects = pictures.get(n, (int(width_px), int(width_px * 540 / 960), []))
        img = Image.new("RGB", (width, height), "white")
        if rects:
            draw = ImageDraw.Draw(img)
            for x0, y0, x1, y1, color in rects:
                draw.rectangle([x0, y0, x1, y1], fill=color)
        target = Path(out_dir) / f"slide_{n:03d}.png"
        img.save(target)
        out[n] = target
    return out


def stub_powerpoint_render(monkeypatch):
    """Patch the COM probe + export used by ``SlideRenderer.render`` (undo with the fixture)."""
    monkeypatch.setattr("app.qpn_renderer.powerpoint_available", lambda: True)
    monkeypatch.setattr("app.qpn_renderer.render_with_powerpoint", _fake_export)


@contextmanager
def powerpoint_render_stubbed():
    """Context-manager form for batch harnesses that have no ``monkeypatch`` fixture."""
    with pytest.MonkeyPatch.context() as mp:
        stub_powerpoint_render(mp)
        yield
