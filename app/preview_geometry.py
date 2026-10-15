"""Authoritative slide-coordinate → preview-coordinate mapping (PROMPT-024R §28/§52).

All target geometry comes from the PPTX (EMU). The review UI displays ONE rendered full-slide image
and must highlight the reviewed target at its TRUE authored position — never moved, re-centred or
independently enlarged. This module owns the math so the mapping is deterministic and headlessly
testable:

    EMU box  --(slide_fraction_box)-->  fractions of the slide (0..1)
    fractions --(map_fraction_box_to_px)--> pixels of the DISPLAYED slide image
    container --(fit_slide_box)--> letterboxed display size inside a container, honouring zoom

Because the overlay is positioned with the same fractions inside the exact display box of the slide
image, resizing the panel, Windows DPI scaling and zoom can never desynchronize the highlight.
"""
from __future__ import annotations

from typing import Dict, Tuple

Box = Tuple[float, float, float, float]  # (x, y, width, height)


def _clamp01(v: float) -> float:
    return max(0.0, min(1.0, v))


def slide_fraction_box(box_emu: Box, slide_width_emu: float, slide_height_emu: float) -> Box:
    """Convert an EMU (left, top, width, height) box into fractions of the authored slide.

    Degenerate slide sizes fall back to a full-slide box; out-of-slide geometry is clamped so the
    highlight can never point outside the rendered preview.
    """
    x, y, w, h = (float(v) for v in box_emu)
    sw = max(1.0, float(slide_width_emu))
    sh = max(1.0, float(slide_height_emu))
    fx, fy = _clamp01(x / sw), _clamp01(y / sh)
    fw = _clamp01((x + w) / sw) - fx
    fh = _clamp01((y + h) / sh) - fy
    return (fx, fy, max(0.0, fw), max(0.0, fh))


def fit_slide_box(container_w: float, container_h: float, slide_width: float, slide_height: float,
                  zoom: float = 1.0) -> Dict[str, float]:
    """Letterbox-aware "Vừa khung" fit of the slide inside a preview container.

    Returns the displayed slide size (``width``/``height``) and its top-left offset
    (``offset_x``/``offset_y``) inside the container. Aspect ratio is always preserved; zoom scales
    the fitted result uniformly (the container then scrolls when the result overflows).
    """
    cw, ch = max(1.0, float(container_w)), max(1.0, float(container_h))
    sw, sh = max(1.0, float(slide_width)), max(1.0, float(slide_height))
    scale = min(cw / sw, ch / sh) * max(0.05, float(zoom))
    width, height = sw * scale, sh * scale
    return {"width": width, "height": height,
            "offset_x": max(0.0, (cw - width) / 2.0), "offset_y": max(0.0, (ch - height) / 2.0)}


def map_fraction_box_to_px(fraction_box: Box, display_w: float, display_h: float) -> Dict[str, float]:
    """Fractions of the slide → pixels of the displayed slide image (identity under resize/zoom:
    the caller passes the CURRENT displayed size, so panel resize / DPI / zoom stay aligned)."""
    fx, fy, fw, fh = fraction_box
    return {"x": fx * display_w, "y": fy * display_h, "width": fw * display_w, "height": fh * display_h}


def map_slide_box_to_preview(box_emu: Box, slide_width_emu: float, slide_height_emu: float,
                             display_w: float, display_h: float) -> Dict[str, float]:
    """One-shot mapping used by tests/tooling: EMU box → px rect on the displayed slide image."""
    return map_fraction_box_to_px(slide_fraction_box(box_emu, slide_width_emu, slide_height_emu),
                                  display_w, display_h)
