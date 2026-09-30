"""Improvement image extraction.

Production rule: "Hình ảnh cải tiến" = ONLY the "Sau cải tiến" pictures selected by
:mod:`improvement_pictures`, stacked vertically in source order
(:func:`build_after_pictures_image`).  The older whole-slide contact sheet
(:func:`build_improvement_image`) is kept for diagnostics / audit only.
"""
from __future__ import annotations

import logging
import tempfile
from io import BytesIO
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

from PIL import Image, ImageDraw

from .pptx_parser import ReportData
from .improvement_pictures import PictureRef
from .qpn_renderer import SlideRenderer, get_font

LOG = logging.getLogger("report_extractor.images")


def combine_vertically(images: Sequence[Image.Image], width: int = 1600, gap: int = 16,
                       labels: Optional[Sequence[str]] = None) -> Image.Image:
    """Contact sheet: all images scaled to the same width, stacked top to bottom."""
    scaled: List[Image.Image] = []
    for im in images:
        im = im.convert("RGB")
        if im.width != width:
            h = max(1, int(im.height * width / im.width))
            im = im.resize((width, h), Image.LANCZOS)
        scaled.append(im)
    label_h = 34 if labels else 0
    total_h = sum(im.height + label_h for im in scaled) + gap * (len(scaled) - 1 if scaled else 0)
    sheet = Image.new("RGB", (width, max(1, total_h)), "white")
    draw = ImageDraw.Draw(sheet)
    font = get_font(22, bold=True)
    y = 0
    for i, im in enumerate(scaled):
        if labels:
            draw.rectangle([0, y, width, y + label_h], fill=(240, 240, 240))
            draw.text((8, y + 6), labels[i] if i < len(labels) else "", fill="black", font=font)
            y += label_h
        sheet.paste(im, (0, y))
        y += im.height
        if i < len(scaled) - 1:
            draw.line([(0, y + gap // 2), (width, y + gap // 2)], fill=(200, 200, 200), width=2)
            y += gap
    return sheet


def save_original_pictures(report: ReportData, slide_numbers: Sequence[int], out_dir: Path) -> List[Path]:
    """Dump the original embedded pictures of the given slides (audit / debugging)."""
    out_dir.mkdir(parents=True, exist_ok=True)
    saved: List[Path] = []
    for n in slide_numbers:
        s = report.slide(n)
        if not s:
            continue
        for i, p in enumerate(s.pictures, start=1):
            if not p.image_blob:
                continue
            ext = (p.image_ext or "png").lower()
            if ext in ("emf", "wmf"):
                # Pillow cannot decode these; keep the raw blob
                target = out_dir / f"slide{n:02d}_pic{i:02d}.{ext}"
                target.write_bytes(p.image_blob)
                saved.append(target)
                continue
            try:
                im = Image.open(BytesIO(p.image_blob))
                target = out_dir / f"slide{n:02d}_pic{i:02d}.png"
                im.convert("RGB").save(target, "PNG")
                saved.append(target)
            except Exception as e:  # noqa: BLE001
                LOG.debug("cannot decode picture %s/%s: %s", n, i, e)
    return saved


def build_improvement_image(report: ReportData, slide_numbers: Sequence[int], target_jpg: Path,
                            renderer: Optional[SlideRenderer] = None, width: int = 1600,
                            pictures_dir: Optional[Path] = None) -> Tuple[Optional[Path], str]:
    """Render improvement slides and combine them into ``target_jpg``.

    Returns (path or None if nothing to render, backend).
    """
    nums = [n for n in slide_numbers if report.slide(n)]
    if not nums:
        return None, ""
    renderer = renderer or SlideRenderer()
    with tempfile.TemporaryDirectory(prefix="re_imp_") as tmp:
        res = renderer.render(report, nums, Path(tmp))
        imgs = [Image.open(res[n]).convert("RGB") for n in nums if n in res]
        labels = [f"Slide {n}" for n in nums if n in res]
        sheet = combine_vertically(imgs, width=width, labels=labels if len(imgs) > 1 else None)
    target_jpg.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(target_jpg, "JPEG", quality=88, optimize=True)
    if pictures_dir is not None:
        try:
            save_original_pictures(report, nums, pictures_dir)
        except Exception as e:  # noqa: BLE001
            LOG.debug("save_original_pictures failed: %s", e)
    return target_jpg, renderer.last_backend


def build_after_pictures_image(report: ReportData, refs: Sequence["PictureRef"], target_jpg: Path,
                               width: int = 1600, pictures_dir: Optional[Path] = None
                               ) -> Tuple[Optional[Path], List[str]]:
    """Stack ONLY the selected "Sau cải tiến" pictures vertically (source order) into ``target_jpg``.

    Returns (path or None when nothing could be embedded, list of problems for 'Cần kiểm tra').
    Original pictures are decoded from the PPTX blobs – nothing is re-rendered or altered.
    """
    problems: List[str] = []
    images: List[Image.Image] = []
    saved_dir = None
    if pictures_dir is not None:
        pictures_dir.mkdir(parents=True, exist_ok=True)
        saved_dir = pictures_dir
    for i, ref in enumerate(refs, start=1):
        blob = ref.block.image_blob
        ext = (ref.block.image_ext or "").lower()
        if not blob:
            problems.append(f"Ảnh Sau cải tiến tại slide {ref.slide} không đọc được dữ liệu – cần bổ sung thủ công")
            continue
        if ext in ("emf", "wmf"):
            if saved_dir is not None:
                (saved_dir / f"after{i:02d}_slide{ref.slide:02d}.{ext}").write_bytes(blob)
            problems.append(f"Ảnh Sau cải tiến tại slide {ref.slide} ở định dạng {ext.upper()} không chèn được – "
                            f"cần bổ sung thủ công")
            continue
        try:
            im = Image.open(BytesIO(blob))
            im.load()
            im = im.convert("RGB")
        except Exception as e:  # noqa: BLE001
            problems.append(f"Ảnh Sau cải tiến tại slide {ref.slide} không giải mã được ({type(e).__name__}) – "
                            f"cần bổ sung thủ công")
            continue
        images.append(im)
        if saved_dir is not None:
            try:
                im.save(saved_dir / f"after{i:02d}_slide{ref.slide:02d}.png", "PNG")
            except Exception as e:  # noqa: BLE001
                LOG.debug("cannot save after picture: %s", e)
    if not images:
        return None, problems
    max_w = max(im.width for im in images)
    sheet = combine_vertically(images, width=min(width, max(400, max_w)))
    target_jpg.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(target_jpg, "JPEG", quality=90, optimize=True)
    return target_jpg, problems
