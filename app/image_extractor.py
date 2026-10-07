"""Improvement image extraction.

Production rule: "Hình ảnh cải tiến" = one rendered After visual-region crop per eligible logical improvement item.
Selection/ownership remain in :mod:`improvement_pictures`; final slide rendering and cropping live in
:func:`export_after_pictures`. Embedded picture files are used only if rendering fails.
"""
from __future__ import annotations

import logging
import re
import tempfile
from io import BytesIO
from pathlib import Path
from typing import Callable, List, Optional, Sequence, Tuple

from PIL import Image, ImageDraw

from .pptx_parser import ReportData
from .improvement_pictures import PictureRef
from .improvement_visual import (ImprovementVisualRegion, build_improvement_visual_regions, crop_box_px)
from .qpn_renderer import SlideRenderer, get_font
from .report_identity import report_scope_key

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
    """Legacy compatibility helper that stacks selected After pictures without slide annotations.

    The production Excel path is :func:`export_after_pictures`, which renders one complete logical-item crop. This
    helper remains for older callers/tests and diagnostics; it does not preserve the authored visual composition.
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


class GroupedImagePaths(list):
    """Flat path list for API compatibility, plus logical-item rows for grouped Excel placement."""
    def __init__(self, groups: Sequence[Sequence[Path]], owners: Optional[Sequence[str]] = None):
        self.groups = [list(group) for group in groups if group]
        self.owners = list(owners or [])[:len(self.groups)]
        super().__init__(path for group in self.groups for path in group)


def _region_stem(region: ImprovementVisualRegion, index: int) -> str:
    scope = re.sub(r"[^A-Za-z0-9_-]+", "_", region.report_scope_id or "local")[:32]
    return f"region_{scope}_{index:03d}_s{region.slide_index:03d}"


def _fallback_picture_group(region: ImprovementVisualRegion, out_dir: Path, stem: str,
                            problems: List[str], report_name: str) -> List[Path]:
    """PROMPT-021 fallback: preserve the old eligible-picture export, still isolated to this item/report."""
    saved: List[Path] = []
    failed = False
    for index, ref in enumerate(region.pictures, start=1):
        blob = ref.block.image_blob
        ext = (ref.block.image_ext or "").lower()
        if not blob:
            failed = True
            problems.append(f"REGION_FALLBACK MN={region.management_number or '-'} report={report_name} "
                            f"scope={region.report_scope_id} slide={region.slide_index} "
                            f"item={region.improvement_item_id}: picture {ref.label} has no embedded bytes")
            continue
        if ext in ("emf", "wmf"):
            failed = True
            problems.append(f"REGION_FALLBACK MN={region.management_number or '-'} report={report_name} "
                            f"scope={region.report_scope_id} slide={region.slide_index} "
                            f"item={region.improvement_item_id}: {ext.upper()} is not Pillow-decodable")
            continue
        try:
            with Image.open(BytesIO(blob)) as source:
                source.load()
                image = source.convert("RGB")
            target = out_dir / f"{stem}_fallback_{index:02d}.png"
            image.save(target, "PNG", optimize=True)
            saved.append(target)
        except Exception as exc:  # noqa: BLE001
            failed = True
            problems.append(f"REGION_FALLBACK MN={region.management_number or '-'} report={report_name} "
                            f"scope={region.report_scope_id} slide={region.slide_index} "
                            f"item={region.improvement_item_id}: picture {ref.label} decode failed ({type(exc).__name__})")
    if failed:
        for path in saved:
            try:
                path.unlink(missing_ok=True)
            except OSError:
                LOG.debug("could not remove incomplete visual fallback %s", path)
        return []
    return saved


def export_after_pictures(report: ReportData, refs: Sequence["PictureRef"], out_dir: Path,
                          renderer: Optional[SlideRenderer] = None, management_number: str = "",
                          should_cancel: Optional[Callable[[], bool]] = None
                          ) -> Tuple[List[Path], List[str]]:
    """Render one faithful After-region PNG per eligible logical item/slide.

    All refs are rechecked against the PROMPT-015 gate. A failed/missing slide render falls back to the existing
    embedded-picture export for that region; if that also fails, no image is returned and a diagnostic is added.
    The returned :class:`GroupedImagePaths` keeps one logical item per Excel visual row.

    PROMPT-025: ``should_cancel`` is an optional cooperative-cancellation hook checked between item-level
    crops (the slide itself is rendered once and shared by every item region of that slide). The batch
    keeps its file-atomic stop semantics ("Dừng sau file hiện tại"): the hook only shortens work that
    would otherwise be redone, never an Excel commit.
    """
    problems: List[str] = []
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    regions = build_improvement_visual_regions(report, refs, management_number=management_number)
    if not regions:
        expected_scope = report_scope_key(report.path)
        for ref in refs:
            scope_matches = not ref.report_scope_id or ref.report_scope_id == expected_scope
            if not scope_matches:
                reason = "report-scope-mismatch"
            elif not (ref.excel_output_eligible and ref.temporal_role == "AFTER"
                      and ref.semantic_role == "PRODUCTION_IMPROVEMENT" and ref.confident_owner and ref.owner_id):
                reason = "final-semantic-eligibility-gate"
            else:
                reason = "no-structural-visual-region"
            LOG.info("REGION_EXCLUDE MN=%s report=%s scope=%s slide=%s item=%s reason=%s",
                     management_number or "-", report.filename, ref.report_scope_id or "local",
                     ref.slide, ref.owner_id or "unknown", reason)
            if reason in ("report-scope-mismatch", "no-structural-visual-region"):
                problems.append(f"REGION_EXCLUDE MN={management_number or '-'} report={report.filename} "
                                f"scope={ref.report_scope_id or 'local'} slide={ref.slide} "
                                f"item={ref.owner_id or 'unknown'}: {reason}")
        return GroupedImagePaths([], []), problems

    renderer = renderer or SlideRenderer()
    render_paths = {}
    render_error = ""
    slides = sorted({region.slide_index for region in regions})
    with tempfile.TemporaryDirectory(prefix="re_visual_region_") as tmp:
        try:
            render_paths = renderer.render(report, slides, Path(tmp))
            LOG.info("REGION_RENDER MN=%s report=%s scope=%s slides=%s backend=%s result=ok",
                     management_number or "-", report.filename, report_scope_key(report.path), slides,
                     renderer.last_backend or "unknown")
        except Exception as exc:  # noqa: BLE001 – renderer errors must not stop the batch
            render_error = f"{type(exc).__name__}: {exc}"
            LOG.warning("REGION_RENDER MN=%s report=%s scope=%s slides=%s result=failed error=%s",
                        management_number or "-", report.filename, report_scope_key(report.path), slides, render_error)

        path_groups: List[List[Path]] = []
        owners: List[str] = []
        for index, region in enumerate(regions, start=1):
            if should_cancel is not None and should_cancel():
                # cooperative stop between item-level crops (PROMPT-025 §43): finish nothing half-done,
                # record the stop, and let the caller decide (the batch keeps file-atomic semantics).
                problems.append(f"REGION_CANCELLED MN={management_number or '-'} report={report.filename} "
                                f"scope={region.report_scope_id} slide={region.slide_index} "
                                f"item={region.improvement_item_id}: đã dừng giữa các mục cải tiến")
                LOG.info("REGION_CANCELLED MN=%s report=%s scope=%s slide=%s item=%s",
                         management_number or "-", report.filename, region.report_scope_id,
                         region.slide_index, region.improvement_item_id)
                break
            stem = _region_stem(region, index)
            rendered = render_paths.get(region.slide_index)
            group: List[Path] = []
            if rendered is not None and Path(rendered).exists():
                try:
                    with Image.open(rendered) as slide_image:
                        slide_image.load()
                        full_slide = slide_image.convert("RGB")
                    slide_width = int(report.slide_width or report.slide(region.slide_index).width)
                    slide_height = int(report.slide_height or report.slide(region.slide_index).height)
                    crop_box = crop_box_px(region.bbox, slide_width, slide_height,
                                           full_slide.width, full_slide.height)
                    crop = full_slide.crop(crop_box)
                    target = out_dir / f"{stem}.png"
                    crop.save(target, "PNG", optimize=True)
                    group = [target]
                    LOG.info("REGION_CROP MN=%s report=%s scope=%s slide=%s item=%s bbox=%s px=%s size=%s path=%s",
                             management_number or "-", report.filename, region.report_scope_id, region.slide_index,
                             region.improvement_item_id, region.bbox, crop_box, crop.size, target.name)
                except Exception as exc:  # noqa: BLE001 – try the safe picture-only fallback
                    render_error = f"{type(exc).__name__}: {exc}"
                    LOG.warning("REGION_CROP MN=%s report=%s scope=%s slide=%s item=%s result=failed error=%s",
                                management_number or "-", report.filename, region.report_scope_id, region.slide_index,
                                region.improvement_item_id, render_error)
            else:
                render_error = render_error or "renderer returned no image for this slide"

            if not group:
                problems.append(f"REGION_FALLBACK MN={management_number or '-'} report={report.filename} "
                                f"scope={region.report_scope_id} slide={region.slide_index} "
                                f"item={region.improvement_item_id}: {render_error}")
                group = _fallback_picture_group(region, out_dir, stem, problems, report.filename)
                if group:
                    LOG.warning("REGION_FALLBACK MN=%s report=%s scope=%s slide=%s item=%s pictures=%d",
                                management_number or "-", report.filename, region.report_scope_id,
                                region.slide_index, region.improvement_item_id, len(group))
                else:
                    problems.append(f"REGION_EXCLUDE MN={management_number or '-'} report={report.filename} "
                                    f"scope={region.report_scope_id} slide={region.slide_index} "
                                    f"item={region.improvement_item_id}: render and eligible-picture fallback "
                                    "both failed")
                    continue
            path_groups.append(group)
            owners.append(region.block_id)

    return GroupedImagePaths(path_groups, owners), problems
