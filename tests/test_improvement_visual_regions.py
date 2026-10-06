"""PROMPT-021 visual-region crops: complete After groups, overlays, semantics and report isolation."""
from __future__ import annotations

from io import BytesIO
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from app.classifier import heuristic_classify
from app.image_extractor import GroupedImagePaths, export_after_pictures
from app.improvement_pictures import SEMANTIC_PRODUCTION, select_after_pictures
from app.improvement_visual import build_improvement_visual_regions
from app.pptx_parser import parse_pptx
from app.qpn_renderer import SlideRenderer
from tests.test_after_evidence import _deck, _pic, _pics, _prod_head, _tb

BEFORE_A = (241, 180, 70)
BEFORE_B = (212, 112, 145)
AFTER_A = (42, 154, 82)
AFTER_B = (28, 104, 186)
AFTER_C = (80, 180, 140)
BEFORE_MARK = (218, 35, 45)
AFTER_MARK = (220, 25, 35)
TRANSITION = (175, 15, 195)
AFTER_ARROW = (0, 185, 220)
AFTER_BORDER = (20, 95, 220)


def _caption(slide, text, left, top, width, height):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top),
                                   Inches(width), Inches(height))
    shape.text_frame.text = text
    shape.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.color.rgb = RGBColor(*AFTER_BORDER)
    shape.line.width = Pt(1.25)
    return shape


def _make_visual_layout(slide, width, after_colors=("#2a9a52", "#1c68ba"), include_inspection=False):
    _prod_head(slide, width)
    _tb(slide, "Cải tiến lỗi Hằn", 1.7, 0.9, 5.4, 0.4, 12)

    # Before side: two eligible-looking picture objects, a red markup and a caption.
    _caption(slide, "Trước cải tiến", 1.45, 1.35, 1.65, 0.34)
    _pics(slide, ["#f1b446", "#d47091"], 1.45, 1.8, w=1.45, h=1.05, gap=0.18)
    before_circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.75), Inches(2.0), Inches(0.48), Inches(0.48))
    before_circle.fill.background()
    before_circle.line.color.rgb = RGBColor(*BEFORE_MARK)
    before_circle.line.width = Pt(3)

    # Large transition arrow is between the two visual groups and must stay outside the crop.
    transition = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(5.45), Inches(2.05),
                                        Inches(1.45), Inches(0.65))
    transition.fill.solid()
    transition.fill.fore_color.rgb = RGBColor(*TRANSITION)
    transition.line.fill.background()

    # A transparent blue frame groups the authored After objects and caption.
    border = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.55), Inches(1.52),
                                    Inches(3.9), Inches(2.35))
    border.fill.background()
    border.line.color.rgb = RGBColor(*AFTER_BORDER)
    border.line.width = Pt(2)

    after_positions = [7.82, 9.25, 10.68] if len(after_colors) == 3 else [7.82, 9.55]
    photo_width = 1.18 if len(after_colors) == 3 else 1.45
    _pics(slide, list(after_colors), after_positions[0], 1.78, w=photo_width, h=1.12,
          gap=(after_positions[1] - after_positions[0] - photo_width))
    _caption(slide, "Sau cải tiến", 8.55, 3.15, 1.95, 0.38)

    # All four annotation kinds below are non-picture objects on the After side.
    after_circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.97), Inches(1.96), Inches(0.46), Inches(0.46))
    after_circle.fill.background()
    after_circle.line.color.rgb = RGBColor(*AFTER_MARK)
    after_circle.line.width = Pt(3)
    after_arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(8.85), Inches(2.42),
                                         Inches(0.52), Inches(0.30))
    after_arrow.fill.solid()
    after_arrow.fill.fore_color.rgb = RGBColor(*AFTER_ARROW)
    after_arrow.line.fill.background()
    callout = slide.shapes.add_textbox(Inches(9.82), Inches(1.93), Inches(0.82), Inches(0.30))
    callout.text_frame.text = "CHECK"
    callout.text_frame.paragraphs[0].runs[0].font.size = Pt(8)
    callout.text_frame.paragraphs[0].runs[0].font.bold = True

    if include_inspection:
        _tb(slide, "Cải tiến trong kiểm tra", 1.7, 4.5, 5.4, 0.4, 12)
        _caption(slide, "Sau cải tiến", 7.82, 4.5, 1.65, 0.34)
        _pics(slide, ["#b3e5fc"], 7.82, 4.95, w=1.45, h=1.05)


def _selected_report(path: Path):
    report = parse_pptx(path)
    classification = heuristic_classify(report)
    selection = select_after_pictures(report, classification.improvement_image_slides)
    return report, selection


def _has_rgb(path_or_image, rgb, tolerance=3, minimum=15):
    opened_by_us = not isinstance(path_or_image, Image.Image)
    image = Image.open(path_or_image) if opened_by_us else path_or_image
    try:
        pixels = image.convert("RGB").tobytes()
        count = 0
        for offset in range(0, len(pixels), 3):
            if all(abs(pixels[offset + i] - int(rgb[i])) <= tolerance for i in range(3)):
                count += 1
                if count >= minimum:
                    return True
        return False
    finally:
        if opened_by_us:
            image.close()


def _write_visual_report(path: Path, after_colors=("#2a9a52", "#1c68ba"), include_inspection=False):
    return _deck(path, [lambda slide, width: _make_visual_layout(slide, width, after_colors, include_inspection)])


def test_inherited_theme_fill_and_line_are_included_in_the_after_crop(tmp_path):
    presentation_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    drawing_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    styled_shape_id = {}

    def layout(slide, width):
        _make_visual_layout(slide, width)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.65), Inches(2.28),
                                       Inches(0.40), Inches(0.38))
        styled_shape_id["id"] = shape.shape_id
        for child in list(shape._element.spPr):
            if etree.QName(child).localname in ("solidFill", "noFill", "gradFill", "blipFill", "pattFill",
                                                "grpFill", "ln"):
                shape._element.spPr.remove(child)
        style = shape._element.find(f"{{{presentation_ns}}}style")
        assert style is not None
        for child in list(style):
            style.remove(child)
        for tag, color in (("fillRef", "accent1"), ("lnRef", "accent2")):
            reference = etree.SubElement(style, f"{{{drawing_ns}}}{tag}", idx="2")
            etree.SubElement(reference, f"{{{drawing_ns}}}schemeClr", val=color)

    deck = _deck(tmp_path / "inherited-theme-style.pptx", [layout])
    report, selection = _selected_report(deck)
    annotation = next(block for block in report.slide(3).annotations if block.shape_id == styled_shape_id["id"])
    assert annotation.fill_visible and annotation.fill_color == "#4f81bd"
    assert annotation.line_visible and annotation.line_color == "#c0504d"

    paths, problems = export_after_pictures(report, selection.after, tmp_path / "inherited-style-region",
                                            renderer=SlideRenderer(prefer=("builtin",), width_px=1200),
                                            management_number="260920207-VOC")
    assert len(paths) == 1 and problems == []
    assert _has_rgb(paths[0], (79, 129, 189)) and _has_rgb(paths[0], (192, 80, 77))


def test_rendered_after_region_is_one_complete_crop_without_before_or_transition(tmp_path):
    deck = _write_visual_report(tmp_path / "visual.pptx")
    report, selection = _selected_report(deck)
    assert len(selection.after) == 2
    assert all(ref.excel_output_eligible and ref.temporal_role == "AFTER"
               and ref.semantic_role == SEMANTIC_PRODUCTION and ref.confident_owner for ref in selection.after)

    regions = build_improvement_visual_regions(report, selection.after, management_number="260920201-VOC")
    assert len(regions) == 1
    region = regions[0]
    assert region.anchor_text == "Sau cải tiến"
    assert region.excel_output_eligible and region.temporal_role == "AFTER"
    assert region.semantic_role == SEMANTIC_PRODUCTION
    included_text = " ".join(block.text for block in region.included_objects)
    assert "Sau cải tiến" in included_text and "CHECK" in included_text
    assert any(block.prst == "ellipse" for block in region.included_objects)  # red circle
    assert any(block.prst == "rightArrow" for block in region.included_objects)  # After arrow
    assert all(ref.block.left >= Inches(7.0) for ref in region.pictures)

    paths, problems = export_after_pictures(report, selection.after, tmp_path / "regions",
                                            renderer=SlideRenderer(prefer=("builtin",), width_px=1440),
                                            management_number="260920201-VOC")
    assert isinstance(paths, GroupedImagePaths)
    assert len(paths) == 1 and len(paths.groups) == 1 and len(paths.groups[0]) == 1
    assert problems == []
    crop = paths[0]
    assert _has_rgb(crop, AFTER_A) and _has_rgb(crop, AFTER_B)
    assert _has_rgb(crop, AFTER_MARK) and _has_rgb(crop, AFTER_ARROW) and _has_rgb(crop, AFTER_BORDER)
    assert not _has_rgb(crop, BEFORE_A) and not _has_rgb(crop, BEFORE_B)
    assert not _has_rgb(crop, BEFORE_MARK) and not _has_rgb(crop, TRANSITION)
    with Image.open(crop) as image:
        assert image.width > 0 and image.height > 0
        # Text descenders/caption area are part of the output, not an adjacent Excel label.
        full_height = int(report.slide_height * 1440 / report.slide_width)
        expected_crop_height = int(region.height * full_height / report.slide_height)
        assert image.height >= expected_crop_height
        pixels = image.convert("RGB").tobytes()
        assert any(pixels[i] < 80 and pixels[i + 1] < 80 and pixels[i + 2] < 80
                   for i in range(0, len(pixels), 3))


def test_safe_padding_does_not_pull_in_a_nearby_transition_arrow(tmp_path):
    def layout(slide, width):
        _make_visual_layout(slide, width)
        arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(7.15), Inches(2.22),
                                       Inches(0.32), Inches(0.30))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = RGBColor(*TRANSITION)
        arrow.line.fill.background()

    deck = _deck(tmp_path / "nearby-transition.pptx", [layout])
    report, selection = _selected_report(deck)
    paths, problems = export_after_pictures(report, selection.after, tmp_path / "nearby-transition-region",
                                            renderer=SlideRenderer(prefer=("builtin",), width_px=1200),
                                            management_number="260920203-VOC")
    assert len(paths) == 1 and problems == []
    assert not _has_rgb(paths[0], TRANSITION)


def test_three_authored_after_pictures_make_one_visual_evidence_image(tmp_path):
    deck = _write_visual_report(tmp_path / "three-after.pptx", after_colors=("#2a9a52", "#1c68ba", "#50b48c"))
    report, selection = _selected_report(deck)
    assert len(selection.after) == 3
    paths, problems = export_after_pictures(report, selection.after, tmp_path / "three-regions",
                                            renderer=SlideRenderer(prefer=("builtin",), width_px=1200),
                                            management_number="260920202-VOC")
    assert len(paths) == 1 and len(paths.groups) == 1 and len(paths.groups[0]) == 1
    assert problems == []
    assert all(_has_rgb(paths[0], rgb) for rgb in (AFTER_A, AFTER_B, AFTER_C))


def test_distinct_defects_on_one_slide_keep_separate_after_crops(tmp_path):
    def layout(slide, width):
        _prod_head(slide, width)
        _tb(slide, "Cải tiến lỗi A", 1.7, 1.05, 5.0, 0.42, 16, True)
        _caption(slide, "Sau cải tiến", 7.8, 1.05, 1.7, 0.34)
        _pics(slide, ["#2a9a52"], 7.8, 1.52, w=1.5, h=1.0)
        _tb(slide, "Cải tiến lỗi B", 1.7, 3.8, 5.0, 0.42, 16, True)
        _caption(slide, "Sau cải tiến", 7.8, 3.8, 1.7, 0.34)
        _pics(slide, ["#ffcc80"], 7.8, 4.27, w=1.5, h=1.0)

    deck = _deck(tmp_path / "two-defects.pptx", [layout])
    report, selection = _selected_report(deck)
    assert len(selection.after) == 2 and len({ref.owner_id for ref in selection.after}) == 2
    paths, problems = export_after_pictures(report, selection.after, tmp_path / "two-defect-regions",
                                            renderer=SlideRenderer(prefer=("builtin",), width_px=1200),
                                            management_number="260920206-VOC")
    assert len(paths) == 2 and len(paths.groups) == 2
    assert all(len(group) == 1 for group in paths.groups) and problems == []
    crop_colors = [(_has_rgb(group[0], AFTER_A), _has_rgb(group[0], (255, 204, 128)))
                   for group in paths.groups]
    assert sorted(crop_colors) == [(False, True), (True, False)]


def test_production_after_crop_excludes_inspection_after_on_same_slide(tmp_path):
    deck = _write_visual_report(tmp_path / "mixed-semantics.pptx", include_inspection=True)
    report, selection = _selected_report(deck)
    assert selection.after and all(ref.semantic_role == SEMANTIC_PRODUCTION for ref in selection.after)
    paths, problems = export_after_pictures(report, selection.after, tmp_path / "production-only",
                                            renderer=SlideRenderer(prefer=("builtin",), width_px=1200),
                                            management_number="260920203-VOC")
    assert len(paths) == 1 and len(paths.groups) == 1
    assert problems == []
    assert _has_rgb(paths[0], AFTER_A) and _has_rgb(paths[0], AFTER_B)
    assert not _has_rgb(paths[0], (179, 229, 252))


def test_renderer_failure_falls_back_all_or_none_to_eligible_pictures(tmp_path):
    deck = _write_visual_report(tmp_path / "fallback.pptx")
    report, selection = _selected_report(deck)

    class BrokenRenderer:
        last_backend = "broken-test-renderer"

        def render(self, *_args, **_kwargs):
            raise OSError("renderer unavailable")

    fallback_dir = tmp_path / "fallback-assets"
    paths, problems = export_after_pictures(report, selection.after, fallback_dir,
                                            renderer=BrokenRenderer(),
                                            management_number="260920205-VOC")
    assert len(paths.groups) == 1 and len(paths.groups[0]) == 2
    assert all(path.name.endswith(("_fallback_01.png", "_fallback_02.png")) for path in paths.groups[0])
    assert any("REGION_FALLBACK" in problem and "renderer unavailable" in problem for problem in problems)

    ref = selection.after[0]
    original_blob = ref.block.image_blob
    try:
        ref.block.image_blob = b"not-a-decodable-image"
        failed_paths, failed_problems = export_after_pictures(
            report, selection.after, tmp_path / "failed-fallback-assets", renderer=BrokenRenderer(),
            management_number="260920205-VOC")
    finally:
        ref.block.image_blob = original_blob
    assert len(failed_paths) == 0 and failed_paths.groups == []
    assert any("decode failed" in problem for problem in failed_problems)
    assert any("both failed" in problem for problem in failed_problems)
    assert not list((tmp_path / "failed-fallback-assets").glob("*_fallback_*.png"))


def test_report_scoped_visual_paths_do_not_overwrite_same_basename_reports(tmp_path):
    colors = (("#2a9a52", "#1c68ba"), ("#ffcc80", "#ce93d8"))
    reports = []
    for index, palette in enumerate(colors):
        reports.append(_write_visual_report(tmp_path / f"folder-{index}" / "same-name.pptx", palette))
    outputs = []
    shared_assets = tmp_path / "shared-assets"
    for index, source in enumerate(reports):
        report, selection = _selected_report(source)
        paths, problems = export_after_pictures(report, selection.after, shared_assets,
                                                renderer=SlideRenderer(prefer=("builtin",), width_px=960),
                                                management_number=f"26092020{index + 4}-VOC")
        assert len(paths) == 1 and len(paths.groups) == 1 and problems == []
        outputs.append(paths[0])
    assert outputs[0] != outputs[1] and outputs[0].exists() and outputs[1].exists()
    assert _has_rgb(outputs[0], AFTER_A) and _has_rgb(outputs[0], AFTER_B)
    assert _has_rgb(outputs[1], (255, 204, 128)) and _has_rgb(outputs[1], (206, 147, 216))
    assert len(list(shared_assets.glob("region_*.png"))) == 2
