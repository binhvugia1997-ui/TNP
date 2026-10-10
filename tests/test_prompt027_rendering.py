"""PROMPT-027: PowerPoint-faithful rendering, true After-region preview, No Fill / No Line, version 1.3.4.

Real-PPTX acceptance exposed several DISTINCT problems that must not be conflated (§58):

* A — the full-slide preview looked vertically shifted because the built-in renderer pinned every text
  run to the top-left of its box, ignoring the authored vertical anchor / insets / spacing (§11).
* B — a shape the author set to ``No Fill`` + ``No Line`` was painted as a BLACK rectangle, because
  ``<a:noFill/>`` is reported by python-pptx as ``MSO_FILL.BACKGROUND`` (not ``None``) and the line
  branch tested only ``is not None``; the unresolved colour then fell back to ``#000000`` (§7/§8/§9).
* C/D — the Learning highlight showed ONE picture while Excel exports the whole item-scoped After
  region, so the UI had to expose two clearly different geometries (§13/§15/§16).

Everything here is headless and deterministic: the built-in renderer is forced so no test depends on
PowerPoint or LibreOffice being installed.  PowerPoint-faithful pixels and the final Excel crop on the
REAL report still require the Windows acceptance pass (§56).
"""
from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from types import SimpleNamespace

import pytest
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

import app as app_pkg
import app.qpn_renderer as qrn
from app.image_learning import ImageCandidate
from app.improvement_visual import crop_box_px
from app.pptx_parser import parse_pptx
from app.preview_geometry import fit_slide_box, map_fraction_box_to_px, slide_fraction_box
from app.qpn_renderer import (FAITHFUL_BACKENDS, RENDER_SCHEMA_VERSION, SlideRenderer, _safe_reason,
                              is_faithful_backend, render_builtin, renderer_status)
from tests.test_after_evidence import _deck, _pics, _tb
from tests.test_prompt004 import MGMT
from tests.test_prompt025_multi_item import (NAME, _builtin_renderer, _export, _regions, _report,
                                             _selection, _two_item_slide)

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
EMU = 914400
AFTER1_RGB = (200, 230, 201)          # item #1 After pictures  (#c8e6c9)
AFTER2_RGB = (179, 229, 252)          # item #2 After pictures  (#b3e5fc)


# ================================================================== fixture helpers
def _force_no_fill_no_line(shape):
    """Configure ``shape`` exactly as PowerPoint does for ``No Fill`` + ``No Line``.

    Writes ``<a:noFill/>`` on the shape body and ``<a:ln><a:noFill/></a:ln>`` for the outline — the real
    OOXML the acceptance deck contains, and the structure that used to be misread as a visible black line.
    """
    sp_pr = shape._element.spPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill", "a:ln"):
        for element in sp_pr.findall(qn(tag)):
            sp_pr.remove(element)
    etree.SubElement(sp_pr, qn("a:noFill"))
    line = etree.SubElement(sp_pr, qn("a:ln"))
    etree.SubElement(line, qn("a:noFill"))
    return shape


def _invisible_content_frame(slide, left, top, width, height, text=""):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top),
                                   Inches(width), Inches(height))
    if text:
        shape.text_frame.text = text
        shape.text_frame.word_wrap = True
    return _force_no_fill_no_line(shape)


def _simple_slide(path: Path, build) -> Path:
    """One-slide 16:9 deck (no report furniture) so rendering assertions stay unambiguous."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    build(prs.slides.add_slide(prs.slide_layouts[6]))
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    return path


def _ring_pixels(image: Image.Image, box, band: int = 2):
    """Pixels of the outline band of ``box`` — where a fake stroke would appear (text stays inside)."""
    x0, y0, x1, y1 = box
    out = []
    for x in range(max(0, x0), min(image.width, x1)):
        for y in list(range(max(0, y0), min(image.height, y0 + band))) + \
                list(range(max(0, y1 - band), min(image.height, y1))):
            out.append(image.getpixel((x, y)))
    for y in range(max(0, y0), min(image.height, y1)):
        for x in list(range(max(0, x0), min(image.width, x0 + band))) + \
                list(range(max(0, x1 - band), min(image.width, x1))):
            out.append(image.getpixel((x, y)))
    return out


def _is_dark(pixel, limit: int = 90) -> bool:
    return max(pixel[:3]) < limit


def _emu_box_to_px(box, slide_w, slide_h, image):
    sx, sy = image.width / slide_w, image.height / slide_h
    return (int(box[0] * sx), int(box[1] * sy), int((box[0] + box[2]) * sx), int((box[1] + box[3]) * sy))


def _ink_rows(image, box):
    x0, y0, x1, y1 = box
    rows = [y for y in range(max(0, y0), min(image.height, y1))
            if any(image.getpixel((x, y)) != (255, 255, 255)
                   for x in range(max(0, x0), min(image.width, x1)))]
    return rows


def _annotations(report, slide_number):
    return report.slide(slide_number).annotations


def _find_annotation(report, slide_number, name):
    for block in _annotations(report, slide_number):
        if block.shape_name == name:
            return block
    raise AssertionError(f"annotation {name!r} not found")


# ================================================================== §7/§8 No Fill / No Line — PARSER
def test_explicit_no_fill_no_line_parses_as_invisible_not_as_missing_colour(tmp_path):
    """``No Fill`` + ``No Line`` must mean "paints nothing", never "colour unknown → visible"."""
    deck = _simple_slide(tmp_path / "nofill.pptx",
                         lambda s: _invisible_content_frame(s, 1.0, 3.0, 4.0, 1.5, "Nội dung ẩn"))
    report = parse_pptx(str(deck))
    block = _find_annotation(report, 1, _invisible_shape_name(report, 1))

    assert block.fill_visible is False and block.line_visible is False
    assert block.fill_explicit_none is True and block.line_explicit_none is True
    assert block.fill_color == "" and block.line_color == ""
    assert block.line_width == 0


def _invisible_shape_name(report, slide_number):
    return _annotations(report, slide_number)[0].shape_name


def test_no_line_is_distinguished_from_unspecified_line(tmp_path):
    """Only an AUTHORED No Line suppresses the theme ``lnRef`` fallback; an unspecified style may inherit."""
    def build(slide):
        hidden = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1))
        hidden.name = "Hidden"
        _force_no_fill_no_line(hidden)
        # An explicitly coloured outline must stay visible (no over-suppression).
        shown = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(3), Inches(3), Inches(1))
        shown.name = "Shown"
        shown.fill.solid()
        shown.fill.fore_color.rgb = RGBColor(0xFF, 0x00, 0x00)
        shown.line.color.rgb = RGBColor(0x00, 0x60, 0xC0)
        shown.line.width = Pt(2)

    report = parse_pptx(str(_simple_slide(tmp_path / "mixed.pptx", build)))
    hidden, shown = (_find_annotation(report, 1, n) for n in ("Hidden", "Shown"))

    assert (hidden.fill_visible, hidden.line_visible) == (False, False)
    assert hidden.line_explicit_none is True
    assert shown.fill_visible is True and shown.line_visible is True
    assert shown.fill_color == "#ff0000" and shown.line_color == "#0060c0"
    assert shown.line_width == Pt(2)


def test_background_fill_never_falls_back_to_style_reference(tmp_path):
    """``fill.background()`` is an authored No Fill, so ``fillRef`` inheritance must not resurrect it."""
    def build(slide):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1))
        shape.name = "NoFillRect"
        shape.fill.background()
        shape.line.fill.background()

    report = parse_pptx(str(_simple_slide(tmp_path / "bgfill.pptx", build)))
    block = _find_annotation(report, 1, "NoFillRect")
    assert block.fill_visible is False and block.line_visible is False
    assert block.fill_explicit_none is True and block.line_explicit_none is True


# ================================================================== §9 built-in renderer paints nothing
def test_builtin_renderer_paints_no_black_frame_for_invisible_shape(tmp_path):
    deck = _simple_slide(tmp_path / "nofill.pptx",
                         lambda s: _invisible_content_frame(s, 1.0, 3.0, 4.0, 1.5, "Nội dung ẩn"))
    report = parse_pptx(str(deck))
    slide = report.slide(1)
    image = render_builtin(report, slide, 1200)
    block = _annotations(report, 1)[0]
    box = _emu_box_to_px((block.left, block.top, block.width, block.height),
                         report.slide_width, report.slide_height, image)

    ring = _ring_pixels(image, box)
    assert ring, "expected a measurable outline band"
    assert not any(_is_dark(p) for p in ring), "No Fill / No Line shape painted a dark (fake) outline"
    # the interior is untouched too: no invented solid fill
    centre = image.getpixel(((box[0] + box[2]) // 2, (box[1] + box[3]) // 2 + 6))
    assert not _is_dark(centre) or centre == (0, 0, 0)      # text glyph or white, never a fill block
    assert sum(1 for p in image.convert("RGB").tobytes()[::3] if p < 90) > 0   # the TEXT is still drawn


def test_visible_shape_still_renders_fill_and_outline(tmp_path):
    """The fix must not hide shapes that really are painted (no blanket "hide all rectangles")."""
    def build(slide):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2), Inches(2), Inches(4), Inches(2))
        shape.name = "Visible"
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0x2E, 0x7D, 0x32)
        shape.line.color.rgb = RGBColor(0x10, 0x10, 0x10)
        shape.line.width = Pt(3)

    report = parse_pptx(str(_simple_slide(tmp_path / "visible.pptx", build)))
    image = render_builtin(report, report.slide(1), 1200)
    block = _find_annotation(report, 1, "Visible")
    box = _emu_box_to_px((block.left, block.top, block.width, block.height),
                         report.slide_width, report.slide_height, image)
    raw = image.crop(box).convert("RGB").tobytes()
    triples = [(raw[i], raw[i + 1], raw[i + 2]) for i in range(0, len(raw), 3)]
    green = sum(1 for p in triples if p[1] > 90 and p[1] > p[0] + 25 and p[1] > p[2] + 25)
    assert green > len(triples) * 0.5, "visible fill was not painted"
    assert any(_is_dark(p) for p in _ring_pixels(image, box)), "visible outline was not painted"


def test_grouped_child_style_never_leaks_into_invisible_child(tmp_path):
    """§10: a group/parent style must not turn an invisible child into a visible black rectangle."""
    def build(slide):
        group = slide.shapes.add_group_shape()
        hidden = group.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1.5))
        hidden.name = "GroupedHidden"
        _force_no_fill_no_line(hidden)
        shown = group.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5), Inches(1), Inches(3), Inches(1.5))
        shown.name = "GroupedShown"
        shown.fill.solid()
        shown.fill.fore_color.rgb = RGBColor(0xC6, 0x28, 0x28)
        shown.line.color.rgb = RGBColor(0x00, 0x00, 0x00)

    report = parse_pptx(str(_simple_slide(tmp_path / "group.pptx", build)))
    slide = report.slide(1)
    hidden = _find_annotation(report, 1, "GroupedHidden")
    shown = _find_annotation(report, 1, "GroupedShown")
    assert (hidden.fill_visible, hidden.line_visible) == (False, False)
    assert shown.fill_visible is True and shown.line_visible is True
    # PROMPT-011 grouped geometry must survive: the child transform places both boxes apart
    assert hidden.left + hidden.width <= shown.left

    image = render_builtin(report, slide, 1200)
    hidden_box = _emu_box_to_px((hidden.left, hidden.top, hidden.width, hidden.height),
                                report.slide_width, report.slide_height, image)
    shown_box = _emu_box_to_px((shown.left, shown.top, shown.width, shown.height),
                               report.slide_width, report.slide_height, image)
    assert not any(_is_dark(p) for p in _ring_pixels(image, hidden_box))
    assert any(_is_dark(p) for p in _ring_pixels(image, shown_box))


# ================================================================== §34 full-slide geometry — no trimming
def test_full_slide_render_preserves_aspect_whitespace_and_object_position(tmp_path):
    """The bitmap is the COMPLETE authored slide: same plane as the EMU bbox, nothing trimmed or moved."""
    picture_top_inches, picture_left_inches = 5.4, 9.0        # deliberately LOW and RIGHT on the slide

    def build(slide):
        _tb(slide, "Tiêu đề trên cùng", 0.5, 0.3, 6, 0.5, 20, True)
        _pics(slide, ["#c8e6c9"], picture_left_inches, picture_top_inches, w=2.0, h=1.2)

    deck = _simple_slide(tmp_path / "geom.pptx", build)
    report = parse_pptx(str(deck))
    slide_w, slide_h = report.slide_width, report.slide_height
    width_px = 1600
    image = render_builtin(report, report.slide(1), width_px)

    # 1. exact requested width, height derived from the slide aspect ratio (never letterboxed/trimmed)
    assert image.width == width_px
    assert image.height == pytest.approx(int(slide_h * width_px / slide_w), abs=1)
    assert image.width / image.height == pytest.approx(slide_w / slide_h, rel=0.01)

    # 2. authored whitespace preserved: the empty margin bands are still blank slide, not cropped away
    assert all(image.getpixel((x, 2)) == (255, 255, 255) for x in range(0, image.width, 97))
    assert all(image.getpixel((2, y)) == (255, 255, 255) for y in range(0, image.height, 97))

    # 3. object pixel position matches the slide geometry within tolerance
    picture = [b for b in report.slide(1).blocks if b.kind == "picture"][0]
    box = _emu_box_to_px((picture.left, picture.top, picture.width, picture.height), slide_w, slide_h, image)
    assert box[1] == pytest.approx(int(picture_top_inches * EMU * image.height / slide_h), abs=2)
    assert box[0] == pytest.approx(int(picture_left_inches * EMU * image.width / slide_w), abs=2)
    inside = image.crop(box).convert("RGB").tobytes()
    matches = sum(1 for i in range(0, len(inside), 3) if tuple(inside[i:i + 3]) == AFTER1_RGB)
    assert matches > (len(inside) // 3) * 0.5


def test_rendered_slide_is_never_white_trimmed_or_repositioned(tmp_path):
    """§23/§27: the preview/evidence bitmap keeps the full slide plane — no content-aware auto-trim."""
    def build(slide):
        _pics(slide, ["#c8e6c9"], 5.5, 3.2, w=2.0, h=1.4)      # one small object, lots of whitespace

    report = parse_pptx(str(_simple_slide(tmp_path / "trim.pptx", build)))
    renderer = _builtin_renderer()
    with pytest.MonkeyPatch.context() as mp:
        mp.setattr(qrn, "trim_white_margins", lambda im, **kw: pytest.fail("renderer must not white-trim"))
        out = renderer.render(report, [1], tmp_path / "out", purpose="learning_preview")
    with Image.open(out[1]) as image:
        assert image.width == 1200
        assert image.height == pytest.approx(int(report.slide_height * 1200 / report.slide_width), abs=1)


# ================================================================== §11 vertical geometry / text position
@pytest.mark.parametrize("anchor,expected", [("t", "top"), ("ctr", "centre"), ("b", "bottom")])
def test_builtin_renderer_honours_vertical_anchor(tmp_path, anchor, expected):
    """The authored vertical anchor decides WHERE text sits in its box — no fixed offsets anywhere."""
    def build(slide):
        box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(6), Inches(3))
        tf = box.text_frame
        tf.text = "Nội dung căn giữa"
        tf.word_wrap = True
        tf.paragraphs[0].runs[0].font.size = Pt(20)
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "ctr": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
        body = tf._txBody.find(qn("a:bodyPr"))
        for tag in ("a:spAutoFit", "a:normAutofit", "a:noAutofit"):
            element = body.find(qn(tag))
            if element is not None:
                body.remove(element)

    report = parse_pptx(str(_simple_slide(tmp_path / f"anchor_{anchor}.pptx", build)))
    slide = report.slide(1)
    block = [b for b in slide.blocks if b.kind == "paragraph"][0]
    assert block.vertical_anchor == anchor
    image = render_builtin(report, slide, 1200)
    box = _emu_box_to_px((block.left, block.top, block.width, block.height),
                         report.slide_width, report.slide_height, image)
    rows = _ink_rows(image, box)
    assert rows, "text was not rendered"
    ink_centre = (rows[0] + rows[-1]) / 2
    box_centre = (box[1] + box[3]) / 2
    box_height = box[3] - box[1]
    offset = (ink_centre - box_centre) / box_height
    if expected == "top":
        assert offset < -0.2
    elif expected == "centre":
        assert abs(offset) < 0.12
    else:
        assert offset > 0.2


def test_text_insets_and_spacing_come_from_the_pptx(tmp_path):
    """Insets / line spacing are read from the file, not assumed — the fallback renderer stays honest."""
    def build(slide):
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(3))
        tf = box.text_frame
        tf.text = "Dòng một"
        tf.word_wrap = True
        tf.margin_left = Inches(0.9)
        tf.margin_top = Inches(0.7)
        tf.margin_right = Inches(0.3)
        tf.margin_bottom = Inches(0.4)
        tf.paragraphs[0].line_spacing = 1.75
        tf.paragraphs[0].space_before = Pt(10)

    report = parse_pptx(str(_simple_slide(tmp_path / "insets.pptx", build)))
    block = [b for b in report.slide(1).blocks if b.kind == "paragraph"][0]
    assert block.inset_left == Inches(0.9) and block.inset_top == Inches(0.7)
    assert block.inset_right == Inches(0.3) and block.inset_bottom == Inches(0.4)
    assert block.line_spacing[0] == pytest.approx(1.75)
    assert block.space_before[0] == pytest.approx(10.0)

    image = render_builtin(report, report.slide(1), 1200)
    box = _emu_box_to_px((block.left, block.top, block.width, block.height),
                         report.slide_width, report.slide_height, image)
    rows = _ink_rows(image, box)
    # a 0.7" top inset on a 7.5" slide at 1200px ≈ 112px: the text must start clearly below the box top
    assert rows[0] - box[1] > 60


def test_autofit_font_scale_is_read_from_normautofit(tmp_path):
    def build(slide):
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        tf = box.text_frame
        tf.text = "Chữ bị thu nhỏ"
        body = tf._txBody.find(qn("a:bodyPr"))
        autofit = etree.SubElement(body, qn("a:normAutofit"))
        autofit.set("fontScale", "62000")

    report = parse_pptx(str(_simple_slide(tmp_path / "autofit.pptx", build)))
    block = [b for b in report.slide(1).blocks if b.kind == "paragraph"][0]
    assert block.autofit_scale == pytest.approx(0.62)


def test_no_global_position_hack_exists_in_the_renderer():
    """§12: no fixed X/Y offset, percentage correction, slide- or vendor-specific special case."""
    source = Path(qrn.__file__).read_text(encoding="utf-8")
    lowered = source.lower()
    for forbidden in ("translatey", "magic offset", "slide_number ==", "if slide ==", "vendor"):
        assert forbidden not in lowered, f"renderer contains a forbidden correction: {forbidden!r}"
    # every text position derives from the parsed box + insets + anchor
    assert "b.inset_top" in source and "b.vertical_anchor" in source


# ================================================================== §26 EMU → pixel crop conversion
def test_crop_box_px_uses_the_actual_rendered_dimensions_not_a_hardcoded_width():
    slide_w, slide_h = 12192000, 6858000
    bbox = (3048000, 1714500, 9144000, 5143500)            # left/top 25% → right/bottom 75%
    for image_w, image_h in ((1600, 900), (1920, 1080), (1234, 694), (800, 450)):
        left, top, right, bottom = crop_box_px(bbox, slide_w, slide_h, image_w, image_h)
        assert left == pytest.approx(image_w * 0.25, abs=1)
        assert top == pytest.approx(image_h * 0.25, abs=1)
        assert right == pytest.approx(image_w * 0.75, abs=1)
        assert bottom == pytest.approx(image_h * 0.75, abs=1)
        assert 0 <= left < right <= image_w and 0 <= top < bottom <= image_h


def test_crop_box_px_clamps_out_of_slide_geometry_and_rejects_degenerate_input():
    slide_w, slide_h = 12192000, 6858000
    left, top, right, bottom = crop_box_px((-1000000, -1000000, slide_w * 2, slide_h * 2),
                                           slide_w, slide_h, 1600, 900)
    assert (left, top, right, bottom) == (0, 0, 1600, 900)
    with pytest.raises(ValueError):
        crop_box_px((0, 0, 0, 0), slide_w, slide_h, 1600, 900)      # empty region → fail loudly
    with pytest.raises(ValueError):
        crop_box_px((0, 0, 100, 100), 0, 0, 1600, 900)


def test_crop_and_bbox_share_one_coordinate_plane(tmp_path):
    """§27: the crop is taken from the FULL-slide bitmap, so region EMU → px is a pure linear map."""
    deck = _simple_slide(tmp_path / "plane.pptx", lambda s: _pics(s, ["#c8e6c9"], 4.0, 2.5, w=3.0, h=2.0))
    report = parse_pptx(str(deck))
    slide_w, slide_h = report.slide_width, report.slide_height
    picture = [b for b in report.slide(1).blocks if b.kind == "picture"][0]
    renderer = _builtin_renderer()
    out = renderer.render(report, [1], tmp_path / "out", purpose="after_evidence")
    with Image.open(out[1]) as full:
        full = full.convert("RGB")
        box = crop_box_px((picture.left, picture.top, picture.right, picture.bottom),
                          slide_w, slide_h, full.width, full.height)
        crop = full.crop(box)
    raw = crop.convert("RGB").tobytes()
    matches = sum(1 for i in range(0, len(raw), 3) if tuple(raw[i:i + 3]) == AFTER1_RGB)
    assert matches > (len(raw) // 3) * 0.8


# ================================================================== §35 picture vs final evidence region
def test_evidence_region_is_larger_than_one_picture_candidate(tmp_path):
    """§21/§35: three After pictures of one item form ONE region — the candidate bbox is not the region."""
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    regions = _regions(report, selection)
    assert len(regions) == 2

    item1 = regions[0]
    assert len(item1.pictures) == 3, "item #1 has three After pictures"
    for ref in item1.pictures:
        pic = ref.block
        # every single picture is strictly SMALLER than the region that owns it
        assert (pic.right - pic.left) < item1.width
        # and each picture lies fully INSIDE its own item's region
        assert item1.bbox[0] <= pic.left and pic.right <= item1.bbox[2]
        assert item1.bbox[1] <= pic.top and pic.bottom <= item1.bbox[3]
    picture_ids = {ref.block.shape_id for ref in item1.pictures}
    assert len(picture_ids) == 3


def test_region_bbox_and_single_picture_bbox_differ_as_dto_geometries(tmp_path):
    """The two DTO geometries the UI must keep apart really are different numbers for a real item."""
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    region = _regions(report, selection)[0]
    slide_w, slide_h = report.slide_width, report.slide_height

    picture = region.pictures[0].block
    pic_box = (picture.left, picture.top, picture.width, picture.height)
    reg_box = (region.bbox[0], region.bbox[1], region.width, region.height)
    assert pic_box != reg_box
    assert reg_box[2] > pic_box[2] and reg_box[3] >= pic_box[3]

    pic_pct = slide_fraction_box(pic_box, slide_w, slide_h)
    reg_pct = slide_fraction_box(reg_box, slide_w, slide_h)
    assert reg_pct[2] > pic_pct[2]
    assert all(0.0 <= v <= 1.0 for v in pic_pct + reg_pct)


# ================================================================== §36/§20/§22 multi-item region ownership
def test_two_items_two_regions_two_owners(tmp_path):
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    regions = _regions(report, selection)

    assert len(regions) == 2
    ids = {r.improvement_item_id for r in regions}
    assert len(ids) == 2, "each improvement item owns its own region"
    assert sorted(r.item_index for r in regions) == [0, 1]
    assert regions[0].slide_index == regions[1].slide_index == 3
    assert regions[0].temporal_role == "AFTER" and regions[1].temporal_role == "AFTER"
    assert all(r.excel_output_eligible for r in regions)


def test_no_cross_item_contamination_in_region_geometry_or_pixels(tmp_path):
    """§37: item #1's crop excludes item #2 and vice versa — proximity alone never merges evidence."""
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    regions = _regions(report, selection)
    first, second = regions

    # geometry: the two regions never overlap vertically on the slide
    assert first.bbox[3] <= second.bbox[1] or second.bbox[3] <= first.bbox[1]

    grouped, problems = _export(report, selection, tmp_path / "crops", renderer=_builtin_renderer())
    assert problems == [] and len(grouped.groups) == 2
    crop1, crop2 = grouped.groups[0][0], grouped.groups[1][0]

    def _count(path, rgb, tolerance=6):
        with Image.open(path) as handle:
            data = handle.convert("RGB").tobytes()
        return sum(1 for i in range(0, len(data), 3)
                   if all(abs(data[i + c] - rgb[c]) <= tolerance for c in range(3)))

    assert _count(crop1, AFTER1_RGB) > 100, "item #1 crop lost its own After pictures"
    assert _count(crop1, AFTER2_RGB) == 0, "item #1 crop leaked item #2 evidence"
    assert _count(crop2, AFTER2_RGB) > 100, "item #2 crop lost its own After pictures"
    assert _count(crop2, AFTER1_RGB) == 0, "item #2 crop leaked item #1 evidence"


def test_every_region_picture_belongs_to_its_own_item(tmp_path):
    """§22: ownership is by item, so no region absorbs the other item's pictures."""
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    regions = _regions(report, selection)
    seen = set()
    for region in regions:
        for ref in region.pictures:
            assert ref.owner_id == region.improvement_item_id
            assert ref.temporal_role == "AFTER" and ref.excel_output_eligible
            key = (region.slide_index, ref.block.shape_id)
            assert key not in seen, "one picture was absorbed by two regions"
            seen.add(key)
    assert len(seen) == 5                                   # 3 pictures (item #1) + 2 (item #2)


# ================================================================== §38/§24 authored whitespace and padding
def test_authored_whitespace_between_after_pictures_is_preserved(tmp_path):
    """§23/§38: the region spans the authored gap; no non-white-pixel trimming tightens it."""
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    region = _regions(report, selection)[0]
    pictures = sorted((ref.block for ref in region.pictures), key=lambda b: b.left)
    assert len(pictures) >= 2
    gap_left = pictures[0].right
    gap_right = pictures[1].left
    assert gap_right > gap_left, "fixture must author a real gap between two After pictures"
    # the gap lies strictly INSIDE the region, so the exported crop keeps the authored spacing
    assert region.bbox[0] < gap_left and gap_right < region.bbox[2]

    grouped, problems = _export(report, selection, tmp_path / "ws", renderer=_builtin_renderer())
    assert problems == []
    with Image.open(grouped.groups[0][0]) as crop:
        crop = crop.convert("RGB")
        sx = crop.width / region.width
        x0, x1 = int((gap_left - region.bbox[0]) * sx), int((gap_right - region.bbox[0]) * sx)
        assert x1 - x0 > 3, "authored gap collapsed in the crop"
        band = [crop.getpixel((x, crop.height // 2)) for x in range(x0 + 2, max(x0 + 3, x1 - 2))]
        assert any(p == (255, 255, 255) or max(p) > 235 for p in band), \
            "authored whitespace between the two After pictures was trimmed away"


def test_region_padding_is_bounded_and_does_not_reach_the_next_item(tmp_path):
    """§24: padding is deterministic, geometry-based and bounded — never report-specific magic numbers."""
    from app.improvement_visual import DEFAULT_SAFE_PADDING_FRACTION
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    regions = _regions(report, selection)
    slide_w, slide_h = report.slide_width, report.slide_height

    for region in regions:
        union = [min(r.block.left for r in region.pictures), min(r.block.top for r in region.pictures),
                 max(r.block.right for r in region.pictures), max(r.block.bottom for r in region.pictures)]
        pad_left = union[0] - region.bbox[0]
        pad_top = union[1] - region.bbox[1]
        pad_right = region.bbox[2] - union[2]
        pad_bottom = region.bbox[3] - union[3]
        assert all(p >= 0 for p in (pad_left, pad_top, pad_right, pad_bottom))
        # bounded by the documented slide-relative fraction plus the associated caption/annotation span
        assert pad_left <= slide_w * DEFAULT_SAFE_PADDING_FRACTION + 1
        assert pad_right <= slide_w * DEFAULT_SAFE_PADDING_FRACTION + 1
        assert pad_bottom <= slide_h * 0.30
    # deterministic: rebuilding yields byte-identical geometry
    again = _regions(report, selection)
    assert [r.bbox for r in again] == [r.bbox for r in regions]
    assert regions[0].bbox[3] <= regions[1].bbox[1], "item #1 padding grew into item #2"


def test_region_geometry_is_never_pixel_trimmed(tmp_path):
    """§23: the crop is pure geometry — a region of blank slide stays blank, it is not shrunk to ink."""
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    region = _regions(report, selection)[0]
    renderer = _builtin_renderer()
    out = renderer.render(report, [3], tmp_path / "r", purpose="after_evidence")
    with Image.open(out[3]) as full:
        full = full.convert("RGB")
        box = crop_box_px(region.bbox, report.slide_width, report.slide_height, full.width, full.height)
        crop = full.crop(box)
    expected_w = box[2] - box[0]
    expected_h = box[3] - box[1]
    assert crop.size == (expected_w, expected_h), "crop was resized/trimmed away from the region geometry"


# ================================================================== §31 semantic gate preserved
def test_before_and_ineligible_pictures_never_form_a_region(tmp_path):
    """§31: the AFTER / PRODUCTION_IMPROVEMENT / confident-owner / excel-eligible gate still decides."""
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    regions = _regions(report, selection)
    owned = {ref.block.shape_id for region in regions for ref in region.pictures}
    before_ids = {b.shape_id for b in report.slide(3).blocks
                  if b.kind == "picture" and b.shape_id not in owned}
    assert before_ids, "fixture must contain Before pictures that are excluded"
    for region in regions:
        assert region.temporal_role == "AFTER"
        assert region.semantic_role == "PRODUCTION_IMPROVEMENT"
        assert not (owned & before_ids)
    # rejected refs are still reported, never silently promoted
    assert all(not (r.excel_output_eligible and r.temporal_role == "AFTER"
                    and r.semantic_role == "PRODUCTION_IMPROVEMENT" and r.confident_owner and r.owner_id)
               for r in selection.rejected)


# ================================================================== §3 render backend diagnostics
def test_slide_render_logs_purpose_backend_and_dimensions(tmp_path, caplog):
    deck = _simple_slide(tmp_path / "log.pptx", lambda s: _pics(s, ["#c8e6c9"], 4, 3, w=2, h=1.4))
    report = parse_pptx(str(deck))
    renderer = _builtin_renderer()
    with caplog.at_level(logging.INFO, logger="report_extractor.renderer"):
        renderer.render(report, [1], tmp_path / "a", purpose="learning_preview")
        renderer.render(report, [1], tmp_path / "b", purpose="after_evidence")

    renders = [r.getMessage() for r in caplog.records if r.getMessage().startswith("SLIDE_RENDER ")]
    assert len(renders) == 2
    assert "purpose=learning_preview" in renders[0] and "purpose=after_evidence" in renders[1]
    for line in renders:
        assert "backend=builtin" in line and "slide=1" in line
    for line in renders:
        assert "faithful=false" in line, "builtin must be reported as reduced fidelity"
        assert "rendered=1200x675" in line
        assert f"slide_emu={report.slide_width}x{report.slide_height}" in line
        assert f"schema=v{RENDER_SCHEMA_VERSION}" in line


def test_backend_failure_is_logged_not_silently_swallowed(tmp_path, caplog):
    deck = _simple_slide(tmp_path / "fail.pptx", lambda s: _pics(s, ["#c8e6c9"], 4, 3, w=2, h=1.4))
    report = parse_pptx(str(deck))
    renderer = SlideRenderer(prefer=("powerpoint", "builtin"), width_px=1200)
    with caplog.at_level(logging.WARNING, logger="report_extractor.renderer"):
        out = renderer.render(report, [1], tmp_path / "o", purpose="learning_preview")

    assert out and renderer.last_backend == "builtin"
    failures = [r.getMessage() for r in caplog.records
                if r.getMessage().startswith("SLIDE_RENDER_BACKEND_FAILED")]
    assert len(failures) == 1
    assert "purpose=learning_preview backend=powerpoint" in failures[0]
    assert "reason=unavailable" in failures[0]


def test_backend_exception_reason_is_safe_and_path_free(tmp_path):
    report = SimpleNamespace(path=Path(r"D:\secret\reports\real_report.pptx"), filename="real_report.pptx")
    reason = _safe_reason(report, RuntimeError(r"cannot open D:\secret\reports\real_report.pptx"))
    assert reason.startswith("RuntimeError:")
    assert "D:\\secret" not in reason and "real_report.pptx" not in reason
    assert "<report>" in reason
    assert len(reason) < 220
    assert _safe_reason(report, ValueError()) == "ValueError"


def test_every_backend_failure_path_is_diagnosed(tmp_path, caplog, monkeypatch):
    """A backend that raises must be reported before the chain falls through (§3)."""
    deck = _simple_slide(tmp_path / "raise.pptx", lambda s: _pics(s, ["#c8e6c9"], 4, 3, w=2, h=1.4))
    report = parse_pptx(str(deck))
    monkeypatch.setattr(qrn, "powerpoint_available", lambda: True)

    def boom(*_a, **_k):
        raise RuntimeError("COM server busy")

    monkeypatch.setattr(qrn, "render_with_powerpoint", boom)
    renderer = SlideRenderer(prefer=("powerpoint", "builtin"), width_px=1200)
    with caplog.at_level(logging.WARNING, logger="report_extractor.renderer"):
        renderer.render(report, [1], tmp_path / "o", purpose="after_evidence")
    failures = [r.getMessage() for r in caplog.records
                if r.getMessage().startswith("SLIDE_RENDER_BACKEND_FAILED")]
    assert any("backend=powerpoint" in f and "COM server busy" in f for f in failures)


# ================================================================== §4/§5/§6 fidelity classification
def test_only_powerpoint_is_a_faithful_backend():
    assert FAITHFUL_BACKENDS == frozenset({"powerpoint"})
    assert is_faithful_backend("powerpoint") is True
    for backend in ("libreoffice", "builtin", "", "unknown"):
        assert is_faithful_backend(backend) is False
    assert set(renderer_status()) == {"powerpoint", "libreoffice", "builtin"}


def test_powerpoint_export_is_used_when_available(tmp_path, monkeypatch):
    """§5: on Windows with PowerPoint, COM export wins and the result is the complete authored slide."""
    deck = _simple_slide(tmp_path / "pp.pptx", lambda s: _pics(s, ["#c8e6c9"], 4, 3, w=2, h=1.4))
    report = parse_pptx(str(deck))
    calls = {}

    def fake_powerpoint(pptx, slide_numbers, out_dir, width_px=1920):
        calls["width_px"] = width_px
        out = {}
        for n in slide_numbers:
            target = Path(out_dir) / f"slide_{n:03d}.png"
            # a stand-in "export": the FULL slide plane at the requested width, aspect preserved
            height = max(1, int(report.slide_height * width_px / report.slide_width))
            Image.new("RGB", (width_px, height), "white").save(target, "PNG")
            out[n] = target
        return out

    monkeypatch.setattr(qrn, "powerpoint_available", lambda: True)
    monkeypatch.setattr(qrn, "render_with_powerpoint", fake_powerpoint)
    renderer = SlideRenderer(width_px=1600)
    out = renderer.render(report, [1], tmp_path / "o", purpose="learning_preview")
    assert renderer.last_backend == "powerpoint" and calls["width_px"] == 1600
    with Image.open(out[1]) as image:
        assert image.width == 1600
        assert image.height == pytest.approx(int(report.slide_height * 1600 / report.slide_width), abs=1)
    assert is_faithful_backend(renderer.last_backend) is True


def test_fallback_chain_still_works_when_higher_backends_are_unavailable(tmp_path, monkeypatch):
    """§6/§28: PowerPoint → LibreOffice → built-in remains supported; nothing fails the report."""
    deck = _simple_slide(tmp_path / "chain.pptx", lambda s: _pics(s, ["#c8e6c9"], 4, 3, w=2, h=1.4))
    report = parse_pptx(str(deck))
    monkeypatch.setattr(qrn, "powerpoint_available", lambda: False)
    monkeypatch.setattr(qrn, "find_soffice", lambda: None)
    renderer = SlideRenderer()
    out = renderer.render(report, [1], tmp_path / "o", purpose="learning_preview")
    assert renderer.last_backend == "builtin" and out[1].exists()


# ================================================================== §29 cache identity / report isolation
def _candidate(deck: Path, picture_id=11, slide=3, item_index=0, owner="item:one"):
    W, H = int(13.333 * EMU), int(7.5 * EMU)
    return ImageCandidate(
        candidate_id=f"{MGMT}|S{slide}|#{picture_id}|1", management_number=MGMT,
        source_name=deck.name, source_file=str(deck), slide=slide, picture_id=picture_id, order=1,
        bounds=(1554480, 2514600, 1188720, 777240), slide_size=(W, H), block_bounds=(0, 0, W, H),
        features={}, evidence=[], confidence=0.9, logical_item_owner=owner,
        owner_heading="Mục một", item_index=item_index)


def _candidate_from_region(deck: Path, report, region, index: int = 0) -> ImageCandidate:
    """A learning candidate for the ``index``-th picture a REAL production region owns."""
    picture = sorted(region.pictures, key=lambda ref: ref.block.left)[index].block
    return ImageCandidate(
        candidate_id=f"{MGMT}|S{region.slide_index}|#{picture.shape_id}|1", management_number=MGMT,
        source_name=deck.name, source_file=str(deck), slide=region.slide_index,
        picture_id=picture.shape_id, order=index + 1,
        bounds=(picture.left, picture.top, picture.width, picture.height),
        slide_size=(report.slide_width, report.slide_height),
        block_bounds=(0, 0, report.slide_width, report.slide_height),
        features={}, evidence=[], confidence=0.9,
        logical_item_owner=region.improvement_item_id, owner_heading=region.improvement_heading,
        item_index=region.item_index)


def _preview_service():
    from app.application_service import ApplicationService
    service = ApplicationService.__new__(ApplicationService)
    service._slide_preview_cache = {}
    service._preview_cache = {}
    return service


def test_preview_cache_survives_renderer_schema_and_backend_change(tmp_path):
    """§29: a preview rendered by older geometry must never be reused after a renderer/backend change."""
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    service = _preview_service()
    candidate = _candidate(deck)

    first = service._slide_preview_for(candidate)
    assert first["src"].startswith("data:image/jpeg;base64,") and first["backend"] == "builtin"
    assert len(service._slide_preview_cache) == 1
    assert service._slide_preview_for(candidate)["src"] == first["src"]       # cache hit, same bytes
    assert len(service._slide_preview_cache) == 1

    # a renderer schema bump invalidates the entry instead of serving the stale (wrong) preview
    import app.application_service as svc
    with pytest.MonkeyPatch.context() as mp:
        mp.setattr(qrn, "RENDER_SCHEMA_VERSION", RENDER_SCHEMA_VERSION + 1)
        bumped = service._slide_preview_for(candidate)
    assert len(service._slide_preview_cache) == 2
    assert bumped["src"].startswith("data:image/jpeg;base64,")

    # a different backend preference chain is a different cache identity too
    class OnlyBuiltin(SlideRenderer):
        def __init__(self, width_px=1600, **kw):
            super().__init__(prefer=("builtin",), width_px=width_px)

    with pytest.MonkeyPatch.context() as mp:
        mp.setattr(qrn, "SlideRenderer", OnlyBuiltin)
        service._slide_preview_for(candidate)
    assert len(service._slide_preview_cache) == 3
    assert svc.LEARNING_PREVIEW_WIDTH_PX == 1600


def test_preview_cache_isolated_per_report_and_invalidated_by_mtime(tmp_path):
    """§29/§39 + PROMPT-020: identical layouts in different reports never share a cached render."""
    deck_a = _deck(tmp_path / "a" / NAME, [_two_item_slide()])
    deck_b = _deck(tmp_path / "b" / NAME, [_two_item_slide()])
    service = _preview_service()

    service._slide_preview_for(_candidate(deck_a))
    assert len(service._slide_preview_cache) == 1
    service._slide_preview_for(_candidate(deck_a, picture_id=16, item_index=1, owner="item:two"))
    assert len(service._slide_preview_cache) == 1, "one render must be shared by all candidates of a slide"
    service._slide_preview_for(_candidate(deck_b))
    assert len(service._slide_preview_cache) == 2, "cross-report cache reuse (PROMPT-020 regression)"

    before = len(service._slide_preview_cache)
    stat = deck_a.stat()
    os.utime(deck_a, (stat.st_atime + 500, stat.st_mtime + 500))
    service._slide_preview_for(_candidate(deck_a))
    assert len(service._slide_preview_cache) == before + 1, "changing the source mtime must invalidate"


def test_evidence_region_truth_survives_a_failed_preview_render(tmp_path, monkeypatch):
    """§6/§15: a failed/reduced preview still reports the real evidence geometry, and is retried later."""
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    item1 = _regions(report, selection)[0]
    candidate = _candidate_from_region(deck, report, item1, index=0)

    service = _preview_service()

    def boom(*_args, **_kwargs):
        raise RuntimeError("renderer exploded")

    monkeypatch.setattr(SlideRenderer, "render", boom)
    preview = service._slide_preview_for(candidate)
    assert preview["src"] == "" and preview["backend"] == "" and preview["faithful"] is False
    assert preview["regions"], "the region truth must survive a render failure"
    region = service._evidence_region_for(candidate, preview["regions"])
    assert region is not None
    assert region["bbox"] == {"x": item1.bbox[0], "y": item1.bbox[1],
                              "width": item1.width, "height": item1.height}
    assert region["itemId"] == item1.improvement_item_id
    # a raising render is NOT cached, so the preview is retried once the backend recovers
    assert len(service._slide_preview_cache) == 0

    monkeypatch.undo()
    recovered = service._slide_preview_for(candidate)
    assert recovered["src"].startswith("data:image/jpeg;base64,")
    assert recovered["regions"] == preview["regions"]


def test_preview_cache_never_leaks_a_filesystem_path(tmp_path):
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    service = _preview_service()
    service._slide_preview_for(_candidate(deck))
    dump = json.dumps(list(service._slide_preview_cache.values()), ensure_ascii=False, default=str)
    for key in service._slide_preview_cache:
        for part in key:
            if isinstance(part, str) and "/" in part:
                assert part not in dump


# ================================================================== §13/§15 candidate vs evidence in the DTO
def test_dto_exposes_backend_and_authoritative_evidence_region(sample_tree, tmp_path):
    from tests.test_prompt024r_learning_workspace import _service_with_candidates

    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    regions = _regions(report, selection)
    item1 = regions[0]
    candidate = _candidate_from_region(deck, report, item1, index=2)     # ONE of the three pictures

    service, _ = _service_with_candidates(sample_tree, tmp_path, [candidate])
    dto = service.learning_state()["images"][0]

    # §4/§6: backend metadata, and builtin must be flagged as NOT pixel-faithful
    assert dto["slidePreviewBackend"] == "builtin"
    assert dto["slidePreviewFaithful"] is False
    assert dto["slidePreview"].startswith("data:image/jpeg;base64,")

    # §14/§15: the evidence region is the production ImprovementVisualRegion bbox, in EMU and in %
    assert dto["targetKind"] == "picture"
    assert dto["evidenceRegionKind"] == "after_region"
    assert dto["evidenceRegionBbox"] == {"x": item1.bbox[0], "y": item1.bbox[1],
                                         "width": item1.width, "height": item1.height}
    assert dto["evidenceRegionItemId"] == item1.improvement_item_id
    assert dto["evidenceRegionItemIndex"] == item1.item_index
    assert dto["evidenceRegionPictureCount"] == 3
    pct = dto["evidenceRegionBboxPct"]
    expected = slide_fraction_box((item1.bbox[0], item1.bbox[1], item1.width, item1.height),
                                  report.slide_width, report.slide_height)
    assert pct["x"] == pytest.approx(expected[0] * 100, abs=0.01)
    assert pct["w"] == pytest.approx(expected[2] * 100, abs=0.01)

    # §13: the two geometries are DIFFERENT — the region is strictly larger than the one picture
    target = dto["targetBbox"]
    assert dto["evidenceRegionBbox"]["width"] > target["width"]
    assert dto["evidenceRegionBbox"] != {**target}
    assert dto["itemId"] == item1.improvement_item_id and dto["itemIndex"] == item1.item_index

    # §30: candidate identity is unchanged by the new additive metadata
    assert dto["id"] == candidate.candidate_id
    # §39/§42: still no filesystem path anywhere in the DTO
    assert str(Path(deck).resolve()) not in json.dumps(dto, ensure_ascii=False)
    assert str(tmp_path) not in json.dumps(dto, ensure_ascii=False)


def test_dto_evidence_region_is_empty_when_no_region_owns_the_picture(sample_tree, tmp_path):
    """A candidate that would not reach Excel must NOT borrow another item's region (§20/§31)."""
    from tests.test_prompt024r_learning_workspace import _service_with_candidates

    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    # a picture id that no eligible After region owns (and that the slide does not even contain)
    assert not any(b.shape_id == 999999 for b in report.slide(3).blocks)
    candidate = _candidate(deck, picture_id=999999)
    service, _ = _service_with_candidates(sample_tree, tmp_path, [candidate])
    dto = service.learning_state()["images"][0]
    assert dto["evidenceRegionBbox"] is None
    assert dto["evidenceRegionBboxPct"] is None
    assert dto["evidenceRegionKind"] == ""
    assert dto["evidenceRegionPictureCount"] == 0
    # the picture candidate geometry is still present and untouched
    assert dto["targetBbox"] and dto["targetKind"] == "picture"


def test_candidates_of_the_same_item_share_its_region_and_never_another_items(sample_tree, tmp_path):
    """§20/§56-L: switching between two pictures of Mục #1 keeps Mục #1's region — never Mục #2's."""
    from tests.test_prompt024r_learning_workspace import _service_with_candidates

    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    regions = _regions(report, selection)
    item1, item2 = regions

    a = _candidate_from_region(deck, report, item1, index=0)
    b = _candidate_from_region(deck, report, item1, index=1)
    other = _candidate_from_region(deck, report, item2, index=0)
    service, _ = _service_with_candidates(sample_tree, tmp_path, [a, b, other])
    d_a, d_b, d_other = service.learning_state()["images"]

    assert d_a["evidenceRegionBbox"] == d_b["evidenceRegionBbox"] == {
        "x": item1.bbox[0], "y": item1.bbox[1], "width": item1.width, "height": item1.height}
    assert d_a["evidenceRegionItemId"] == d_b["evidenceRegionItemId"] == item1.improvement_item_id
    assert d_other["evidenceRegionItemId"] == item2.improvement_item_id
    assert d_other["evidenceRegionBbox"] != d_a["evidenceRegionBbox"]
    # the picture candidate overlay still differs per candidate (two different pictures)
    assert d_a["targetBbox"] != d_b["targetBbox"]


# ================================================================== §40 overlay math for BOTH geometries
@pytest.mark.parametrize("container,zoom", [((1600, 900), 1.0), ((820, 700), 1.0),
                                            ((820, 700), 2.0), ((400, 900), 1.0),
                                            ((1280, 480), 1.5)])
def test_both_overlays_stay_aligned_under_resize_letterbox_zoom_and_dpi(container, zoom):
    """§19/§40: candidate and evidence overlays use ONE slide plane, so both stay registered.

    Simulates fit-to-view, a resized panel, letterboxing, zoom and Windows DPI scaling (100/125/150%)
    by scaling the container the way a DPI-scaled layout would.
    """
    slide_w, slide_h = 12192000, 6858000
    pic_box = (7_000_000, 2_500_000, 1_200_000, 800_000)
    region_box = (6_500_000, 2_200_000, 3_500_000, 1_900_000)

    for dpi in (1.0, 1.25, 1.5):
        cw, ch = container[0] * dpi, container[1] * dpi
        fit = fit_slide_box(cw, ch, slide_w, slide_h, zoom=zoom)
        pic = map_fraction_box_to_px(slide_fraction_box(pic_box, slide_w, slide_h),
                                     fit["width"], fit["height"])
        region = map_fraction_box_to_px(slide_fraction_box(region_box, slide_w, slide_h),
                                        fit["width"], fit["height"])
        # both map through the identical display box, so their relative geometry is invariant
        for box, mapped in ((pic_box, pic), (region_box, region)):
            fx, fy, fw, fh = slide_fraction_box(box, slide_w, slide_h)
            assert mapped["x"] / fit["width"] == pytest.approx(fx)
            assert mapped["y"] / fit["height"] == pytest.approx(fy)
            assert mapped["width"] / fit["width"] == pytest.approx(fw)
            assert mapped["height"] / fit["height"] == pytest.approx(fh)
        # the picture stays inside the region at every size / zoom / DPI
        assert region["x"] <= pic["x"] and region["y"] <= pic["y"]
        assert region["x"] + region["width"] >= pic["x"] + pic["width"]
        assert region["y"] + region["height"] >= pic["y"] + pic["height"]
        # letterboxing is handled by the fit box, never by shifting the overlay
        assert fit["offset_x"] >= 0 and fit["offset_y"] >= 0


def test_overlay_percentages_are_relative_to_the_slide_image_not_the_panel():
    """§19: an overlay must not be positioned against the outer panel when the slide is letterboxed."""
    slide_w, slide_h = 12192000, 6858000
    fit = fit_slide_box(900, 900, slide_w, slide_h)          # tall container → vertical letterbox
    assert fit["offset_y"] > 0 and fit["width"] == pytest.approx(900)
    region = map_fraction_box_to_px(slide_fraction_box((0, 0, slide_w, slide_h), slide_w, slide_h),
                                    fit["width"], fit["height"])
    assert region["height"] == pytest.approx(fit["height"])
    assert region["height"] < 900, "overlay used the panel height instead of the slide image height"


# ================================================================== §49/§50/§51 version consistency
def test_backend_version_is_the_authoritative_1_3_4_build_017():
    assert app_pkg.__version__ == "1.3.4"
    assert app_pkg.BUILD_NUMBER == 17
    assert app_pkg.BUILD_ID == "017" and app_pkg.BUILD_LABEL == "Build 017"
    assert app_pkg.APP_TITLE == "Report Extractor v1.3.4"
    assert app_pkg.VERSION_LABEL == "1.3.4 — Build 017"
    assert app_pkg.VERSION_LINE == "version=1.3.4 build=017"


def test_frontend_and_backend_version_metadata_agree():
    """§51: the regression test that FAILS if the two sides drift (e.g. backend 1.3.4 / frontend 1.3.2)."""
    frontend = Path(__file__).resolve().parent.parent / "frontend"
    package = json.loads((frontend / "package.json").read_text(encoding="utf-8"))
    lock = json.loads((frontend / "package-lock.json").read_text(encoding="utf-8"))

    assert package["version"] == app_pkg.__version__, \
        f"frontend/package.json {package['version']} != backend {app_pkg.__version__}"
    assert lock["version"] == app_pkg.__version__
    assert lock["packages"][""]["version"] == app_pkg.__version__

    # the visible HTML title and the mock asset carry version AND build for the tester (§50)
    index_html = (frontend / "index.html").read_text(encoding="utf-8")
    assert f"Report Extractor — {app_pkg.__version__} — {app_pkg.BUILD_LABEL}" in index_html

    # unrelated dependency versions must not be swept up by a version bump (§49)
    assert lock["packages"]["node_modules/fast-fifo"]["version"] == "1.3.2"

    # the single canonical backend source: the frontend must not hardcode its own copy
    header = (frontend / "src" / "components" / "AppHeader.tsx").read_text(encoding="utf-8")
    assert app_pkg.__version__ not in header and "Build 017" not in header
    assert "appInfo.version" in header and "appInfo.build" in header


def test_bridge_reports_one_authoritative_version_to_the_ui(sample_tree, tmp_path):
    from tests.test_application_service import _make_service

    service, _ = _make_service(sample_tree, tmp_path)
    payload = service.app_version()
    assert payload["version"] == app_pkg.__version__ == "1.3.4"
    assert payload["build"] == app_pkg.BUILD_ID == "017"
    assert payload["buildNumber"] == app_pkg.BUILD_NUMBER == 17
    assert payload["title"] == app_pkg.APP_TITLE == "Report Extractor v1.3.4"
    config = service.current_config()["app"]
    assert (config["version"], config["build"]) == (app_pkg.__version__, app_pkg.BUILD_ID)


def test_version_regression_test_would_catch_a_drift(tmp_path):
    """Prove the §51 guard is real: a frontend left behind at 1.3.2/015 makes the comparison fail."""
    frontend = tmp_path / "frontend"
    frontend.mkdir()
    (frontend / "package.json").write_text(json.dumps({"version": "1.3.2"}), encoding="utf-8")
    stale = json.loads((frontend / "package.json").read_text(encoding="utf-8"))
    assert stale["version"] != app_pkg.__version__
    with pytest.raises(AssertionError):
        assert stale["version"] == app_pkg.__version__, "frontend/backend version drift"


def test_stable_backup_is_untouched():
    """§63: the v1.0.4 / Build 004 stable backup keeps its own historical version."""
    init = (Path(__file__).resolve().parent.parent / "backup"
            / "ReportExtractor_v1.0.4_Build004_STABLE" / "app" / "__init__.py").read_text(encoding="utf-8")
    assert '__version__ = "1.0.4"' in init and "BUILD_NUMBER = 4" in init
    assert "1.3.4" not in init and "BUILD_NUMBER = 17" not in init


# ================================================================== §18 preview never mutates evidence
def test_learning_preview_does_not_modify_evidence_or_region_geometry(tmp_path):
    """§18: dim/blur is UI only — building a preview leaves the production region bytes/geometry intact."""
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    before = [r.bbox for r in _regions(report, selection)]

    service = _preview_service()
    preview = service._slide_preview_for(_candidate(deck))
    assert preview["src"] and preview["regions"]

    grouped, problems = _export(report, selection, tmp_path / "crops", renderer=_builtin_renderer())
    assert problems == [] and len(grouped.groups) == 2
    after = [r.bbox for r in _regions(report, selection)]
    assert before == after
    # the exported crop still matches the region geometry exactly
    region = _regions(report, selection)[0]
    with Image.open(grouped.groups[0][0]) as crop:
        renderer = _builtin_renderer()
        out = renderer.render(report, [3], tmp_path / "r2", purpose="after_evidence")
        with Image.open(out[3]) as full:
            box = crop_box_px(region.bbox, report.slide_width, report.slide_height,
                              full.width, full.height)
        assert crop.size == (box[2] - box[0], box[3] - box[1])


def test_preview_regions_match_the_production_regions_exactly(tmp_path):
    """§15: the DTO region source IS the production builder — same bboxes, same item scoping."""
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    _, _, selection = _selection(report)
    production = _regions(report, selection)

    service = _preview_service()
    preview = service._slide_preview_for(_candidate(deck))
    exposed = {(r["itemId"], r["afterBlockIndex"]): r["bbox"] for r in preview["regions"]}
    assert len(exposed) == len(production) == 2
    for region in production:
        bbox = exposed[(region.improvement_item_id, region.after_block_index)]
        assert bbox == {"x": region.bbox[0], "y": region.bbox[1],
                        "width": region.width, "height": region.height}
        assert region.item_index in (0, 1)


def test_slide_dimensions_and_rendered_dimensions_are_both_reported(tmp_path, caplog):
    """§42: Windows acceptance can read report/slide/rendered/backend/item facts from safe logs."""
    deck = _deck(tmp_path / "in" / NAME, [_two_item_slide()])
    report = _report(deck)
    with caplog.at_level(logging.INFO):
        _, _, selection = _selection(report)
        _export(report, selection, tmp_path / "crops", renderer=_builtin_renderer())
    text = "\n".join(r.getMessage() for r in caplog.records)
    assert "SLIDE_RENDER purpose=after_evidence backend=builtin" in text
    assert f"slide_emu={report.slide_width}x{report.slide_height}" in text
    assert "REGION_CROP" in text and "REGION_BBOX" in text
    assert "ITEM_SEGMENTATION" in text
    assert str(Path(deck).resolve()) not in text.replace(deck.name, "")
