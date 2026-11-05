"""Last-slide effectiveness exclusion (Bug C): a slide titled "Hiệu quả cải tiến"
contributes ZERO text and ZERO pictures, in every processing path.

Regression fixtures are simulated SlideData/Block objects shaped like real countermeasure
decks (title placeholder + body restating the countermeasure).  The exclusion vocabulary
already existed; the segmentation layer lost the title's verify kind to the body's first
improvement-like heading (see .specify/bugs/last-slide-effectiveness-leak/assessment.md).
"""
from pathlib import Path

from app.classifier import Classification, first_heading_kind, heuristic_classify
from app.extractor import collect_sections, excluded_bands, extract_record, join_sections
from app.pptx_parser import Block, ReportData, SlideData

H = 5143500


def _slide(n, blocks):
    return SlideData(number=n, blocks=blocks, width=9144000, height=H)


def _title(text, order=0):
    return Block(kind="title", text=text, top=100000, height=500000, shape_id=1, order=order,
                 bold=True, size_pt=28)


def _para(text, order=1, top=800000):
    return Block(kind="paragraph", text=text, top=top, height=2000000, shape_id=2, order=order,
                 bold=False, size_pt=18)


def _lead_slides():
    return [
        _slide(1, [_title("Quality Problem Notice"), _para("Defect: scratch on rear cover")]),
        _slide(2, [_title("Nguyên nhân"), _para("Nguyên nhân: chân đỡ màn hình dài chạm vào sản phẩm.")]),
        _slide(3, [_title("4. Cải tiến trong sản xuất"),
                   _para("Trước: chân đỡ màn hình dài chạm vào sản phẩm.\nSau: cắt ngắn chân đỡ 3cm.")]),
    ]


def _effectiveness_slide():
    """Title placeholder + body whose first line is an improvement-like heading."""
    return _slide(4, [_title("5. Hiệu quả cải tiến"),
                      _para("Đối sách cải tiến:\nCắt ngắn chân đỡ màn hình 3cm.\n"
                            "Sau cải tiến: không phát sinh lỗi.")])



def _report(slides):
    return ReportData(path=Path("fake.pptx"), slides=slides, slide_width=9144000, slide_height=H)


def _improvement_text(report, cls):
    secs = collect_sections(report, cls)
    imp = [s for s in secs if s.kind in ("improvement", "standard")]
    return join_sections(imp), secs


def test_effectiveness_slide_excluded_end_to_end():
    report = _report(_lead_slides() + [_effectiveness_slide()])
    cls = heuristic_classify(report)
    # The slide is listed in BOTH lists (its body restates the countermeasure) — the
    # section layer, not the slide lists, is what must exclude it.
    assert cls.improvement_slides == [3, 4]
    assert 4 in cls.verify_slides
    text, secs = _improvement_text(report, cls)
    assert "không phát sinh lỗi" not in text
    assert "Cắt ngắn chân đỡ màn hình 3cm" not in text
    assert "cắt ngắn chân đỡ 3cm" in text.lower()
    assert all(s.kind in ("verify", "followup") for s in secs if s.slide == 4)
    assert any(e.startswith("S4 verify") for e in
               [f"S{s.slide} {s.kind}: {(s.heading or s.lines[0]).strip()[:70]}" for s in secs
                if s.kind in ("inspection", "followup", "verify", "temporary") and s.lines])


def test_effectiveness_slide_excluded_when_classifier_lists_it_as_improvement():
    """The LLM may only point at slides; segmentation still drops the effectiveness slide,
    even when the classification lists it in BOTH improvement_slides and verify_slides."""
    report = _report(_lead_slides() + [_effectiveness_slide()])
    cls = Classification(management_number="260918080-VOC", qpn_slide=1, cause_slides=[2],
                         improvement_slides=[3, 4], verify_slides=[4], source="qwen", confidence=0.9)
    rec = extract_record(report, cls, {"rear": "Rear"}, ["A185"])
    imp = rec.improvement
    assert "không phát sinh lỗi" not in imp
    assert "Cắt ngắn chân đỡ màn hình 3cm" not in imp
    assert "cắt ngắn chân đỡ 3cm" in imp.lower()
    assert any(e.startswith("S4 verify") for e in rec.excluded_sections)


def test_effectiveness_body_before_title_shape_still_excluded():
    """Body shapes ordered ahead of the title shape: pre-heading lines take the
    verify-led default instead of leaking as improvement text."""
    slide4 = _slide(4, [_para("Đối sách cải tiến:\nCắt ngắn chân đỡ màn hình 3cm.\n"
                              "Sau cải tiến: không phát sinh lỗi.", order=0),
                        _title("5. Hiệu quả cải tiến", order=1)])
    report = _report(_lead_slides() + [slide4])
    cls = heuristic_classify(report)
    assert 4 in cls.verify_slides
    text, _ = _improvement_text(report, cls)
    assert "không phát sinh lỗi" not in text
    assert "Cắt ngắn chân đỡ màn hình 3cm" not in text


def test_trailing_effectiveness_block_on_improvement_slide_still_splits():
    """Mid-slide effectiveness heading still terminates the improvement section (PROMPT-001)."""
    report = _report([
        _slide(1, [_title("Quality Problem Notice"), _para("Defect: scratch")]),
        _slide(2, [_title("4. Cải tiến trong sản xuất"),
                   _para("Trước: chân đỡ dài chạm sản phẩm.\nSau: cắt ngắn 3cm.\n"
                         "Hiệu quả cải tiến:\nKhông phát sinh lỗi trong 5 ngày.")]),
    ])
    cls = heuristic_classify(report)
    text, secs = _improvement_text(report, cls)
    assert "cắt ngắn 3cm" in text.lower()
    assert "Không phát sinh lỗi" not in text
    assert [s.kind for s in secs] == ["improvement", "verify"]


def test_effectiveness_slide_pictures_excluded_by_bands():
    """The geometric twin of the segmentation: the whole title-led effectiveness slide
    is one excluded band, so no picture on it can be selected as After evidence."""
    slide = _effectiveness_slide()
    bands = excluded_bands(slide, H)
    assert bands, "an excluded band must cover the effectiveness slide"
    top, bottom, kind = bands[0]
    assert kind == "verify"
    assert top <= 600000 and bottom >= H - 1, f"band must span the slide body: {(top, bottom, kind)}"
    # The production improvement slide keeps its pictures: no band at all.
    prod = _slide(3, [_title("4. Cải tiến trong sản xuất"),
                      _para("Trước: chân đỡ dài.\nSau: cắt ngắn 3cm.")])
    assert excluded_bands(prod, H) == []


def test_mid_slide_excluded_block_still_closes_at_next_heading():
    """A mid-slide (non-title) effectiveness block does NOT swallow the next section."""
    slide = _slide(2, [_title("4. Cải tiến trong sản xuất"),
                       _para("Trước: chân đỡ dài.\nSau: cắt ngắn 3cm.\n"
                             "Hiệu quả cải tiến:\nKhông phát sinh lỗi.\n"
                             "Cải tiến bổ sung:\nThay chân đỡ mới.")])
    bands = excluded_bands(slide, H)
    assert len(bands) == 1
    _, bottom, kind = bands[0]
    assert kind == "verify" and bottom < H, "the band must end where the next section starts"
    secs = [s for s in collect_sections(
        _report([_slide(1, [_title("Quality Problem Notice")]), slide]),
        Classification(qpn_slide=1, improvement_slides=[2], verify_slides=[2])) if s.slide == 2]
    assert [s.kind for s in secs] == ["improvement", "verify", "improvement"]


def test_first_heading_kind_of_effectiveness_slide_is_verify():
    assert first_heading_kind(_effectiveness_slide()) == "verify"


def _header_band(text, order, top):
    """A heading textbox in the slide header band (no title placeholder): the A1285 shape."""
    return Block(kind="paragraph", text=text, left=274320, top=top, width=7315200,
                 height=400000, shape_id=order + 1, order=order, bold=True, size_pt=24)


def test_header_band_effectiveness_heading_does_not_swallow_later_long_term_heading():
    """A1285 shape: slide 8 has NO title placeholder — a header-band "HIỆU QUẢ" heading,
    then a LATER "ĐỐI SÁCH LÂU DÀI" heading whose body is production text.  Only the
    effectiveness block is excluded; the long-term countermeasure text stays."""
    # Slide number 2 positionally (ReportData.slide is positional); shaped like A1285 slide 8.
    slide = _slide(2, [
        _header_band("7. HIỆU QUẢ SAU CẢI TIẾN", 0, 100000),
        Block(kind="paragraph", text="Theo dõi 4 tuần: chưa có dữ liệu", left=914400,
              top=914400, width=7315200, height=400000, shape_id=2, order=1, size_pt=14),
        _header_band("ĐỐI SÁCH LÂU DÀI", 2, 3657600),
        Block(kind="paragraph", text="- Cập nhật SOP Rev.03, triển khai ngang dòng A175",
              left=914400, top=4292600, width=7315200, height=400000, shape_id=4,
              order=3, size_pt=14),
    ])
    report = _report([_slide(1, [_title("Quality Problem Notice")]), slide])
    cls = Classification(qpn_slide=1, improvement_slides=[2], verify_slides=[2])
    secs = [s for s in collect_sections(report, cls) if s.slide == 2]
    assert [s.kind for s in secs] == ["verify", "improvement"]
    text, _ = _improvement_text(report, cls)
    assert "Cập nhật SOP Rev.03" in text
    assert "chưa có dữ liệu" not in text
