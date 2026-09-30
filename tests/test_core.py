"""Scanner, config, Ollama helpers, PPTX parsing, classification, extraction."""
import json

import pytest

from app.classifier import (build_prompt, classify, find_qpn_slide, heuristic_classify,
                            merge_llm_with_heuristic, section_kind_of_heading)
from app.config import AppConfig, normalize_ollama_url
from app.extractor import extract_record
from app.ollama_client import OllamaError, parse_json_response, parse_models, preferred_model
from app.pptx_parser import parse_pptx
from app.scanner import parse_dnd_paths, scan_folder, scan_inputs
from make_samples import IMPROVEMENT_TEXT, IMPROVEMENT_TEXT_2, LONG_TERM_TEXT, TEMP_TEXT, CAUSE_TEXT_1, CAUSE_TEXT_2


# ---------------------------------------------------------------- scanning
def test_recursive_scan_finds_all_reports(sample_tree):
    files = scan_folder(sample_tree["reports"])
    names = sorted(p.name for p in files)
    assert len(files) == 4
    assert any(n.endswith("report3.pptx") for n in names) and any("260918083-VOC" in n for n in names)
    assert any("September/A185" in str(p).replace("\\", "/") for p in files)


def test_scan_skips_temp_and_non_pptx(tmp_path):
    (tmp_path / "a.pptx").write_bytes(b"x")
    (tmp_path / "~$a.pptx").write_bytes(b"x")
    (tmp_path / "b.docx").write_bytes(b"x")
    (tmp_path / "sub").mkdir()
    (tmp_path / "sub" / "c.PPTX").write_bytes(b"x")
    assert sorted(p.name for p in scan_folder(tmp_path)) == ["a.pptx", "c.PPTX"]


def test_scan_inputs_mix_and_dedupe(sample_tree):
    f = sample_tree["files"][0]
    files = scan_inputs([sample_tree["reports"], f, str(f)])
    assert len(files) == 4


def test_parse_dnd_paths():
    assert parse_dnd_paths("{C:/a b/x.pptx} C:/y.pptx") == ["C:/a b/x.pptx", "C:/y.pptx"]
    assert parse_dnd_paths("/tmp/folder") == ["/tmp/folder"]


# ---------------------------------------------------------------- config / ollama
@pytest.mark.parametrize("raw,expected", [
    ("192.168.1.50:11434", "http://192.168.1.50:11434"),
    ("http://192.168.1.50:11434", "http://192.168.1.50:11434"),
    ("http://192.168.1.50:11434/", "http://192.168.1.50:11434"),
    ("  192.168.1.50 ", "http://192.168.1.50:11434"),
    ("https://ai-box:11434/api/tags", "https://ai-box:11434"),
    ("", "http://127.0.0.1:11434"),
])
def test_normalize_ollama_url(raw, expected):
    assert normalize_ollama_url(raw) == expected


def test_parse_models_and_preference():
    tags = {"models": [{"name": "llama3:8b"}, {"name": "qwen2.5:4b-instruct"}, {"name": "qwen3:8b"},
                       {"model": "nomic-embed-text"}]}
    models = parse_models(tags)
    assert models == ["llama3:8b", "qwen2.5:4b-instruct", "qwen3:8b", "nomic-embed-text"]
    assert preferred_model(models) == "qwen2.5:4b-instruct"
    assert preferred_model(["llama3:8b", "qwen3:8b"]) == "qwen3:8b"
    assert preferred_model([]) is None


def test_parse_json_response_tolerates_fences_and_prose():
    assert parse_json_response('```json\n{"a": 1}\n```')["a"] == 1
    assert parse_json_response('Sure: {"qpn_slide": 2, "x": [1,2]} done')["qpn_slide"] == 2
    with pytest.raises(OllamaError):
        parse_json_response("no json here")


def test_config_persistence(tmp_path):
    p = tmp_path / "config.json"
    cfg = AppConfig()
    cfg.ollama_server = "192.168.1.50:11434"
    cfg.model = "qwen3:4b"
    cfg.last_report_folder = r"D:\Reports"
    cfg.last_template = r"D:\Templates\Verification.xlsx"
    cfg.last_output_folder = r"D:\Output"
    cfg.save(p)
    data = json.loads(p.read_text(encoding="utf-8"))
    assert data["ollama_server"] == "http://192.168.1.50:11434"
    assert "password" not in json.dumps(data).lower()
    loaded = AppConfig.load(p)
    assert loaded.model == "qwen3:4b"
    assert loaded.last_report_folder == r"D:\Reports"
    assert loaded.last_template.endswith("Verification.xlsx")
    assert loaded.last_output_folder == r"D:\Output"


def test_config_load_missing_file_gives_defaults(tmp_path):
    cfg = AppConfig.load(tmp_path / "nope.json")
    assert cfg.model == "qwen3:4b" and cfg.ollama_server.startswith("http://")


# ---------------------------------------------------------------- pptx / classification
def test_parse_pptx_preserves_unicode_and_order(a185_report):
    r = parse_pptx(a185_report)
    assert len(r.slides) == 8
    assert "Quality Problem Notice" in r.slides[1].text
    assert "NGUYÊN NHÂN" in r.slides[2].text
    assert r.slides[1].text_blocks[0].text == "Quality Problem Notice"   # title first (top)
    assert len(r.slides[4].pictures) == 2
    assert "Đèn kiểm tra tại OQC không đủ sáng (800 lux)" in r.slides[2].text


def test_qpn_slide_detection(a185_report, report_factory):
    r = parse_pptx(a185_report)
    assert find_qpn_slide(r) == 2
    no_qpn = parse_pptx(report_factory("noqpn.pptx", with_qpn=False))
    assert find_qpn_slide(no_qpn) is None


def test_heading_kinds():
    assert section_kind_of_heading("2. NGUYÊN NHÂN") == "cause"
    assert section_kind_of_heading("XỬ LÝ TẠM THỜI") == "temporary"
    assert section_kind_of_heading("PHƯƠNG PHÁP XỬ LÝ TẠM THỜI") == "temporary"
    assert section_kind_of_heading("Đối sách tạm thời") == "temporary"
    assert section_kind_of_heading("3. CẢI TIẾN TRONG SẢN XUẤT") == "improvement"
    assert section_kind_of_heading("ĐỐI SÁCH LÂU DÀI") == "improvement"
    assert section_kind_of_heading("BÁO CÁO ĐỐI SÁCH LỖI XƯỚC") is None
    assert section_kind_of_heading("Quality Problem Notice") == "qpn"


def test_heuristic_classification(a185_report):
    c = heuristic_classify(parse_pptx(a185_report))
    assert c.qpn_slide == 2
    assert c.cause_slides == [3]
    assert c.temporary_slides == [4]
    assert c.improvement_slides == [5, 6, 7]
    assert c.improvement_image_slides == [5, 6]
    assert 1 not in c.improvement_slides


def test_llm_merge_guards_qpn_and_validates_indices(a185_report):
    r = parse_pptx(a185_report)
    heur = heuristic_classify(r)
    llm = {"management_number": "260918080-VOC", "model": "SM-A185", "item": "Rear",
           "qpn_slide": 4, "cause_slides": ["3"], "improvement_slides": [5, 6, 99],
           "improvement_image_slides": [5], "temporary_slides": []}
    c = merge_llm_with_heuristic(llm, heur, r)
    assert c.qpn_slide == 2                       # exact text wins over LLM
    assert 99 not in c.improvement_slides
    assert 7 in c.improvement_slides              # heading-detected slide appended
    assert 4 in c.temporary_slides
    assert c.improvement_image_slides == [5, 6]           # LLM can never narrow the structural set


class FakeOllama:
    def __init__(self, answer=None, fail=False):
        self.answer = answer or {}
        self.fail = fail
        self.prompts = []

    def generate_json(self, model, prompt, system="", **kw):
        self.prompts.append(prompt)
        if self.fail:
            raise OllamaError("boom")
        return self.answer


def test_classify_uses_only_slide_text_and_falls_back(a185_report):
    r = parse_pptx(a185_report)
    fake = FakeOllama({"qpn_slide": 2, "cause_slides": [3], "improvement_slides": [5, 6, 7],
                       "improvement_image_slides": [5, 6], "temporary_slides": [4], "model": "SM-A185"})
    c = classify(r, fake, "qwen3:4b")
    assert c.source.startswith("qwen")
    assert "\nS2 | Quality Problem Notice" in fake.prompts[0] and "Quality Problem Notice" in fake.prompts[0]
    assert "PK\x03" not in fake.prompts[0]        # no binary sent
    c2 = classify(r, FakeOllama(fail=True), "qwen3:4b")
    assert c2.source == "heuristic" and c2.qpn_slide == 2


# ---------------------------------------------------------------- extraction (data integrity)
def test_extraction_preserves_original_text_and_excludes_temporary(a185_report):
    r = parse_pptx(a185_report)
    rec = extract_record(r, heuristic_classify(r), {"rear": "Rear", "Rear": "Rear"}, ["A185", "A175"])
    assert rec.management_number == "260918080-VOC"
    assert rec.model == "A185"
    assert rec.item == "Rear"
    assert rec.defect_content == "Xước: 15ea\nLệch ANT: 5ea\nMẻ: 13ea"
    # full original improvement text, in source order
    for line in IMPROVEMENT_TEXT.split("\n") + IMPROVEMENT_TEXT_2.split("\n") + LONG_TERM_TEXT.split("\n"):
        if line.strip():
            assert line in rec.improvement, line
    assert rec.improvement.index("3. CẢI TIẾN TRONG SẢN XUẤT") < rec.improvement.index("4. CẢI TIẾN TRONG KIỂM TRA") \
        < rec.improvement.index("5. ĐỐI SÁCH LÂU DÀI")
    # temporary handling excluded from improvement
    for line in TEMP_TEXT.split("\n"):
        assert line not in rec.improvement
    assert "Sorting 100%" not in rec.improvement
    assert "XỬ LÝ TẠM THỜI" not in rec.improvement
    assert "Sorting 100%" in rec.temporary_excluded
    # root cause keeps both sub-sections
    for line in (CAUSE_TEXT_1 + "\n" + CAUSE_TEXT_2).split("\n"):
        assert line in rec.root_cause
    assert len(rec.cause_sections) >= 2
    # vendor: no "công đoạn assy <Vendor>" phrase in this deck -> blank + review; date derived from mgmt no.
    assert rec.vendor == "Doaltech"                      # source 'Daoltech' -> canonical controlled-list spelling
    assert rec.occurrence_date_text == "18/09/2026"      # from 260918080-VOC, not from slide text
    assert all(v == "" for v in rec.weeks.values())
    assert rec.review_reasons == []


def test_temporary_phrase_not_removed_globally(report_factory):
    """'Sorting 100%' inside an improvement section must be kept (removal is by section)."""
    from pptx import Presentation
    from pptx.util import Inches
    p = report_factory("x.pptx")
    prs = Presentation(str(p))
    s = prs.slides[4]
    tb = s.shapes.add_textbox(Inches(0.5), Inches(6.6), Inches(7), Inches(0.5))
    tb.text_frame.text = "- Áp dụng Sorting 100% bằng máy AOI (thiết bị mới)"
    prs.save(str(p))
    r = parse_pptx(p)
    rec = extract_record(r, heuristic_classify(r))
    assert "Sorting 100% bằng máy AOI" in rec.improvement
    assert "hàng tồn kho" not in rec.improvement


def test_llm_values_only_used_when_verbatim_in_source(a185_report):
    r = parse_pptx(a185_report)
    c = heuristic_classify(r)
    c.model, c.item, c.management_number = "Z999", "Bracket", "000000000-XXX"
    rec = extract_record(r, c)
    assert rec.model == "A185" and rec.management_number == "260918080-VOC"
    assert rec.item == "Rear"


def test_missing_sections_flag_needs_review(report_factory):
    r = parse_pptx(report_factory("m.pptx", with_qpn=False, with_improvement=False))
    rec = extract_record(r, heuristic_classify(r))
    assert rec.improvement == ""
    assert any("QPN" in x for x in rec.review_reasons)
    assert any("đối sách cải tiến" in x for x in rec.review_reasons)


def test_prompt_is_reasonably_small(a185_report):
    prompt = build_prompt(parse_pptx(a185_report))
    assert len(prompt) < 20000
