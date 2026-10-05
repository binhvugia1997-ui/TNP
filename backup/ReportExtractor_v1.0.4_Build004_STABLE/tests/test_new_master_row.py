"""Rule #19: a valid Management Number absent from the master gets ONE automatically created report row."""
import json
import shutil
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image

import app.batch_processor as bp
from app.batch_processor import BatchOptions, BatchProcessor
from app.config import AppConfig
from app.excel_writer import ExcelWriter
from app.gui_controller import GuiController
from app.prescan import (ACTION_OUTSIDE_PERIOD, ACTION_PROCESS, ACTION_PROCESS_NEW_ROW, ACTION_SOURCE_DUPLICATE,
                         CACHE_FILE_NAME, FastScanCache, MasterLookup, month_period, prescan)

SHEET = "Kiểm chứng"
SEP = month_period(2026, 9)
RUN_TODAY = date(2026, 9, 20)


def _fake(folder: Path, name: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    p = folder / name
    p.write_bytes(b"not a pptx")
    return p


def _run(files, template, out, **kw):
    events = []
    opts = BatchOptions(files=[Path(f) for f in files], template=template, output_file=out, use_ollama=False,
                        **{"today": RUN_TODAY, "period": SEP, **kw})
    proc = BatchProcessor(opts, on_file=lambda i, s, d: events.append((i, s, d)))
    return proc.run(), events, proc


def _rich_previous_row(template: Path, mgmt="260901001-VOC"):
    """Row 4 = a fully populated previous report with distinctive formatting, images and WEEK values;
    rows 4 merge G:H? no – keep real-template style: single-row merge in the note area (R4:S4)."""
    wb = load_workbook(template)
    ws = wb[SHEET]
    thick = Side(style="medium", color="FF0000")
    ws.cell(row=4, column=2, value=mgmt)
    ws.cell(row=4, column=3, value="Teawon")
    ws.cell(row=4, column=4, value=date(2026, 9, 1))
    ws.cell(row=4, column=5, value="A999")
    ws.cell(row=4, column=6, value="Rear")
    ws.cell(row=4, column=7, value="lỗi cũ")
    ws.cell(row=4, column=9, value="nguyên nhân cũ")
    ws.cell(row=4, column=10, value="đối sách cũ")
    for k in range(12, 20):
        ws.cell(row=4, column=k, value="OK-cũ")
    for c in range(1, 20):
        cell = ws.cell(row=4, column=c)
        cell.border = Border(left=thick, right=thick, top=thick, bottom=thick)
        cell.fill = PatternFill("solid", fgColor="FFF2CC")
        cell.font = Font(name="Arial", size=9, italic=True)
        cell.alignment = Alignment(vertical="top", wrap_text=True, horizontal="left")
        cell.number_format = "@"
    ws.cell(row=4, column=4).number_format = "dd/mm/yyyy"
    ws.row_dimensions[4].height = 120
    pic = template.with_name("pic.png")
    Image.new("RGB", (30, 30), "blue").save(pic)
    for col_letter in ("H", "K"):
        img = XLImage(str(pic))
        img.anchor = f"{col_letter}4"
        ws.add_image(img)
    wb.save(template)


# ---------------------------------------------------------------- 1-6: creation + processing path
def test_new_row_created_key_written_and_processed(sample_tree, template, tmp_path, monkeypatch):
    _rich_previous_row(template)
    calls = {"parse": 0}
    real = bp.parse_pptx
    monkeypatch.setattr(bp, "parse_pptx", lambda p: calls.__setitem__("parse", calls["parse"] + 1) or real(p))
    out = tmp_path / "Output" / "Kiem_chung_09_2026.xlsx"
    s, ev, proc = _run([sample_tree["files"][0]], template, out)
    assert proc.prescan_result.items[0].action == ACTION_PROCESS_NEW_ROW and s.candidates == 1      # 3
    assert calls["parse"] == 1 and s.completed == 1 and s.new_rows == 1                               # 6
    fr = proc.results[0]
    assert fr.new_row and fr.excel_row == 5
    ws = load_workbook(out)[SHEET]
    assert ws.cell(row=5, column=2).value == "260918080-VOC"                                          # 2
    assert ws.cell(row=5, column=4).value.strftime("%d/%m/%Y") == "18/09/2026"                        # 4 (YYMMDD)
    assert ws.cell(row=5, column=5).value == "A185" and ws.cell(row=5, column=9).value               # 5
    assert ws.cell(row=5, column=10).value and ws.cell(row=5, column=1).value == 2
    assert ws.cell(row=6, column=2).value is None                                                     # 1: exactly one
    assert "Dòng Excel         : Tạo mới  – Đã tạo dòng mới: 5  – Management Number mới: " in bp.format_file_diagnostics(fr)


# ---------------------------------------------------------------- 7-10: when NOT to create
def test_no_row_for_invalid_key_invalid_date_outside_period_or_ignored_duplicate(template, tmp_path):
    nokey = _fake(tmp_path / "in", "bao_cao_khong_ma.pptx")
    baddate = _fake(tmp_path / "in", "(CTMS)_11107_261345001_ Đối sách.pptx")          # month 13 -> parser rejects
    octo = _fake(tmp_path / "in", "261001002-VOC_october.pptx")
    d1 = _fake(tmp_path / "in" / "A", "260915003-VOC_a.pptx")
    d2 = _fake(tmp_path / "in" / "B", "260915003-VOC_b.pptx")
    import os
    os.utime(d1, (1_700_000_000, 1_700_000_000))
    os.utime(d2, (1_700_000_900, 1_700_000_900))
    out = tmp_path / "Output" / "k.xlsx"
    s, ev, proc = _run([nokey, baddate, octo, d1, d2], template, out)
    acts = [it.action for it in proc.prescan_result.items]
    assert acts[0] == "INVALID_MANAGEMENT_NUMBER" and acts[1] == "INVALID_MANAGEMENT_NUMBER"          # 7, 8
    assert acts[2] == ACTION_OUTSIDE_PERIOD and acts[3] == ACTION_SOURCE_DUPLICATE                   # 9, 10
    assert acts[4] == ACTION_PROCESS_NEW_ROW and s.new_rows == 1                                      # the survivor only
    ws = load_workbook(out)[SHEET]
    keys = [ws.cell(row=r, column=2).value for r in range(4, 8)]
    assert keys == ["260915003-VOC", None, None, None]


# ---------------------------------------------------------------- 11-15: existing rows unchanged
def test_existing_partial_complete_and_duplicate_rows_never_get_new_rows(sample_tree, template, tmp_path):
    wb = load_workbook(template)
    ws = wb[SHEET]
    ws.cell(row=4, column=2, value="260918080-VOC")                               # partial (key only)
    ws.cell(row=5, column=2, value="260918081-VOC")                               # duplicate pair
    ws.cell(row=6, column=2, value="260918081-VOC")
    wb.save(template)
    out = tmp_path / "Output" / "k.xlsx"
    s, ev, proc = _run([sample_tree["files"][0], sample_tree["files"][1]], template, out)
    assert [it.action for it in proc.prescan_result.items] == [ACTION_PROCESS, ACTION_PROCESS]
    assert s.new_rows == 0 and all(not r.new_row for r in proc.results)                            # 11, 12, 14
    ws2 = load_workbook(out)[SHEET]
    assert ws2.cell(row=4, column=5).value == "A185" and ws2.cell(row=5, column=5).value == "A185"
    assert ws2.cell(row=6, column=5).value is None and ws2.cell(row=7, column=2).value is None
    assert ws2.cell(row=6, column=2).fill.fgColor.rgb.endswith("FFC7CE")                           # 15
    # complete row -> skipped, still no new row                                                      # 13
    s2, ev2, proc2 = _run([sample_tree["files"][0]], template, out)
    assert s2.skipped == 1 and s2.new_rows == 0 and load_workbook(out)[SHEET].cell(row=7, column=2).value is None


# ---------------------------------------------------------------- 16-25: formatting cloned, data not
def test_new_row_clones_format_but_never_business_data_images_or_weeks(template, tmp_path):
    _rich_previous_row(template)
    wb = load_workbook(template)
    ws = wb[SHEET]
    ws.merge_cells(start_row=4, start_column=18, end_row=4, end_column=19)         # single-row merge R4:S4
    wb.save(template)
    out = tmp_path / "Output" / "k.xlsx"
    w = ExcelWriter(template, out)
    row = w.create_row("260923130")
    assert row == 5
    assert w.create_row("260923130") == 5                                           # idempotent – no 2nd row
    w.save()
    ws = load_workbook(out)[SHEET]
    src, dst = ws.cell(row=4, column=9), ws.cell(row=5, column=9)
    assert dst.border.left.style == "medium" and dst.border.left.color.rgb == src.border.left.color.rgb   # 17
    assert dst.fill.fgColor.rgb == src.fill.fgColor.rgb and dst.font.name == "Arial" and dst.font.italic
    assert dst.alignment.wrap_text and dst.alignment.vertical == "top"                                 # 18
    assert ws.cell(row=5, column=4).number_format == "dd/mm/yyyy"
    assert ws.row_dimensions[5].height == 120                                                          # 16
    merges = [str(r) for r in ws.merged_cells.ranges]
    assert "R5:S5" in merges and "R4:S4" in merges                                                     # 19
    cells = {}
    for r in ws.merged_cells.ranges:
        for c in r.cells:
            assert c not in cells, "overlapping merge"                                                 # 20
            cells[c] = True
    assert ws.cell(row=5, column=2).value == "260923130" and ws.cell(row=5, column=1).value == 2
    for col in (3, 4, 5, 6, 7, 9, 10):
        assert ws.cell(row=5, column=col).value is None                                                # 21
    for k in range(12, 18):
        assert ws.cell(row=5, column=k).value is None and ws.cell(row=5, column=k).border.left.style == "medium"   # 24, 25
    assert ws.cell(row=5, column=18).value is None and ws.cell(row=5, column=19).value is None
    rows_with_images = sorted(i.anchor._from.row + 1 for i in ws._images)
    assert rows_with_images == [4, 4]                                                                  # 22, 23
    assert ws.column_dimensions["J"].width == 70                                                       # widths intact


def test_blank_preformatted_form_row_is_reused_not_skipped(template, tmp_path):
    # the sample template has a pre-formatted blank row 4 -> the first new key lands there (no gap rows)
    w = ExcelWriter(template, tmp_path / "k.xlsx")
    assert w.create_row("260923130") == 4 and w.create_row("260923131") == 5
    w.save()
    ws = load_workbook(tmp_path / "k.xlsx")[SHEET]
    assert ws.cell(row=5, column=9).border.left.style == "thin"                   # format copied from row 4


# ---------------------------------------------------------------- 26-28: failure after creation
def test_failure_after_creation_leaves_one_retryable_row(sample_tree, template, tmp_path, monkeypatch):
    out = tmp_path / "Output" / "k.xlsx"
    broken = tmp_path / "in" / "260918080-VOC_report.pptx"
    broken.parent.mkdir()
    broken.write_bytes(b"corrupt")
    s, ev, proc = _run([broken], template, out)
    assert s.failed == 1 and s.new_rows == 1 and proc.results[0].status == "error"
    ws = load_workbook(out)[SHEET]
    assert ws.cell(row=4, column=2).value == "260918080-VOC" and ws.cell(row=5, column=2).value is None   # 26
    cache_file = out.parent / "logs" / CACHE_FILE_NAME
    assert not cache_file.exists() or "260918080-VOC" not in json.loads(cache_file.read_text("utf-8"))["entries"]
    # retry with the fixed report: same row reused, no second row, incremental fill
    shutil.copy(sample_tree["files"][0], broken)
    s2, ev2, proc2 = _run([broken], template, out)
    assert s2.completed == 1 and s2.new_rows == 0 and not proc2.results[0].new_row
    assert proc2.prescan_result.items[0].action == ACTION_PROCESS                  # 27
    ws2 = load_workbook(out)[SHEET]
    assert ws2.cell(row=4, column=5).value == "A185" and ws2.cell(row=5, column=2).value is None   # 28
    cache2 = json.loads((out.parent / "logs" / CACHE_FILE_NAME).read_text(encoding="utf-8"))
    assert cache2["entries"]["260918080-VOC"]["status"] == "completed"              # 33


# ---------------------------------------------------------------- 29-30: cache vs. missing row
def test_cache_hit_with_missing_row_is_new_row_not_fast_skip(tmp_path):
    f = _fake(tmp_path / "in", "260915001-VOC_r.pptx")
    cache = FastScanCache(tmp_path / CACHE_FILE_NAME)
    cache.record("260915001-VOC", f, "completed")                                   # success in LAST month's workbook
    res = prescan([f], SEP, cache, MasterLookup(lambda m: [], lambda r: []), today=RUN_TODAY)
    it = res.items[0]
    assert it.action == ACTION_PROCESS_NEW_ROW and it.cache_decision == "miss:master_row_missing"   # 29
    assert len(res.candidates) == 1 and res.counts()["fast_skipped"] == 0                              # 30


def test_new_workbook_not_suppressed_by_cache_from_previous_workbook(sample_tree, template, tmp_path):
    out1 = tmp_path / "Output" / "Kiem_chung_08_2026.xlsx"
    s1, _, _ = _run([sample_tree["files"][0]], template, out1, period=month_period(2026, 9))
    assert s1.completed == 1
    # new monthly workbook in the same Output folder (shared cache file) – the key is absent there
    out2 = tmp_path / "Output" / "Kiem_chung_09_2026.xlsx"
    s2, ev2, proc2 = _run([sample_tree["files"][0]], template, out2)
    assert s2.fast_skipped == 0 and s2.new_rows == 1 and s2.completed == 1
    assert load_workbook(out2)[SHEET].cell(row=4, column=2).value == "260918080-VOC"


# ---------------------------------------------------------------- 31-32: GUI
def test_gui_prescan_shows_new_rows_and_counts_them_in_denominator(sample_tree, template, tmp_path, monkeypatch):
    class _D(date):
        @classmethod
        def today(cls):
            return RUN_TODAY
    monkeypatch.setattr(bp, "date", _D)
    reports = tmp_path / "Bao_cao" / "2026-09"
    reports.mkdir(parents=True)
    shutil.copy(sample_tree["files"][0], reports / "260918080-VOC_a.pptx")
    shutil.copy(sample_tree["files"][1], reports / "260918081-VOC_b.pptx")
    wb = load_workbook(template)
    wb[SHEET].cell(row=4, column=2, value="260918080-VOC")
    wb.save(template)
    ctl = GuiController(AppConfig(), config_path=tmp_path / "c.json")
    ctl.set_report_folder(str(reports))
    ctl.set_template(str(template))
    ctl.set_output(str(tmp_path / "Output" / "Kiem_chung_09_2026.xlsx"))
    assert ctl.start(use_ollama=False)
    ctl.processor._thread.join(120)
    ctl.pump()
    lines = ctl.summary_lines()
    assert "Management Number mới: 1" in lines and "Cần bổ sung dữ liệu: 1" in lines and "Cần xử lý thực tế: 2" in lines
    assert ctl.progress.total == 2 and ctl.progress.text == "Đã xử lý: 2 / 2 — 100%"
    new = next(ctl.result_for(i) for i in range(2) if ctl.result_for(i).management_number == "260918081-VOC")
    assert new.new_row and new.status in ("completed", "needs_review")
    assert ctl.diagnostics_for(ctl.files.index(Path(new.source_file)))["Dòng sử dụng"] == str(new.excel_row)
