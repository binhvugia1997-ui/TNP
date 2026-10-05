"""PROMPT-004D – automatic new production row for a valid Management Number absent from the master.

Spec §14 checklist (one test per bullet, numbered in the test names).  Everything runs with use_ollama=False:
no Qwen is ever needed to decide the destination row.
"""
import shutil
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

import app.batch_processor as bp
import app.excel_writer as xw
from app.batch_processor import BatchOptions, BatchProcessor, format_file_diagnostics
from app.excel_writer import ExcelWriter
from app.prescan import ACTION_PROCESS_NEW_ROW, MasterLookup, prescan
from app.extractor import management_number_from_filename

SHEET = "Kiểm chứng"
COL = {"stt": 1, "mgmt": 2, "vendor": 3, "date": 4, "model": 5, "item": 6, "defect": 7, "qpn": 8, "cause": 9,
       "improvement": 10, "images": 11}
WEEK_COLS = range(12, 20)


def _run(files, template, out, **kw):
    events = []
    opts = BatchOptions(files=[Path(f) for f in files], template=template, output_file=out, use_ollama=False, **kw)
    proc = BatchProcessor(opts, on_file=lambda i, s, d: events.append((i, s, d)))
    return proc.run(), events, proc


def _prefill(template: Path, rows, complete=True):
    """rows = list of Management Numbers written from row 4; complete rows get every managed field + WEEKs."""
    wb = load_workbook(template)
    ws = wb[SHEET]
    for i, mgmt in enumerate(rows):
        r = 4 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=mgmt)
        if complete:
            for key in ("vendor", "model", "item", "defect", "cause", "improvement"):
                ws.cell(row=r, column=COL[key], value=f"old-{key[0].upper()}-{mgmt}")   # no header-like words
            ws.cell(row=r, column=COL["date"], value=date(2026, 9, 1))
            ws.cell(row=r, column=COL["qpn"], value="x")
            ws.cell(row=r, column=COL["images"], value="x")
            for k in WEEK_COLS:
                ws.cell(row=r, column=k, value="OK")
    wb.save(template)


def _keys(ws):
    return [ws.cell(row=r, column=2).value for r in range(4, ws.max_row + 1) if ws.cell(row=r, column=2).value]


def _copy_as(src: Path, folder: Path, name: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    dst = folder / name
    shutil.copy(src, dst)
    return dst


# ---------------------------------------------------------------- 1. new numeric key -> exactly one new row
def test_01_new_numeric_key_creates_exactly_one_row(sample_tree, template, tmp_path):
    rep = _copy_as(sample_tree["files"][0], tmp_path / "in", "(CTMS)_20506_260925015_A185_Rear.pptx")
    _prefill(template, ["999999999-VOC"])
    out = tmp_path / "out" / "k.xlsx"
    summary, events, proc = _run([rep], template, out)
    fr = proc.results[0]
    assert fr.management_number == "260925015" and fr.new_row and fr.excel_row == 5
    assert fr.status == "completed" and summary.new_rows == 1 and summary.not_written == 0 and summary.failed == 0
    ws = load_workbook(out)[SHEET]
    assert _keys(ws) == ["999999999-VOC", "260925015"]
    assert ws.cell(row=5, column=COL["model"]).value == "A185"            # processed normally into the new row
    # final GUI stage + Vietnamese status text
    assert events[-1][1] == "completed_new"
    assert bp.STAGE_LABELS_VI["completed_new"] == "Hoàn thành — đã thêm Management Number mới"
    diag = format_file_diagnostics(fr)
    assert "Đã tạo dòng mới: 5" in diag and "Management Number mới: 260925015" in diag


# ---------------------------------------------------------------- 2. new -VOC key -> exactly one new row
def test_02_new_voc_key_creates_exactly_one_row(sample_tree, template, tmp_path):
    _prefill(template, ["999999999-VOC"])
    out = tmp_path / "out" / "k.xlsx"
    summary, _, proc = _run([sample_tree["files"][0]], template, out)      # 260918080-VOC
    assert proc.results[0].new_row and proc.results[0].excel_row == 5 and summary.new_rows == 1
    assert _keys(load_workbook(out)[SHEET]) == ["999999999-VOC", "260918080-VOC"]


# ---------------------------------------------------------------- 3. no valid key -> no row, no modification
def test_03_no_management_number_creates_no_row_and_leaves_excel_untouched(report_factory, template, tmp_path):
    nokey = report_factory("Bao cao doi sach A185 21.09.2026.pptx")
    assert not management_number_from_filename(nokey.name)
    _prefill(template, ["999999999-VOC"])
    out = tmp_path / "out" / "k.xlsx"
    summary, events, proc = _run([nokey], template, out)
    fr = proc.results[0]
    assert fr.status == "not_written" and fr.excel_row is None and not fr.new_row
    assert "Không xác định được Management Number từ tên file" in fr.error
    assert summary.not_written == 1 and summary.new_rows == 0 and not summary.backup_file   # nothing modified
    assert _keys(load_workbook(out)[SHEET]) == ["999999999-VOC"]
    assert bp.STAGE_LABELS_VI["not_written"] == "Cần kiểm tra — Không xác định được Management Number từ tên file"


# ---------------------------------------------------------------- 4. the exact filename key is written
@pytest.mark.parametrize("name,expected", [
    ("(CTMS)_20506_260925015_A185_Rear.pptx", "260925015"),
    ("(CTMS)_20509_260918080-VOC_A185 Rear 21.09.2026.pptx", "260918080-VOC"),
])
def test_04_exact_management_number_from_filename_is_written(sample_tree, template, tmp_path, name, expected):
    rep = _copy_as(sample_tree["files"][0], tmp_path / "in", name)
    out = tmp_path / "out" / "k.xlsx"
    _, _, proc = _run([rep], template, out)
    ws = load_workbook(out)[SHEET]
    assert ws.cell(row=proc.results[0].excel_row, column=2).value == expected


# ---------------------------------------------------------------- 5 + 6. date from the key, never from the filename
def test_05_06_occurrence_date_from_key_not_from_filename_date(sample_tree, template, tmp_path):
    rep = _copy_as(sample_tree["files"][0], tmp_path / "in", "(CTMS)_20506_260925015_A185_Rear 24.09.2026.pptx")
    out = tmp_path / "out" / "k.xlsx"
    _, _, proc = _run([rep], template, out)
    ws = load_workbook(out)[SHEET]
    d = ws.cell(row=proc.results[0].excel_row, column=COL["date"]).value
    assert isinstance(d, datetime) and d.date() == date(2026, 9, 25)       # 260925015 -> 25/09/2026, not 24.09.2026
    assert proc.results[0].occurrence_date == "25/09/2026"


# ---------------------------------------------------------------- 7 + 8 + 9 + 10. formatting copied, values / WEEK / manual blank
def test_07_10_format_cloned_but_no_business_values_weeks_or_manual_fields(sample_tree, template, tmp_path):
    _prefill(template, ["260901001-VOC"])
    wb = load_workbook(template)
    ws = wb[SHEET]
    thick = Side(style="medium", color="FF0000")
    for c in range(1, 20):
        cell = ws.cell(row=4, column=c)
        cell.border = Border(left=thick, right=thick, top=thick, bottom=thick)
        cell.fill = PatternFill("solid", fgColor="FFF2CC")
        cell.font = Font(name="Arial", size=9, italic=True)
        cell.alignment = Alignment(vertical="top", wrap_text=True, horizontal="left")
    ws.cell(row=4, column=COL["date"]).number_format = "dd/mm/yyyy"
    ws.row_dimensions[4].height = 120
    wb.save(template)
    out = tmp_path / "out" / "k.xlsx"
    _, _, proc = _run([sample_tree["files"][0]], template, out)
    ws = load_workbook(out)[SHEET]
    r = proc.results[0].excel_row
    assert r == 5
    for c in range(1, 20):
        cell = ws.cell(row=r, column=c)
        assert cell.border.left.style == "medium" and cell.fill.fgColor.rgb.endswith("FFF2CC")
        assert cell.font.italic and cell.font.name == "Arial" and cell.alignment.wrap_text
    assert ws.cell(row=r, column=COL["date"]).number_format.lower() == "dd/mm/yyyy"
    # 8. nothing from row 4 leaked into row 5 (only the report's own values are there)
    for key in ("vendor", "model", "item", "defect", "cause", "improvement"):
        assert ws.cell(row=r, column=COL[key]).value != f"old-{key[0].upper()}-260901001-VOC"
    assert ws.cell(row=r, column=2).value == "260918080-VOC"
    # 9 + 10. WEEK+1..+8 and manual fields stay blank; the previous row is untouched
    assert all(ws.cell(row=r, column=k).value is None for k in WEEK_COLS)
    assert all(ws.cell(row=4, column=k).value == "OK" for k in WEEK_COLS)
    assert ws.cell(row=4, column=COL["model"]).value == "old-M-260901001-VOC"


# ---------------------------------------------------------------- 11. backup BEFORE the row is created
def test_11_backup_exists_before_new_row_is_created(sample_tree, template, tmp_path, monkeypatch):
    out = tmp_path / "out" / "k.xlsx"
    seen = []
    orig_create = xw.ExcelWriter.create_row
    orig_set = xw.ExcelWriter._set_cell

    def spy_create(self, mgmt):
        seen.append(("create_row:start", self.backup_path is not None and Path(self.backup_path).exists()))
        return orig_create(self, mgmt)

    def spy_set(self, row, field, value, **kw):
        seen.append((f"first_write:{field}", self.backup_path is not None and Path(self.backup_path).exists()))
        return orig_set(self, row, field, value, **kw)

    monkeypatch.setattr(xw.ExcelWriter, "create_row", spy_create)
    monkeypatch.setattr(xw.ExcelWriter, "_set_cell", spy_set)
    summary, _, proc = _run([sample_tree["files"][0]], template, out)
    assert proc.results[0].new_row and summary.backup_file and Path(summary.backup_file).exists()
    assert seen[0] == ("create_row:start", False)                                   # no backup for a skip-only batch
    assert seen[1][0].startswith("first_write:") and seen[1][1] is True               # backup exists before 1st write
    # the backup is the pre-change master: no new key inside
    assert _keys(load_workbook(summary.backup_file)[SHEET]) == []
    assert "260918080-VOC" in _keys(load_workbook(out)[SHEET])


# ---------------------------------------------------------------- 12. backup failure -> no row, no modification
def test_12_backup_failure_creates_no_row(sample_tree, template, tmp_path, monkeypatch):
    out = tmp_path / "out" / "k.xlsx"
    _prefill(template, ["999999999-VOC"])

    orig = xw.shutil.copyfile

    def boom(src, dst, *a, **k):
        if "_backup_" in Path(dst).name:
            raise OSError("disk full")
        return orig(src, dst, *a, **k)

    monkeypatch.setattr(xw.shutil, "copyfile", boom)
    summary, _, proc = _run([sample_tree["files"][0]], template, out)
    fr = proc.results[0]
    assert fr.status == "error" and not fr.new_row and fr.excel_row is None
    assert "Không tạo được bản sao lưu Excel" in fr.error and "Không tạo được dòng mới" in fr.error
    assert summary.new_rows == 0 and not summary.backup_file
    assert _keys(load_workbook(out)[SHEET]) == ["999999999-VOC"]                    # output = untouched master copy


# ---------------------------------------------------------------- 13. save after every record (key on disk early)
def test_13_save_after_record_new_row_key_is_on_disk_before_parsing(sample_tree, template, tmp_path, monkeypatch):
    out = tmp_path / "out" / "k.xlsx"
    seen = {}
    orig = bp.parse_pptx

    def spy(path, *a, **k):
        seen["keys_at_parse"] = _keys(load_workbook(out)[SHEET])
        return orig(path, *a, **k)

    monkeypatch.setattr(bp, "parse_pptx", spy)
    summary, _, proc = _run([sample_tree["files"][0]], template, out)
    assert seen["keys_at_parse"] == ["260918080-VOC"]                                # created + saved before parse
    assert proc.results[0].status == "completed"
    ws = load_workbook(out)[SHEET]
    assert ws.cell(row=4, column=COL["model"]).value == "A185"                         # final save after the record


# ---------------------------------------------------------------- 14. same report twice in one batch -> one row
def test_14_same_report_twice_in_one_batch_creates_one_row_then_skips(sample_tree, template, tmp_path):
    out = tmp_path / "out" / "k.xlsx"
    rep = sample_tree["files"][0]
    twin = _copy_as(rep, tmp_path / "copy", rep.name)
    summary, _, proc = _run([rep, twin], template, out)
    done = [r for r in proc.results if r.new_row]
    assert len(done) == 1 and done[0].excel_row == 4                                 # exactly one row created
    other = [r for r in proc.results if not r.new_row]
    assert len(other) == 1 and other[0].excel_row in (None, 4) and other[0].status != "completed"
    assert summary.new_rows == 1 and _keys(load_workbook(out)[SHEET]) == ["260918080-VOC"]
    # the writer itself is idempotent too (same key asked twice inside one workbook session -> same row)
    w = ExcelWriter(template, tmp_path / "out2" / "k.xlsx")
    assert w.create_row("260918080-VOC") == w.create_row("260918080-VOC") == 4


# ---------------------------------------------------------------- 15 + 16. second run -> found, skipped before Qwen, no duplicate
def test_15_16_second_run_finds_row_and_skips_before_any_parse(sample_tree, template, tmp_path, monkeypatch):
    out = tmp_path / "out" / "k.xlsx"
    s1, _, p1 = _run([sample_tree["files"][0]], template, out)
    assert p1.results[0].new_row and s1.new_rows == 1
    calls = {"parse": 0, "classify": 0}
    monkeypatch.setattr(bp, "parse_pptx", lambda *a, **k: calls.__setitem__("parse", calls["parse"] + 1) or (_ for _ in ()).throw(AssertionError("parsed")))
    s2, events, p2 = _run([sample_tree["files"][0]], template, out)
    assert p2.results[0].status == "skipped" and p2.results[0].excel_row == 4 and not p2.results[0].new_row
    assert s2.new_rows == 0 and s2.skipped == 1 and calls["parse"] == 0 and not s2.backup_file
    assert _keys(load_workbook(out)[SHEET]) == ["260918080-VOC"]
    assert all(stage not in ("analyzing", "reading") for _, stage, _ in events)


# ---------------------------------------------------------------- 17. single-row / duplicate-row behaviour unchanged
def test_17_single_and_duplicate_existing_rows_never_get_a_new_row(sample_tree, template, tmp_path):
    out = tmp_path / "out" / "k.xlsx"
    _prefill(template, ["260918080-VOC", "260918081-VOC", "260918081-VOC"], complete=False)
    summary, _, proc = _run(sample_tree["files"][:2], template, out)
    single, dup = proc.results
    assert single.excel_row == 4 and not single.new_row and single.duplicate_rows == []
    assert dup.excel_row == 5 and not dup.new_row and dup.duplicate_rows == [6] and dup.status == "completed"
    assert summary.new_rows == 0
    ws = load_workbook(out)[SHEET]
    assert _keys(ws) == ["260918080-VOC", "260918081-VOC", "260918081-VOC"]
    assert ws.cell(row=6, column=2).fill.fgColor.rgb.endswith("FFC7CE") and ws.cell(row=6, column=COL["model"]).value is None


# ---------------------------------------------------------------- placement: never header / footer, deterministic
def test_18_placement_is_next_data_row_and_deterministic(template, tmp_path):
    out = tmp_path / "out" / "k.xlsx"
    _prefill(template, ["260901001-VOC", "260901002-VOC"])
    wb = load_workbook(template)
    wb[SHEET].cell(row=9, column=1, value="Ghi chú cuối bảng")                       # footer-like text below the data
    wb.save(template)
    w = ExcelWriter(template, out)
    r1 = w.create_row("260925015")
    r2 = w.create_row("260925015")                                                   # idempotent inside a batch
    r3 = w.create_row("260925016-VOC")
    assert (r1, r2, r3) == (6, 6, 7) and r1 >= w.data_start
    w.save()
    ws = load_workbook(out)[SHEET]
    assert ws.cell(row=2, column=2).value == "Management Number" or ws.cell(row=2, column=2).value is not None  # header intact
    assert [ws.cell(row=r, column=2).value for r in (6, 7)] == ["260925015", "260925016-VOC"]
    assert ws.cell(row=9, column=1).value == "Ghi chú cuối bảng"


# ---------------------------------------------------------------- prescan / GUI planning strings
def test_19_prescan_plans_new_row_as_processable_not_error(tmp_path, template):
    rep = tmp_path / "in" / "260925015_new.pptx"
    rep.parent.mkdir(parents=True)
    rep.write_bytes(b"x")                                                              # the pre-scan must not open it
    res = prescan([rep], None, None, MasterLookup(lambda m: [], lambda r: []), today=date(2026, 9, 30))
    item = res.items[0]
    assert item.action == ACTION_PROCESS_NEW_ROW and item in res.candidates
    assert item.reason == "Sẽ thêm dòng mới cho Management Number 260925015"
    assert res.counts()["new_rows"] == 1 and res.counts()["candidates"] == 1
    assert "Management Number mới (sẽ thêm dòng): 1" in res.summary_lines_vi()
    # an invalid YYMMDD prefix is still rejected without creating anything
    bad = tmp_path / "in" / "261345999_bad.pptx"
    bad.write_bytes(b"x")
    res2 = prescan([bad], None, None, MasterLookup(lambda m: [], lambda r: []), today=date(2026, 9, 30))
    assert res2.items[0].action != ACTION_PROCESS_NEW_ROW and res2.counts()["new_rows"] == 0
