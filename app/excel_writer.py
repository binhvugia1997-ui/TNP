"""Write extracted records into the user's Excel verification template.

* Never overwrites the template: output workbook = copy of template (or the
  existing output workbook is re-opened so batches can be resumed / appended).
* Sheet ``Kiểm chứng`` is located by (accent-insensitive) name; header cells
  are matched by field aliases so the exact column layout of the template does
  not matter.
* Formatting of the previous data row (borders, fonts, alignment, single-row
  merges) is copied to every new row.
* Images are embedded (not linked) and scaled to fit the target column.
"""
from __future__ import annotations

import logging
import math
import re
import shutil
from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

from .pptx_parser import norm_key

LOG = logging.getLogger("report_extractor.excel")

TARGET_SHEET_KEY = "kiem chung"
MAPPING_SHEET_KEYS = ("phan loai", "classification", "mapping", "danh muc")

# field -> list of accent-insensitive aliases (norm_key form)
FIELD_ALIASES: Dict[str, List[str]] = {
    "stt": ["stt", "no", "so tt", "tt"],
    "management_number": ["management number", "management no", "mgmt no", "ma quan ly", "so quan ly",
                          "management"],
    "vendor": ["ten vendor", "vendor", "nha cung cap", "ncc"],
    "occurrence_date": ["ngay phat sinh", "occurrence date", "ngay xay ra", "ngay"],
    "model": ["model", "model name"],
    "item": ["item", "hang muc", "linh kien", "part"],
    "defect_content": ["noi dung loi", "defect content", "noi dung khong phu hop", "hien tuong", "defect"],
    "qpn": ["qpn", "quality problem notice", "hinh anh qpn"],
    "root_cause": ["nguyen nhan", "root cause", "cause"],
    "improvement": ["noi dung doi sach cai tien", "noi dung doi sach", "doi sach cai tien", "doi sach",
                    "countermeasure", "corrective action", "cai tien"],
    "improvement_image": ["hinh anh cai tien", "hinh anh doi sach", "hinh anh", "image", "improvement image"],
    "status": ["trang thai", "status", "ket qua"],
    "note": ["ghi chu", "note", "remark", "remarks"],
}
for _i in range(1, 9):
    FIELD_ALIASES[f"week_{_i}"] = [f"week +{_i}", f"week+{_i}", f"week {_i}", f"w +{_i}", f"w+{_i}",
                                   f"tuan +{_i}", f"tuan {_i}", f"w{_i}"]

MAX_ROW_HEIGHT_PT = 409.0
EXCEL_DEFAULT_COL_WIDTH = 8.43


def _alias_score(cell_key: str, field: str) -> int:
    if not cell_key:
        return 0
    best = 0
    for alias in FIELD_ALIASES[field]:
        if cell_key == alias:
            best = max(best, 3)
        elif field.startswith("week_"):
            # week aliases must match exactly-ish ("week +1" != "week +10")
            if re.fullmatch(re.escape(alias) + r"(?:\D.*)?", cell_key):
                best = max(best, 2)
        elif len(alias) >= 4 and re.search(r"(?<![a-z0-9])" + re.escape(alias) + r"(?![a-z0-9])", cell_key):
            best = max(best, 2 if len(alias) > 5 else 1)
    return best


class TemplateError(RuntimeError):
    pass


class ExcelWriter:
    def __init__(self, template: Path, output: Path):
        self.template = Path(template)
        self.output = Path(output)
        if not self.template.exists():
            raise TemplateError(f"Không tìm thấy form Excel: {self.template}")
        if self.template.resolve() == self.output.resolve():
            raise TemplateError("File kết quả không được trùng với form Excel gốc")
        self.output.parent.mkdir(parents=True, exist_ok=True)
        if not self.output.exists():
            shutil.copyfile(self.template, self.output)
        self.wb = load_workbook(self.output)
        self.ws = self._find_target_sheet()
        self.columns: Dict[str, int] = {}
        self.header_row = 0
        self.data_start = 0
        self._detect_headers()
        self.item_mapping, self.known_models = self._read_mapping_sheet()
        self._dirty = False

    # ------------------------------------------------------------------
    # Template analysis
    # ------------------------------------------------------------------
    def _find_target_sheet(self):
        for ws in self.wb.worksheets:
            if norm_key(ws.title) == TARGET_SHEET_KEY:
                return ws
        for ws in self.wb.worksheets:
            if TARGET_SHEET_KEY in norm_key(ws.title):
                return ws
        raise TemplateError("Không tìm thấy sheet 'Kiểm chứng' trong form Excel")

    def _detect_headers(self) -> None:
        ws = self.ws
        max_scan = min(ws.max_row, 30)
        row_matches: Dict[int, Dict[str, Tuple[int, int]]] = {}
        for r in range(1, max_scan + 1):
            matches: Dict[str, Tuple[int, int]] = {}
            for c in range(1, min(ws.max_column, 80) + 1):
                v = ws.cell(row=r, column=c).value
                if not isinstance(v, str) or not v.strip():
                    continue
                if len(v) > 60 or "\n" in v.strip():      # data text, not a header label
                    continue
                key = norm_key(v)
                for field in FIELD_ALIASES:
                    s = _alias_score(key, field)
                    if s and (field not in matches or s > matches[field][1]):
                        matches[field] = (c, s)
            if matches:
                row_matches[r] = matches
        if not row_matches:
            raise TemplateError("Không nhận diện được dòng tiêu đề trong sheet 'Kiểm chứng'")
        # header row = row with most strong matches
        def strength(r: int) -> Tuple[int, int]:
            m = row_matches[r]
            core = sum(1 for f in ("improvement", "root_cause", "qpn", "model", "item", "defect_content") if f in m)
            return (core, len(m))
        header = max(row_matches, key=strength)
        self.header_row = header
        cols: Dict[str, int] = {}
        # merge header row + up to 2 sub-header rows (WEEK +1.. often sit below a merged title)
        for r in (header, header + 1, header + 2):
            if r in row_matches:
                for f, (c, s) in row_matches[r].items():
                    if f not in cols:
                        cols[f] = c
        last_header = header
        for r in (header + 1, header + 2):
            if r in row_matches and any(f.startswith("week_") or f in ("model", "item") for f in row_matches[r]):
                last_header = r
        # if row below header has any string cells that look like headers (all strings, no numbers), skip it
        self.columns = cols
        self.data_start = last_header + 1
        missing = [f for f in ("improvement", "root_cause") if f not in cols]
        if missing:
            LOG.warning("Template columns not found: %s (columns found: %s)", missing, cols)

    def _read_mapping_sheet(self) -> Tuple[Dict[str, str], List[str]]:
        mapping: Dict[str, str] = {}
        models: List[str] = []
        for ws in self.wb.worksheets:
            if not any(k in norm_key(ws.title) for k in MAPPING_SHEET_KEYS):
                continue
            rows = [[c for c in row] for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 2000), values_only=True)]
            # detect header
            hdr_idx, item_col, key_col, model_col = None, None, None, None
            for i, row in enumerate(rows[:10]):
                keys = [norm_key(str(v)) if v is not None else "" for v in row]
                for ci, k in enumerate(keys):
                    if k in ("item", "hang muc", "linh kien", "ten item", "phan loai item") and item_col is None:
                        item_col = ci
                    elif k in ("model", "ten model", "model name") and model_col is None:
                        model_col = ci
                    elif k in ("tu khoa", "keyword", "keywords", "ten", "ten linh kien", "mo ta", "noi dung",
                               "ten goi", "alias") and key_col is None:
                        key_col = ci
                if item_col is not None or model_col is not None:
                    hdr_idx = i
                    break
            body = rows[hdr_idx + 1:] if hdr_idx is not None else rows
            for row in body:
                vals = [str(v).strip() for v in row if v is not None and str(v).strip()]
                if hdr_idx is not None:
                    if item_col is not None and item_col < len(row) and row[item_col]:
                        item = str(row[item_col]).strip()
                        mapping.setdefault(item, item)
                        if key_col is not None and key_col < len(row) and row[key_col]:
                            for kw in re.split(r"[;,/\n]+", str(row[key_col])):
                                kw = kw.strip()
                                if kw:
                                    mapping.setdefault(kw, item)
                    if model_col is not None and model_col < len(row) and row[model_col]:
                        for mv in re.split(r"[;,/\n]+", str(row[model_col])):
                            mv = mv.strip()
                            if mv and mv not in models:
                                models.append(mv)
                else:
                    if len(vals) >= 2:
                        mapping.setdefault(vals[0], vals[1])
                        mapping.setdefault(vals[1], vals[1])
                    elif len(vals) == 1:
                        mapping.setdefault(vals[0], vals[0])
            break
        if mapping:
            LOG.info("Item mapping loaded: %d keywords, %d models", len(mapping), len(models))
        return mapping, models

    # ------------------------------------------------------------------
    # Row helpers
    # ------------------------------------------------------------------
    def _anchor(self, row: int, col: int):
        """Top-left cell of the merged range containing (row, col), or the cell itself."""
        for rng in self.ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                return self.ws.cell(row=rng.min_row, column=rng.min_col), rng
        return self.ws.cell(row=row, column=col), None

    def _row_used(self, row: int) -> bool:
        check = [f for f in self.columns if f not in ("stt", "vendor", "occurrence_date", "status", "note")
                 and not f.startswith("week_")]
        for f in check:
            cell, _ = self._anchor(row, self.columns[f])
            if cell.value not in (None, ""):
                return True
        # rows holding only an image count as used too
        for img in getattr(self.ws, "_images", []):
            try:
                if img.anchor._from.row + 1 == row:
                    return True
            except Exception:
                pass
        return False

    def next_row(self) -> int:
        r = self.data_start
        limit = max(self.ws.max_row, self.data_start) + 1
        while r <= limit and self._row_used(r):
            r += 1
        return r

    def used_count(self) -> int:
        return sum(1 for r in range(self.data_start, self.next_row()) if self._row_used(r))

    def _copy_row_style(self, src_row: int, dst_row: int) -> None:
        ws = self.ws
        if src_row < self.data_start or src_row == dst_row:
            return
        max_col = max(list(self.columns.values()) + [ws.max_column])
        for c in range(1, max_col + 1):
            s = ws.cell(row=src_row, column=c)
            d = ws.cell(row=dst_row, column=c)
            if isinstance(d, MergedCell):
                continue
            if s.has_style:
                d._style = copy(s._style)
        # replicate single-row merges of the source row
        for rng in list(ws.merged_cells.ranges):
            if rng.min_row == src_row and rng.max_row == src_row:
                try:
                    ws.merge_cells(start_row=dst_row, start_column=rng.min_col,
                                   end_row=dst_row, end_column=rng.max_col)
                except Exception:
                    pass
        if ws.row_dimensions[src_row].height:
            ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height

    def _col_width_px(self, col: int, rng=None) -> int:
        cols = range(rng.min_col, rng.max_col + 1) if rng is not None else [col]
        total = 0.0
        for c in cols:
            w = self.ws.column_dimensions[get_column_letter(c)].width or EXCEL_DEFAULT_COL_WIDTH
            total += w
        return int(total * 7 + 5)

    def _set_cell(self, row: int, field: str, value: Any, wrap: bool = True) -> Optional[Any]:
        if field not in self.columns:
            return None
        cell, _ = self._anchor(row, self.columns[field])
        cell.value = value if value not in ("",) else None
        if wrap and isinstance(value, str):
            al = copy(cell.alignment) if cell.alignment else Alignment()
            cell.alignment = Alignment(horizontal=al.horizontal or "left", vertical="top", wrap_text=True,
                                       indent=al.indent, text_rotation=al.text_rotation)
        return cell

    def _embed_image(self, row: int, field: str, path: Path, max_height_pt: float = 400.0) -> float:
        """Embed picture in the field cell; returns the required row height (pt)."""
        if field not in self.columns or not path or not Path(path).exists():
            return 0.0
        col = self.columns[field]
        cell, rng = self._anchor(row, col)
        col_px = max(40, self._col_width_px(col, rng) - 6)
        with PILImage.open(path) as im:
            w, h = im.size
        scale = min(col_px / w, (max_height_pt / 0.75) / h, 1.0)
        img = XLImage(str(path))
        img.width = int(w * scale)
        img.height = int(h * scale)
        img.anchor = f"{get_column_letter(cell.column)}{cell.row}"
        self.ws.add_image(img)
        return img.height * 0.75 + 6

    def _text_height_pt(self, row: int, field: str, text: str) -> float:
        if field not in self.columns or not text:
            return 0.0
        col = self.columns[field]
        _, rng = self._anchor(row, col)
        width_chars = max(8.0, self._col_width_px(col, rng) / 7.0)
        lines = 0
        for ln in text.split("\n"):
            lines += max(1, math.ceil(len(ln) / max(1.0, width_chars * 1.05)))
        return lines * 15.0 + 4

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def append_record(self, rec, qpn_png: Optional[Path] = None, improvement_jpg: Optional[Path] = None,
                      status_text: str = "", note_text: str = "") -> int:
        """Write one report as one new row. Returns the Excel row number."""
        row = self.next_row()
        src = row - 1 if row - 1 >= self.data_start else self.data_start
        if src != row:
            self._copy_row_style(src, row)
        if "stt" in self.columns:
            self._set_cell(row, "stt", self.used_count() + 1, wrap=False)
        self._set_cell(row, "management_number", rec.management_number)
        # Business rules: user fills manually -> ALWAYS blank
        self._set_cell(row, "vendor", "")
        self._set_cell(row, "occurrence_date", "")
        self._set_cell(row, "model", rec.model)
        self._set_cell(row, "item", rec.item)
        self._set_cell(row, "defect_content", rec.defect_content)
        self._set_cell(row, "root_cause", rec.root_cause)
        self._set_cell(row, "improvement", rec.improvement)
        for i in range(1, 9):   # WEEK +1..+8: blank unless real source data (none parsed) -> blank
            self._set_cell(row, f"week_{i}", rec.weeks.get(i, "") or "", wrap=False)
        if status_text:
            self._set_cell(row, "status", status_text)
        if note_text:
            self._set_cell(row, "note", note_text)

        heights = [self.ws.row_dimensions[row].height or 15.0]
        heights.append(self._text_height_pt(row, "improvement", rec.improvement))
        heights.append(self._text_height_pt(row, "root_cause", rec.root_cause))
        heights.append(self._text_height_pt(row, "defect_content", rec.defect_content))
        if qpn_png:
            heights.append(self._embed_image(row, "qpn", Path(qpn_png)))
        if improvement_jpg:
            heights.append(self._embed_image(row, "improvement_image", Path(improvement_jpg)))
        self.ws.row_dimensions[row].height = min(MAX_ROW_HEIGHT_PT, max(heights))
        self._dirty = True
        return row

    def _refresh_images(self) -> None:
        """Work around openpyxl closing in-memory image buffers after a save.

        Images read from an existing workbook keep a BytesIO ``ref`` that openpyxl
        closes during ``save``; the next save would fail with
        'I/O operation on closed file'.  We cache the bytes once and hand a fresh
        buffer to openpyxl before every save.
        """
        for ws in self.wb.worksheets:
            for img in getattr(ws, "_images", []):
                ref = img.ref
                if isinstance(ref, (str, Path)):
                    continue
                cached = getattr(img, "_re_bytes", None)
                if cached is None:
                    try:
                        if hasattr(ref, "getvalue"):
                            cached = ref.getvalue()
                        else:
                            ref.seek(0)
                            cached = ref.read()
                    except Exception:
                        continue
                    img._re_bytes = cached  # type: ignore[attr-defined]
                img.ref = BytesIO(cached)

    def save(self) -> Path:
        tmp = self.output.with_name(self.output.stem + ".saving.xlsx")
        self._refresh_images()
        self.wb.save(tmp)
        shutil.move(str(tmp), str(self.output))
        self._dirty = False
        return self.output

    def close(self) -> None:
        try:
            self.wb.close()
        except Exception:
            pass


def validate_template(template: Path) -> Tuple[bool, str]:
    """Quick diagnostic: can we find the sheet and the key columns?"""
    try:
        wb = load_workbook(template, read_only=False)
    except Exception as e:  # noqa: BLE001
        return False, f"Không mở được file: {e}"
    try:
        ws = None
        for w in wb.worksheets:
            if TARGET_SHEET_KEY in norm_key(w.title):
                ws = w
                break
        if ws is None:
            return False, "Thiếu sheet 'Kiểm chứng'"
        found = set()
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 30)):
            for c in row:
                if isinstance(c.value, str):
                    k = norm_key(c.value)
                    for f in FIELD_ALIASES:
                        if _alias_score(k, f):
                            found.add(f)
        need = {"improvement", "root_cause", "qpn"}
        missing = need - found
        if missing:
            return False, "Thiếu cột: " + ", ".join(sorted(missing))
        return True, f"OK ({len(found)} cột nhận diện)"
    finally:
        wb.close()
