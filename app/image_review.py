"""PROMPT-006 – re-apply user-confirmed improvement-image labels to the Excel result.

After "Lưu xác nhận" in the review window the affected reports are re-evaluated (deterministic rules + confirmed
labels; no Ollama) and ONLY the ``improvement_image`` cell of the matching Management-Number row is rewritten.
Every other business field, WEEK columns, formulas, other rows and sheets stay untouched; the writer's normal
backup-before-first-modification rule applies.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence

from .excel_writer import ExcelWriter
from .image_extractor import export_after_pictures
from .image_learning import ImageCandidate, ImageLearning, select_with_learning
from .improvement_pictures import select_after_pictures
from .pptx_parser import parse_pptx

LOG = logging.getLogger("report_extractor.image_review")


@dataclass
class ReapplyResult:
    updated_rows: List[int] = field(default_factory=list)
    pictures: int = 0
    messages: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors


def group_by_file(cands: Sequence[ImageCandidate]) -> Dict[str, List[ImageCandidate]]:
    out: Dict[str, List[ImageCandidate]] = {}
    for c in cands:
        if c.source_file:
            out.setdefault(c.source_file, []).append(c)
    return out


def reapply_labels(template: Path, output: Path, cands: Sequence[ImageCandidate], learning: ImageLearning,
                   assets_dir: Optional[Path] = None, writer: Optional[ExcelWriter] = None) -> ReapplyResult:
    """Recompute the After pictures of every report represented in ``cands`` and rewrite only their
    improvement-image cells.  Reports whose Management Number is not found (or found several times) are reported
    and skipped – the batch rules for duplicates are not changed here."""
    res = ReapplyResult()
    groups = group_by_file(cands)
    if not groups:
        res.messages.append("Không có ảnh nào cần cập nhật.")
        return res
    own_writer = writer is None
    try:
        if writer is None:
            writer = ExcelWriter(Path(template), Path(output))
    except Exception as e:  # noqa: BLE001
        res.errors.append(f"Không mở được file Excel kết quả: {e}")
        return res
    assets_dir = Path(assets_dir) if assets_dir else Path(output).parent / "assets"
    try:
        for src, group in groups.items():
            mn = group[0].management_number
            try:
                report = parse_pptx(src)
            except Exception as e:  # noqa: BLE001
                res.errors.append(f"{Path(src).name}: không đọc lại được PPTX ({e})")
                continue
            slides = sorted({c.slide for c in group})
            sel = select_after_pictures(report, slides)
            refs, new_cands, _ = select_with_learning(report, slides, sel, learning, mn, src)
            rows = writer.find_rows_by_management_number(mn) if mn else []
            if not rows:
                res.errors.append(f"{Path(src).name}: không tìm thấy dòng Management Number {mn!r} trong Excel")
                continue
            row = rows[0]                                             # canonical = topmost (PROMPT-004 rule)
            paths, problems = export_after_pictures(report, refs, assets_dir / f"{Path(src).stem}_IMPROVEMENT_pics")
            try:
                n = writer.replace_improvement_images(row, paths)
            except Exception as e:  # noqa: BLE001
                res.errors.append(f"{Path(src).name}: không ghi được ảnh vào dòng {row} ({e})")
                continue
            res.updated_rows.append(row)
            res.pictures += n
            res.messages.append(f"{Path(src).name}: dòng {row} – {n} ảnh Sau cải tiến"
                                + (f" ({'; '.join(problems)})" if problems else ""))
            # keep the live candidate objects in sync for the GUI
            by_id = {c.candidate_id: c for c in new_cands}
            for c in group:
                nc = by_id.get(c.candidate_id)
                if nc is not None:
                    c.decision, c.decision_source, c.user_label = nc.decision, nc.decision_source, nc.user_label
                    c.evidence = list(nc.evidence)
        if res.updated_rows:
            writer.save()
    except Exception as e:  # noqa: BLE001
        res.errors.append(f"Lỗi khi lưu Excel: {e}")
    finally:
        if own_writer:
            writer.close()
    return res
