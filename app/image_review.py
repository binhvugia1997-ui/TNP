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

from .classifier import heuristic_classify
from .content_learning import ContentCandidate, select_content_with_learning
from .excel_writer import ExcelLockedError, ExcelWriter
from .extractor import collect_sections, join_sections
from .image_extractor import export_after_pictures
from .image_learning import ImageCandidate, ImageLearning, select_with_learning
from .improvement_pictures import select_after_pictures
from .pptx_parser import parse_pptx
from .report_identity import report_scope_key

LOG = logging.getLogger("report_extractor.image_review")


@dataclass
class ReapplyResult:
    """Transaction view of an Excel re-apply.

    ``prepared_rows`` are rows modified in memory; ``updated_rows`` (== committed rows) is filled ONLY after the
    workbook has been written and atomically moved onto the target.  A locked target leaves ``updated_rows`` empty,
    sets ``locked_path`` and keeps the friendly message in ``errors`` (technical detail in ``technical_error`` +
    log)."""
    updated_rows: List[int] = field(default_factory=list)          # committed rows
    prepared_rows: List[int] = field(default_factory=list)
    pictures: int = 0
    messages: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    locked_path: Optional[str] = None
    technical_error: str = ""

    @property
    def ok(self) -> bool:
        return not self.errors

    @property
    def committed_rows(self) -> List[int]:
        return self.updated_rows

    @property
    def locked(self) -> bool:
        return self.locked_path is not None


def _commit(writer: ExcelWriter, res: ReapplyResult, kind: str) -> None:
    """Commit boundary shared by both re-apply flows: rows become 'updated' only after a successful save."""
    target = str(writer.output)
    LOG.info("%s_REAPPLY_PREPARED rows=%d target=%s", kind, len(res.prepared_rows), target)
    if not res.prepared_rows:
        return
    try:
        writer.save()
    except ExcelLockedError as e:
        res.locked_path = str(e.path)
        res.technical_error = repr(e.original)
        res.errors.append(str(e))
        LOG.error("%s_REAPPLY_COMMIT_FAILED target=%s prepared_rows=%d error=%r", kind, e.path,
                  len(res.prepared_rows), e.original)
        return
    except Exception as e:  # noqa: BLE001
        res.technical_error = repr(e)
        res.errors.append(f"Lỗi khi lưu Excel: {e}")
        LOG.error("%s_REAPPLY_COMMIT_FAILED target=%s prepared_rows=%d error=%r", kind, target,
                  len(res.prepared_rows), e)
        return
    res.updated_rows = list(res.prepared_rows)
    LOG.info("%s_REAPPLY_COMMIT_OK rows=%d target=%s", kind, len(res.updated_rows), target)


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
    LOG.info("IMAGE_REAPPLY_START target=%s files=%d", output, len(groups))
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
            scope = report_scope_key(src)
            report_assets = assets_dir / f"{Path(src).stem}_{scope}_IMPROVEMENT_pics"
            paths, problems = export_after_pictures(report, refs, report_assets)
            try:
                n = writer.replace_improvement_images(row, paths)
            except Exception as e:  # noqa: BLE001
                res.errors.append(f"{Path(src).name}: không ghi được ảnh vào dòng {row} ({e})")
                continue
            res.prepared_rows.append(row)
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
        _commit(writer, res, "IMAGE")
    except Exception as e:  # noqa: BLE001
        res.technical_error = repr(e)
        res.errors.append(f"Lỗi khi cập nhật Excel: {e}")
    finally:
        if own_writer:
            writer.close()
    return res


# ============================================================================ PROMPT-006B content re-apply
def group_content_by_file(cands: Sequence[ContentCandidate]) -> Dict[str, List[ContentCandidate]]:
    out: Dict[str, List[ContentCandidate]] = {}
    for c in cands:
        if c.source_file:
            out.setdefault(c.source_file, []).append(c)
    return out


def reapply_content_labels(template: Path, output: Path, cands: Sequence[ContentCandidate], learning,
                           writer: Optional[ExcelWriter] = None) -> ReapplyResult:
    """Rebuild the improvement TEXT of every report represented in ``cands`` from the deterministic sections plus
    the confirmed block labels, and rewrite only the ``improvement`` cell of the matching Management-Number row.
    ``learning`` is the :class:`app.image_learning.ImageLearning` facade (``.content``) or a ContentLearning."""
    res = ReapplyResult()
    groups = group_content_by_file(cands)
    if not groups:
        res.messages.append("Không có khối nội dung nào cần cập nhật.")
        return res
    content = getattr(learning, "content", learning)
    own_writer = writer is None
    LOG.info("CONTENT_REAPPLY_START target=%s files=%d", output, len(groups))
    try:
        if writer is None:
            writer = ExcelWriter(Path(template), Path(output))
    except Exception as e:  # noqa: BLE001
        res.errors.append(f"Không mở được file Excel kết quả: {e}")
        return res
    try:
        for src, group in groups.items():
            mn = group[0].management_number
            try:
                report = parse_pptx(src)
                cls = heuristic_classify(report)                       # structural only – no Ollama on re-apply
            except Exception as e:  # noqa: BLE001
                res.errors.append(f"{Path(src).name}: không đọc lại được PPTX ({e})")
                continue
            sections = collect_sections(report, cls)
            imp_sections = [s for s in sections if s.kind in ("improvement", "standard")]
            slides = sorted(set(cls.improvement_slides) | {s.slide for s in imp_sections} | {c.slide for c in group})
            text, new_cands, _ = select_content_with_learning(report, sections, slides, content, mn, src)
            if text is None:
                text = join_sections(imp_sections)
            rows = writer.find_rows_by_management_number(mn) if mn else []
            if not rows:
                res.errors.append(f"{Path(src).name}: không tìm thấy dòng Management Number {mn!r} trong Excel")
                continue
            row = rows[0]
            try:
                changed = writer.replace_improvement_text(row, text)
            except Exception as e:  # noqa: BLE001
                res.errors.append(f"{Path(src).name}: không ghi được nội dung vào dòng {row} ({e})")
                continue
            res.prepared_rows.append(row)
            res.messages.append(f"{Path(src).name}: dòng {row} – nội dung cải tiến "
                                + ("đã chuẩn bị" if changed else "không thay đổi"))
            by_id = {c.candidate_id: c for c in new_cands}
            for c in group:
                nc = by_id.get(c.candidate_id)
                if nc is not None:
                    c.decision, c.decision_source, c.user_label = nc.decision, nc.decision_source, nc.user_label
                    c.evidence = list(nc.evidence)
        _commit(writer, res, "CONTENT")
    except Exception as e:  # noqa: BLE001
        res.technical_error = repr(e)
        res.errors.append(f"Lỗi khi cập nhật Excel: {e}")
    finally:
        if own_writer:
            writer.close()
    return res
