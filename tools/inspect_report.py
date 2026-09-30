"""Dump a real PPTX report: slide text, pictures, heuristic classification, extracted
record and (optionally) the raw + normalised Qwen answer.

    python tools/inspect_report.py report.pptx [--template T.xlsx] [--server 192.168.1.50:11434 --model qwen3:4b]
"""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.classifier import (SYSTEM_PROMPT, build_prompt, heuristic_classify,  # noqa: E402
                            merge_llm_with_heuristic, normalize_llm_response)
from app.extractor import extract_record  # noqa: E402
from app.pptx_parser import parse_pptx  # noqa: E402


def inspect_report(path: Path, template: Path | None = None, server: str = "", model: str = "",
                   full_text: bool = True) -> str:
    out = [f"REPORT: {path}"]
    r = parse_pptx(path)
    out.append(f"Slides: {len(r.slides)}  size: {r.slide_width}x{r.slide_height} EMU")
    for s in r.slides:
        out.append(f"\n=== Slide {s.number}  ({len(s.text_blocks)} text blocks, {len(s.pictures)} pictures) ===")
        txt = s.text if full_text else s.text[:600]
        out.append(txt or "(no text)")
    heur = heuristic_classify(r)
    out.append("\n--- Heuristic classification ---")
    out.append(json.dumps(heur.to_dict(), ensure_ascii=False, indent=1))

    item_mapping, known_models = {}, []
    if template and Path(template).exists():
        from app.excel_writer import ExcelWriter
        with tempfile.TemporaryDirectory() as tmp:
            w = ExcelWriter(Path(template), Path(tmp) / "probe.xlsx")
            item_mapping, known_models = w.item_mapping, w.known_models
            w.close()

    cls = heur
    if server and model:
        from app.ollama_client import OllamaClient
        out.append(f"\n--- Qwen ({server} / {model}) ---")
        try:
            raw = OllamaClient(server).generate_json(model, build_prompt(r), system=SYSTEM_PROMPT)
            out.append("RAW JSON:\n" + json.dumps(raw, ensure_ascii=False, indent=1))
            out.append("NORMALISED:\n" + json.dumps(normalize_llm_response(raw), ensure_ascii=False, indent=1, default=str))
            cls = merge_llm_with_heuristic(raw, heur, r)
            out.append("MERGED:\n" + json.dumps(cls.to_dict(), ensure_ascii=False, indent=1))
        except Exception as e:  # noqa: BLE001
            out.append(f"Qwen FAILED -> heuristic fallback: {type(e).__name__}: {e}")

    rec = extract_record(r, cls, item_mapping, known_models)
    out.append("\n--- Extracted record (what would go to Excel) ---")
    out.append(f"classifier          : {cls.source}")
    out.append(f"Management number   : {rec.management_number!r}")
    out.append("Tên vendor          : '' (always blank)")
    out.append("Ngày phát sinh      : '' (always blank)")
    out.append(f"Model               : {rec.model!r}")
    out.append(f"Item                : {rec.item!r}")
    out.append(f"Nội dung lỗi        :\n{rec.defect_content}")
    out.append(f"QPN slide           : {rec.qpn_slide}")
    out.append(f"Nguyên nhân ({len(rec.root_cause)} chars):\n{rec.root_cause}")
    out.append(f"Nội dung đối sách cải tiến ({len(rec.improvement)} chars):\n{rec.improvement}")
    out.append(f"[EXCLUDED] Xử lý tạm thời ({len(rec.temporary_excluded)} chars):\n{rec.temporary_excluded}")
    out.append(f"Hình ảnh cải tiến   : slides {rec.improvement_image_slides}")
    out.append("WEEK +1..+8         : blank")
    out.append(f"Blank fields        : {rec.blank_fields}")
    out.append(f"Cần kiểm tra        : {rec.review_reasons}")
    return "\n".join(out)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--template")
    ap.add_argument("--server", default="")
    ap.add_argument("--model", default="")
    ap.add_argument("--short", action="store_true", help="truncate slide text")
    a = ap.parse_args()
    print(inspect_report(Path(a.pptx), Path(a.template) if a.template else None, a.server, a.model, not a.short))
