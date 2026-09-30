"""Entry point.

    ReportExtractor.exe                 -> GUI
    ReportExtractor.exe --diag          -> print diagnostics
    ReportExtractor.exe --cli FOLDER --template T.xlsx --output OUT.xlsx [--server ..] [--model ..]
                                        -> headless batch (same pipeline as the GUI)
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


def force_utf8_stdio() -> None:
    """Windows consoles default to a legacy code page (cp437/cp1252) which turns
    Vietnamese into mojibake.  Re-open stdout/stderr as UTF-8 regardless of the
    console / PowerShell settings.  Never touches the source data."""
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream is not None and hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def write_utf8_report(text: str, path: Path) -> Path:
    """Write diagnostics as UTF-8 with BOM so Notepad/Excel/PowerShell show Vietnamese correctly."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8-sig", newline="\n")
    return path


def _cli(args) -> int:
    from .batch_processor import STAGE_LABELS_VI, BatchOptions, BatchProcessor
    from .config import AppConfig, normalize_ollama_url
    from .scanner import scan_inputs

    cfg = AppConfig.load()
    files = scan_inputs(args.inputs)
    if not files:
        print("Không tìm thấy file .pptx nào.")
        return 2
    server = normalize_ollama_url(args.server or cfg.ollama_server)
    model = args.model if args.model is not None else cfg.model
    opts = BatchOptions(files=files, template=Path(args.template), output_file=Path(args.output),
                        ollama_server=server, model=model or "", force_reprocess=args.force,
                        use_ollama=bool(model) and not args.no_ai,
                        fill_temporary_column=bool(cfg.fill_temporary_column))

    def on_file(i, stage, detail):
        print(f"[{i + 1}/{len(files)}] {files[i].name}: {STAGE_LABELS_VI.get(stage, stage)} {detail}".rstrip(), flush=True)

    proc = BatchProcessor(opts, on_file=on_file, on_log=lambda m: print(m, flush=True))
    s = proc.run()
    print(f"\nTổng: {s.total}\nHoàn thành: {s.completed}\nCần kiểm tra: {s.needs_review}\nLỗi: {s.failed}\nBỏ qua: {s.skipped}")
    print(f"Kết quả: {s.output_file}")
    return 0 if s.failed == 0 else 1


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="ReportExtractor", description="Report Extractor – PPTX → Excel kiểm chứng")
    parser.add_argument("--diag", action="store_true", help="in chẩn đoán hệ thống rồi thoát")
    parser.add_argument("--no-ollama-check", action="store_true")
    parser.add_argument("--cli", dest="inputs", nargs="+", metavar="PATH", help="chạy không giao diện: thư mục / file pptx")
    parser.add_argument("--template", help="form Excel .xlsx")
    parser.add_argument("--output", help="file Excel kết quả")
    parser.add_argument("--server", help="địa chỉ Ollama, vd 192.168.1.50:11434")
    parser.add_argument("--model", help="tên model Ollama (rỗng = không dùng AI)")
    parser.add_argument("--no-ai", action="store_true", help="chỉ dùng nhận diện từ khoá")
    parser.add_argument("--force", action="store_true", help="xử lý lại file đã xử lý")
    parser.add_argument("--inspect", metavar="PATH", help="in cấu trúc file .pptx hoặc form .xlsx để kiểm tra mapping")
    parser.add_argument("--out", metavar="FILE", help="ghi kết quả --inspect ra file UTF-8 (mặc định inspect_template.txt / inspect_report.txt)")
    parser.add_argument("--version", action="store_true")
    args = parser.parse_args(argv)
    force_utf8_stdio()

    if args.version:
        from . import __version__
        print(f"Report Extractor {__version__}")
        return 0
    if args.inspect:
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
        p = Path(args.inspect)
        if p.suffix.lower() in (".xlsx", ".xlsm"):
            from inspect_template import inspect_template
            text = inspect_template(p)
            default_out = "inspect_template.txt"
        else:
            from inspect_report import inspect_report
            text = inspect_report(p, Path(args.template) if args.template else None, args.server or "", args.model or "")
            default_out = "inspect_report.txt"
        out = write_utf8_report(text, Path(args.out or default_out))
        print(text)
        print(f"\n[Đã ghi UTF-8: {out.resolve()}]")
        return 0
    if args.diag:
        from .config import AppConfig
        from .diagnostics import format_diagnostics, run_diagnostics
        print(format_diagnostics(run_diagnostics(AppConfig.load(), check_ollama=not args.no_ollama_check)))
        return 0
    if args.inputs:
        if not args.template or not args.output:
            parser.error("--cli cần --template và --output")
        return _cli(args)

    try:
        from .gui import launch
    except ImportError as e:
        print(f"Không khởi động được giao diện (tkinter): {e}", file=sys.stderr)
        return 3
    launch()
    return 0


if __name__ == "__main__":
    sys.exit(main())
