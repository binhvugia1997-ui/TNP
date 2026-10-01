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
from typing import Optional


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


def _cli_period(args):
    """--all | --month MM/YYYY | --from dd/mm/yyyy --to dd/mm/yyyy | default: month from the Excel file name."""
    from .prescan import (ALL_PERIOD, AUTO_PERIOD_FAIL_VI, auto_period_from_excel, month_period, range_period)
    if getattr(args, "all_periods", False):
        return ALL_PERIOD, ""
    if getattr(args, "month", None):
        try:
            mm, yy = args.month.split("/")
            return month_period(int(yy), int(mm)), ""
        except (ValueError, TypeError):
            return None, f"--month không hợp lệ: {args.month!r} (định dạng MM/YYYY)"
    if getattr(args, "date_from", None) or getattr(args, "date_to", None):
        per, err = range_period(args.date_from or "", args.date_to or "")
        return per, err
    per = auto_period_from_excel(args.output) or auto_period_from_excel(args.template)
    return (per, "") if per else (None, AUTO_PERIOD_FAIL_VI + " (--month MM/YYYY, --from/--to dd/mm/yyyy hoặc --all)")


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
    period, perr = _cli_period(args)
    if perr:
        print(perr)
        return 2
    print(period.label_vi())
    opts = BatchOptions(period=period, files=files, template=Path(args.template), output_file=Path(args.output),
                        ollama_server=server, model=model or "", force_reprocess=args.force,
                        use_ollama=bool(model) and not args.no_ai,
                        fill_temporary_column=bool(cfg.fill_temporary_column),
                        vendors=list(cfg.vendors or []),
                        row_mode="append" if args.append else (cfg.row_mode or "match"))

    def on_file(i, stage, detail):
        print(f"[{i + 1}/{len(files)}] {files[i].name}: {STAGE_LABELS_VI.get(stage, stage)} {detail}".rstrip(), flush=True)

    proc = BatchProcessor(opts, on_file=on_file, on_log=lambda m: print(m, flush=True))
    s = proc.run()
    print(f"\nTổng: {s.total}\nHoàn thành: {s.completed}\nCần kiểm tra: {s.needs_review}\nChưa ghi (không tìm thấy dòng): {s.not_written}\nLỗi: {s.failed}\nBỏ qua: {s.skipped}\n"
          f"Ngoài thời gian xử lý: {s.outside_period}\nTrùng Management Number trong folder: {s.source_duplicates}\n"
          f"Bỏ qua nhanh — đã xử lý gần đây: {s.fast_skipped}\nCần xử lý thực tế: {s.candidates}")
    print(f"Kết quả: {s.output_file}")
    return 0 if s.failed == 0 else 1


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="ReportExtractor", description="Report Extractor – PPTX → Excel kiểm chứng")
    parser.add_argument("--diag", action="store_true", help="in chẩn đoán hệ thống rồi thoát")
    parser.add_argument("--no-ollama-check", action="store_true")
    parser.add_argument("--ollama-test", action="store_true",
                        help="gửi 1 yêu cầu JSON nhỏ tới Ollama (cùng mã nguồn với xử lý thật) và in thời gian")
    parser.add_argument("--cli", dest="inputs", nargs="+", metavar="PATH", help="chạy không giao diện: thư mục / file pptx")
    parser.add_argument("--template", help="form Excel .xlsx")
    parser.add_argument("--output", help="file Excel kết quả")
    parser.add_argument("--server", help="địa chỉ Ollama, vd 192.168.1.50:11434")
    parser.add_argument("--model", help="tên model Ollama (rỗng = không dùng AI)")
    parser.add_argument("--no-ai", action="store_true", help="chỉ dùng nhận diện từ khoá")
    parser.add_argument("--force", action="store_true", help="xử lý lại file đã xử lý")
    parser.add_argument("--month", help="thời gian xử lý: tháng MM/YYYY (mặc định: nhận diện từ tên file Excel)")
    parser.add_argument("--from", dest="date_from", help="thời gian xử lý từ ngày dd/mm/yyyy")
    parser.add_argument("--to", dest="date_to", help="thời gian xử lý đến ngày dd/mm/yyyy")
    parser.add_argument("--all", dest="all_periods", action="store_true", help="không lọc theo thời gian")
    parser.add_argument("--append", action="store_true", help="ghi mỗi báo cáo thành dòng mới (mặc định: tìm dòng theo Management Number)")
    parser.add_argument("--inspect", metavar="PATH", help="in cấu trúc file .pptx hoặc form .xlsx để kiểm tra mapping")
    parser.add_argument("--out", metavar="FILE", help="ghi kết quả --inspect ra file UTF-8 (mặc định inspect_template.txt / inspect_report.txt)")
    parser.add_argument("--version", action="store_true")
    args = parser.parse_args(argv)
    force_utf8_stdio()

    if args.version:
        from . import APP_TITLE, VERSION_LINE
        print(f"{APP_TITLE} ({VERSION_LINE})")
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
    if args.ollama_test:
        from .config import AppConfig, normalize_ollama_url
        from .diagnostics import format_diagnostics, ollama_smoke_rows
        cfg = AppConfig.load()
        server = normalize_ollama_url(args.server or cfg.ollama_server)
        model = args.model if args.model is not None else cfg.model
        if not model:
            print("Chưa chọn model (dùng --model qwen3:4b).")
            return 2
        print(format_diagnostics(ollama_smoke_rows(server, model, timeout=int(cfg.request_timeout or 180))))
        return 0
    if args.diag:
        from .config import AppConfig
        from .diagnostics import format_diagnostics, run_diagnostics
        print(format_diagnostics(run_diagnostics(AppConfig.load(), check_ollama=not args.no_ollama_check, smoke=True)))
        return 0
    if args.inputs:
        if not args.template or not args.output:
            parser.error("--cli cần --template và --output")
        return _cli(args)

    return launch_gui()


STARTUP_ERROR_LOG = "startup_error.log"
STARTUP_ERROR_VI = "Không thể khởi động Report Extractor.\n\nChi tiết đã được ghi tại:\nlogs\\" + STARTUP_ERROR_LOG


def _startup_logging() -> None:
    """Portable dirs + STARTUP line (version/prompt/packaged/executable/portable_root, no secrets)."""
    from . import runtime_paths
    from .logger import LOG, setup_logging
    runtime_paths.ensure_portable_dirs()
    try:
        setup_logging(runtime_paths.logs_dir())
    except OSError as e:                      # read-only location: keep going, the GUI still works
        LOG.warning("Không ghi được log khởi động: %s", e)
    LOG.info("STARTUP_MODE gui %s", runtime_paths.describe())


def record_startup_error(exc: BaseException, log_dir: Optional[Path] = None) -> Optional[Path]:
    """Write the traceback of a fatal pre-GUI failure to ``logs/startup_error.log`` (never silent)."""
    import traceback
    from datetime import datetime
    try:
        if log_dir is None:
            from . import runtime_paths
            log_dir = runtime_paths.logs_dir()
        log_dir = Path(log_dir)
        log_dir.mkdir(parents=True, exist_ok=True)
        from . import VERSION_LINE
        target = log_dir / STARTUP_ERROR_LOG
        with target.open("a", encoding="utf-8") as fh:
            fh.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} {VERSION_LINE} python={sys.version.split()[0]} "
                     f"executable={sys.executable}\n")
            fh.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
        return target
    except Exception:  # noqa: BLE001
        return None


def _show_startup_error_dialog() -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Report Extractor", STARTUP_ERROR_VI)
        root.destroy()
    except Exception:  # noqa: BLE001
        print(STARTUP_ERROR_VI, file=sys.stderr)


def launch_gui() -> int:
    try:
        _startup_logging()
        from .gui import launch
    except ImportError as e:
        record_startup_error(e)
        print(f"Không khởi động được giao diện (tkinter): {e}", file=sys.stderr)
        _show_startup_error_dialog()
        return 3
    except Exception as e:  # noqa: BLE001
        record_startup_error(e)
        _show_startup_error_dialog()
        return 4
    try:
        launch()
    except Exception as e:  # noqa: BLE001 – failure before/while creating the main window
        from .logger import LOG
        LOG.exception("GUI_FATAL %s", e)
        record_startup_error(e)
        _show_startup_error_dialog()
        return 4
    return 0


if __name__ == "__main__":
    sys.exit(main())
