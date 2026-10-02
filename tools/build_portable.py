"""Windows Portable release builder (PROMPT-003) – driven by build_portable.bat.

Steps (fail fast, clear Vietnamese message on error):
  1. Python version check (same tuples as setup.bat / app.diagnostics – never weakened for the build)
  2. build virtualenv .venv-build (or reuse) with requirements-build.txt
  3. pyflakes + pytest (can be skipped with --skip-tests, never skipped by default)
  4. clean build/ and dist/<name>/
  5. PyInstaller ReportExtractor.spec (onedir, windowed, no UPX)
  6. assemble portable folder: Output/, logs/, config/, README.txt, FIRST_RUN.txt, VERSION.txt,
     Install_Ollama_Optional.bat
  7. validate artifact (no tests/.venv/.git/sample data/dev config inside; exe present)
  8. ZIP + SHA256SUMS.txt → release/
The script never bundles Ollama or any model and never writes a developer config into the artifact.
"""
from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import importlib.util
import os
import platform
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SUPPORTED_PYTHON = ((3, 10), (3, 11), (3, 12), (3, 13))   # explicit tuples, identical to setup.bat policy
FORBIDDEN_IN_ARTIFACT = ("tests", ".venv", ".venv-build", ".git", ".pytest_cache", "sample_data", "__pycache__")
FORBIDDEN_SUFFIXES = (".pptx", ".ppt", ".xlsx", ".pyc")
DOCS_DIR = ROOT / "release_docs"


def _version() -> tuple[str, str]:
    spec = importlib.util.spec_from_file_location("app_version", ROOT / "app" / "__init__.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[union-attr]
    return mod.__version__, mod.BUILD_ID


def step(n: int, title: str) -> None:
    print(f"\n[{n:02d}] {title}", flush=True)


def fail(msg: str, code: int = 1) -> None:
    print(f"\nLỖI BUILD: {msg}", file=sys.stderr, flush=True)
    sys.exit(code)


def run(cmd: list[str], cwd: Path = ROOT) -> None:
    print("   $", " ".join(str(c) for c in cmd), flush=True)
    r = subprocess.run([str(c) for c in cmd], cwd=str(cwd))
    if r.returncode != 0:
        fail(f"lệnh thất bại (mã {r.returncode}): {' '.join(str(c) for c in cmd)}", r.returncode or 1)


def check_python(py: Path) -> None:
    out = subprocess.run([str(py), "-c", "import sys;print(sys.version_info[0],sys.version_info[1],sys.maxsize>2**32)"],
                         capture_output=True, text=True)
    if out.returncode != 0:
        fail(f"không chạy được {py}")
    major, minor, is64 = out.stdout.split()
    if (int(major), int(minor)) not in SUPPORTED_PYTHON:
        fail(f"Python {major}.{minor} không được hỗ trợ để build. Hỗ trợ: "
             + ", ".join(f"{a}.{b}" for a, b in SUPPORTED_PYTHON))
    if is64 != "True":
        fail("cần Python 64-bit để đóng gói.")
    print(f"   Python {major}.{minor} 64-bit OK ({py})")


GIT_UNAVAILABLE = "unavailable"


def get_git_revision(root: Path = ROOT, timeout: float = 10.0) -> str:
    """Short git revision of ``root`` – informational metadata ONLY, never a build prerequisite.

    Returns ``GIT_UNAVAILABLE`` ("unavailable") deterministically when git is not installed
    (FileNotFoundError), the folder is not a repository / the command fails (non-zero exit), the call times
    out, or any other OS-level error happens.  Never raises.
    """
    try:
        r = subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=str(root), capture_output=True,
                           text=True, timeout=timeout, check=True)
    except (FileNotFoundError, subprocess.CalledProcessError, subprocess.TimeoutExpired, OSError) as e:
        print(f"[info] Git revision không khả dụng ({type(e).__name__}) – ghi '{GIT_UNAVAILABLE}' vào metadata")
        return GIT_UNAVAILABLE
    except Exception as e:  # noqa: BLE001 – metadata must never abort the build
        print(f"[info] Git revision không khả dụng ({type(e).__name__}: {e}) – ghi '{GIT_UNAVAILABLE}'")
        return GIT_UNAVAILABLE
    rev = (r.stdout or "").strip()
    return rev if rev else GIT_UNAVAILABLE


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def render_doc(doc_name: str, **fields: str) -> str:
    """Fill ``{placeholders}`` of release_docs/<doc_name>.  First positional parameter is deliberately NOT called
    ``name`` – ``name`` is one of the template fields (portable folder name)."""
    text = (DOCS_DIR / doc_name).read_text(encoding="utf-8")
    for k, v in fields.items():
        text = text.replace("{" + k + "}", v)
    return text


def write_release_metadata(folder: Path, version: str, build_id: str, name: str,
                           built: str | None = None, git_rev: str | None = None) -> str:
    """Write README.txt / FIRST_RUN.txt / VERSION.txt / Install_Ollama_Optional.bat into ``folder``.

    ``git_rev`` defaults to ``get_git_revision()``; when git is unavailable the metadata is still complete and valid
    (``git=unavailable`` / ``Git revision: unavailable``).  Returns the revision string that was written.
    """
    built = built or _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    git_rev = git_rev if git_rev else get_git_revision()
    fields = dict(version=version, build_id=build_id, built=built, git=git_rev, name=name,
                  python=platform.python_version())
    (folder / "README.txt").write_text(render_doc("README.txt", **fields), encoding="utf-8-sig")
    (folder / "FIRST_RUN.txt").write_text(render_doc("FIRST_RUN.txt", **fields), encoding="utf-8-sig")
    (folder / "VERSION.txt").write_text(
        f"Report Extractor\nversion={version}\nprompt={build_id}\nbuilt={built}\ngit={git_rev}\n"
        f"Git revision: {git_rev}\n"
        f"python={platform.python_version()}\npackaging=PyInstaller onedir windowed (no UPX)\n"
        f"ollama_bundled=no\nmodel_bundled=no\ndefault_server=http://127.0.0.1:11434\ndefault_model=qwen3:4b\n",
        encoding="utf-8-sig")
    shutil.copyfile(DOCS_DIR / "Install_Ollama_Optional.bat", folder / "Install_Ollama_Optional.bat")
    return git_rev


def validate_artifact(folder: Path, exe_name: str = "ReportExtractor.exe") -> list[str]:
    problems: list[str] = []
    if not (folder / exe_name).exists() and not (folder / "ReportExtractor").exists():
        problems.append(f"thiếu {exe_name}")
    if not (folder / "_internal").is_dir():
        problems.append("thiếu thư mục _internal (không phải build onedir?)")
    for req in ("Output", "logs", "config"):
        if not (folder / req).is_dir():
            problems.append(f"thiếu thư mục {req}")
    for req in ("README.txt", "FIRST_RUN.txt", "VERSION.txt", "Install_Ollama_Optional.bat"):
        if not (folder / req).is_file():
            problems.append(f"thiếu {req}")
    for p in folder.rglob("*"):
        rel = p.relative_to(folder)
        parts = set(rel.parts)
        if parts & set(FORBIDDEN_IN_ARTIFACT):
            problems.append(f"file phát triển trong gói: {rel}")
            continue
        if p.is_file() and p.suffix.lower() in FORBIDDEN_SUFFIXES:
            problems.append(f"dữ liệu mẫu/dev trong gói: {rel}")
        if p.is_file() and (p.name.lower() == "ollama.exe" or p.suffix.lower() == ".gguf"):
            problems.append(f"gói không được chứa Ollama/model: {rel}")
    cfg = folder / "config" / "config.json"
    if cfg.exists():
        problems.append("config/config.json không được có sẵn trong gói (cấu hình của máy dev)")
    for cand in folder.glob("config.json"):
        problems.append(f"{cand.name} của máy dev nằm trong gói")
    return problems


def main() -> int:
    ap = argparse.ArgumentParser(description="Build Report Extractor Windows Portable")
    ap.add_argument("--python", default=sys.executable, help="Python interpreter used to create the build venv")
    ap.add_argument("--skip-tests", action="store_true", help="không chạy pytest (chỉ dùng khi đã chạy riêng)")
    ap.add_argument("--no-venv", action="store_true", help="dùng interpreter hiện tại thay vì .venv-build")
    ap.add_argument("--allow-non-windows", action="store_true", help="cho phép chạy trên Linux/macOS (chỉ thử nghiệm)")
    args = ap.parse_args()

    version, build_id = _version()
    name = f"ReportExtractor_v{version}_Portable"
    step(1, f"Report Extractor {version} ({build_id}) – Windows Portable")
    if platform.system() != "Windows" and not args.allow_non_windows:
        fail("gói Windows Portable phải được build trên Windows (PyInstaller không cross-compile).")

    step(2, "Kiểm tra Python")
    base_py = Path(args.python)
    check_python(base_py)

    step(3, "Môi trường build")
    if args.no_venv:
        py = base_py
    else:
        venv = ROOT / ".venv-build"
        py = venv / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
        if not py.exists():
            run([base_py, "-m", "venv", str(venv)])
        run([py, "-m", "pip", "install", "--upgrade", "pip", "--quiet"])
        run([py, "-m", "pip", "install", "-r", str(ROOT / "requirements-build.txt"), "pytest", "pyflakes", "--quiet"])

    step(4, "Kiểm tra mã nguồn (pyflakes) và test (pytest)")
    flakes = subprocess.run([str(py), "-m", "pyflakes", "app", "tools", "run.py"], cwd=str(ROOT),
                            capture_output=True, text=True)
    # "imported but unused" lines are deliberate availability probes (win32com, fitz, tkinter …) – everything else fails
    real = [ln for ln in (flakes.stdout + flakes.stderr).splitlines() if ln.strip() and "imported but unused" not in ln]
    if real:
        fail("pyflakes:\n   " + "\n   ".join(real))
    print("   pyflakes OK")
    if args.skip_tests:
        print("   (bỏ qua pytest theo yêu cầu --skip-tests)")
    else:
        run([py, "-m", "pytest", "-q"])

    step(5, "Dọn build/ và dist/")
    for d in (ROOT / "build", ROOT / "dist" / name):
        if d.exists():
            shutil.rmtree(d)

    step(6, "PyInstaller (onedir, windowed, no UPX)")
    run([py, "-m", "PyInstaller", "--noconfirm", "--clean", str(ROOT / "ReportExtractor.spec")])
    folder = ROOT / "dist" / name
    if not folder.exists():
        fail(f"PyInstaller không tạo {folder}")

    step(7, "Tạo thư mục Output/, logs/, config/")
    for sub in ("Output", "logs", "config"):
        (folder / sub).mkdir(exist_ok=True)
        (folder / sub / ".keep").write_text("", encoding="utf-8")

    step(8, "Tài liệu README.txt / FIRST_RUN.txt / VERSION.txt / Install_Ollama_Optional.bat")
    git_rev = write_release_metadata(folder, version, build_id, name)
    print(f"  Git revision: {git_rev}")

    step(9, "Kiểm tra gói (không chứa file dev, không config máy dev, không Ollama/model)")
    problems = validate_artifact(folder)
    if problems:
        fail("gói không hợp lệ:\n   - " + "\n   - ".join(problems))
    print("   OK")

    step(10, "Nén ZIP và SHA256SUMS.txt")
    release = ROOT / "release"
    release.mkdir(exist_ok=True)
    zip_path = release / f"{name}.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for p in sorted(folder.rglob("*")):
            zf.write(p, str(Path(name) / p.relative_to(folder)))
    exe = folder / "ReportExtractor.exe"
    lines = []
    if exe.exists():
        lines.append(f"{sha256(exe)} *{name}/ReportExtractor.exe")
    lines.append(f"{sha256(zip_path)} *{zip_path.name}")
    sums = "\n".join(lines) + "\n"
    (folder / "SHA256SUMS.txt").write_text(sums, encoding="utf-8")
    (release / "SHA256SUMS.txt").write_text(sums, encoding="utf-8")
    print("   " + sums.replace("\n", "\n   ").rstrip())

    step(11, "Hoàn tất")
    print(f"   Thư mục portable: {folder}\n   ZIP: {zip_path}\n   SHA256SUMS: {release / 'SHA256SUMS.txt'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
