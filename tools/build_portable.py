"""Windows Portable release builder (PROMPT-003) – driven by build_portable.bat.

Steps (fail fast, clear Vietnamese message on error):
  1. Python version check (same tuples as setup.bat / app.diagnostics – never weakened for the build)
  2. build virtualenv .venv-build (or reuse) with requirements-build.txt (which now also installs
     requirements-webview.txt – pywebview + pythonnet, i.e. the WebView2 runtime the exe launches)
  3. pyflakes + pytest (can be skipped with --skip-tests, never skipped by default)
  4. build the React frontend (`cd frontend && npm run build`) – the exe serves frontend/dist (§2/§3)
  5. clean build/ and dist/<name>/
  6. PyInstaller ReportExtractor.spec (onedir, windowed, no UPX; bundles frontend/dist + webview)
  7. assemble portable folder: Output/, logs/, config/, frontend/dist (documented Portable layout,
     byte-identical to the copy inside _internal), README.txt, FIRST_RUN.txt, VERSION.txt,
     Install_Ollama_Optional.bat
  8. validate artifact (no tests/.venv/.git/sample data/dev config inside; exe present; React bundle and
     the pywebview JS bridge / WebView2 interop assemblies present)
  9. ZIP + SHA256SUMS.txt → release/
 10. publish to the LAN update folder (tools/publish_update.py) – ONLY after every step above succeeded;
     ZIP first (verified), version.json LAST, obsolete ReportExtractor_*.zip removed afterwards (--no-publish to skip)
The script never bundles Ollama or any model and never writes a developer config into the artifact.
"""
from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import importlib.util
import json
import os
import platform
import shutil
import struct
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
from console_safe import child_env, configure_console, safe_text  # noqa: E402
import dotnet_pe  # noqa: E402  (ECMA-335 reader: proves what a packaged .NET assembly IS – PROMPT-030)
import publish_update  # noqa: E402  (centralised update-folder configuration + safe publisher)

PUBLISH_FAILED_RC = 3          # build succeeded, publishing did not (distinct from build failures = 1)
SUPPORTED_PYTHON = ((3, 10), (3, 11), (3, 12), (3, 13))   # explicit tuples, identical to setup.bat policy
FORBIDDEN_IN_ARTIFACT = ("tests", ".venv", ".venv-build", ".git", ".pytest_cache", "sample_data", "__pycache__",
                         "update_staging", "update_backup", "release", "learning_data")
FORBIDDEN_SUFFIXES = (".pptx", ".ppt", ".xlsx", ".pyc")
# Narrow allowlist of dependency RUNTIME resources that legitimately carry a forbidden suffix.  Compared as
# lower-cased path-component tuples (platform independent).  python-pptx needs its default template to build a
# Presentation(); it is not project/user/sample data.  Nothing else under _internal is exempt.
ALLOWED_DEPENDENCY_RESOURCES = (
    ("_internal", "pptx", "templates", "default.pptx"),
)
DOCS_DIR = ROOT / "release_docs"


def _version_module():
    spec = importlib.util.spec_from_file_location("app_version", ROOT / "app" / "__init__.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore[union-attr]
    return mod


def _version() -> tuple[str, str]:
    """(version, zero-padded numeric build) from the single authoritative app/__init__.py."""
    mod = _version_module()
    return mod.__version__, mod.BUILD_ID


def build_number() -> int:
    return int(_version_module().BUILD_NUMBER)


def package_name(version: str) -> str:
    """Release ZIP referenced by version.json: ReportExtractor_<version>.zip"""
    return f"ReportExtractor_{version}.zip"


def write_version_manifest(release_dir: Path, version: str, build: int, zip_path: Path,
                           built: str | None = None) -> Path:
    """release/version.json for the offline updater – sha256 of the EXACT zip; written LAST."""
    data = {"version": version, "build": int(build), "package": zip_path.name, "sha256": sha256(zip_path),
            "size": zip_path.stat().st_size, "built": built or _dt.datetime.now().strftime("%Y-%m-%d %H:%M")}
    out = Path(release_dir) / "version.json"
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return out


def step(n: int, title: str) -> None:
    print(safe_text(f"\n[{n:02d}] {title}", stream=sys.stdout), flush=True)


def fail(msg: str, code: int = 1) -> None:
    print(safe_text(f"\nLỖI BUILD: {msg}", stream=sys.stderr), file=sys.stderr, flush=True)
    sys.exit(code)


def run(cmd: list[str], cwd: Path = ROOT) -> None:
    print(safe_text("   $ " + " ".join(str(c) for c in cmd), stream=sys.stdout), flush=True)
    r = subprocess.run([str(c) for c in cmd], cwd=str(cwd), env=child_env())
    if r.returncode != 0:
        fail(f"lệnh thất bại (mã {r.returncode}): {' '.join(str(c) for c in cmd)}", r.returncode or 1)


def check_python(py: Path) -> None:
    out = subprocess.run([str(py), "-c", "import sys;print(sys.version_info[0],sys.version_info[1],sys.maxsize>2**32)"],
                         capture_output=True, text=True, encoding="utf-8", errors="replace", env=child_env())
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
                           text=True, encoding="utf-8", errors="replace", timeout=timeout, check=True,
                           env=child_env())
    except (FileNotFoundError, subprocess.CalledProcessError, subprocess.TimeoutExpired, OSError) as e:
        print(f"[info] Git revision không khả dụng ({type(e).__name__}) – ghi '{GIT_UNAVAILABLE}' vào metadata")
        return GIT_UNAVAILABLE
    except Exception as e:  # noqa: BLE001 – metadata must never abort the build
        print(f"[info] Git revision không khả dụng ({type(e).__name__}: {e}) – ghi '{GIT_UNAVAILABLE}'")
        return GIT_UNAVAILABLE
    rev = (r.stdout or "").strip()
    return rev if rev else GIT_UNAVAILABLE


FRONTEND_DIR = ROOT / "frontend"
FRONTEND_DIST = FRONTEND_DIR / "dist"


def find_npm() -> str | None:
    """Absolute npm/npm.cmd of the BUILD machine (the Portable target never needs Node – §8)."""
    for candidate in (("npm.cmd", "npm.bat", "npm") if os.name == "nt" else ("npm",)):
        found = shutil.which(candidate)
        if found:
            return found
    return None


def build_frontend(skip: bool) -> None:
    """Produce `frontend/dist` with the SAME command a developer runs (§2).

    A stale or missing bundle is the single most common cause of a Portable exe that opens a blank window,
    so the build stops here rather than packaging whatever happens to be on disk.
    """
    npm = find_npm()
    if skip:
        if not (FRONTEND_DIST / "index.html").is_file():
            fail("--skip-frontend được dùng nhưng thiếu frontend/dist/index.html. "
                 "Chạy: cd frontend && npm run build")
        print("   (bỏ qua build frontend theo --skip-frontend) dùng frontend/dist sẵn có")
        return
    if npm is None:
        fail("không tìm thấy npm trên máy BUILD (máy đích không cần Node, nhưng máy build thì cần). "
             "Cài Node.js LTS rồi chạy lại, hoặc dùng --skip-frontend nếu frontend/dist đã build sẵn.")
    print(f"   npm: {npm}")
    if not (FRONTEND_DIR / "node_modules").is_dir():
        run([npm, "install", "--no-audit", "--no-fund"], cwd=FRONTEND_DIR)
    run([npm, "run", "build"], cwd=FRONTEND_DIR)
    if not (FRONTEND_DIST / "index.html").is_file():
        fail("npm run build không tạo ra frontend/dist/index.html")
    assets = FRONTEND_DIST / "assets"
    if not assets.is_dir() or not any(assets.glob("*.js")):
        fail("frontend/dist/assets thiếu file .js – bản build React không hợp lệ")
    print(f"   frontend/dist OK (index.html + {len(list(assets.glob('*')))} file trong assets/)")


def copy_frontend_into_portable(folder: Path) -> None:
    """Place the React bundle at the documented Portable layout: `<folder>/frontend/dist` (§14).

    `_internal/frontend/dist` is already inside the onedir package (ReportExtractor.spec `datas`);
    `app.desktop.frontend_dist_dir()` accepts either location.  Both copies come from the same build, and
    `validate_artifact` proves they are byte-identical so a tester can never edit a stale duplicate.
    """
    target = folder / "frontend" / "dist"
    if target.exists():
        shutil.rmtree(target)
    shutil.copytree(FRONTEND_DIST, target)
    print(f"   frontend/dist → {target.relative_to(folder).as_posix()}")


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
        f"Report Extractor\nversion={version}\nbuild={build_id}\nbuilt={built}\ngit={git_rev}\n"
        f"Git revision: {git_rev}\n"
        f"python={platform.python_version()}\npackaging=PyInstaller onedir windowed (no UPX)\n"
        f"ollama_bundled=no\nmodel_bundled=no\ndefault_server=http://127.0.0.1:11434\ndefault_model=qwen3:4b\n",
        encoding="utf-8-sig")
    shutil.copyfile(DOCS_DIR / "Install_Ollama_Optional.bat", folder / "Install_Ollama_Optional.bat")
    return git_rev


def is_allowed_dependency_resource(rel: Path | str) -> bool:
    """True only for the exact allow-listed dependency runtime files (relative to the portable folder)."""
    parts = tuple(x.lower() for x in Path(str(rel).replace("\\", "/")).parts)
    return parts in ALLOWED_DEPENDENCY_RESOURCES


#: Files that MUST exist inside the package for the React UI and the JS↔Python bridge to work (§3/§17/§18).
#: Checked relative to the portable folder; ``_internal`` is the PyInstaller onedir resource root.
REQUIRED_FRONTEND_FILES = (
    "_internal/frontend/dist/index.html",      # bundled by ReportExtractor.spec `datas`
    "frontend/dist/index.html",                # documented Portable layout, next to _internal (§14)
)
#: pywebview data files.  Without webview/js/* the window opens but window.pywebview.api never appears, so
#: the bridge silently dies; without the WebView2 assemblies the EdgeChromium backend cannot start (§6/§18).
REQUIRED_WEBVIEW_FILES = (
    "_internal/webview/js/api.js",
    "_internal/webview/js/lib/dom_json.js",
    "_internal/webview/js/state.js",
)
REQUIRED_WEBVIEW_WINDOWS_FILES = (
    "_internal/webview/lib/Microsoft.Web.WebView2.Core.dll",
    "_internal/webview/lib/Microsoft.Web.WebView2.WinForms.dll",
    "_internal/webview/lib/runtimes/win-x64/native/WebView2Loader.dll",
)


# --------------------------------------------------------------------------- .NET runtime (PROMPT-028R)
# A packaged pywebview app dies at start-up with
#     RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize
#                   from _internal\pythonnet\runtime\Python.Runtime.dll
# when pythonnet's managed assembly is not exactly where clr_loader expects, or when a second conflicting
# copy exists.  Both are packaging mistakes, so the BUILD verifies them instead of the tester discovering
# them by double-clicking the exe.
PYTHONNET_RUNTIME_DLL = "Python.Runtime.dll"
PYTHONNET_RUNTIME_RELPATH = Path("_internal") / "pythonnet" / "runtime" / PYTHONNET_RUNTIME_DLL
#: where clr_loader.ffi.load_netfx() dlopen()s its mixed-mode C++/CLI host from, inside the package.
CLR_LOADER_RELPATH = Path("_internal") / "clr_loader" / "ffi" / "dlls"
#: same key app/desktop.py reads to report the installed .NET Framework; kept literal here so the build
#: gate stays importable without the app package.
DOTNET_FX_REG_KEY = r"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
#: pywebview's Windows backend (winforms -> edgechromium) needs these to be reachable at start-up.
REQUIRED_BUILD_PACKAGES = ("pywebview", "pythonnet", "clr-loader", "pyinstaller", "pyinstaller-hooks-contrib")


def _query(py: Path, code: str) -> dict:
    """Run ``code`` in the BUILD interpreter and parse the single JSON object it prints."""
    r = subprocess.run([str(py), "-c", code], cwd=str(ROOT), capture_output=True, text=True,
                       encoding="utf-8", errors="replace", env=child_env())
    if r.returncode != 0:
        fail(f"không truy vấn được môi trường build: {r.stderr.strip()[:400]}")
    try:
        return json.loads((r.stdout or "").strip().splitlines()[-1])
    except (ValueError, IndexError):
        fail(f"môi trường build trả về kết quả không phải JSON: {r.stdout.strip()[:400]}")
        return {}                                       # unreachable – fail() exits


def report_build_dependencies(py: Path) -> dict:
    """Print the exact versions the package will be built with, and fail on a missing Windows essential.

    A reused ``.venv-build`` with stale or half-installed packages is one of the suspected causes of the
    PROMPT-028R failure, so the versions are recorded in the build log before anything is compiled.
    """
    code = (
        "import importlib.metadata as md, importlib.util, json\n"
        f"names = {REQUIRED_BUILD_PACKAGES!r}\n"
        "out = {}\n"
        "for n in names:\n"
        "    try:\n"
        "        out[n] = md.version(n)\n"
        "    except Exception:\n"
        "        out[n] = None\n"
        "out['_python'] = __import__('sys').version.split()[0]\n"
        "print(json.dumps(out))\n"
    )
    info = _query(py, code)
    for name in REQUIRED_BUILD_PACKAGES:
        print(f"   {name}: {info.get(name) or 'KHÔNG CÓ'}")
    print(f"   python (build venv): {info.get('_python')}")
    if platform.system() == "Windows":
        missing = [n for n in ("pywebview", "pythonnet", "clr-loader") if not info.get(n)]
        if missing:
            fail("thiếu gói bắt buộc để đóng gói UI React trên Windows: " + ", ".join(missing)
                 + ". Chạy lại với --fresh-venv, hoặc: pip install -r requirements-webview.txt")
    return info


#: Modules that MUST import cleanly in the build venv before anything is frozen.  A frozen app can only
#: contain what the build venv could import, so this is the cheapest possible reproduction of the
#: PROMPT-028R start-up failure — it runs BEFORE PyInstaller instead of after a double-click on Windows.
RUNTIME_IMPORT_PROBES = ("webview", "pythonnet", "clr", "clr_loader")


def verify_runtime_imports(py: Path) -> dict:
    """Import the pywebview/.NET stack inside the build venv and fail the build if any import breaks.

    Also checks that the installed ``clr-loader`` still satisfies pythonnet's own declared requirement
    (``clr_loader<0.3.0,>=0.2.7`` for pythonnet 3.0.5).  A reused ``.venv-build`` holding an out-of-range
    clr-loader produces exactly the runtime symptom PROMPT-028R had to explain, and nothing else in the
    pipeline would notice.
    """
    code = (
        "import importlib, importlib.metadata as md, json\n"
        f"names = {RUNTIME_IMPORT_PROBES!r}\n"
        "out = {'imports': {}, 'requires': {}}\n"
        "for n in names:\n"
        "    try:\n"
        "        importlib.import_module(n)\n"
        "        out['imports'][n] = ''\n"
        "    except BaseException as e:\n"
        "        out['imports'][n] = f'{type(e).__name__}: {e}'[:300]\n"
        "try:\n"
        "    reqs = md.requires('pythonnet') or []\n"
        "    out['requires'] = {'clr_loader': [r for r in reqs if r.lower().startswith('clr_loader')] or [],\n"
        "                       'clr_loader_installed': md.version('clr-loader')}\n"
        "except Exception as e:\n"
        "    out['requires'] = {'error': f'{type(e).__name__}: {e}'[:200]}\n"
        "print(json.dumps(out))\n"
    )
    info = _query(py, code)
    imports = info.get("imports") or {}
    for name in RUNTIME_IMPORT_PROBES:
        err = imports.get(name)
        print(f"   import {name}: " + ("OK" if err is not None and not err else f"LỖI – {err}"))
    broken = [n for n in RUNTIME_IMPORT_PROBES if imports.get(n)]
    if broken and platform.system() == "Windows":
        detail = "; ".join(f"{n}: {imports[n]}" for n in broken)
        fail("môi trường build không import được UI React/.NET: " + detail
             + ". Chạy lại với --fresh-venv, hoặc: pip install -r requirements-webview.txt")
    req = info.get("requires") or {}
    installed = req.get("clr_loader_installed")
    declared = [r for r in (req.get("clr_loader") or []) if "extra" not in r]
    if installed and declared:
        try:
            from packaging.requirements import Requirement
            ok = all(installed in Requirement(d).specifier for d in declared)
        except Exception:                                  # noqa: BLE001 – an unverifiable constraint is a warning
            ok, declared = True, declared
        print(f"   clr-loader {installed} vs pythonnet yêu cầu {declared or '?'}: "
              + ("OK" if ok else "KHÔNG THOẢ"))
        if not ok and platform.system() == "Windows":
            fail(f"clr-loader {installed} không thoả yêu cầu của pythonnet ({', '.join(declared)}). "
                 f"Chạy lại với --fresh-venv.")
    return info


def source_pythonnet_dll(py: Path) -> dict:
    """Path / size / SHA256 of Python.Runtime.dll in the BUILD environment (the intended source copy)."""
    code = (
        "import hashlib, importlib.util, json, pathlib\n"
        "spec = importlib.util.find_spec('pythonnet')\n"
        "if not spec or not spec.origin:\n"
        "    print(json.dumps({'found': False}))\n"
        "else:\n"
        "    p = pathlib.Path(spec.origin).parent / 'runtime' / 'Python.Runtime.dll'\n"
        "    ok = p.is_file()\n"
        "    print(json.dumps({'found': ok, 'path': str(p),\n"
        "                      'size': p.stat().st_size if ok else 0,\n"
        "                      'sha256': hashlib.sha256(p.read_bytes()).hexdigest() if ok else ''}))\n"
    )
    return _query(py, code)


#: The five native symbols ``clr_loader/ffi/netfx.py`` declares in its cdef.  ``ClrLoader.dll`` is the only
#: thing that can supply them, and a package shipping a host without one of them fails at
#: ``pyclr_get_function`` with exactly the PROMPT-030 error.
PYCLR_EXPORTS = ("pyclr_initialize", "pyclr_create_appdomain", "pyclr_get_function",
                 "pyclr_close_appdomain", "pyclr_finalize")
#: .NET Framework ``Release`` (HKLM\\SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full) that first
#: ships the netstandard facade a ``.NETStandard,Version=v2.0`` assembly needs.  461808 == 4.7.2.
NETFX_RELEASE_BY_TARGET = {"v4.7": 460798, "v4.7.1": 461308, "v4.7.2": 461808,
                           "v4.8": 528040, "v4.8.1": 533320}
#: ``Python.Runtime.dll`` targets ``.NETStandard,Version=v2.0`` (measured, not assumed) — see module comment.
NETSTANDARD_MIN_NETFX_RELEASE = {"v2.0": 461808}


def mark_of_the_web(path: Path) -> str | None:
    """The ``Zone.Identifier`` alternate data stream of ``path``, or None when it is not blocked.

    Windows attaches that stream when a ZIP is extracted or a file arrives from another machine, and the
    .NET Framework then REFUSES ``Assembly.LoadFrom`` on it (``COR_E_FILELOAD``, HRESULT 0x80131515,
    "Operation is not supported") unless the hosting AppDomain enables ``loadFromRemoteSources`` — while
    ``LoadLibrary`` on the NATIVE ``ClrLoader.dll`` ignores the stream completely.  That asymmetry is the
    one mechanism which lets clr_loader create its app domain and still fail to resolve
    ``Python.Runtime.Loader.Initialize`` from a DLL that exists at the right path with the right SHA256,
    i.e. from a package that passes every presence/hash check.  Reading the stream is harmless everywhere:
    on a filesystem without ADS support the open simply fails.
    """
    if not path.is_file():
        return None
    try:
        with open(f"{path}:Zone.Identifier", "r", encoding="utf-8", errors="replace") as handle:
            return handle.read(400).strip() or "present"
    except OSError:
        return None


def required_netfx_release(image) -> tuple[int | None, str]:
    """Minimum .NET Framework ``Release`` that can host ``image``, derived from what it actually targets."""
    for target in image.target_frameworks:
        lowered = target.lower()
        if "netstandard" in lowered:
            version = lowered.split("version=")[-1].strip()
            need = NETSTANDARD_MIN_NETFX_RELEASE.get(version)
            if need is None:
                return None, f"{target} (no .NET Framework release can host it)"
            return need, f"{target} → .NET Framework Release ≥ {need} (4.7.2+)"
        if "netframework" in lowered:
            version = "v" + lowered.split("version=")[-1].strip().lstrip("v")
            return NETFX_RELEASE_BY_TARGET.get(version), f"{target}"
    return None, ", ".join(image.target_frameworks) or "<no TargetFrameworkAttribute>"


def prove_runtime_identity(dll: Path, problems: list[str]) -> None:
    """Prove the packaged payload IS pythonnet's managed runtime and declares the entry point clr_loader binds.

    Presence and SHA256 cannot answer this: they say a file with the right name and the right bytes is
    there, not that the bytes are a .NET assembly declaring ``Python.Runtime.Loader.Initialize`` with the
    signature ``clr_loader/ffi/netfx.py``'s ``entry_point`` typedef requires.  Everything here is read out
    of the PE itself, so it is decided identically on Windows and on Linux.
    """
    try:
        image = dotnet_pe.inspect(dll)
    except dotnet_pe.DotNetPEError as exc:
        problems.append(f"{dll.name} không phải là một .NET assembly đọc được: {exc}")
        return
    print("   " + dotnet_pe.describe(image).replace("\n", "\n   "))
    if image.assembly is None or image.assembly.name != "Python.Runtime":
        problems.append(f"{dll.name} có danh tính assembly '{image.assembly}' chứ không phải "
                        "'Python.Runtime' — file đúng tên/đúng vị trí nhưng không phải runtime của pythonnet")
    ok, why = dotnet_pe.signature_matches_entry_point(image)
    print(f"   entry point    {why}")
    if not ok:
        problems.append(f"{dll.name} không phân giải được entry point: {why}")
    if not image.il_only or image.requires_32bit:
        problems.append(f"{dll.name} không phải IL-only/any-cpu (machine={image.machine}, "
                        f"IL-only={image.il_only}, 32-bit-required={image.requires_32bit}) — kiến trúc "
                        "của nó sẽ xung đột với tiến trình ReportExtractor.exe")
    need, because = required_netfx_release(image)
    if need is None and "no .NET Framework" in because:
        problems.append(f"{dll.name} nhắm tới {because}: .NET Framework (netfx) không bao giờ nạp được nó")
    elif need is not None:
        release = dotnet_framework_release_of_build_machine()
        print(f"   .NET cần       {because}"
              + (f"; máy này Release={release}" if release is not None else ""))
        if release is not None and release < need:
            problems.append(f"{dll.name} cần .NET Framework Release ≥ {need} nhưng máy này có {release} "
                            "(4.7.2+): Assembly.LoadFrom sẽ ném và clr_loader/netfx.py chỉ báo "
                            "'Failed to resolve Python.Runtime.Loader.Initialize'")


def prove_clr_loader(folder: Path, arch: str, problems: list[str]) -> Path | None:
    """Prove the packaged netfx host is the right architecture and exports what clr_loader dlopen()s."""
    host = folder / CLR_LOADER_RELPATH / arch / "ClrLoader.dll"
    if not host.is_file():
        problems.append(f"thiếu clr_loader/ffi/dlls/{arch}/ClrLoader.dll (hook-clr_loader) – "
                        "clr_loader.netfx không nạp được .NET Framework")
        return None
    print(f"   host netfx     {host.relative_to(folder).as_posix()} size={host.stat().st_size} "
          f"sha256={sha256(host)}")
    try:
        image = dotnet_pe.inspect(host)
    except dotnet_pe.DotNetPEError as exc:
        problems.append(f"ClrLoader.dll ({arch}) không đọc được: {exc}")
        return host
    print("   " + dotnet_pe.describe(image).replace("\n", "\n   "))
    missing = [name for name in PYCLR_EXPORTS if name not in image.exports]
    if missing:
        problems.append(f"ClrLoader.dll ({arch}) không export {', '.join(missing)} — cdef của "
                        "clr_loader/ffi/netfx.py sẽ thất bại ngay khi dlopen")
    if image.il_only:
        problems.append(f"ClrLoader.dll ({arch}) là IL-only: nó phải là mixed-mode C++/CLI để cffi "
                        "dlopen() được các hàm pyclr_*")
    if image.machine != arch:
        problems.append(f"ClrLoader.dll trong thư mục {arch}/ thực chất là {image.machine} — tiến trình "
                        f"{arch} sẽ nhận BadImageFormatException")
    return host


#: Runs in the BUILD venv but against the PACKAGED bytes: the exact call chain that failed on Windows
#: (``ffi.dlopen(ClrLoader.dll)`` → ``pyclr_initialize`` → ``pyclr_create_appdomain`` →
#: ``pyclr_get_function(Python.Runtime.dll, "Python.Runtime.Loader", "Initialize")``), so the build proves
#: LOADABILITY instead of only presence/hash.  ``clr_loader`` never checks whether the app-domain handle
#: came back NULL and its netfx cdef has no error accessor, so both are read directly here.
PROBE_PACKAGED_RUNTIME = r"""
import json, sys
from pathlib import Path
from cffi import FFI

folder, arch = Path(sys.argv[1]), sys.argv[2]
out = {"arch": arch}
ffi = FFI()
ffi.cdef('''
typedef void* pyclr_domain;
typedef int (*entry_point)(void* buffer, int size);
void pyclr_initialize();
void* pyclr_create_appdomain(const char* name, const char* config_file);
entry_point pyclr_get_function(pyclr_domain domain, const char* assembly_path,
                               const char* class_name, const char* function);
void pyclr_close_appdomain(pyclr_domain domain);
void pyclr_finalize();
''')
try:
    host = folder / "_internal" / "clr_loader" / "ffi" / "dlls" / arch / "ClrLoader.dll"
    assembly = folder / "_internal" / "pythonnet" / "runtime" / "Python.Runtime.dll"
    out["host"] = str(host)
    out["assembly"] = str(assembly)
    fw = ffi.dlopen(str(host))
    fw.pyclr_initialize()
    domain = fw.pyclr_create_appdomain(ffi.NULL, ffi.NULL)
    out["appdomain"] = "NULL" if domain == ffi.NULL else "created"
    func = fw.pyclr_get_function(domain, str(assembly).encode("utf8"),
                                 b"Python.Runtime.Loader", b"Initialize")
    out["resolved"] = func != ffi.NULL
    if out["resolved"]:
        buffer = ffi.from_buffer("char[]", b"")
        out["initialize_rc"] = int(func(ffi.cast("void*", buffer), 0))
except BaseException as exc:
    out["error"] = f"{type(exc).__name__}: {exc}"[:400]
print(json.dumps(out))
"""


def probe_packaged_runtime(folder: Path, py: Path, arch: str) -> list[str]:
    """Actually resolve ``Python.Runtime.Loader.Initialize`` out of the finished package (Windows only).

    Returns the problems it found.  A failure here means the shipped exe WILL die at start-up, so the build
    stops instead of publishing it.  Where the probe cannot run (non-Windows, no build interpreter) it says
    so explicitly rather than reporting a pass it never performed.
    """
    if platform.system() != "Windows":
        print("   probe loadability: BỎ QUA — .NET Framework chỉ có trên Windows; phần chứng minh "
              "danh tính/kiến trúc/entry point ở trên vẫn chạy trên mọi nền tảng")
        return []
    try:
        proc = subprocess.run([str(py), "-c", PROBE_PACKAGED_RUNTIME, str(folder), arch],
                              cwd=str(folder), capture_output=True, text=True, timeout=240,
                              encoding="utf-8", errors="replace", env=child_env())
    except (OSError, subprocess.SubprocessError) as exc:
        return [f"không chạy được probe nạp runtime đã đóng gói: {type(exc).__name__}: {exc}"]
    try:
        out = json.loads((proc.stdout or "").strip().splitlines()[-1])
    except (ValueError, IndexError):
        return [f"probe nạp runtime đã đóng gói không trả về JSON (rc={proc.returncode}): "
                f"{(proc.stdout or '').strip()[-300:]} {(proc.stderr or '').strip()[-300:]}".strip()]
    for key in ("arch", "appdomain", "resolved", "initialize_rc", "error"):
        if key in out:
            print(f"   probe {key:14} {out[key]}")
    problems: list[str] = []
    if out.get("error"):
        problems.append(f"probe nạp runtime đã đóng gói thất bại: {out['error']}")
    if out.get("appdomain") == "NULL":
        problems.append("pyclr_create_appdomain trả về NULL: clr_loader không tạo được AppDomain .NET "
                        "Framework (clr_loader/netfx.py KHÔNG kiểm tra điều này và vẫn báo "
                        "initialized=True). Nguyên nhân thường gặp: .NET Framework < 4.7.2 hoặc "
                        "ClrLoader.dll sai kiến trúc")
    if out.get("resolved") is False and not out.get("error"):
        problems.append("pyclr_get_function trả về NULL cho Python.Runtime.Loader.Initialize trong gói "
                        "đã build — chính là lỗi làm ReportExtractor.exe chết khi khởi động. Kiểm tra "
                        "Zone.Identifier (file bị Windows chặn) và .NET Framework Release")
    if isinstance(out.get("initialize_rc"), int) and out["initialize_rc"] != 0:
        problems.append(f"Python.Runtime.Loader.Initialize trả về {out['initialize_rc']} (khác 0): "
                        "pythonnet không khởi động được trong gói đã build")
    return problems


def dotnet_framework_release_of_build_machine() -> int | None:
    """``Release`` of the .NET Framework 4.x on THIS machine, or None when it cannot be read."""
    if platform.system() != "Windows":
        return None
    try:
        import winreg                                    # noqa: PLC0415 – Windows only
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, DOTNET_FX_REG_KEY) as key:
            return int(winreg.QueryValueEx(key, "Release")[0])
    except Exception:                                    # noqa: BLE001 – an unreadable key is not a build error
        return None


def validate_pythonnet_runtime(folder: Path, source: dict | None = None,
                               py: Path | None = None) -> list[str]:
    """Prove the packaged .NET runtime is present, unique, unmodified, loadable and not blocked.

    PROMPT-028R proved the first three.  PROMPT-030 adds what the shipped exe actually needed: the payload
    must BE pythonnet's managed assembly declaring the entry point clr_loader binds (``prove_runtime_identity``),
    the netfx host must have the right architecture and exports (``prove_clr_loader``), neither file may be
    blocked by a Mark-of-the-Web stream (``mark_of_the_web``), and on Windows the real
    ``pyclr_get_function`` call must succeed against the packaged bytes (``probe_packaged_runtime``).
    """
    problems: list[str] = []
    arch = "amd64" if struct.calcsize("P") * 8 > 32 else "x86"
    found = sorted(folder.rglob(PYTHONNET_RUNTIME_DLL))
    if source and source.get("found"):
        print(f"   nguồn {PYTHONNET_RUNTIME_DLL}: {source.get('path')}")
        print(f"        size={source.get('size')} sha256={source.get('sha256')}")
    if not found:
        problems.append(f"thiếu {PYTHONNET_RUNTIME_DLL} trong gói – pywebview sẽ không khởi động được "
                        "(hook-clr của PyInstaller phải đóng gói nó dạng BINARY)")
        return problems
    if len(found) > 1:
        problems.append(f"có {len(found)} bản {PYTHONNET_RUNTIME_DLL} trong gói (phải đúng 1): "
                        + ", ".join(p.relative_to(folder).as_posix() for p in found))
    expected = (folder / PYTHONNET_RUNTIME_RELPATH).resolve()
    for dll in found:
        rel = dll.relative_to(folder)
        packaged_sha = sha256(dll)
        print(f"   gói  {rel.as_posix()}: size={dll.stat().st_size} sha256={packaged_sha}")
        if dll.resolve() != expected:
            problems.append(f"{PYTHONNET_RUNTIME_DLL} sai vị trí: {rel.as_posix()} (pythonnet tìm ở "
                            f"{PYTHONNET_RUNTIME_RELPATH.as_posix()})")
        if source and source.get("found") and packaged_sha != source.get("sha256"):
            problems.append(f"{rel.as_posix()} có SHA256 KHÁC bản pythonnet trong môi trường build "
                            f"(gói={packaged_sha[:16]}… nguồn={str(source.get('sha256'))[:16]}…) "
                            "- có thể .venv-build cũ hoặc hook đã sửa file")
        if dll.resolve() == expected:
            # Cross-platform: the proof below reads the PE, it does not need a Windows CLR.
            prove_runtime_identity(dll, problems)
        blocked = mark_of_the_web(dll)
        if blocked:
            problems.append(f"{rel.as_posix()} đang BỊ WINDOWS CHẶN (Zone.Identifier: {blocked[:120]}). "
                            "LoadLibrary bỏ qua cờ này nhưng Assembly.LoadFrom của .NET Framework thì "
                            "KHÔNG, nên clr_loader tạo được AppDomain mà vẫn không phân giải được "
                            "Python.Runtime.Loader.Initialize. Gỡ bằng: chuột phải → Properties → "
                            "Unblock, hoặc PowerShell: Unblock-File -Path '<đường dẫn>'")
    host = prove_clr_loader(folder, arch, problems)
    if host is None and platform.system() != "Windows":
        # hook-clr_loader chỉ thu thập ClrLoader.dll khi is_win/is_cygwin, nên ngoài Windows sự vắng mặt
        # này là đúng chứ không phải lỗi đóng gói.
        problems = [p for p in problems if "ClrLoader.dll" not in p]
    elif host is not None:
        blocked = mark_of_the_web(host)
        if blocked:
            problems.append(f"ClrLoader.dll ({arch}) đang BỊ WINDOWS CHẶN (Zone.Identifier) — gỡ bằng "
                            "Unblock-File cho cả thư mục Portable trước khi chạy")
    problems += probe_packaged_runtime(folder, py, arch) if py is not None else []
    return problems


def _same_tree(left: Path, right: Path) -> bool:
    """True when two directories hold byte-identical files under identical relative paths."""
    if not left.is_dir() or not right.is_dir():
        return False
    rel_left = sorted(p.relative_to(left) for p in left.rglob("*") if p.is_file())
    rel_right = sorted(p.relative_to(right) for p in right.rglob("*") if p.is_file())
    if rel_left != rel_right:
        return False
    return all((left / rel).read_bytes() == (right / rel).read_bytes() for rel in rel_left)


def validate_frontend(folder: Path) -> list[str]:
    """The packaged React UI + pywebview runtime, verified instead of assumed (§16/§17/§18)."""
    problems: list[str] = []
    for rel in REQUIRED_FRONTEND_FILES:
        if not (folder / Path(rel)).is_file():
            problems.append(f"thiếu UI React trong gói: {rel}")
    bundled, portable = folder / "_internal" / "frontend" / "dist", folder / "frontend" / "dist"
    if bundled.is_dir() and portable.is_dir() and not _same_tree(bundled, portable):
        problems.append("frontend/dist trong _internal và ngoài portable root KHÁC nhau "
                        "(hai bản copy phải giống hệt từng byte)")
    assets = portable / "assets"
    if not assets.is_dir():
        problems.append("thiếu frontend/dist/assets (React bundle không đầy đủ)")
    else:
        if not any(assets.glob("*.js")):
            problems.append("frontend/dist/assets không có file .js nào")
        if not any(assets.glob("*.css")):
            problems.append("frontend/dist/assets không có file .css nào")
    index = portable / "index.html"
    if index.is_file():
        html = index.read_text(encoding="utf-8", errors="replace")
        # Vite emits root-relative asset URLs; pywebview serves the dist folder over loopback HTTP, so
        # these must resolve.  A source-tree or repository path here would mean a broken package (§16).
        if "/assets/" not in html:
            problems.append("frontend/dist/index.html không tham chiếu /assets/ – bundle React sai")
        for leak in ("frontend/src", "node_modules", str(ROOT)):
            if leak in html:
                problems.append(f"frontend/dist/index.html còn tham chiếu đường dẫn máy dev: {leak}")
    for rel in REQUIRED_WEBVIEW_FILES:
        if not (folder / Path(rel)).is_file():
            problems.append(f"thiếu pywebview JS bridge trong gói: {rel}")
    if platform.system() == "Windows":
        for rel in REQUIRED_WEBVIEW_WINDOWS_FILES:
            if not (folder / Path(rel)).is_file():
                problems.append(f"thiếu WebView2 interop assembly trong gói: {rel}")
    return problems


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
        if p.is_file() and p.suffix.lower() in FORBIDDEN_SUFFIXES and not is_allowed_dependency_resource(rel):
            problems.append(f"dữ liệu mẫu/dev trong gói: {rel}")
        if p.is_file() and (p.name.lower() == "ollama.exe" or p.suffix.lower() == ".gguf"):
            problems.append(f"gói không được chứa Ollama/model: {rel}")
    cfg = folder / "config" / "config.json"
    if cfg.exists():
        problems.append("config/config.json không được có sẵn trong gói (cấu hình của máy dev)")
    for cand in folder.glob("config.json"):
        problems.append(f"{cand.name} của máy dev nằm trong gói")
    return problems


def validate_existing_package(folder: Path, py: Path) -> int:
    """Run the PROMPT-030 runtime gate against an ALREADY BUILT Portable folder and report the real cause.

    The build gate proves a package PyInstaller just produced.  The crash PROMPT-030 fixes was reported
    from ``H:\\ReportExtractor_v1.3.4_Portable`` — a copy that may have travelled through a ZIP, a USB
    stick or a network share AFTER the build, which is exactly how a Windows Mark-of-the-Web block appears
    on ``Python.Runtime.dll`` while every presence/hash check still passes.  So the same proof has to be
    runnable against the folder that actually fails, without rebuilding anything.
    """
    print(f"[validate-only] thư mục: {folder}")
    if not folder.is_dir():
        fail(f"không thấy thư mục Portable: {folder}")
    if not py.exists() or py == Path(sys.executable):
        candidate = ROOT / ".venv-build" / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
        py = candidate if candidate.exists() else py
    print(f"[validate-only] interpreter so sánh: {py}")
    source = source_pythonnet_dll(py) if py.exists() else None
    problems = validate_pythonnet_runtime(folder, source, py if py.exists() else None)
    if problems:
        print("\n[validate-only] KẾT LUẬN: gói này KHÔNG thể khởi động, vì:")
        for problem in problems:
            print("   - " + problem)
        return 1
    print("\n[validate-only] OK: Python.Runtime.Loader.Initialize phân giải được từ gói này.")
    print("   Nếu ReportExtractor.exe vẫn chết thì nguyên nhân nằm NGOÀI gói (quyền thư mục, antivirus,")
    print("   WebView2 Runtime) — gửi lại logs/app.log và logs/startup_error.log.")
    return 0


def main() -> int:
    configure_console()
    ap = argparse.ArgumentParser(description="Build Report Extractor Windows Portable")
    ap.add_argument("--python", default=sys.executable, help="Python interpreter used to create the build venv")
    ap.add_argument("--skip-tests", action="store_true", help="không chạy pytest (chỉ dùng khi đã chạy riêng)")
    ap.add_argument("--no-venv", action="store_true", help="dùng interpreter hiện tại thay vì .venv-build")
    ap.add_argument("--allow-non-windows", action="store_true", help="cho phép chạy trên Linux/macOS (chỉ thử nghiệm)")
    ap.add_argument("--fresh-venv", action="store_true",
                    help="xoá và tạo lại .venv-build (tránh gói cũ/hỏng trong môi trường build dùng lại)")
    ap.add_argument("--skip-frontend", action="store_true",
                    help="không chạy `npm run build`, dùng frontend/dist đã có sẵn")
    ap.add_argument("--no-publish", action="store_true", help="không tự động xuất bản vào thư mục cập nhật LAN")
    ap.add_argument("--validate-only", metavar="FOLDER", default=None,
                    help="KHÔNG build: chạy cổng chứng minh runtime .NET (PROMPT-030) trên một thư mục "
                         "Portable đã có — ví dụ H:\\ReportExtractor_v1.3.4_Portable — rồi thoát. Đây là "
                         "cách xác định vì sao ReportExtractor.exe chết khi khởi động mà không cần build lại.")
    ap.add_argument("--publish-dir", default=None,
                    help=f"thư mục cập nhật cục bộ (mặc định {publish_update.update_folder()})")
    args = ap.parse_args()

    if args.validate_only:
        return validate_existing_package(Path(args.validate_only), Path(args.python))

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
        if args.fresh_venv and venv.exists():
            print("   --fresh-venv: xoá .venv-build cũ để build lại từ đầu")
            shutil.rmtree(venv)
        if not py.exists():
            run([base_py, "-m", "venv", str(venv)])
        run([py, "-m", "pip", "install", "--upgrade", "pip", "--quiet"])
        run([py, "-m", "pip", "install", "-r", str(ROOT / "requirements-build.txt"), "pytest", "pyflakes", "--quiet"])

    report_build_dependencies(py)
    verify_runtime_imports(py)

    step(4, "Kiểm tra mã nguồn (pyflakes) và test (pytest)")
    flakes = subprocess.run([str(py), "-m", "pyflakes", "app", "tools", "run.py"], cwd=str(ROOT),
                            capture_output=True, text=True, encoding="utf-8", errors="replace", env=child_env())
    # "imported but unused" lines are deliberate availability probes (win32com, fitz, tkinter …) – everything else fails
    real = [ln for ln in (flakes.stdout + flakes.stderr).splitlines() if ln.strip() and "imported but unused" not in ln]
    if real:
        fail("pyflakes:\n   " + "\n   ".join(real))
    print("   pyflakes OK")
    if args.skip_tests:
        print("   (bỏ qua pytest theo yêu cầu --skip-tests)")
    else:
        run([py, "-m", "pytest", "-q"])

    step(5, "Build frontend React (npm run build) – UI của ReportExtractor.exe")
    build_frontend(args.skip_frontend)

    step(6, "Dọn build/ và dist/")
    for d in (ROOT / "build", ROOT / "dist" / name):
        if d.exists():
            shutil.rmtree(d)

    step(7, "PyInstaller (onedir, windowed, no UPX, bundle frontend/dist + pywebview/WebView2)")
    run([py, "-m", "PyInstaller", "--noconfirm", "--clean", str(ROOT / "ReportExtractor.spec")])
    folder = ROOT / "dist" / name
    if not folder.exists():
        fail(f"PyInstaller không tạo {folder}")

    step(8, "Tạo thư mục Output/, logs/, config/ và frontend/dist (layout Portable)")
    for sub in ("Output", "logs", "config"):
        (folder / sub).mkdir(exist_ok=True)
        (folder / sub / ".keep").write_text("", encoding="utf-8")
    copy_frontend_into_portable(folder)

    step(9, "Tài liệu README.txt / FIRST_RUN.txt / VERSION.txt / Install_Ollama_Optional.bat")
    git_rev = write_release_metadata(folder, version, build_id, name)
    print(f"  Git revision: {git_rev}")

    step(10, "Kiểm tra gói (không chứa file dev, không config máy dev, không Ollama/model)")
    source_dll = source_pythonnet_dll(py)
    problems = (validate_artifact(folder) + validate_frontend(folder)
                + validate_pythonnet_runtime(folder, source_dll, py))
    if problems:
        fail("gói không hợp lệ:\n   - " + "\n   - ".join(problems))
    print("   OK")

    step(11, "Nén ZIP, SHA256SUMS.txt và version.json (gói cập nhật offline)")
    release = ROOT / "release"
    release.mkdir(exist_ok=True)
    zip_path = release / package_name(version)
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for p in sorted(folder.rglob("*")):
            # arcname is POSIX on purpose: str(Path(...)) would emit "\" on Windows (PROMPT-029 §5)
            zf.write(p, f"{name}/{p.relative_to(folder).as_posix()}")
    exe = folder / "ReportExtractor.exe"
    lines = []
    if exe.exists():
        lines.append(f"{sha256(exe)} *{name}/ReportExtractor.exe")
    lines.append(f"{sha256(zip_path)} *{zip_path.name}")
    sums = "\n".join(lines) + "\n"
    (folder / "SHA256SUMS.txt").write_text(sums, encoding="utf-8")
    (release / "SHA256SUMS.txt").write_text(sums, encoding="utf-8")
    print("   " + sums.replace("\n", "\n   ").rstrip())
    manifest = write_version_manifest(release, version, build_number(), zip_path)
    print(f"   version.json: {manifest.read_text(encoding='utf-8').strip()}")

    step(12, "Hoàn tất build")
    print(f"   Thư mục portable: {folder}\n   ZIP: {zip_path}\n   SHA256SUMS: {release / 'SHA256SUMS.txt'}\n"
          f"   version.json: {manifest}")

    if args.no_publish:
        print(f"   (bỏ qua xuất bản theo --no-publish) Phát hành thủ công: copy {zip_path.name} vào thư mục Update "
              f"TRƯỚC, sau đó copy version.json SAU CÙNG.")
        return 0
    step(13, "Xuất bản vào thư mục cập nhật LAN (ZIP trước – version.json SAU CÙNG)")
    return publish_step(release, args.publish_dir)


def publish_step(release: Path, publish_dir: str | None = None) -> int:
    """Runs only after a fully successful build; a publish failure never masquerades as a build failure."""
    res = publish_update.publish_release(release, publish_dir)
    if res.ok:
        return 0
    print(f"BUILD THÀNH CÔNG nhưng PUBLISH FAILED: {res.error}\n   Gói build vẫn nằm trong {release}; "
          f"có thể chạy lại: python tools\\publish_update.py")
    return PUBLISH_FAILED_RC


if __name__ == "__main__":
    sys.exit(main())
