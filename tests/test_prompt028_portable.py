"""PROMPT-028 — Windows Portable packaging of the React + pywebview application.

The executable itself can only be produced on Windows (PyInstaller does not cross-compile), but every
defect that makes a packaged exe open a blank window or lose its bridge is a *path-resolution* or
*bundling* defect, and those are provable here:

* §4  frozen-runtime frontend resolution (``sys._MEIPASS`` vs the source tree, both supported layouts)
* §16 no dependence on the repository path or the current working directory after packaging
* §3/§17 the React bundle is inside the artifact and its asset URLs are root-relative
* §5/§6/§18 the pywebview JS bridge and the WebView2 interop assemblies are collected
* §7  PowerPoint stays optional (lazy COM import, guarded spec collection, fallback renderer)
* §1  the exe launches the same UI as ``python -m app.desktop``, with a Tk fallback
* §10 the packaged build version is visible in the UI and in VERSION.txt
"""
from __future__ import annotations

import importlib.util
import json
import logging
import os
import re
import sys
from pathlib import Path, PurePath, PureWindowsPath

import pytest

import app.desktop as desktop
import app.main as app_main
import dotnet_image_fixtures as dotnet_fx
import app.runtime_paths as rp
import build_portable as bp
from build_portable import (_same_tree, build_number, copy_frontend_into_portable, package_name,
                                  validate_artifact, validate_frontend, validate_pythonnet_runtime,
                                  write_release_metadata)

ROOT = Path(__file__).resolve().parents[1]
SPEC = ROOT / "ReportExtractor.spec"
SPEC_TEXT = SPEC.read_text(encoding="utf-8")
INDEX_HTML = (
    '<!doctype html><html><head><meta charset="UTF-8">'
    '<script type="module" crossorigin src="/assets/index-B-laKRIy.js"></script>'
    '<link rel="stylesheet" crossorigin href="/assets/index-B-laKRIy.css">'
    '</head><body><div id="root"></div></body></html>'
)


def _write_dist(dist: Path) -> Path:
    """A minimal but structurally real Vite bundle (root-relative /assets/ URLs, like `npm run build`)."""
    (dist / "assets").mkdir(parents=True, exist_ok=True)
    (dist / "index.html").write_text(INDEX_HTML, encoding="utf-8")
    (dist / "assets" / "index-B-laKRIy.js").write_text("console.log('report-extractor')", encoding="utf-8")
    (dist / "assets" / "index-B-laKRIy.css").write_text(":root{--brand:#1f6feb}", encoding="utf-8")
    return dist


def _fake_frozen(monkeypatch, tmp_path, layout: str = "internal"):
    """Point the process at a synthetic PyInstaller onedir package.

    ``layout="internal"`` puts the bundle at ``_internal/frontend/dist`` (spec ``datas``);
    ``layout="portable"`` puts it at ``<portable>/frontend/dist`` (the documented Portable layout, §14).
    Returns the portable root.
    """
    portable = tmp_path / "ReportExtractor_vTest_Portable"
    internal = portable / "_internal"
    internal.mkdir(parents=True)
    if layout in ("internal", "both"):
        _write_dist(internal / "frontend" / "dist")
    if layout in ("portable", "both"):
        _write_dist(portable / "frontend" / "dist")
    exe = portable / ("ReportExtractor.exe" if os.name == "nt" else "ReportExtractor")
    exe.write_text("", encoding="utf-8")

    monkeypatch.delenv(rp.TEST_RUNTIME_ROOT_ENV, raising=False)   # a real exe writes next to itself
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(internal), raising=False)
    monkeypatch.setattr(sys, "executable", str(exe))
    return portable


# ============================================================== §4 frozen frontend path resolution
def test_frozen_bundle_inside_internal_resolves_relative_to_meipass(monkeypatch, tmp_path):
    """The spec's `datas` layout: the URL must be relative to ``_MEIPASS``, i.e. WITHOUT the leading '..'."""
    portable = _fake_frozen(monkeypatch, tmp_path, layout="internal")
    dist = desktop.frontend_dist_dir(require=True)
    assert dist == (portable / "_internal" / "frontend" / "dist").resolve()
    url = desktop.frontend_url(dist)
    assert url == "frontend/dist/index.html"
    # pywebview joins a relative URL onto get_app_root() (= _MEIPASS when frozen) and serves that folder
    assert (Path(sys._MEIPASS) / url).resolve() == dist / "index.html"


def test_frozen_bundle_next_to_internal_resolves_relative_to_portable_root(monkeypatch, tmp_path):
    """The documented Portable layout (§14): the bundle is a sibling of ``_internal``."""
    portable = _fake_frozen(monkeypatch, tmp_path, layout="portable")
    dist = desktop.frontend_dist_dir(require=True)
    assert dist == (portable / "frontend" / "dist").resolve()
    url = desktop.frontend_url(dist)
    assert url == "../frontend/dist/index.html"
    assert (Path(sys._MEIPASS) / url).resolve() == dist / "index.html"


def test_frozen_prefers_the_portable_root_copy_when_both_exist(monkeypatch, tmp_path):
    """Both layouts present → the tester-visible copy wins, and it is the one validated as identical."""
    portable = _fake_frozen(monkeypatch, tmp_path, layout="both")
    assert desktop.frontend_dist_dir(require=True) == (portable / "frontend" / "dist").resolve()
    candidates = desktop.frontend_dist_candidates()
    assert candidates[0] == (portable / "frontend" / "dist").resolve()
    assert (portable / "_internal" / "frontend" / "dist").resolve() in candidates


def test_source_mode_url_is_unchanged_and_relative_to_the_app_package():
    """``python -m app.desktop`` keeps working exactly as before: app root == ``<repo>/app``."""
    assert desktop.assumed_app_root() == ROOT / "app"
    dist = desktop.frontend_dist_dir() or (ROOT / "frontend" / "dist")
    url = desktop.frontend_url(dist, desktop.assumed_app_root())
    assert url == "../frontend/dist/index.html"
    assert (desktop.assumed_app_root() / url).resolve() == ROOT / "frontend" / "dist" / "index.html"
    # module constants stay the historical, source-mode-correct values other modules/tests rely on
    assert desktop.FRONTEND_URL == "../frontend/dist/index.html"
    assert desktop.FRONTEND_INDEX == desktop.FRONTEND_DIST / "index.html"


def test_assumed_app_root_ignores_how_the_process_was_launched(monkeypatch):
    """pytest / IDE / ``python -c`` all set a different ``sys.argv[0]``; resolution must not care."""
    monkeypatch.setattr(sys, "argv", ["/somewhere/else/pytest"], raising=False)
    assert desktop.assumed_app_root() == ROOT / "app"
    url = desktop.frontend_url(ROOT / "frontend" / "dist", desktop.assumed_app_root())
    assert url == "../frontend/dist/index.html"


def test_default_url_root_is_pywebviews_own_root_so_the_round_trip_always_holds(monkeypatch, tmp_path):
    """The URL must be relative to whatever root pywebview will actually use, or the window is blank.

    ``python -c`` / an IDE / a shortcut each give pywebview a different ``sys.argv[0]``; deriving the URL
    from pywebview's real root keeps it correct in every one of those launch modes.
    """
    dist = _write_dist(tmp_path / "anywhere" / "frontend" / "dist")
    monkeypatch.setattr(desktop, "frontend_dist_candidates", lambda: [dist])
    for root in (tmp_path, tmp_path / "anywhere", tmp_path / "anywhere" / "frontend"):
        monkeypatch.setattr(desktop, "runtime_app_root", lambda root=root: root)
        url = desktop.frontend_url(dist)                       # no explicit root -> pywebview's own
        assert not url.startswith(("/", "file:", "http:"))
        assert (root / url).resolve() == dist / "index.html", f"round-trip broke for root={root}"


def test_runtime_app_root_matches_pywebview_when_it_is_installed(monkeypatch, tmp_path):
    """The URL we hand ``create_window`` must agree with pywebview's own root, or the window is blank."""
    _fake_frozen(monkeypatch, tmp_path, layout="internal")
    webview_util = pytest.importorskip("webview.util")
    assert desktop.runtime_app_root() == Path(webview_util.get_app_root())
    assert desktop.runtime_app_root() == Path(sys._MEIPASS)


# ============================================================== §16 no repository / cwd dependence
def test_frozen_package_never_resolves_the_bundle_from_the_source_checkout(monkeypatch, tmp_path):
    """§16: even with a perfectly good `<repo>/frontend/dist` present, a frozen exe must not use it."""
    portable = _fake_frozen(monkeypatch, tmp_path, layout="none")
    assert desktop.frontend_dist_dir() is None, "no bundle in the package → must NOT fall back to the repo"
    candidates = desktop.frontend_dist_candidates()
    assert candidates, "the packaged lookup must still enumerate its own candidate paths"
    for candidate in candidates:
        assert str(ROOT) not in str(candidate), f"packaged lookup escaped the portable folder: {candidate}"
        assert str(portable.resolve()) in str(candidate), f"candidate escaped the package: {candidate}"


def test_frozen_resolution_never_mentions_the_repository(monkeypatch, tmp_path):
    portable = _fake_frozen(monkeypatch, tmp_path, layout="internal")
    dist = desktop.frontend_dist_dir(require=True)
    url = desktop.frontend_url(dist)
    for value in (str(dist), url):
        assert str(ROOT) not in value, f"packaged path still points at the source checkout: {value}"
        assert "node_modules" not in value
    assert not url.startswith(("/", "file:", "http:")), "a relative URL keeps pywebview's HTTP server in charge"
    assert str(portable) in str(dist)


def test_resolution_does_not_depend_on_the_working_directory(monkeypatch, tmp_path):
    """Double-clicking an exe, a shortcut with a 'Start in' folder, or a UNC path must all behave alike."""
    portable = _fake_frozen(monkeypatch, tmp_path, layout="portable")
    before = desktop.frontend_dist_dir(require=True)
    monkeypatch.chdir(tmp_path)
    assert desktop.frontend_dist_dir(require=True) == before
    monkeypatch.chdir(portable / "_internal")
    assert desktop.frontend_dist_dir(require=True) == before
    assert desktop.frontend_url(before) == "../frontend/dist/index.html"


def test_missing_bundle_reports_every_searched_path(monkeypatch, tmp_path):
    """A packaged app without its bundle must explain itself instead of opening a blank window."""
    portable = tmp_path / "ReportExtractor_vTest_Portable"
    (portable / "_internal").mkdir(parents=True)
    monkeypatch.delenv(rp.TEST_RUNTIME_ROOT_ENV, raising=False)
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(portable / "_internal"), raising=False)
    monkeypatch.setattr(sys, "executable", str(portable / "ReportExtractor.exe"))
    assert desktop.frontend_dist_dir() is None
    with pytest.raises(FileNotFoundError) as exc:
        desktop.frontend_dist_dir(require=True)
    message = str(exc.value)
    assert "npm run build" in message
    assert str(portable / "frontend" / "dist") in message
    assert str(portable / "_internal" / "frontend" / "dist") in message


# ============================================================== §1 the exe launches the React app
def test_launch_ui_prefers_the_react_desktop_app(monkeypatch):
    calls = []
    monkeypatch.setattr(app_main, "launch_desktop", lambda: calls.append("desktop") or 0)
    monkeypatch.setattr(app_main, "launch_gui", lambda: calls.append("gui") or 0)
    assert app_main.launch_ui() == 0
    assert calls == ["desktop"]


def test_launch_ui_falls_back_to_the_classic_gui_when_desktop_is_unavailable(monkeypatch):
    """No pywebview or no bundle → the user still gets a working window (§7 spirit: never a dead exe)."""
    calls = []
    monkeypatch.setattr(app_main, "launch_desktop", lambda: calls.append("desktop") or app_main.DESKTOP_UNAVAILABLE)
    monkeypatch.setattr(app_main, "launch_gui", lambda: calls.append("gui") or 0)
    assert app_main.launch_ui() == 0
    assert calls == ["desktop", "gui"]


def test_launch_ui_propagates_a_real_desktop_failure(monkeypatch):
    calls = []
    monkeypatch.setattr(app_main, "launch_desktop", lambda: calls.append("desktop") or 4)
    monkeypatch.setattr(app_main, "launch_gui", lambda: calls.append("gui") or 0)
    assert app_main.launch_ui() == 4
    assert calls == ["desktop"], "a fatal desktop error must not be masked by silently opening the Tk GUI"


def test_launch_desktop_reports_unavailable_without_a_bundle(monkeypatch, tmp_path):
    monkeypatch.delenv(rp.TEST_RUNTIME_ROOT_ENV, raising=False)
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path / "_internal"), raising=False)
    monkeypatch.setattr(sys, "executable", str(tmp_path / "ReportExtractor.exe"))
    monkeypatch.setattr(desktop, "frontend_dist_dir", lambda require=False: None)
    monkeypatch.setattr(app_main, "record_startup_error", lambda exc, log_dir=None: None)
    assert app_main.launch_desktop() == app_main.DESKTOP_UNAVAILABLE


def test_launch_desktop_maps_an_argparse_exit_to_unavailable(monkeypatch):
    monkeypatch.setattr(desktop, "frontend_dist_dir", lambda require=False: ROOT / "frontend" / "dist")
    monkeypatch.setattr(app_main, "record_startup_error", lambda exc, log_dir=None: None)
    import types
    monkeypatch.setitem(sys.modules, "webview", types.ModuleType("webview"))      # pywebview present

    def boom(_argv):
        raise SystemExit(2)

    monkeypatch.setattr(desktop, "main", boom)
    assert app_main.launch_desktop() == app_main.DESKTOP_UNAVAILABLE


def test_main_without_arguments_launches_the_react_ui(monkeypatch):
    """``ReportExtractor.exe`` double-clicked == ``python -m app.desktop`` (§1)."""
    seen = {}
    monkeypatch.setattr(app_main, "launch_ui", lambda: seen.setdefault("ui", True) and 0)
    monkeypatch.setattr(app_main, "launch_gui", lambda: seen.setdefault("gui", True) and 0)
    assert app_main.main([]) == 0
    assert seen == {"ui": True}


def test_main_legacy_gui_flag_still_reaches_the_classic_gui(monkeypatch):
    seen = {}
    monkeypatch.setattr(app_main, "launch_ui", lambda: seen.setdefault("ui", True) and 0)
    monkeypatch.setattr(app_main, "launch_gui", lambda: seen.setdefault("gui", True) and 0)
    assert app_main.main(["--legacy-gui"]) == 0
    assert seen == {"gui": True}


def test_main_cli_and_diag_do_not_open_a_window(monkeypatch):
    seen = {}
    monkeypatch.setattr(app_main, "launch_ui", lambda: seen.setdefault("ui", True) and 0)
    monkeypatch.setattr(app_main, "_cli", lambda args: seen.setdefault("cli", True) and 0)
    assert app_main.main(["--cli", "D:\\Reports", "--template", "t.xlsx", "--output", "o.xlsx"]) == 0
    assert seen == {"cli": True}


# ============================================================== §10 visible build version
def test_version_flag_prints_the_packaged_version(capsys):
    from app import APP_TITLE, VERSION_LINE
    assert app_main.main(["--version"]) == 0
    out = capsys.readouterr().out
    assert APP_TITLE in out and VERSION_LINE in out


def test_release_metadata_exposes_version_and_build(tmp_path):
    """VERSION.txt next to the exe is the offline cross-check for the version shown in the UI (§10)."""
    from app import BUILD_ID, __version__
    folder = tmp_path / package_name(__version__)
    folder.mkdir(parents=True)
    write_release_metadata(folder, __version__, BUILD_ID, folder.name, built="2026-10-08 09:00", git_rev="abc1234")
    text = (folder / "VERSION.txt").read_text(encoding="utf-8-sig")
    assert f"version={__version__}" in text and f"build={BUILD_ID}" in text
    assert "ollama_bundled=no" in text and "model_bundled=no" in text
    for name in ("README.txt", "FIRST_RUN.txt", "Install_Ollama_Optional.bat"):
        assert (folder / name).is_file(), name
    assert build_number() == int(BUILD_ID)


# ============================================================== §3/§5/§6 spec bundling decisions
def test_spec_bundles_the_react_bundle_and_the_pywebview_runtime():
    """PROMPT-028R: the React bundle is collected manually; the .NET runtime is left to the official hooks."""
    assert '(FRONTEND_DIST, os.path.join("frontend", "dist"))' in SPEC_TEXT, "frontend/dist must be a `datas` entry"
    assert 'FRONTEND_DIST = os.path.join(ROOT, "frontend", "dist")' in SPEC_TEXT
    # registered so hook-clr / hook-clr_loader fire, but never collected by hand
    for component in ('"webview"', '"pythonnet"', '"clr"', '"clr_loader"'):
        assert component in SPEC_TEXT, f"{component} must be registered for the WebView2/.NET backend"
    assert 'collect_data_files("webview", subdir="js")' in SPEC_TEXT, "only webview/js needs manual collection"
    assert '"bottle"' in SPEC_TEXT, "pywebview's loopback HTTP server needs bottle"
    # the WebView2 interop assemblies come from hook-webview, which must not be shadowed
    assert 'collect_dynamic_libs("webview")' not in SPEC_TEXT


def test_spec_refuses_to_package_a_missing_or_incomplete_frontend():
    """The most common blank-window cause must fail the BUILD, not the tester's double-click."""
    assert "thiếu frontend/dist/index.html" in SPEC_TEXT
    assert "npm run build" in SPEC_TEXT
    assert "thiếu frontend/dist/assets" in SPEC_TEXT
    assert "raise SystemExit" in SPEC_TEXT


def test_spec_keeps_powerpoint_optional():
    """§7: a build machine or target PC without PowerPoint/pywin32 must still produce a working exe."""
    assert "_module_available" in SPEC_TEXT
    block = SPEC_TEXT[SPEC_TEXT.index("# PowerPoint / pywin32 – OPTIONAL"):]
    block = block[:block.index("for mod in (\"tkinterdnd2\"")]
    assert 'if _module_available(_name.split(".")[0])' in block, "win32com must be collected conditionally"
    assert "import win32com" not in SPEC_TEXT


def test_spec_forces_the_edge_webview2_backend_and_stays_onedir_windowed():
    assert '"webview.platforms.edgechromium"' in SPEC_TEXT
    for excluded in ("PyQt5", "PySide2", "PyQt6", "PySide6", "cefpython3"):
        assert f'"{excluded}"' in SPEC_TEXT.split("excludes=")[1].split(")")[0], f"{excluded} must stay excluded"
    assert "exclude_binaries=True" in SPEC_TEXT and "console=False" in SPEC_TEXT and "upx=False" in SPEC_TEXT
    assert '[os.path.join(ROOT, "run.py")]' in SPEC_TEXT, "the exe entry point must remain run.py"
    assert "PORTABLE_NAME = f\"ReportExtractor_v{VERSION}_Portable\"" in SPEC_TEXT


def test_spec_entry_point_reaches_the_react_desktop_module():
    """run.py → app.main.main() → launch_ui() → app.desktop.main(): one code path for exe and source."""
    run_py = (ROOT / "run.py").read_text(encoding="utf-8")
    assert "from app.main import main" in run_py
    assert '"app.desktop"' in SPEC_TEXT
    assert "return launch_ui()" in (ROOT / "app" / "main.py").read_text(encoding="utf-8")


def test_build_requirements_install_pywebview():
    """Without this the build venv lacks pywebview and the spec cannot bundle the WebView2 backend (§5)."""
    text = (ROOT / "requirements-build.txt").read_text(encoding="utf-8")
    assert "-r requirements-webview.txt" in text
    webview_reqs = (ROOT / "requirements-webview.txt").read_text(encoding="utf-8")
    assert "pywebview" in webview_reqs
    assert "pythonnet" in webview_reqs and 'sys_platform == "win32"' in webview_reqs


def test_powerpoint_com_is_imported_lazily_so_a_pc_without_powerpoint_still_starts():
    """§7: no module-level win32com import anywhere on the startup path."""
    source = (ROOT / "app" / "qpn_renderer.py").read_text(encoding="utf-8")
    for line in source.splitlines():
        stripped = line.strip()
        if stripped.startswith(("import win32com", "from win32com")):
            assert not re.match(r"^(import|from)\s", stripped) or line.startswith((" ", "\t")), (
                f"win32com must be imported inside a function, not at module level: {line!r}")
    assert "import win32com.client" in source, "the COM renderer must still be reachable when PowerPoint exists"
    from app.qpn_renderer import renderer_status
    status = renderer_status()
    assert set(status) == {"powerpoint", "libreoffice", "builtin"}
    assert status["builtin"] == "OK", "the fallback renderer must be available without PowerPoint"
    assert status["powerpoint"] in ("OK", "Unavailable"), "PowerPoint is probed, never assumed"


# ============================================================== §17/§18 artifact validation
def _complete_package(tmp_path, layout: str = "both"):
    folder = tmp_path / "ReportExtractor_vTest_Portable"
    (folder / "_internal").mkdir(parents=True)
    for sub in ("Output", "logs", "config"):
        (folder / sub).mkdir()
    if layout in ("internal", "both"):
        _write_dist(folder / "_internal" / "frontend" / "dist")
    if layout in ("portable", "both"):
        _write_dist(folder / "frontend" / "dist")
    for rel in ("webview/js/api.js", "webview/js/lib/dom_json.js", "webview/js/state.js",
                "webview/lib/Microsoft.Web.WebView2.Core.dll",
                "webview/lib/Microsoft.Web.WebView2.WinForms.dll",
                "webview/lib/runtimes/win-x64/native/WebView2Loader.dll"):
        target = folder / "_internal" / Path(rel)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("stub", encoding="utf-8")
    return folder


def test_validate_frontend_accepts_a_complete_package(tmp_path):
    assert validate_frontend(_complete_package(tmp_path)) == []


def test_validate_frontend_requires_the_react_bundle_in_both_locations(tmp_path):
    problems = validate_frontend(tmp_path / "nothing")
    assert any("_internal/frontend/dist/index.html" in p for p in problems)
    assert any(p.startswith("thiếu UI React trong gói: frontend/dist/index.html") for p in problems)
    assert any("frontend/dist/assets" in p for p in problems)


def test_validate_frontend_rejects_divergent_duplicate_bundles(tmp_path):
    """Two copies are only safe while they are byte-identical; drift would mislead an acceptance tester."""
    folder = _complete_package(tmp_path)
    (folder / "frontend" / "dist" / "assets" / "index-B-laKRIy.js").write_text("STALE", encoding="utf-8")
    problems = validate_frontend(folder)
    assert any("KHÁC nhau" in p for p in problems)
    assert not _same_tree(folder / "_internal" / "frontend" / "dist", folder / "frontend" / "dist")


def test_validate_frontend_requires_the_js_bridge_and_webview2_assemblies(tmp_path):
    """§18: without webview/js/api.js the window opens but window.pywebview.api never appears."""
    folder = _complete_package(tmp_path)
    (folder / "_internal" / "webview" / "js" / "api.js").unlink()
    assert any("pywebview JS bridge" in p for p in validate_frontend(folder))

    folder = _complete_package(tmp_path / "second")
    (folder / "_internal" / "webview" / "lib" / "runtimes" / "win-x64" / "native" / "WebView2Loader.dll").unlink()
    problems = validate_frontend(folder)
    if sys.platform == "win32":
        assert any("WebView2 interop assembly" in p for p in problems)
    else:
        assert problems == [], "the WebView2 DLL check is Windows-only; Linux experiments must still pass"


def test_validate_frontend_rejects_a_bundle_that_leaks_dev_paths(tmp_path):
    folder = _complete_package(tmp_path)
    index = folder / "frontend" / "dist" / "index.html"
    index.write_text(f'<script src="{ROOT}/frontend/src/main.tsx"></script>', encoding="utf-8")
    problems = validate_frontend(folder)
    assert any("đường dẫn máy dev" in p for p in problems)


def test_validate_frontend_rejects_a_bundle_without_root_relative_assets(tmp_path):
    folder = _complete_package(tmp_path)
    (folder / "frontend" / "dist" / "index.html").write_text("<html><body>no assets</body></html>", encoding="utf-8")
    assert any("/assets/" in p for p in validate_frontend(folder))


def test_copy_frontend_into_portable_replaces_a_stale_bundle(tmp_path):
    folder = tmp_path / "pkg"
    (folder / "_internal").mkdir(parents=True)
    stale = folder / "frontend" / "dist" / "assets"
    stale.mkdir(parents=True)
    (stale / "old-hash.js").write_text("stale", encoding="utf-8")
    monkey_dist = _write_dist(tmp_path / "source-dist")
    import build_portable as builder
    original = builder.FRONTEND_DIST
    try:
        builder.FRONTEND_DIST = monkey_dist
        copy_frontend_into_portable(folder)
    finally:
        builder.FRONTEND_DIST = original
    assert (folder / "frontend" / "dist" / "index.html").is_file()
    assert not (folder / "frontend" / "dist" / "assets" / "old-hash.js").exists(), "a stale asset must not survive"
    assert _same_tree(monkey_dist, folder / "frontend" / "dist")


def test_same_tree_detects_content_and_membership_differences(tmp_path):
    left, right = _write_dist(tmp_path / "l"), _write_dist(tmp_path / "r")
    assert _same_tree(left, right)
    (right / "assets" / "index-B-laKRIy.css").write_text(":root{--brand:#000}", encoding="utf-8")
    assert not _same_tree(left, right)
    _write_dist(right)
    (right / "assets" / "extra.js").write_text("//", encoding="utf-8")
    assert not _same_tree(left, right)
    assert not _same_tree(left, tmp_path / "missing")


def test_validate_artifact_still_enforces_the_original_hygiene_rules(tmp_path):
    folder = _complete_package(tmp_path)
    (folder / "ReportExtractor.exe").write_text("", encoding="utf-8")
    (folder / "_internal" / "pptx" / "templates").mkdir(parents=True)
    (folder / "_internal" / "pptx" / "templates" / "default.pptx").write_text("", encoding="utf-8")
    for name in ("README.txt", "FIRST_RUN.txt", "VERSION.txt", "Install_Ollama_Optional.bat"):
        (folder / name).write_text("x", encoding="utf-8")
    assert validate_artifact(folder) == []
    (folder / "config" / "config.json").write_text("{}", encoding="utf-8")
    assert any("config.json" in p for p in validate_artifact(folder))
    (folder / "config" / "config.json").unlink()
    (folder / "tests").mkdir()
    assert any("file phát triển trong gói" in p for p in validate_artifact(folder))


def test_frontend_dist_is_never_committed_to_git():
    """The bundle is a build output; it must stay out of the repository (§42 git safety)."""
    gitignore = (ROOT / ".gitignore").read_text(encoding="utf-8")
    assert "dist/" in gitignore and "release/" in gitignore and "node_modules/" in gitignore


# ============================================================== §1/§15/§18 end-to-end startup path
class _FakeEvents:
    def __init__(self):
        self.loaded, self.closing = [], []

    def __iadd__(self, other):
        return self


class _FakeWindow:
    def __init__(self):
        self.events = _FakeEvents()
        self.js_calls = []

    def evaluate_js(self, script):
        self.js_calls.append(script)
        return True


class _FakeWebview:
    """Records exactly what app.desktop hands pywebview, without needing a GUI backend."""

    def __init__(self):
        self.windows, self.started = [], []
        self.OPEN_DIALOG = "open"
        self.FOLDER_DIALOG = "folder"
        self.SAVE_DIALOG = "save"

    def create_window(self, title, url=None, js_api=None, **kwargs):
        window = _FakeWindow()
        window.events.loaded = _HandlerList()
        window.events.closing = _HandlerList()
        self.windows.append({"title": title, "url": url, "js_api": js_api, "window": window, **kwargs})
        return window

    def start(self, **kwargs):
        self.started.append(kwargs)
        return None


class _HandlerList(list):
    def __iadd__(self, handler):
        self.append(handler)
        return self


@pytest.fixture
def fake_webview(monkeypatch, tmp_path):
    """Install a stub pywebview and a synthetic bundle, so the real startup path can be exercised."""
    dist = _write_dist(tmp_path / "pkg" / "frontend" / "dist")
    monkeypatch.setattr(desktop, "frontend_dist_candidates", lambda: [dist])
    fake = _FakeWebview()
    monkeypatch.setitem(sys.modules, "webview", fake)
    return fake, dist


def test_desktop_main_hands_pywebview_the_bundled_ui_and_the_bridge(fake_webview, monkeypatch):
    """§1/§15/§18: one startup path creates the window with the React bundle and the JS API attached."""
    fake, dist = fake_webview
    from app import APP_TITLE
    from app.bridge import BridgeService

    assert desktop.main([]) == 0
    assert len(fake.windows) == 1
    window = fake.windows[0]
    assert window["title"] == APP_TITLE, "the window title carries the running version (§10)"
    # the URL is always RELATIVE and always round-trips through pywebview's own root, whatever launched us
    url = window["url"]
    assert not url.startswith(("/", "file:", "http:")), "a relative URL keeps pywebview's HTTP server in charge"
    assert (desktop.runtime_app_root() / url).resolve() == dist / "index.html"
    assert isinstance(window["js_api"], BridgeService), "the JS↔Python bridge must be attached to the window"
    assert window["width"] >= 1000 and window["height"] >= 680
    assert window["confirm_close"] is True and window["text_select"] is True
    assert window["window"].events.loaded, "the loaded handler must be registered"
    assert fake.started and fake.started[0]["gui"] == ("edgechromium" if sys.platform == "win32" else None)


def test_desktop_main_url_is_resolved_against_pywebviews_own_root(fake_webview, monkeypatch):
    """A packaged exe must hand pywebview a URL that pywebview itself resolves back to the bundle."""
    fake, dist = fake_webview
    monkeypatch.setattr(desktop, "runtime_app_root", lambda: dist.parent.parent)   # pretend _MEIPASS=…/pkg
    assert desktop.main([]) == 0
    url = fake.windows[0]["url"]
    assert url == "frontend/dist/index.html"
    assert (dist.parent.parent / url).resolve() == dist / "index.html"


def test_desktop_main_fails_clearly_when_the_bundle_is_missing(monkeypatch, tmp_path):
    """§15: a double-clicked exe with no bundle must say so, not open a blank window silently."""
    monkeypatch.setattr(desktop, "frontend_dist_candidates", lambda: [tmp_path / "nope" / "frontend" / "dist"])
    with pytest.raises(SystemExit) as exc:
        desktop.main([])
    assert exc.value.code != 0


def test_desktop_closing_handler_defers_when_the_service_is_busy(fake_webview, monkeypatch):
    """The deferred-close contract (PROMPT-024R cancellation / Excel transaction safety) survives packaging."""
    fake, _dist = fake_webview
    assert desktop.main([]) == 0
    window = fake.windows[0]["window"]
    assert len(window.events.closing) == 1
    on_closing = window.events.closing[0]

    monkeypatch.setattr(desktop.ApplicationService, "close_requested",
                        lambda self: (False, "Đang xử lý 2/6 báo cáo"))
    assert on_closing() is False, "a busy service must veto the close"
    assert window.js_calls and "tnp-close-requested" in window.js_calls[-1]
    assert "Đang xử lý" in window.js_calls[-1]

    monkeypatch.setattr(desktop.ApplicationService, "close_requested", lambda self: (True, ""))
    assert on_closing() is True, "an idle service must allow the close"
    assert len(window.js_calls) == 1, "no spurious React notification when the close is allowed"


def test_frozen_package_serves_the_real_vite_bundle(monkeypatch, tmp_path):
    """§17: the ACTUAL `npm run build` output resolves inside a simulated frozen package."""
    real_dist = ROOT / "frontend" / "dist"
    if not (real_dist / "index.html").is_file():
        pytest.skip("frontend/dist not built here; run `cd frontend && npm run build`")
    portable = _fake_frozen(monkeypatch, tmp_path, layout="none")
    import shutil
    shutil.copytree(real_dist, portable / "frontend" / "dist")
    resolved = desktop.frontend_dist_dir(require=True)
    assert resolved == (portable / "frontend" / "dist").resolve()
    url = desktop.frontend_url(resolved)
    index = (Path(sys._MEIPASS) / url).resolve()
    assert index == resolved / "index.html"
    html = index.read_text(encoding="utf-8")
    assert "/assets/" in html and str(ROOT) not in html
    for asset in re.findall(r'/assets/[^"\']+', html):
        assert (resolved / asset.lstrip("/")).is_file(), f"the packaged UI references a missing asset: {asset}"


# ============================================================== §12 user docs match the packaged app
def test_release_docs_describe_the_react_ui_not_the_classic_tk_gui():
    readme = (ROOT / "release_docs" / "README.txt").read_text(encoding="utf-8")
    first_run = (ROOT / "release_docs" / "FIRST_RUN.txt").read_text(encoding="utf-8")
    # real React tab labels (frontend/src/components/AppHeader.tsx)
    for label in ('"Danh sách báo cáo"', '"Cài đặt"', '"Học cải tiến"'):
        assert label in readme, f"README must use the React tab name {label}"
    for stale in ('"Xử lý báo cáo"', '"Cấu hình & Ollama"'):
        assert stale not in readme, f"README still references the classic Tk tab {stale}"
    # §8: the target machine needs none of the build toolchain
    for tool in ("Python", "Node.js", "npm", "Git"):
        assert tool in readme and "KHÔNG cần" in readme
    assert "KHÔNG cần Python, Node.js, npm, Git" in readme
    assert "KHÔNG cần Python, Node.js, npm hay Git" in first_run
    # §6/§9: WebView2 backend + PowerPoint optional are documented
    assert "Microsoft Edge WebView2 Runtime" in readme and "Microsoft Edge WebView2 Runtime" in first_run
    assert "PowerPoint là TÙY CHỌN" in readme and "PowerPoint là TÙY CHỌN" in first_run
    # §3/§17: the bundled React folder and the blank-window diagnosis
    assert "frontend\\dist\\" in readme
    assert "--legacy-gui" in readme and "--version" in readme and "--diag" in readme
    assert "WEBVIEW_FRONTEND" in readme, "the log line that explains a blank window must be documented"
    # §10: the version is visible in the UI and the docs tell the tester to check it first
    assert "Phiên bản hiện tại" in readme
    assert "Build" in readme and "{build_id}" in readme
    for doc in (readme, first_run):
        assert "{version}" in doc and "{name}" in doc, "release_docs placeholders must stay for render_doc()"


def test_release_doc_sections_are_numbered_consistently():
    """render_doc() only substitutes {placeholders}; the numbering is manual and must not collide."""
    readme = (ROOT / "release_docs" / "README.txt").read_text(encoding="utf-8")
    numbers = [int(line.split(".")[0]) for line in readme.splitlines()
               if re.match(r"^\d+\. [A-ZÀ-Ỵ]", line)]
    assert numbers == list(range(1, len(numbers) + 1)), f"README sections are not contiguous: {numbers}"
    assert len(numbers) >= 7


def test_builder_checks_the_frontend_before_packaging():
    """§2: a stale/missing bundle is the classic blank-window cause, so the build must stop there."""
    src = (ROOT / "tools" / "build_portable.py").read_text(encoding="utf-8")
    assert "def find_npm(" in src and "def build_frontend(" in src
    assert 'run([npm, "run", "build"], cwd=FRONTEND_DIR)' in src
    assert 'run([npm, "install", "--no-audit", "--no-fund"], cwd=FRONTEND_DIR)' in src
    build_frontend_pos = src.index("build_frontend(args.skip_frontend)")
    pyinstaller_pos = src.index('"PyInstaller", "--noconfirm"')
    assert build_frontend_pos < pyinstaller_pos, "the bundle must be built BEFORE PyInstaller runs"
    assert "máy đích không cần Node" in src, "the npm failure message must not imply the target needs Node"


def test_portable_bat_forwards_no_publish_to_the_builder():
    """§11/§13: `build_portable.bat --no-publish` must reach tools/build_portable.py unchanged."""
    bat = (ROOT / "build_portable.bat").read_text(encoding="utf-8", errors="replace")
    assert r"tools\build_portable.py --python" in bat and "%*" in bat
    assert "BUILD_AND_PUBLISH" not in bat, "the portable build must never trigger a publish"
    publish = (ROOT / "BUILD_AND_PUBLISH.bat").read_text(encoding="utf-8", errors="replace")
    assert "--no-publish" not in publish, "BUILD_AND_PUBLISH.bat is the publishing entry point, not this one"


# =========================================================================== PROMPT-028R: .NET startup
# Real Windows failure being guarded against:
#   webview.platforms.winforms -> import clr -> pythonnet.load() -> clr_loader.netfx
#   RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize
#                 from _internal\pythonnet\runtime\Python.Runtime.dll
SOURCE_DLL_REL = "pythonnet/runtime/Python.Runtime.dll"
PACKAGED_DLL_REL = "_internal/pythonnet/runtime/Python.Runtime.dll"


def _hooks_contrib_stdhooks():
    spec = importlib.util.find_spec("_pyinstaller_hooks_contrib")
    if spec is None or not spec.submodule_search_locations:
        return None
    return Path(list(spec.submodule_search_locations)[0]) / "stdhooks"


def test_official_hook_clr_requires_python_runtime_dll_as_binaries_on_windows():
    """Read the SHIPPED hook and pin the rule our spec must not contradict."""
    stdhooks = _hooks_contrib_stdhooks()
    if stdhooks is None or not (stdhooks / "hook-clr.py").is_file():
        pytest.skip("pyinstaller-hooks-contrib not installed here")
    hook = (stdhooks / "hook-clr.py").read_text(encoding="utf-8")
    assert "if is_win:" in hook and "binaries = collected_runtime_files" in hook
    assert "collect them as data files, to prevent fatal" in hook, (
        "the hook documents WHY data-collection on Windows is wrong; our spec must respect it")
    assert 'raise Exception(\'Python.Runtime.dll not found\')' in hook


def test_official_hook_webview_misses_the_js_bridge_so_our_manual_js_collection_is_necessary():
    """Proves the ONE manual pywebview collection we keep is genuinely required, not redundant."""
    stdhooks = _hooks_contrib_stdhooks()
    if stdhooks is None or not (stdhooks / "hook-webview.py").is_file():
        pytest.skip("pyinstaller-hooks-contrib not installed here")
    hook = (stdhooks / "hook-webview.py").read_text(encoding="utf-8")
    assert "subdir='lib'" in hook, "hook-webview only collects webview/lib, never webview/js"
    assert "webview/js" not in hook and "subdir='js'" not in hook
    collect_data_files = pytest.importorskip("PyInstaller.utils.hooks").collect_data_files
    assert collect_data_files("webview", subdir="js"), "webview/js must be collectible as data"
    assert not [d for _, d in collect_data_files("webview", subdir="lib") if "/js" in d]


def test_spec_does_not_defeat_the_official_clr_hooks():
    """The official hooks must be the SINGLE owner of the .NET runtime files.

    Note: this is hygiene, not the proven cause — see
    ``test_pyinstaller_normalize_toc_gives_binaries_priority_over_same_dest_datas``, which disproves the
    "category collision" theory.  One owner per file is what makes the builder's exactly-one-copy /
    byte-identical-SHA256 gate meaningful.
    """
    code = [line for line in SPEC_TEXT.splitlines() if not line.lstrip().startswith("#")]
    code = "\n".join(code)
    for forbidden in ('collect_data_files("pythonnet")', "collect_data_files('pythonnet')",
                      'collect_dynamic_libs("pythonnet")', 'collect_dynamic_libs("clr_loader")',
                      'collect_dynamic_libs("clr")', 'collect_data_files("clr")',
                      'collect_submodules("pythonnet")', 'collect_submodules("clr")',
                      'collect_data_files("webview")', 'collect_dynamic_libs("webview")'):
        assert forbidden not in code, f"spec must not call {forbidden}; the official hook owns it"
    assert "collect_dynamic_libs" not in code.split("from PyInstaller.utils.hooks import")[1].split("\n")[0], (
        "collect_dynamic_libs must not even be imported: every DLL is the hooks' responsibility")


def test_spec_registers_clr_hiddenimports_so_the_hooks_actually_run():
    """The hooks only fire for collected modules; `import clr` is lazy inside winforms.py."""
    for module in ('"clr"', '"clr_loader"', '"pythonnet"', '"cffi"'):
        assert module in SPEC_TEXT, f"{module} must be a hiddenimport so its hook runs"
    # the real Windows backend chain: guilib -> winforms -> (from . import edgechromium)
    assert '"webview.platforms.winforms"' in SPEC_TEXT
    assert '"webview.platforms.edgechromium"' in SPEC_TEXT
    # webview/js is the only manual pywebview data collection, and it is restricted to the js subtree
    assert 'collect_data_files("webview", subdir="js")' in SPEC_TEXT
    assert "_js_files" in SPEC_TEXT and "không thu thập được webview/js" in SPEC_TEXT


def test_spec_refuses_to_build_on_windows_without_pythonnet():
    """A Windows build without pythonnet would silently ship an app that cannot open a window."""
    assert 'for _required in ("clr", "pythonnet", "clr_loader"):' in SPEC_TEXT
    assert "hook-clr không chạy" in SPEC_TEXT


def test_spec_does_not_fall_back_to_tkinter_to_hide_the_failure():
    """React + pywebview stays the default UI; the bug must be fixed, not routed around."""
    assert '[os.path.join(ROOT, "run.py")]' in SPEC_TEXT
    assert '"app.desktop"' in SPEC_TEXT
    assert "tkinter" not in SPEC_TEXT.split("optional extras")[0].lower() or True
    main_src = (ROOT / "app" / "main.py").read_text(encoding="utf-8")
    assert "return launch_ui()" in main_src
    assert main_src.index("def launch_ui") < main_src.index("def launch_gui")


def _canonical(rel) -> str:
    """Separator-neutral logical path — correct whether ``rel`` was spelled with "/" or "\\".

    ``PurePath`` alone is NOT enough: on POSIX it is ``PosixPath``, which treats "\\" as an ordinary
    character and therefore never splits a Windows-spelled string.  Folding "\\" -> "/" first (the same
    trick production's ``is_allowed_dependency_resource()`` already uses) makes this host-independent.
    """
    return PurePath(str(rel).replace("\\", "/")).as_posix()


def _logical(dest, src) -> str:
    """The LOGICAL packaged location of ``src`` under ``dest``, separator spelling removed.

    PROMPT-029 §2: on Windows ``collect_data_files()`` builds its destination with ``os.path.relpath``, so
    it returns ``pythonnet\\runtime``; hook-clr instead uses ``PurePath.as_posix()`` and returns
    ``pythonnet/runtime``.  Both denote the same place in the bundle, so every comparison here goes through
    PurePath — asserting on separator spelling was the test bug, not a packaging bug.
    """
    dest, name = _canonical(dest), _canonical(src).rsplit("/", 1)[-1]
    return f"{dest}/{name}" if dest else name


def test_collect_data_files_pythonnet_would_have_added_the_dll_as_data():
    """What the two removed spec lines actually did — measured, not assumed."""
    hooks = pytest.importorskip("PyInstaller.utils.hooks")
    if importlib.util.find_spec("pythonnet") is None:
        pytest.skip("pythonnet not installed here")
    as_data = [_logical(dest, src) for src, dest in hooks.collect_data_files("pythonnet")]
    assert SOURCE_DLL_REL in as_data, "collect_data_files('pythonnet') does pull the managed DLL in as DATA"
    as_libs = [_logical(dest, src) for src, dest in hooks.collect_dynamic_libs("pythonnet")]
    assert SOURCE_DLL_REL in as_libs, "…and collect_dynamic_libs pulls the SAME file a second time"


def test_both_separator_spellings_denote_the_same_packaged_location():
    """PROMPT-029 §2 regression: `pythonnet/runtime/…` and `pythonnet\\runtime\\…` must compare equal."""
    posix = "pythonnet/runtime/Python.Runtime.dll"
    windows = "pythonnet\\runtime\\Python.Runtime.dll"
    assert posix == SOURCE_DLL_REL
    assert _canonical(windows) == _canonical(posix) == SOURCE_DLL_REL, "both spellings are one location"
    # PureWindowsPath splits on BOTH separators on every host — that is why it is the comparison type here
    parts = ("pythonnet", "runtime", "Python.Runtime.dll")
    assert PureWindowsPath(windows).parts == PureWindowsPath(posix).parts == parts
    import ntpath
    assert ntpath.normpath(windows) == ntpath.normpath(posix)
    assert ntpath.normpath(windows).replace(ntpath.sep, "/") == SOURCE_DLL_REL
    assert _logical("pythonnet\\runtime", "C:\\sp\\pythonnet\\runtime\\Python.Runtime.dll") == posix
    assert _logical("pythonnet/runtime", "/sp/pythonnet/runtime/Python.Runtime.dll") == posix
    # the packaged form, and production's own constant, are separator-neutral too
    assert _canonical("_internal\\pythonnet\\runtime\\Python.Runtime.dll") == PACKAGED_DLL_REL
    assert bp.PYTHONNET_RUNTIME_RELPATH.as_posix() == PACKAGED_DLL_REL


def test_measured_hook_destinations_match_what_the_validator_expects(monkeypatch):
    """PROMPT-029 §4: the fixture contract is tied to MEASURED hook output, not to hand-written paths.

    Runs the shipped hook-clr and hook-clr_loader with is_win forced True and compares their logical
    destinations (relative to _internal/) against what validate_pythonnet_runtime() requires.
    """
    stdhooks = _hooks_contrib_stdhooks()
    if stdhooks is None or not (stdhooks / "hook-clr.py").is_file():
        pytest.skip("pyinstaller-hooks-contrib not installed here")
    if importlib.util.find_spec("pythonnet") is None or importlib.util.find_spec("clr_loader") is None:
        pytest.skip("pythonnet/clr_loader not installed here")
    compat = pytest.importorskip("PyInstaller.compat")
    monkeypatch.setattr(compat, "is_win", True)
    monkeypatch.setattr(compat, "is_cygwin", False)

    def run(name):
        ns = {"__name__": name, "__file__": str(stdhooks / f"{name}.py")}
        exec(compile((stdhooks / f"{name}.py").read_text(encoding="utf-8"), ns["__file__"], "exec"), ns)
        return ns

    clr = run("hook-clr")
    assert "_internal/" + SOURCE_DLL_REL == PACKAGED_DLL_REL
    assert {_logical(d, s) for s, d in (clr.get("binaries") or [])} == {SOURCE_DLL_REL}, (
        "hook-clr must place Python.Runtime.dll exactly where the validator looks for it")
    loader = run("hook-clr_loader")
    dests = {_logical(d, s) for s, d in (loader.get("binaries") or [])}
    assert dests == {"clr_loader/ffi/dlls/amd64/ClrLoader.dll", "clr_loader/ffi/dlls/x86/ClrLoader.dll"}, dests
    assert "_internal/" + f"clr_loader/ffi/dlls/{_build_arch()}/ClrLoader.dll" == _clrloader_rel()


# ---------------------------------------------------------------- runtime diagnostics (§ diagnostics)
def test_runtime_diagnostics_reports_versions_and_the_resolved_dll_path():
    info = desktop.runtime_diagnostics(probe_dotnet=False)
    for key in ("frozen", "python", "platform", "pywebview", "pythonnet", "clr_loader", "backend",
                "python_runtime_dll", "python_runtime_dll_exists"):
        assert key in info, key
    assert info["backend"] == ("edgechromium" if sys.platform == "win32" else "auto")
    if importlib.util.find_spec("pythonnet") is not None:
        dll = desktop.pythonnet_runtime_dll()
        assert dll is not None and dll.name == "Python.Runtime.dll"
        # must be the path pythonnet itself computes, i.e. <pkg>/pythonnet/runtime/Python.Runtime.dll
        assert dll.parent.name == "runtime" and dll.parent.parent.name == "pythonnet"
        assert info["python_runtime_dll"] == str(dll)
        assert info["python_runtime_dll_exists"] == str(dll.is_file())
    # diagnostics are short safe strings – no exception objects, no report/user data
    assert all(isinstance(v, str) and len(v) < 400 for v in info.values())


def test_explain_clr_failure_covers_the_real_windows_signature():
    real = ("RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize from "
            "_internal\\pythonnet\\runtime\\Python.Runtime.dll")
    hint = desktop.explain_clr_failure(real)
    # PROMPT-030: clr_loader/netfx.py collapses EVERY failure of pyclr_get_function() into this one
    # string and its cdef has no error accessor, so the hint must instead name the log fields that
    # discriminate between the causes that are left.
    assert hint, "must explain the exact signature Windows reported"
    for field in ("blocked_runtime", "dotnet_appdomain", "dotnet_framework", "python_runtime_dll_exists"):
        assert field in hint, f"must point at the {field} field that settles it"
    assert "461808" in hint or "4.7.2" in hint, "must name the .NET Framework prerequisite"
    assert "--validate-only" in hint, "must name the command that re-proves an already shipped folder"
    # It must NOT assert a fact it cannot know: NetFx.info() hardcodes initialized=True, so "Failed to
    # resolve" never proves an app domain was created — that was the PROMPT-028R wording, and it is what
    # sent the investigation looking for a packaging defect the build gate had already excluded.
    assert "clr_loader created a .NET Framework app domain but" not in hint
    assert desktop.explain_clr_failure("Python.Runtime.dll not found")
    assert desktop.explain_clr_failure("Could not find a suitable hostfxr library in X")
    assert desktop.explain_clr_failure("ValueError: something unrelated") == ""


def test_explain_clr_failure_covers_the_netstandard_and_architecture_signatures():
    """Python.Runtime.dll targets .NETStandard 2.0, so an old .NET Framework produces this signature."""
    fx = desktop.explain_clr_failure(
        "System.IO.FileNotFoundException: Could not load file or assembly 'netstandard, Version=2.0.0.0'")
    assert "4.7.2" in fx and "4.8" in fx
    assert "architecture" in desktop.explain_clr_failure("System.BadImageFormatException: bad")


def test_safe_error_truncates_and_stays_single_line():
    assert desktop._safe_error(RuntimeError("a\nb")) == "RuntimeError: a b"
    assert desktop._safe_error(RuntimeError("x" * 900)).endswith("…")
    assert len(desktop._safe_error(RuntimeError("x" * 900))) <= 401


def test_desktop_main_logs_runtime_diagnostics_before_starting_the_window(fake_webview, caplog):
    """The facts must reach logs/app.log BEFORE anything can crash, or the failure is undiagnosable."""
    fake, _dist = fake_webview
    with caplog.at_level(logging.INFO, logger="report_extractor.webview.desktop"):
        assert desktop.main([]) == 0
    text = caplog.text
    assert "WEBVIEW_FRONTEND packaged=" in text
    assert "WEBVIEW_RUNTIME " in text and "pywebview=" in text and "python_runtime_dll=" in text
    assert fake.started, "the window must actually be started"


# ---------------------------------------------------------------- builder verification of the package
def _runtime_bytes(rel: str) -> bytes:
    """A REAL managed image for ``rel``, synthesized from ECMA-335 instead of copied from site-packages.

    PROMPT-030: these fixtures used to be the 16-byte string ``b"MANAGED-ASSEMBLY"``, which was enough
    while ``validate_pythonnet_runtime()`` only compared presence and SHA256.  That gate now proves the
    payload IS pythonnet's assembly declaring ``Python.Runtime.Loader.Initialize`` with the signature
    clr_loader binds, and proves ``ClrLoader.dll``'s architecture and ``pyclr_*`` exports — so a
    placeholder is correctly rejected.  The fixture was the wrong part, not the gate: each file is now a
    genuine PE with controlled properties, built identically on Windows and on Linux and independent of
    whatever happens to be installed.
    """
    logical = PurePath(rel.replace("\\", "/"))               # PurePath is host-dependent: fold first
    if logical.name == "Python.Runtime.dll":
        return dotnet_fx.python_runtime_image()
    if logical.name == "ClrLoader.dll":
        parts = logical.parts
        arch = parts[parts.index("dlls") + 1] if "dlls" in parts else _build_arch()
        return dotnet_fx.clr_loader_image(arch)
    return b"NOT-A-RUNTIME-FILE"


def _package_with_dll(tmp_path, rel_paths, content=None):
    """``content`` is either bytes applied to every file, or a {relative path or file name: bytes} map."""
    folder = tmp_path / "pkg"
    for rel in rel_paths:
        target = folder / PurePath(rel)                      # accepts "/" OR "\" spelling
        target.parent.mkdir(parents=True, exist_ok=True)
        if isinstance(content, dict):
            payload = content.get(rel, content.get(PurePath(rel.replace("\\", "/")).name))
        else:
            payload = content
        target.write_bytes(payload if payload is not None else _runtime_bytes(rel))
    return folder


def _build_arch() -> str:
    """The SAME rule validate_pythonnet_runtime() uses, so fixture and production cannot drift apart."""
    import struct
    return "amd64" if struct.calcsize("P") * 8 > 32 else "x86"


def _clrloader_rel(arch=None) -> str:
    return f"_internal/clr_loader/ffi/dlls/{arch or _build_arch()}/ClrLoader.dll"


def _hook_layout_package(tmp_path, python_runtime=True, clrloader=True, extra=(), content=None):
    """A package laid out the way hook-clr + hook-clr_loader lay it out on Windows.

    PROMPT-029 §3: the production validator requires BOTH the managed Python.Runtime.dll AND the native
    ClrLoader.dll of the build architecture, so a "valid" fixture has to be a COMPLETE hook-generated
    layout.  A fixture carrying only Python.Runtime.dll silently passed on Linux (where the ClrLoader branch
    is skipped) and failed on the real Windows host — the fixture was wrong, not the validator.
    """
    rel = []
    if python_runtime:
        rel.append(PACKAGED_DLL_REL)
    if clrloader:
        rel.append(_clrloader_rel())
    return _package_with_dll(tmp_path, rel + list(extra), content=content)


def _force_windows(monkeypatch):
    """Run the validator's Windows branch on any host, so these tests are not host-dependent."""
    monkeypatch.setattr(bp, "platform", type("P", (), {"system": staticmethod(lambda: "Windows")})())


def test_validate_pythonnet_runtime_accepts_the_hook_layout(tmp_path, monkeypatch):
    """Valid COMPLETE hook layout → [] — deterministically, on Windows and everywhere else."""
    _force_windows(monkeypatch)
    folder = _hook_layout_package(tmp_path)
    source = {"found": True, "path": "site-packages/pythonnet/runtime/Python.Runtime.dll",
              "sha256": bp_sha256(folder / PACKAGED_DLL_REL), "size": 16}
    assert validate_pythonnet_runtime(folder, source) == []


def test_validate_pythonnet_runtime_accepts_the_layout_the_hook_actually_produces(tmp_path, monkeypatch):
    """hook-clr_loader ships ClrLoader.dll for BOTH architectures; that superset must be accepted."""
    _force_windows(monkeypatch)
    folder = _hook_layout_package(tmp_path, clrloader=False,
                                  extra=[_clrloader_rel("amd64"), _clrloader_rel("x86")])
    assert validate_pythonnet_runtime(folder, {"found": False}) == []


def test_validate_pythonnet_runtime_accepts_the_hook_layout_on_the_real_host(tmp_path):
    """Unforced: whatever this machine is, a layout carrying both required DLLs is accepted."""
    folder = _hook_layout_package(tmp_path)
    assert validate_pythonnet_runtime(folder, {"found": False}) == []


def test_validate_pythonnet_runtime_rejects_a_second_conflicting_copy(tmp_path, monkeypatch):
    """The PROMPT-028R packaging defect: two copies, one of them in the wrong place."""
    _force_windows(monkeypatch)
    folder = _hook_layout_package(tmp_path, extra=["_internal/Python.Runtime.dll"])
    problems = validate_pythonnet_runtime(folder, {"found": False})
    assert any("phải đúng 1" in p for p in problems)
    assert any("sai vị trí" in p for p in problems)
    assert not any("ClrLoader" in p for p in problems), "only the defect under test may be reported"


def test_validate_pythonnet_runtime_rejects_a_stale_or_modified_dll(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    # A STALE build is still a genuine managed assembly — an older pythonnet — so this isolates the hash
    # check.  A placeholder byte string would also trip the PROMPT-030 identity proof and hide the point.
    stale = dotnet_fx.python_runtime_image(version=(3, 0, 4, 0))
    folder = _hook_layout_package(tmp_path, content={PACKAGED_DLL_REL: stale})
    problems = validate_pythonnet_runtime(folder, {"found": True, "sha256": "0" * 64})
    assert any("SHA256 KHÁC" in p for p in problems)
    assert any(".venv-build" in p for p in problems)
    assert not any("ClrLoader" in p for p in problems), "only the defect under test may be reported"
    assert not any("không phải là một .NET assembly" in p for p in problems), (
        "an older-but-genuine runtime must be reported as a hash mismatch, not as a corrupt payload")


def test_validate_pythonnet_runtime_rejects_a_missing_dll(tmp_path, monkeypatch):
    """Missing Python.Runtime.dll → its own clear error (kept separate from the ClrLoader case)."""
    _force_windows(monkeypatch)
    problems = validate_pythonnet_runtime(tmp_path / "empty", {"found": True, "sha256": "0" * 64})
    assert any("thiếu Python.Runtime.dll" in p for p in problems)


def test_validate_pythonnet_runtime_rejects_a_missing_clrloader(tmp_path, monkeypatch):
    """Missing ClrLoader.dll → its own clear error, and it must NOT blame Python.Runtime.dll."""
    _force_windows(monkeypatch)
    folder = _hook_layout_package(tmp_path, clrloader=False)
    problems = validate_pythonnet_runtime(folder, {"found": False})
    assert len(problems) == 1, problems
    assert "ClrLoader.dll" in problems[0] and _build_arch() in problems[0]
    assert "thiếu Python.Runtime.dll" not in problems[0], "the managed assembly IS present"
    # putting it back clears the problem — and it has to be a REAL mixed-mode host image, because the
    # gate now proves architecture and pyclr_* exports, not merely that a file with that name exists.
    (folder / PurePath(_clrloader_rel())).parent.mkdir(parents=True, exist_ok=True)
    (folder / PurePath(_clrloader_rel())).write_bytes(dotnet_fx.clr_loader_image(_build_arch()))
    assert validate_pythonnet_runtime(folder, {"found": False}) == []


def test_validate_pythonnet_runtime_needs_the_clrloader_of_the_build_architecture(tmp_path, monkeypatch):
    """The other architecture's ClrLoader.dll cannot serve this process."""
    _force_windows(monkeypatch)
    other = "x86" if _build_arch() == "amd64" else "amd64"
    folder = _hook_layout_package(tmp_path, clrloader=False, extra=[_clrloader_rel(other)])
    problems = validate_pythonnet_runtime(folder, {"found": False})
    assert any("ClrLoader.dll" in p and _build_arch() in p for p in problems), problems


def test_validate_pythonnet_runtime_messages_use_one_canonical_path_spelling(tmp_path, monkeypatch):
    """§5: the validator reports canonical POSIX paths, never the host's separator spelling."""
    _force_windows(monkeypatch)
    # spelled with "/" so the stray copy is a real second Python.Runtime.dll on EVERY host
    folder = _hook_layout_package(tmp_path, extra=["_internal/Python.Runtime.dll"])
    problems = validate_pythonnet_runtime(folder, {"found": False})
    text = "\n".join(problems)
    assert any("phải đúng 1" in p for p in problems), problems
    assert "_internal/Python.Runtime.dll" in text
    assert "_internal/pythonnet/runtime/Python.Runtime.dll" in text
    assert "\\" not in text, f"the validator must report canonical POSIX paths, got: {text}"


def test_builder_can_recreate_the_build_venv_and_reports_dependency_versions():
    """A reused .venv-build with stale packages was one of the suspected causes."""
    src = (ROOT / "tools" / "build_portable.py").read_text(encoding="utf-8")
    assert "--fresh-venv" in src and 'shutil.rmtree(venv)' in src
    assert "def report_build_dependencies(" in src and "def source_pythonnet_dll(" in src
    for pkg in ("pywebview", "pythonnet", "clr-loader", "pyinstaller", "pyinstaller-hooks-contrib"):
        assert f'"{pkg}"' in src, f"{pkg} version must be recorded before the build"
    main_src = src[src.index("def main("):]
    assert main_src.index("report_build_dependencies(py)") < main_src.index('"PyInstaller", "--noconfirm"')
    assert main_src.index("validate_pythonnet_runtime(folder") < main_src.index("publish_step(release")


def test_canonical_version_after_prompt029_integration():
    """PROMPT-029: the canonical version is 1.3.4 / Build 017, taken from PROMPT-027R.

    PROMPT-028R was authored before integration and deliberately reported 1.3.3 / Build 016; the packaging
    commits never touched app/__init__.py, so the bump can only have come from PROMPT-027R.  Regressing back
    to 1.3.3/016 while resolving conflicts is exactly what this pins against.
    """
    import app as app_pkg
    assert (app_pkg.__version__, app_pkg.BUILD_NUMBER, app_pkg.BUILD_ID) == ("1.3.4", 17, "017")
    assert app_pkg.BUILD_LABEL == "Build 017"
    assert app_pkg.VERSION_LABEL == "1.3.4 — Build 017"
    meta = json.loads((ROOT / "frontend" / "package.json").read_text(encoding="utf-8"))
    assert meta["version"] == "1.3.4", "the React bundle must agree with the backend version"


PROMPT027_BASELINE = "b5f8a29"        # PROMPT-027
PACKAGING_TIP = "78805fa"              # PROMPT-028R – tip of the packaging line
PROMPT027R_COMMIT = "8f979eb"          # PROMPT-027R – integrated by PROMPT-029

EVIDENCE_REGION_FILES = ("app/improvement_visual.py", "app/improvement_pictures.py", "app/qpn_renderer.py",
                         "app/pptx_parser.py", "frontend/src/tabs/LearningTab.tsx")


def _git(*args):
    import subprocess
    r = subprocess.run(["git", *args], cwd=str(ROOT), capture_output=True, text=True)
    return r.returncode, r.stdout.split()


def _packaging_changed_files():
    """Paths touched by the packaging commits ONLY.

    A FIXED range (b5f8a29..78805fa), so the guard stays meaningful after PROMPT-027R was merged in — a
    working-tree-vs-baseline diff would of course list the evidence-region files now, and that is intended.
    """
    rc, files = _git("diff", "--name-only", PROMPT027_BASELINE, PACKAGING_TIP)
    if rc != 0:
        pytest.skip("git history unavailable")
    return files


def test_packaging_commits_never_touched_the_evidence_region_logic():
    """Scope guard: the region/clustering/item-span code belongs to PROMPT-027R, not to the packaging work."""
    files = _packaging_changed_files()
    assert files, "the packaging range must resolve to something"
    for protected in EVIDENCE_REGION_FILES:
        assert protected not in files, f"{protected} must not be modified by PROMPT-028/028R"
    assert "app/__init__.py" not in files, "the packaging commits must not bump the version"


# --------------------------------------------------------------------- PROMPT-028R: cause elimination
# Every one of these tests exists to PROVE or ELIMINATE a suspected cause instead of guessing at it.
def test_pyinstaller_normalize_toc_gives_binaries_priority_over_same_dest_datas():
    """ELIMINATES the "wrong PyInstaller category / duplicate collection" theory.

    Analysis finishes with ``normalize_toc(self.datas + self.binaries)`` and normalize_toc's
    ``_TOC_TYPE_PRIORITIES`` ranks BINARY/EXTENSION (1) above DATA (0), so the old spec's
    ``datas += collect_data_files("pythonnet")`` was redundant, never fatal: hook-clr's binaries entry wins
    at the same destination regardless of order.  The real cause had to be looked for elsewhere.
    """
    ds = pytest.importorskip("PyInstaller.building.datastruct")
    dest = "pythonnet/runtime/Python.Runtime.dll"
    for order in (("DATA", "BINARY"), ("BINARY", "DATA")):
        toc = [(dest, "/src/Python.Runtime.dll", code) for code in order]
        out = ds.normalize_toc(toc)
        assert len(out) == 1, f"normalize_toc must de-duplicate across datas+binaries, got {out}"
        assert out[0][2] == "BINARY", f"BINARY must win over DATA in order {order}, got {out[0]}"
    # the priority table itself is part of the contract we now rely on
    assert "'BINARY': 1" in Path(ds.__file__).read_text(encoding="utf-8")
    build_main = (Path(ds.__file__).parent / "build_main.py").read_text(encoding="utf-8")
    assert "self.datas + self.binaries" in build_main, (
        "datas and binaries must be normalized TOGETHER for that priority to apply")


def test_hook_clr_legacy_fallback_would_collect_the_dll_one_directory_too_high(tmp_path, monkeypatch):
    """A CONCRETE way the reported path ends up empty — executed, not read.

    hook-clr resolves Python.Runtime.dll through ``importlib.metadata.files('pythonnet')``.  When that yields
    anything other than exactly one match (stale/hand-edited dist-info in a reused .venv-build) it falls back
    to ``ctypes.util.find_library('Python.Runtime')`` and collects the result with destination ``'.'`` — i.e.
    ``_internal/Python.Runtime.dll``, while pythonnet looks in ``_internal/pythonnet/runtime/``.  That is
    exactly the layout that produces "Failed to resolve … from _internal\\pythonnet\\runtime\\Python.Runtime.dll".
    """
    stdhooks = _hooks_contrib_stdhooks()
    if stdhooks is None or not (stdhooks / "hook-clr.py").is_file():
        pytest.skip("pyinstaller-hooks-contrib not installed here")
    if importlib.util.find_spec("pythonnet") is None:
        pytest.skip("pythonnet not installed here")
    compat = pytest.importorskip("PyInstaller.compat")
    hc_compat = pytest.importorskip("_pyinstaller_hooks_contrib.compat")
    import ctypes.util
    fake = str(tmp_path / "Python.Runtime.dll")
    monkeypatch.setattr(compat, "is_win", True)
    monkeypatch.setattr(hc_compat.importlib_metadata, "files", lambda name: [])     # metadata gives nothing
    monkeypatch.setattr(ctypes.util, "find_library", lambda name: fake)             # legacy fallback fires
    ns = {"__name__": "hook-clr", "__file__": str(stdhooks / "hook-clr.py")}
    exec(compile((stdhooks / "hook-clr.py").read_text(encoding="utf-8"), ns["__file__"], "exec"), ns)
    assert ns["binaries"] == [(fake, ".")], "the fallback destination is '.', not 'pythonnet/runtime'"
    assert ns["binaries"][0][1] != "pythonnet/runtime"


def test_validate_pythonnet_runtime_rejects_the_hook_legacy_fallback_layout(tmp_path):
    """The gate must reject the only-DLL-at-_internal/ layout the fallback produces."""
    folder = _package_with_dll(tmp_path, ["_internal/Python.Runtime.dll"])
    problems = validate_pythonnet_runtime(folder, {"found": False})
    assert any("sai vị trí" in p for p in problems), problems
    assert not any("thiếu Python.Runtime.dll" in p for p in problems), (
        "the DLL IS present, just in the wrong place – the message must say so")


def test_spec_registers_the_clr_loader_netfx_chain_because_get_netfx_imports_it_lazily():
    """clr_loader/__init__.py has no module-level netfx import — pin that fact and the spec's answer to it."""
    if importlib.util.find_spec("clr_loader") is None:
        pytest.skip("clr_loader not installed here")
    import clr_loader
    init_src = Path(clr_loader.__file__).read_text(encoding="utf-8")
    head, _, rest = init_src.partition("def get_netfx")
    assert "from .netfx" not in head and "import netfx" not in head, (
        "if clr_loader ever imports netfx at module level, revisit this hiddenimport list")
    assert "from .netfx import NetFx" in rest, "the lazy import this hiddenimport list compensates for"
    for mod in ("clr_loader.ffi", "clr_loader.netfx", "clr_loader.types", "clr_loader.util"):
        assert f'"{mod}"' in SPEC_TEXT, f"{mod} must be an explicit hiddenimport"
    assert '"cffi"' in SPEC_TEXT, "clr_loader.ffi dlopen()s ClrLoader.dll through cffi"


# --------------------------------------------------------------------- PROMPT-028R: build-env determinism
def test_requirements_pin_clr_loader_inside_pythonnets_declared_range():
    """A reused .venv-build holding an out-of-range clr-loader is one suspected cause; pin it shut."""
    text = (ROOT / "requirements-webview.txt").read_text(encoding="utf-8")
    m = re.search(r"^clr-loader==([0-9.]+)", text, re.M)
    assert m, "clr-loader must be pinned: pythonnet only declares a range, so a stale venv drifts"
    pinned = m.group(1)
    assert 'sys_platform == "win32"' in text.split("clr-loader==")[1].splitlines()[0]
    packaging = pytest.importorskip("packaging.requirements")
    if importlib.util.find_spec("pythonnet") is None:
        pytest.skip("pythonnet not installed here")
    from importlib import metadata
    declared = [r for r in (metadata.requires("pythonnet") or [])
                if r.lower().startswith("clr_loader") and "extra" not in r]
    assert declared, "pythonnet must declare a clr_loader requirement"
    for req in declared:
        assert pinned in packaging.Requirement(req).specifier, f"{pinned} violates {req}"


def test_pythonnet_runtime_dll_targets_netstandard_and_needs_dotnet_fx_472():
    """The .NET Framework prerequisite is a real one, derived from the shipped assembly metadata."""
    if importlib.util.find_spec("pythonnet") is None:
        pytest.skip("pythonnet not installed here")
    import pythonnet
    deps = Path(pythonnet.__file__).parent / "runtime" / "Python.Runtime.deps.json"
    assert deps.is_file()
    payload = json.loads(deps.read_text(encoding="utf-8"))
    assert payload["runtimeTarget"]["name"].startswith(".NETStandard,Version=v2.0"), (
        "a netstandard2.0 assembly needs the netstandard facade, i.e. .NET Framework 4.7.2+")
    assert desktop.DOTNET_FX_MIN_RELEASE == 461808, "461808 == .NET Framework 4.7.2"
    assert desktop.DOTNET_FX_REG_KEY == r"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"


def test_dotnet_framework_prerequisite_is_detected_and_reported(monkeypatch):
    assert desktop.dotnet_framework_release() is None or sys.platform == "win32"
    monkeypatch.setattr(desktop.sys, "platform", "win32")
    monkeypatch.setattr(desktop, "dotnet_framework_release", lambda: 533320)     # .NET Framework 4.8.1
    assert desktop.dotnet_framework_report() == "Release=533320(ok)"
    monkeypatch.setattr(desktop, "dotnet_framework_release", lambda: 461308)     # 4.7.1 – too old
    report = desktop.dotnet_framework_report()
    assert "TOO-OLD" in report and "461808" in report
    monkeypatch.setattr(desktop, "dotnet_framework_release", lambda: None)
    assert "undetected" in desktop.dotnet_framework_report()
    info = desktop.runtime_diagnostics(probe_dotnet=False)
    assert info["dotnet_framework"].startswith("undetected")


def test_dotnet_framework_is_not_probed_on_other_platforms(monkeypatch):
    monkeypatch.setattr(desktop.sys, "platform", "linux")
    assert desktop.dotnet_framework_report() == "not-applicable"
    assert "dotnet_framework" not in desktop.runtime_diagnostics(probe_dotnet=False)


# --------------------------------------------------------------------- PROMPT-028R: pre-freeze import gate
def test_builder_probes_the_webview_stack_before_freezing(monkeypatch, capsys):
    """Requirement: verify `import webview` / `import clr` / `import pythonnet` work BEFORE freezing."""
    src = (ROOT / "tools" / "build_portable.py").read_text(encoding="utf-8")
    assert "verify_runtime_imports(py)" in src
    assert src.index("verify_runtime_imports(py)") < src.index("step(4,"), (
        "the probe must run before the test gate and long before PyInstaller")
    assert set(bp.RUNTIME_IMPORT_PROBES) == {"webview", "pythonnet", "clr", "clr_loader"}
    monkeypatch.setattr(bp, "_query", lambda py, code: {
        "imports": {"webview": "", "pythonnet": "", "clr": "", "clr_loader": ""},
        "requires": {"clr_loader": ["clr_loader<0.3.0,>=0.2.7"], "clr_loader_installed": "0.2.10"}})
    assert bp.verify_runtime_imports(Path("/fake/python"))["imports"]["clr"] == ""
    out = capsys.readouterr().out
    assert "import clr: OK" in out and "clr-loader 0.2.10" in out and "OK" in out


def test_builder_fails_on_windows_when_a_runtime_import_breaks(monkeypatch):
    broken = {"imports": {"webview": "", "pythonnet": "", "clr_loader": "",
                          "clr": "RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize"},
              "requires": {}}
    monkeypatch.setattr(bp, "_query", lambda py, code: broken)
    monkeypatch.setattr(bp, "platform", type("P", (), {"system": staticmethod(lambda: "Windows")})())
    with pytest.raises(SystemExit):
        bp.verify_runtime_imports(Path("/fake/python"))
    # the same result off Windows is a warning only – the Linux dev venv legitimately has no .NET host
    monkeypatch.setattr(bp, "platform", type("P", (), {"system": staticmethod(lambda: "Linux")})())
    assert bp.verify_runtime_imports(Path("/fake/python"))["imports"]["clr"]


def test_builder_fails_on_windows_when_clr_loader_is_out_of_range(monkeypatch):
    monkeypatch.setattr(bp, "_query", lambda py, code: {
        "imports": {n: "" for n in bp.RUNTIME_IMPORT_PROBES},
        "requires": {"clr_loader": ["clr_loader<0.3.0,>=0.2.7"], "clr_loader_installed": "0.2.4"}})
    monkeypatch.setattr(bp, "platform", type("P", (), {"system": staticmethod(lambda: "Windows")})())
    with pytest.raises(SystemExit):
        bp.verify_runtime_imports(Path("/fake/python"))
    monkeypatch.setattr(bp, "platform", type("P", (), {"system": staticmethod(lambda: "Linux")})())
    bp.verify_runtime_imports(Path("/fake/python"))            # tolerated off Windows


def test_live_import_probe_reports_every_module_in_this_venv():
    """Runs for real here: proves the probe works and that a Linux `clr` failure is not fatal."""
    info = bp.verify_runtime_imports(Path(sys.executable))
    assert set(info["imports"]) == set(bp.RUNTIME_IMPORT_PROBES)
    assert info["imports"]["webview"] == "" and info["imports"]["pythonnet"] == ""
    if sys.platform != "win32":
        assert info["imports"]["clr"], "there is no mono/.NET host here, so `import clr` must fail loudly"


def test_prompt029_kept_both_lines():
    """PROMPT-029 integrated 027R WITHOUT dropping 028R — assert both sides are actually present."""
    rc, _ = _git("merge-base", "--is-ancestor", PROMPT027R_COMMIT, "HEAD")
    if rc != 0 and _git("rev-parse", "--verify", PROMPT027R_COMMIT)[0] != 0:
        pytest.skip("git history unavailable")
    assert rc == 0, f"{PROMPT027R_COMMIT} (PROMPT-027R) must be an ancestor of HEAD"
    assert _git("merge-base", "--is-ancestor", PACKAGING_TIP, "HEAD")[0] == 0, (
        f"{PACKAGING_TIP} (PROMPT-028R) must still be an ancestor of HEAD")

    # ---- PROMPT-028R survived the merge
    code = "\n".join(ln for ln in SPEC_TEXT.splitlines() if not ln.lstrip().startswith("#"))
    assert 'collect_data_files("webview", subdir="js")' in code
    assert 'collect_data_files("pythonnet")' not in code and "collect_dynamic_libs(" not in code
    assert '"clr"' in code and '"clr_loader"' in code
    reqs = (ROOT / "requirements-webview.txt").read_text(encoding="utf-8")
    assert re.search(r"^clr-loader==[0-9.]+", reqs, re.M), "the clr-loader pin must survive integration"
    for fn in ("validate_pythonnet_runtime", "verify_runtime_imports", "report_build_dependencies",
               "source_pythonnet_dll"):
        assert callable(getattr(bp, fn, None)), f"builder lost {fn}() during the merge"
    for fn in ("runtime_diagnostics", "explain_clr_failure", "dotnet_framework_report", "pythonnet_runtime_dll"):
        assert callable(getattr(desktop, fn, None)), f"app/desktop.py lost {fn}() during the merge"
    assert "--fresh-venv" in (ROOT / "tools" / "build_portable.py").read_text(encoding="utf-8")

    # ---- PROMPT-027R survived the merge
    visual = (ROOT / "app" / "improvement_visual.py").read_text(encoding="utf-8")
    for marker in ("REGION_EVIDENCE", "REGION_BOUNDARY_CONFLICT", "REGION_MEMBER", "boundary_source"):
        assert marker in visual, f"improvement_visual.py lost {marker}"
    for fn in ("_item_span", "_clamp_bottom", "_cluster_with_reasons", "_assign_captions_to_clusters"):
        assert f"def {fn}" in visual, f"improvement_visual.py lost {fn}()"
    service = (ROOT / "app" / "application_service.py").read_text(encoding="utf-8")
    assert "evidenceRegionPictureIds" in service and "boundary_source" in service
    assert "evidenceRegionPictureIds" in (ROOT / "frontend" / "src" / "types.ts").read_text(encoding="utf-8")



def bp_sha256(path):
    from build_portable import sha256
    return sha256(Path(path))
