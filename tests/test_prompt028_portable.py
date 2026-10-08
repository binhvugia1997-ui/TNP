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

import os
import re
import sys
from pathlib import Path

import pytest

import app.desktop as desktop
import app.main as app_main
import app.runtime_paths as rp
from build_portable import (_same_tree, build_number, copy_frontend_into_portable, package_name,
                                  validate_artifact, validate_frontend, write_release_metadata)

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
    assert '(FRONTEND_DIST, os.path.join("frontend", "dist"))' in SPEC_TEXT, "frontend/dist must be a `datas` entry"
    assert 'FRONTEND_DIST = os.path.join(ROOT, "frontend", "dist")' in SPEC_TEXT
    for component in ('"webview"', '"pythonnet"', '"clr"', '"clr_loader"'):
        assert component in SPEC_TEXT, f"{component} must be collected for the WebView2 backend"
    assert "collect_data_files(_name)" in SPEC_TEXT, "webview/js and webview/lib are DATA files"
    assert "collect_dynamic_libs(_name)" in SPEC_TEXT, "WebView2Loader.dll must be collected"
    assert '"bottle"' in SPEC_TEXT, "pywebview's loopback HTTP server needs bottle"


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
