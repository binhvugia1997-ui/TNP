"""Run the React frontend inside a local pywebview desktop window.

This is the entry point the Windows Portable ``ReportExtractor.exe`` launches (PROMPT-028): the frozen
executable must open the SAME React application as ``python -m app.desktop``.

Frontend location is resolved for BOTH runtimes, because pywebview resolves a *relative* window URL
against its own application root, which differs between the two:

* source mode  – ``webview.util.get_app_root()`` == ``dirname(realpath(sys.argv[0]))`` == ``<repo>/app``,
  so the bundle is addressed as ``../frontend/dist/index.html``.
* frozen mode  – ``get_app_root()`` == ``sys._MEIPASS`` == ``<portable>/_internal``, so the very same
  file is addressed as ``frontend/dist/index.html`` (or ``../frontend/dist/index.html`` when the builder
  placed the bundle next to ``_internal``).

The relative form is required on purpose: pywebview serves it over a loopback HTTP server rooted at the
dist folder.  An absolute ``file://`` URL would break Chromium ES-module loading of ``assets/*``.
Nothing here depends on the current working directory or on the repository path (§16).
"""
from __future__ import annotations

import argparse
import logging
import os
import platform
import sys
from pathlib import Path
from typing import List, Optional

from . import APP_TITLE
from .application_service import ApplicationService
from .bridge import BridgeService
from .runtime_paths import is_packaged, portable_root, resource_root

LOG = logging.getLogger("report_extractor.webview.desktop")

#: Where the Vite bundle lives, relative to a root.  Kept as ONE definition so the packaged layout and
#: the source layout can never drift apart.
FRONTEND_DIST_PARTS = ("frontend", "dist")
FRONTEND_INDEX_NAME = "index.html"
#: pywebview resolves a relative window URL against its own app root; see the module docstring.
SOURCE_URL_RELATIVE_TO_APP_DIR = "../frontend/dist/index.html"


def assumed_app_root() -> Path:
    """The directory pywebview resolves relative window URLs against, computed deterministically.

    pywebview's own ``get_app_root()`` returns ``sys._MEIPASS`` when frozen and
    ``dirname(realpath(sys.argv[0]))`` otherwise.  For the two launch modes that matter those are:

    * ``ReportExtractor.exe``            → ``<portable>/_internal``   (``sys._MEIPASS``)
    * ``python -m app.desktop``          → ``<repo>/app``             (``sys.argv[0]`` is that module)

    ``sys.argv[0]`` is deliberately NOT read here: under pytest, an IDE runner or ``python -c`` it points at
    the launcher rather than the app, which would make this value (and the module constant below) depend on
    how the process happened to start.  :func:`runtime_app_root` re-checks against pywebview itself.
    """
    if is_packaged():
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def runtime_app_root() -> Path:
    """pywebview's REAL app root, used when actually creating the window.

    Prefers ``webview.util.get_app_root()`` so the URL handed to ``create_window`` can never disagree with
    the one pywebview resolves; falls back to :func:`assumed_app_root` when pywebview is not importable.
    """
    try:
        from webview.util import get_app_root          # noqa: PLC0415 – optional dependency
        return Path(get_app_root())
    except Exception:                                  # noqa: BLE001 – pywebview may be absent (tests, CLI)
        return assumed_app_root()


def frontend_dist_candidates() -> List[Path]:
    """Every location the built React bundle may legitimately occupy, most specific first.

    Packaged builds accept BOTH supported layouts: the bundle next to ``_internal`` (the documented
    Portable layout, easy for a tester to inspect) and the bundle inside ``_internal`` (PyInstaller
    ``datas``).  Source builds resolve to ``<repo>/frontend/dist``.
    """
    seen: List[Path] = []

    def add(path: Optional[Path]) -> None:
        if path is None:
            return
        try:
            resolved = Path(path).resolve()
        except OSError:
            return
        if resolved not in seen:
            seen.append(resolved)

    dist = Path(*FRONTEND_DIST_PARTS)
    if is_packaged():
        # A packaged exe resolves ONLY inside its own folder.  Falling back to a source checkout would
        # hide a broken package behind a developer machine that happens to have `frontend/dist` (§16).
        add(portable_root() / dist)                     # <portable>/frontend/dist  (documented layout)
        add(resource_root() / dist)                     # <portable>/_internal/frontend/dist (spec datas)
        add(resource_root().parent / dist)              # defensive: sibling of the bundle root
        return seen
    add(resource_root() / dist)                         # source mode: <repo>/frontend/dist
    add(Path(__file__).resolve().parent.parent / dist)  # defensive: next to the app/ package
    return seen


def frontend_dist_dir(require: bool = False) -> Optional[Path]:
    """The first candidate that actually contains ``index.html`` (None when the bundle is missing)."""
    for candidate in frontend_dist_candidates():
        if (candidate / FRONTEND_INDEX_NAME).is_file():
            return candidate
    if require:
        searched = "\n  ".join(str(c) for c in frontend_dist_candidates()) or "(none)"
        raise FileNotFoundError(
            "React bundle not found. Run `cd frontend && npm run build` first.\nSearched:\n  " + searched)
    return None


def frontend_url(dist_dir: Optional[Path] = None, root: Optional[Path] = None) -> str:
    """Relative window URL for pywebview, derived from the RESOLVED bundle directory.

    Never hardcoded per runtime: it is ``relpath(dist_dir, root)/index.html``, which yields
    ``../frontend/dist/index.html`` in source mode and ``frontend/dist/index.html`` when the bundle is
    inside ``_internal`` — so assets keep loading over pywebview's loopback HTTP server in both cases.
    ``root`` defaults to :func:`runtime_app_root` — pywebview's OWN root — because that is the directory
    pywebview will resolve the URL against.  Callers that need a value independent of how the process was
    launched (the module constant below, tests) pass :func:`assumed_app_root` explicitly.
    """
    dist_dir = Path(dist_dir) if dist_dir else frontend_dist_dir()
    if dist_dir is None:
        return SOURCE_URL_RELATIVE_TO_APP_DIR            # keep the historical, source-mode-correct value
    try:
        relative = os.path.relpath(dist_dir, Path(root) if root else runtime_app_root())
    except ValueError:                                   # different drives on Windows
        relative = str(dist_dir)
    url = Path(relative).as_posix() + "/" + FRONTEND_INDEX_NAME
    resolved = (Path(root) if root else runtime_app_root()) / url
    if Path(os.path.normpath(str(resolved))) != dist_dir / FRONTEND_INDEX_NAME:
        LOG.warning("WEBVIEW_FRONTEND_URL_MISMATCH root=%s url=%s resolved=%s expected=%s",
                    root, url, resolved, dist_dir / FRONTEND_INDEX_NAME)
    return url


def frontend_index(dist_dir: Optional[Path] = None) -> Optional[Path]:
    resolved = Path(dist_dir) if dist_dir else frontend_dist_dir()
    return (resolved / FRONTEND_INDEX_NAME) if resolved else None


# Import-time convenience values for the SOURCE runtime.  They stay module-level constants because other
# modules/tests reference them, but ``main()`` always re-resolves through the helpers above so a frozen
# executable is never tied to the repository layout.
FRONTEND_DIST = frontend_dist_dir() or (resource_root() / Path(*FRONTEND_DIST_PARTS))
FRONTEND_INDEX = FRONTEND_DIST / FRONTEND_INDEX_NAME
# Computed against assumed_app_root() so this constant is stable no matter how the process was launched
# (pytest, an IDE, `python -c`).  main() re-resolves against pywebview's real root before creating a window.
FRONTEND_URL = frontend_url(FRONTEND_DIST if FRONTEND_INDEX.is_file() else None, assumed_app_root())


def _module_version(name: str) -> str:
    """Best-effort version string for diagnostics; never raises and never imports heavy side effects."""
    try:
        from importlib import metadata
        return metadata.version(name)
    except Exception:                                    # noqa: BLE001
        try:
            module = __import__(name, fromlist=["__version__"])
            return str(getattr(module, "__version__", "unknown"))
        except Exception:                                # noqa: BLE001
            return "absent"


def pythonnet_runtime_dll() -> Optional[Path]:
    """Where pythonnet.load() will look for Python.Runtime.dll — computed the same way pythonnet does.

    ``pythonnet/__init__.py`` uses ``Path(__file__).parent / "runtime" / "Python.Runtime.dll"``, so in a
    frozen build this is ``_internal/pythonnet/runtime/Python.Runtime.dll``.  Reporting the exact path is
    what makes a packaging mistake diagnosable from ``logs/app.log`` alone.
    """
    try:
        import pythonnet                                 # noqa: PLC0415 – optional, Windows-only in practice
        return Path(pythonnet.__file__).resolve().parent / "runtime" / "Python.Runtime.dll"
    except Exception:                                    # noqa: BLE001
        return None


#: Python.Runtime.dll (pythonnet 3.0.5) targets .NETStandard,Version=v2.0 — see its deps.json.  Loading a
#: netstandard2.0 assembly into the .NET Framework CLR needs the `netstandard.dll` facade, which ships with
#: .NET Framework 4.7.2 == registry Release 461808.  Below that, Assembly.LoadFrom throws and
#: clr_loader/netfx.py surfaces it as "Failed to resolve Python.Runtime.Loader.Initialize".
DOTNET_FX_MIN_RELEASE = 461808
#: The registry key Windows uses to report the installed .NET Framework 4.x version.
DOTNET_FX_REG_KEY = r"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"


def dotnet_framework_release() -> Optional[int]:
    """Installed .NET Framework 4.x `Release` number, or None when it cannot be determined.

    Never raises: a missing/unreadable key must not be the reason the app fails to start.
    """
    if sys.platform != "win32":
        return None
    try:
        import winreg                                    # noqa: PLC0415 – Windows only
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, DOTNET_FX_REG_KEY) as key:
            value, _ = winreg.QueryValueEx(key, "Release")
            return int(value)
    except Exception:                                    # noqa: BLE001
        return None


def dotnet_framework_report() -> str:
    """Human-readable .NET Framework status, including whether it can host Python.Runtime.dll."""
    release = dotnet_framework_release()
    if sys.platform != "win32":
        return "not-applicable"
    if release is None:
        return "undetected(.NET Framework 4.x not found or registry unreadable)"
    return f"Release={release}({'ok' if release >= DOTNET_FX_MIN_RELEASE else 'TOO-OLD: need 4.7.2+/461808'})"


# --------------------------------------------------------------------------- Mark-of-the-Web (PROMPT-030)
# The packaged exe died with one opaque line:
#
#     RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize from
#                   H:\...\_internal\pythonnet\runtime\Python.Runtime.dll
#
# ``clr_loader/ffi/netfx.py`` declares exactly five native functions — pyclr_initialize,
# pyclr_create_appdomain, pyclr_get_function, pyclr_close_appdomain, pyclr_finalize — and NO error
# accessor, so whatever the CLR really threw is discarded and every distinct failure collapses into that
# string.  ``NetFx.__init__`` also stores whatever ``pyclr_create_appdomain`` returned without checking it
# for NULL — which, since clr-loader 0.2.10 ANSWERS an unnamed (default) request with NULL on purpose to
# mean "the root/current AppDomain" (ClrLoader.cs: ``_domains[0]``), is harmless for clr_loader itself and
# only became a defect when PROMPT-030's tooling read that NULL as failure (fixed in PROMPT-030S-R).
#
# The classes that a BUILD can settle are now settled there (``tools/build_portable.py`` proves the payload
# is pythonnet's own managed assembly — ``Python.Runtime 3.0.5.0``, IL-only/any-cpu, targeting
# ``.NETStandard,Version=v2.0`` — declaring ``Python.Runtime.Loader.Initialize`` as
# ``static int32 (native int, int32)``, which is precisely clr_loader's ``entry_point`` typedef, and that
# the packaged ``ClrLoader.dll`` is mixed-mode, matches the interpreter's architecture and exports all five
# ``pyclr_*`` symbols).  A wrong or renamed payload, a pythonnet↔clr-loader signature skew and an
# architecture mismatch therefore stop the build instead of the exe.
#
# What remains is a property of the machine the package LANDED on, and a leading CANDIDATE that fits the
# observed facts is a blocked assembly — a hypothesis, NOT a proven cause: no A/B evidence (same folder,
# blocked vs unblocked) was ever captured on the failing H:\\ machine, and PROMPT-030S-R neither confirms
# nor refutes it.  Windows attaches a ``Zone.Identifier`` alternate data stream when a folder is
# extracted from ``release/ReportExtractor_<version>.zip`` or copied from another PC or USB stick —
# i.e. AFTER the build gate ran.  ``LoadLibrary`` ignores that stream, so the native ``ClrLoader.dll``
# loads and clr_loader reaches the root AppDomain (its unnamed default, see ``appdomain_report``);
# ``Assembly.LoadFrom`` does NOT, so the managed ``Python.Runtime.dll`` WOULD be refused with
# ``COR_E_FILELOAD`` / 0x80131515 "Operation is not supported".  That profile matches a DLL which exists
# at the right path with the right SHA256 and still cannot initialize, and is why the same bytes work from
# ``.venv-build`` (pip never attaches the stream); whether it explains the original H:\\ death is
# unproven — which is exactly why the block is reported and cleared up front instead of assumed.
ZONE_IDENTIFIER_STREAM = "Zone.Identifier"


def _read_zone_identifier(path: Path) -> Optional[str]:
    """The ``Zone.Identifier`` stream of ``path``, or None when it is not blocked.

    Never raises: a filesystem without alternate data streams simply reports "not blocked".
    """
    try:
        with open(f"{path}:{ZONE_IDENTIFIER_STREAM}", "r", encoding="utf-8", errors="replace") as handle:
            return handle.read(400).strip() or "present"
    except OSError:
        return None


def _remove_zone_identifier(path: Path) -> bool:
    """Delete the ``Zone.Identifier`` stream — what Windows' "Unblock" checkbox and Unblock-File do."""
    try:
        os.remove(f"{path}:{ZONE_IDENTIFIER_STREAM}")
        return True
    except OSError:
        return False


def packaged_runtime_files() -> List[Path]:
    """The managed assembly and the native netfx host the frozen app must load — never anything outside it.

    Both live under ``sys._MEIPASS``: ``pythonnet`` computes ``<bundle>/pythonnet/runtime/Python.Runtime.dll``
    and ``clr_loader.ffi.load_netfx()`` computes ``<bundle>/clr_loader/ffi/dlls/<arch>/ClrLoader.dll``.
    """
    dll = pythonnet_runtime_dll()
    if dll is None:
        return []
    arch = "amd64" if sys.maxsize > 2 ** 32 else "x86"
    bundle = dll.parents[2]                     # …/pythonnet/runtime/Python.Runtime.dll → …/_internal
    return [dll, bundle / "clr_loader" / "ffi" / "dlls" / arch / "ClrLoader.dll"]


def unblock_packaged_runtime(remediate: bool = True) -> dict:
    """Report — and by default clear — a Mark-of-the-Web block on the packaged .NET runtime files.

    Only ever runs on Windows inside a frozen build, and only touches the two runtime files under
    ``sys._MEIPASS``; a development install gets files from pip, which never attaches the stream, and
    site-packages is left alone.  Clearing it is the same act as ticking "Unblock" on the folder, and
    without it the exe cannot start at all, so it is done up front and recorded in the log.  Every failure
    path degrades to a report: a read-only or locked file must never be the reason the app does not start.
    """
    report: dict = {"checked": 0, "blocked": [], "unblocked": [], "failed": []}
    if sys.platform != "win32" or not is_packaged():
        return report
    for path in packaged_runtime_files():
        try:
            present = path.is_file()
        except OSError:
            continue
        if not present:
            continue
        report["checked"] += 1
        zone = _read_zone_identifier(path)
        if zone is None:
            continue
        report["blocked"].append(path.name)
        if remediate and _remove_zone_identifier(path):
            report["unblocked"].append(path.name)
        elif remediate:
            report["failed"].append(path.name)
    return report


def appdomain_report() -> str:
    """What clr_loader REALLY holds as its .NET Framework AppDomain — with clr-loader's own semantics.

    ``pythonnet.get_runtime_info()`` cannot answer this: ``clr_loader/netfx.py::NetFx.info()`` returns
    ``RuntimeInfo(kind=".NET Framework", version="<undefined>", initialized=True, …)`` with
    ``initialized`` HARDCODED and no version at all.  A log line reading "Runtime: .NET Framework /
    Initialized: True" is therefore NOT evidence about the domain — the PROMPT-028R hint text asserted
    exactly that and misled the diagnosis.  Read the handle, but read it the way clr-loader 0.2.10 MEANS
    it (verified against the exact upstream tag, PROMPT-030S-R): ``get_netfx()`` defaults to
    ``domain=None``, ``NetFx.__init__`` hands ``ffi.NULL`` to ``pyclr_create_appdomain``, and
    ``ClrLoader.cs`` registers ``AppDomain.CurrentDomain`` as index 0 and DELIBERATELY answers an unnamed
    request with ``IntPtr.Zero``.  A NULL handle on the default path is therefore the ROOT/CURRENT
    AppDomain — a normal success value — and NOT "the CLR itself would not start"; inferring failure from
    it was PROMPT-030's own defect.  The distinct results:

      * ``no-runtime-selected``      – pythonnet never built a runtime: THAT is the startup-failure case;
      * ``not-a-netfx-runtime(…)``   – a coreclr/mono runtime, which has no ``_domain`` attribute;
      * ``root(…)``                  – the netfx root/current AppDomain (clr-loader's unnamed default);
      * ``created(…)``               – a named AppDomain (ClrLoader.cs gives those index ≥ 1);
      * ``NULL(…)``                  – a NAMED domain whose handle came back zero: a real creation
        failure — impossible in clr-loader 0.2.10's normal flow (named domains get an index or the
        constructor raises), so it stays reported as an anomaly;
      * ``unknown``                  – the diagnostic itself failed; diagnostics must never be the crash.
    """
    if sys.platform != "win32":
        return "not-applicable"
    try:
        import pythonnet                                 # noqa: PLC0415 – optional, Windows-only in practice
        runtime = getattr(pythonnet, "_RUNTIME", None)
        if runtime is None:
            return "no-runtime-selected"
        if not hasattr(runtime, "_domain"):
            return f"not-a-netfx-runtime({type(runtime).__name__})"
        domain = runtime._domain                         # noqa: SLF001 – the only place the truth lives
        name = getattr(runtime, "_domain_name", None)   # noqa: SLF001 – the domain NetFx was ASKED for
        try:
            from clr_loader.ffi import ffi               # noqa: PLC0415
            is_null = bool(domain == ffi.NULL)
        except Exception:                                # noqa: BLE001 – cffi missing must not hide the answer
            is_null = not domain
        if is_null:
            if name:
                return f"NULL(pyclr_create_appdomain failed for named domain {name!r})"
            return "root(AppDomain.CurrentDomain — clr-loader's unnamed default; a normal, successful value)"
        return f"created(named AppDomain {name!r})" if name else "created(named AppDomain)"
    except Exception:                                    # noqa: BLE001 – diagnostics must never be the crash
        return "unknown"


def runtime_diagnostics(probe_dotnet: bool = True, remediate_block: bool = True) -> dict:
    """Everything needed to explain a pywebview/.NET start-up failure, as safe short strings.

    Versions and the resolved DLL path only — no report data, no user documents.  Paths do appear here
    because this goes to ``logs/app.log`` and ``logs/startup_error.log``, never to the React UI.

    Has one deliberate side effect: ``unblock_packaged_runtime()`` clears a Mark-of-the-Web block on the
    packaged runtime files before anything tries to load them.  This function already probes ``import clr``
    (which loads the runtime), so the remediation happens here — before that probe — rather than after the
    failure it exists to prevent.  Pass ``remediate_block=False`` for a read-only report.
    """
    dll = pythonnet_runtime_dll()
    info = {
        "frozen": str(is_packaged()),
        "python": platform.python_version(),
        "platform": sys.platform,
        "pywebview": _module_version("pywebview"),
        "pythonnet": _module_version("pythonnet"),
        "clr_loader": _module_version("clr-loader"),
        "backend": "edgechromium" if sys.platform == "win32" else "auto",
        "python_runtime_dll": str(dll) if dll else "pythonnet-absent",
        "python_runtime_dll_exists": str(bool(dll and dll.is_file())),
    }
    # On Windows the default pythonnet runtime is .NET Framework ("netfx"); name it explicitly so a missing
    # or unsuitable .NET installation is obvious instead of surfacing as an opaque resolve error.
    if sys.platform == "win32":
        try:
            import pythonnet
            spec = pythonnet.get_runtime_info()
            info["dotnet_runtime"] = (spec.kind if spec else "not-loaded-yet")
        except Exception:                                # noqa: BLE001
            info["dotnet_runtime"] = "netfx(default-on-windows)"
        info["dotnet_framework"] = dotnet_framework_report()
        # Clear a Windows block BEFORE anything loads the assembly: LoadLibrary ignores Zone.Identifier
        # but Assembly.LoadFrom does not — the mechanism by which a DLL that exists and hashes correctly
        # CAN still fail to initialize (PROMPT-030's leading, unproven hypothesis for the original
        # H:\ death; the remediation is kept because it is safe and the block is directly observable).
        block = unblock_packaged_runtime(remediate=remediate_block)
        info["runtime_files_checked"] = str(block["checked"])
        info["blocked_runtime"] = ",".join(block["blocked"]) or "none"
        if block["unblocked"]:
            info["unblocked_runtime"] = ",".join(block["unblocked"])
        if block["failed"]:
            info["unblock_failed"] = ",".join(block["failed"])
        if probe_dotnet:
            try:
                import clr  # noqa: F401  – triggers pythonnet.load(); idempotent once loaded
                info["clr_import"] = "ok"
            except BaseException as exc:                 # noqa: BLE001 – the whole point is to capture it
                info["clr_import"] = f"{type(exc).__name__}"
                info["clr_import_error"] = _safe_error(exc)
            # Only recorded after the attempt, so it reflects the handle NetFx really holds; for the
            # default unnamed request that handle is NULL and means the ROOT/CURRENT AppDomain (see
            # appdomain_report) — it must never be read as a CLR startup failure.
            info["dotnet_appdomain"] = appdomain_report()
    return info


#: The two failure signatures that mean "the packaged .NET runtime is wrong", translated for a human.
_CLR_HINTS = (
    ("Failed to resolve Python.Runtime.Loader.Initialize",
     "clr_loader/netfx.py maps EVERY failure of the native pyclr_get_function() onto this one string and "
     "its cdef declares no error accessor, so the real .NET exception is discarded. It does NOT prove an "
     "app domain was created (NetFx.info() hardcodes initialized=True and version='<undefined>'), and it "
     "prints the path pythonnet ASKED for, not a path that exists. Read the same WEBVIEW_RUNTIME log line "
     "instead, in this order: (1) blocked_runtime=<name> — Windows attached a Zone.Identifier stream when "
     "this folder came out of a ZIP or from another PC/USB stick; LoadLibrary ignores it so ClrLoader.dll "
     "loads, Assembly.LoadFrom refuses it with 0x80131515. The app clears it automatically when it can "
     "(see unblocked_runtime / unblock_failed); otherwise right-click the Portable folder → Properties → "
     "Unblock, or Unblock-File on it. (2) dotnet_appdomain — root(AppDomain.CurrentDomain …) IS the "
     "normal value: clr-loader 0.2.10's default netfx path requests the UNNAMED domain and ClrLoader.cs "
     "answers it with its root-domain sentinel, so NULL here (or 'root(...)') is NOT a startup failure and "
     "says NOTHING about the .NET Framework version (that reading was PROMPT-030's bug, fixed in "
     "PROMPT-030S-R); created(...) is a named AppDomain, equally alive. A root/created value proves the "
     "native host loaded, so the refusal sits on the managed-load side — keep reading (1), (3) and (4). "
     "Only no-runtime-selected / unknown / not-a-netfx-runtime describe the runtime never starting. "
     "(3) dotnet_framework=Release=…(TOO-OLD) — Python.Runtime.dll "
     "targets .NETStandard 2.0 and needs Release >= 461808; install the .NET Framework 4.8 Runtime. "
     "(4) python_runtime_dll_exists=False, or a SHA256 that differs from the build venv — a packaging "
     "mistake; run `python tools/build_portable.py --validate-only <portable folder>` to see which. "
     "Assembly identity, architecture and the Initialize signature are proven at build time, so they are "
     "not candidate causes here."),
    ("Python.Runtime.dll not found",
     "pythonnet is not packaged correctly; _internal/pythonnet/runtime/Python.Runtime.dll is missing."),
    ("Could not find a suitable hostfxr",
     "clr_loader tried the .NET Core host; the packaged build should use .NET Framework (netfx) on Windows."),
    ("netstandard",
     "The .NET Framework on this PC is older than 4.7.2, so it cannot load the .NETStandard 2.0 assembly "
     "Python.Runtime.dll. Install the .NET Framework 4.8 Runtime and start again."),
    ("BadImageFormatException",
     "Python.Runtime.dll or ClrLoader.dll does not match this process' architecture (32- vs 64-bit)."),
)


def _safe_error(exc: BaseException, limit: int = 400) -> str:
    """Exception text for logs: type + message, truncated, with no embedded traceback objects."""
    text = f"{type(exc).__name__}: {exc}".replace("\n", " ")
    return text if len(text) <= limit else text[:limit] + "…"


def explain_clr_failure(text: str) -> str:
    """Actionable hint for a known .NET/pywebview packaging failure, or '' when the text is unfamiliar."""
    for needle, hint in _CLR_HINTS:
        if needle.lower() in text.lower():
            return hint
    return ""


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Run the local React + pywebview desktop application.")
    parser.add_argument("--debug", action="store_true", help="Enable pywebview developer tools and verbose logs.")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    dist_dir = frontend_dist_dir()
    index = frontend_index(dist_dir)
    if index is None or not index.is_file():
        try:
            frontend_dist_dir(require=True)
        except FileNotFoundError as exc:
            parser.error(str(exc))
        parser.error(f"React bundle not found: {index}")
    try:
        import webview
    except ImportError:
        parser.error("pywebview is missing. Install `requirements-webview.txt` into the development environment.")

    # PROMPT-028R: record the .NET/pywebview facts BEFORE anything can fail, so a packaged app that dies
    # while initialising the Windows backend explains itself in logs/app.log instead of vanishing.
    diagnostics = runtime_diagnostics()
    LOG.info("WEBVIEW_RUNTIME %s", " ".join(f"{k}={v}" for k, v in diagnostics.items()))
    if diagnostics.get("blocked_runtime", "none") != "none":
        LOG.warning("WEBVIEW_DOTNET_BLOCKED %s unblocked=%s failed=%s — Windows had marked the packaged "
                    ".NET runtime as having come from another computer; Assembly.LoadFrom refuses such a "
                    "file even though it exists at the right path with the right SHA256",
                    diagnostics.get("blocked_runtime"), diagnostics.get("unblocked_runtime", "none"),
                    diagnostics.get("unblock_failed", "none"))
    hint = explain_clr_failure(" ".join(str(v) for v in diagnostics.values()))
    if diagnostics.get("clr_import") not in (None, "ok") or diagnostics.get("python_runtime_dll_exists") == "False":
        LOG.error("WEBVIEW_DOTNET_PROBLEM %s%s", diagnostics.get("clr_import_error", ""),
                  (" | HINT: " + hint) if hint else "")

    # Resolve against pywebview's OWN app root (sys._MEIPASS when frozen) and log both, so a packaged app
    # that cannot find its bundle explains itself in logs/ instead of opening a blank window.
    app_root = runtime_app_root()
    url = frontend_url(dist_dir, app_root)
    LOG.info("WEBVIEW_FRONTEND packaged=%s app_root=%s dist=%s url=%s",
             is_packaged(), app_root, dist_dir, url)
    if not (Path(app_root) / url).resolve() == Path(dist_dir) / FRONTEND_INDEX_NAME:
        LOG.error("WEBVIEW_FRONTEND_MISMATCH pywebview would resolve %r to %s, expected %s",
                  url, (Path(app_root) / url), Path(dist_dir) / FRONTEND_INDEX_NAME)

    service = ApplicationService()
    bridge = BridgeService(service)
    window = webview.create_window(
        APP_TITLE,
        url=url,
        js_api=bridge,
        width=1440,
        height=940,
        min_size=(1000, 680),
        confirm_close=True,
        text_select=True,
    )
    bridge.bind_window(window)
    LOG.info("WEBVIEW_BRIDGE bound")

    def on_loaded(*_args):
        LOG.info("WEBVIEW_BRIDGE ready")

    window.events.loaded += on_loaded

    def on_closing(*_args):
        allowed, message = service.close_requested()
        if allowed:
            return True
        try:
            window.evaluate_js("window.dispatchEvent(new CustomEvent('tnp-close-requested', {detail: "
                               + __import__("json").dumps(message, ensure_ascii=False) + "}))")
        except Exception:
            LOG.debug("Could not notify React about a deferred close", exc_info=True)
        return False

    window.events.closing += on_closing
    try:
        webview.start(debug=args.debug, gui="edgechromium" if sys.platform == "win32" else None)
    except BaseException as exc:                         # noqa: BLE001 – re-raised with context attached
        detail = _safe_error(exc)
        advice = explain_clr_failure(detail)
        LOG.error("WEBVIEW_START_FAILED %s%s", detail, (" | HINT: " + advice) if advice else "")
        if advice:
            raise RuntimeError(f"{detail}\n\n{advice}") from exc
        raise
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
