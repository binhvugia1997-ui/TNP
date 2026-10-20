"""PROMPT-030 — the packaged .NET runtime must be PROVEN LOADABLE, not merely present.

The real Windows Portable died before its first window::

    RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize from:
    H:\\ReportExtractor_v1.3.4_Portable\\_internal\\pythonnet\\runtime\\Python.Runtime.dll

That string is the *only* thing ``clr_loader/netfx.py`` reports: its cdef declares
``pyclr_initialize`` / ``pyclr_create_appdomain`` / ``pyclr_get_function`` / ``pyclr_close_appdomain`` /
``pyclr_finalize`` and **no error accessor**, so whatever the CLR actually threw is discarded, and
``NetFx.__init__`` stores whatever ``pyclr_create_appdomain`` returned without checking it for NULL while
``NetFx.info()`` hardcodes ``initialized=True``.  PROMPT-030S-R fixes what that first PROMPT-030 gate then
got WRONG: in clr-loader 0.2.10 ``get_netfx()`` defaults to ``domain=None``, ``NetFx.__init__`` passes
``ffi.NULL``, and ``ClrLoader.cs`` deliberately answers the unnamed request with index 0 — the
ROOT/CURRENT ``AppDomain`` it registered in ``Initialize()``.  So the probe's ``appdomain NULL`` is the
root-domain SENTINEL, never by itself an AppDomain creation failure; success and failure are decided by
positive evidence (resolved pointer + ``Initialize()==0``, matching pythonnet's own ``load()`` contract)
and the gate fails CLOSED on missing or malformed probe output.

These tests pin what PROMPT-030 adds (as corrected by PROMPT-030R and PROMPT-030S-R):

* the packaged payload's **managed identity** — assembly name, IL-only/any-cpu, target framework, and a
  ``Python.Runtime.Loader.Initialize`` that is ``static int32 (native int, int32)``, which is precisely the
  ``entry_point`` typedef clr_loader turns into a delegate;
* the packaged netfx host's **architecture, mixed-mode-ness and five ``pyclr_*`` exports**;
* **Mark-of-the-Web reporting**, a candidate post-build cause (a folder extracted from ``release/*.zip``
  or copied from another PC): ``LoadLibrary`` ignores ``Zone.Identifier`` so the native host loads and
  clr_loader reaches the root AppDomain, while ``Assembly.LoadFrom`` can refuse the managed assembly with
  0x80131515.  This is an evidence-based HYPOTHESIS — no A/B proof exists that it caused the original
  H:\\ failure — so the tests pin the safe observe-and-unblock behavior, not a proven causal claim;
* a **real loadability probe** that performs the failing call chain against the packaged bytes, reads the
  app-domain handle that ``clr_loader`` itself never inspects, and treats its NULL root-domain sentinel as
  the NORMAL case while demanding positive evidence (``resolved True`` + ``initialize_rc == 0``) and
  failing CLOSED on missing, malformed, erroring, nonzero-exit or timed-out probe output;
* PROMPT-030T — that probe must run pythonnet's **whole** lifecycle, not just its start.  pythonnet 3.0.5's
  ``load()`` registers ``atexit(unload)`` the moment ``Loader.Initialize`` returns 0, ``unload()`` calls
  ``Loader.Shutdown(b"full_shutdown")`` before ``_RUNTIME.shutdown()``, and clr-loader's own ``atexit``
  ``_release()`` then calls ``pyclr_finalize()``.  Stopping after ``Initialize`` leaves
  ``AppDomain.CurrentDomain.ProcessExit`` subscribed and pythonnet's managed Finalizer queue full, so the
  CLR tears the runtime down out of order and ``Py.cs:41``'s ``~GILState()`` throws
  ``GIL must always be released…`` off the finalizer thread — unhandled on .NET Framework, i.e. a nonzero
  subprocess exit AFTER healthy-looking JSON was printed.  The probe therefore reports
  ``shutdown_resolved`` / ``shutdown_rc`` / ``finalized`` / ``stage`` and the gate demands all of them;
  the probe source is also EXECUTED here against a stub ``cffi`` so its stage/cleanup control flow is
  proven on any host instead of merely grepped;
* startup diagnostics that report the discriminating facts — root/current vs named AppDomain, blocked
  runtime, .NET Framework release — instead of asserting an unprovable one.

Nothing here needs Windows, a .NET installation, PyInstaller or a built artifact: the managed images are
synthesized from ECMA-335 by ``tests/dotnet_image_fixtures.py``, and the real shipped DLLs are asserted
against the same reader whenever ``pythonnet`` happens to be installed.
"""
from __future__ import annotations

import importlib.util
import json
import struct
import subprocess
import sys
import types
from pathlib import Path, PurePath

import pytest

import app.desktop as desktop
import build_portable as bp
import dotnet_image_fixtures as fx
import dotnet_pe
from build_portable import validate_pythonnet_runtime
from dotnet_image_fixtures import clr_loader_image, python_runtime_image

ROOT = Path(__file__).resolve().parents[1]
PACKAGED_DLL_REL = "_internal/pythonnet/runtime/Python.Runtime.dll"


def _build_arch() -> str:
    """The SAME rule validate_pythonnet_runtime() uses, so fixture and production cannot drift apart."""
    return "amd64" if struct.calcsize("P") * 8 > 32 else "x86"


def _clrloader_rel(arch: str | None = None) -> str:
    return f"_internal/clr_loader/ffi/dlls/{arch or _build_arch()}/ClrLoader.dll"


def _force_windows(monkeypatch):
    """Run the validator's Windows-only branches on any host, so these tests are not host-dependent."""
    monkeypatch.setattr(bp, "platform", types.SimpleNamespace(system=staticmethod(lambda: "Windows")))


def _package(tmp_path, runtime=python_runtime_image(), host=clr_loader_image(),
             runtime_rel=PACKAGED_DLL_REL, host_rel=None, extra=()):
    """A Portable folder laid out the way hook-clr + hook-clr_loader lay it out on Windows."""
    folder = tmp_path / "pkg"
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "ReportExtractor.exe").write_bytes(b"MZ")
    payloads = {}
    if runtime is not None:
        payloads[runtime_rel] = runtime
    payloads[host_rel or _clrloader_rel()] = host
    for rel, blob in list(payloads.items()) + list(extra):
        target = folder / PurePath(rel.replace("\\", "/"))
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(blob)
    return folder


def _real_dll(relative: str):
    """The shipped DLL from site-packages, or None when the package is not installed on this host."""
    for name in ("pythonnet", "clr_loader"):
        spec = importlib.util.find_spec(name)
        if spec and spec.origin:
            candidate = Path(spec.origin).parent.parent / relative
            if candidate.is_file():
                return candidate
    return None


def _validate_only_interpreter() -> Path:
    """Mirror validate_existing_package()'s preference for the build venv on Windows."""
    candidate = ROOT / ".venv-build" / "Scripts" / "python.exe"
    return candidate if candidate.is_file() else Path(sys.executable)


def _package_from_installed_runtime(tmp_path, py: Path):
    """Stage exact DLL bytes discoverable by the same interpreter used for the live probe."""
    arch = _build_arch()
    runtime = bp.source_pythonnet_dll(py)
    loader_code = (
        "import importlib.util, json\n"
        "from pathlib import Path\n"
        "spec = importlib.util.find_spec('clr_loader')\n"
        "host = (Path(spec.origin).parent / 'ffi' / 'dlls' / "
        f"{arch!r} / 'ClrLoader.dll') if spec and spec.origin else None\n"
        "print(json.dumps({'found': bool(host and host.is_file()), "
        "'path': str(host) if host else None, "
        "'cffi': importlib.util.find_spec('cffi') is not None}))\n"
    )
    loader = bp._query(py, loader_code)
    missing = []
    if not runtime.get("found"):
        missing.append("pythonnet/runtime/Python.Runtime.dll")
    if not loader.get("found"):
        missing.append(f"clr_loader/ffi/dlls/{arch}/ClrLoader.dll")
    if not loader.get("cffi"):
        missing.append("cffi in the live-probe interpreter")
    if missing:
        pytest.fail("Windows live-runtime integration prerequisites are missing from "
                    f"{py}: {', '.join(missing)}")

    return _package(tmp_path,
                    runtime=Path(runtime["path"]).read_bytes(),
                    host=Path(loader["path"]).read_bytes())


# =========================================================================== the ECMA-335 reader itself
def test_reader_rejects_a_file_that_is_not_a_pe(tmp_path):
    target = tmp_path / "Python.Runtime.dll"
    target.write_bytes(b"MANAGED-ASSEMBLY")                    # exactly the PROMPT-028R placeholder
    with pytest.raises(dotnet_pe.DotNetPEError) as error:
        dotnet_pe.inspect(target)
    assert "not a PE image" in str(error.value)


def test_reader_rejects_a_native_dll_with_no_cli_header(tmp_path):
    """A native payload renamed to Python.Runtime.dll must be named as "not a .NET assembly"."""
    image = bytearray(python_runtime_image())
    pe = struct.unpack_from("<I", bytes(image), 0x3C)[0]
    magic = struct.unpack_from("<H", bytes(image), pe + 24)[0]
    directory = pe + 24 + (112 if magic == 0x20B else 96) + 14 * 8
    struct.pack_into("<II", image, directory, 0, 0)             # erase the CLI header directory
    target = tmp_path / "Python.Runtime.dll"
    target.write_bytes(bytes(image))
    with pytest.raises(dotnet_pe.DotNetPEError) as error:
        dotnet_pe.inspect(target)
    assert "no CLI header" in str(error.value)


def test_reader_rejects_a_missing_file(tmp_path):
    with pytest.raises(dotnet_pe.DotNetPEError) as error:
        dotnet_pe.inspect(tmp_path / "absent.dll")
    assert "cannot read" in str(error.value)


def test_reader_rejects_truncated_metadata(tmp_path):
    target = tmp_path / "Python.Runtime.dll"
    target.write_bytes(python_runtime_image()[:900])
    with pytest.raises(dotnet_pe.DotNetPEError):
        dotnet_pe.inspect(target)


def test_reader_decodes_compressed_integers_per_ecma_335():
    for blob, expected in ((b"\x03", 3), (b"\x7f", 127), (b"\x81\x00", 256),
                           (b"\xbf\xff", 16383), (b"\xc1\x00\x00\x00", 16777216)):
        assert dotnet_pe._Cursor(blob).compressed() == expected


def test_reader_decodes_the_signatures_clr_loader_cares_about():
    decode = dotnet_pe._decode_method_signature
    assert decode(fx._signature_blob("int32 (native int, int32)")) == "int32 (native int, int32)"
    assert decode(fx._signature_blob("int32 (string)")) == "int32 (string)"
    assert decode(fx._signature_blob("void ()")) == "void ()"
    assert decode(fx._signature_blob("void (string)", has_this=True)) == "void (string)"


def test_reader_derives_the_public_key_token_the_way_dotnet_does(tmp_path):
    """A strong-named assembly reports an 8-byte token: SHA1(public key)[-8:] reversed."""
    image = dotnet_pe.inspect(_write(tmp_path / "signed.dll", python_runtime_image(strong_name=True)))
    assert image.strong_name_signed
    assert len(image.assembly.public_key_token) == 16
    plain = dotnet_pe.inspect(_write(tmp_path / "unsigned.dll", python_runtime_image(strong_name=False)))
    assert plain.assembly.public_key_token == ""


def _write(path: Path, blob: bytes) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(blob)
    return path


# =============================================== measured facts about the DLLs this build really ships
REAL_RUNTIME = _real_dll("pythonnet/runtime/Python.Runtime.dll")
REAL_HOST_AMD64 = _real_dll("clr_loader/ffi/dlls/amd64/ClrLoader.dll")
REAL_HOST_X86 = _real_dll("clr_loader/ffi/dlls/x86/ClrLoader.dll")


@pytest.mark.skipif(REAL_RUNTIME is None, reason="pythonnet is not installed on this host")
def test_real_pythonnet_runtime_declares_exactly_what_clr_loader_binds():
    """The measured facts that eliminate whole classes of cause for the PROMPT-030 crash.

    Recorded so the conclusion survives without the package installed: Python.Runtime.dll is pythonnet's
    own strong-named assembly, it is IL-only/any-cpu (so an architecture mismatch is impossible), it
    targets .NETStandard 2.0 (so .NET Framework 4.7.2+ is required), and its Initialize is static with the
    signature clr_loader's entry_point typedef demands (so a pythonnet↔clr-loader skew is impossible).
    """
    image = dotnet_pe.inspect(REAL_RUNTIME)
    assert image.is_managed and image.is_dll
    assert image.assembly.name == "Python.Runtime"
    assert image.assembly.version[:3] == (3, 0, 5)
    assert image.assembly.public_key_token, "pythonnet's runtime is strong-named"
    assert image.il_only and not image.requires_32bit
    assert image.arch_text == "any-cpu"
    assert image.target_frameworks == (".NETStandard,Version=v2.0",)
    assert "netstandard" in [ref.name for ref in image.assembly_refs]
    assert image.has_type(dotnet_pe.ENTRY_POINT_TYPE)
    ok, why = dotnet_pe.signature_matches_entry_point(image)
    assert ok, why
    method = image.find_method(dotnet_pe.ENTRY_POINT_TYPE, dotnet_pe.ENTRY_POINT_METHOD)
    assert method.is_static and method.signature == dotnet_pe.ENTRY_POINT_SIGNATURE
    assert bp.required_netfx_release(image) == (461808, image.target_frameworks[0]
                                                + " → .NET Framework Release ≥ 461808 (4.7.2+)")


@pytest.mark.skipif(REAL_HOST_AMD64 is None or REAL_HOST_X86 is None,
                    reason="clr_loader is not installed on this host")
def test_real_clr_loader_host_is_mixed_mode_and_exports_the_five_pyclr_symbols():
    for arch, path in (("amd64", REAL_HOST_AMD64), ("x86", REAL_HOST_X86)):
        image = dotnet_pe.inspect(path)
        assert image.assembly.name == "ClrLoader"
        assert image.machine == arch, f"{path} must be the {arch} host"
        assert not image.il_only, "cffi dlopen()s it, so it must be mixed-mode C++/CLI"
        assert image.target_frameworks == (".NETFramework,Version=v4.7.2",)
        assert set(image.exports) == set(bp.PYCLR_EXPORTS), image.exports
    # The two architectures are NOT interchangeable — which is why the gate proves the right one shipped.
    assert dotnet_pe.inspect(REAL_HOST_X86).requires_32bit
    assert not dotnet_pe.inspect(REAL_HOST_AMD64).requires_32bit


def test_pyclr_exports_match_the_cdef_clr_loader_actually_declares():
    """The probe and the gate both assume these five names; they come from clr_loader/ffi/netfx.py."""
    spec = importlib.util.find_spec("clr_loader")
    if spec is None:
        pytest.skip("clr_loader is not installed on this host")
    cdef = (Path(spec.origin).parent / "ffi" / "netfx.py").read_text(encoding="utf-8")
    for name in bp.PYCLR_EXPORTS:
        assert name in cdef, f"{name} is declared by clr_loader's cdef and must stay in PYCLR_EXPORTS"


# =========================================================================== the build gate: identity
def test_gate_accepts_a_faithful_hook_layout(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    assert validate_pythonnet_runtime(_package(tmp_path), {"found": False}) == []


def test_gate_rejects_a_managed_assembly_that_is_not_pythonnet(tmp_path, monkeypatch):
    """A correct-looking .NET assembly under the right name in the right place is STILL the wrong payload.

    "Do not assume matching filename means correct DLL": presence, uniqueness, location and SHA256 all
    agree here, and only the assembly identity says otherwise.
    """
    _force_windows(monkeypatch)
    impostor = python_runtime_image(assembly_name="Something.Else")
    problems = validate_pythonnet_runtime(_package(tmp_path, runtime=impostor), {"found": False})
    assert any("danh tính assembly" in p and "Something.Else" in p for p in problems), problems


def test_gate_rejects_a_runtime_that_does_not_declare_the_loader_type(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    problems = validate_pythonnet_runtime(
        _package(tmp_path, runtime=python_runtime_image(include_type=False)), {"found": False})
    assert any("không phân giải được entry point" in p and "Python.Runtime.Loader" in p
               for p in problems), problems


def test_gate_rejects_a_runtime_whose_loader_has_no_initialize(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    problems = validate_pythonnet_runtime(
        _package(tmp_path, runtime=python_runtime_image(include_method=False)), {"found": False})
    assert any("declares no method named Initialize" in p for p in problems), problems


def test_gate_rejects_an_entry_point_clr_loader_cannot_turn_into_a_delegate(tmp_path, monkeypatch):
    """clr_loader/types.py requires ``static int Func(IntPtr, int)``; anything else yields a NULL handle."""
    _force_windows(monkeypatch)
    cases = (
        (dict(method_flags=fx.METHOD_PUBLIC | fx.METHOD_HIDE_BY_SIG), "not static"),
        (dict(signature="int32 (string)"), "clr_loader needs"),
        (dict(signature="void (native int, int32)"), "clr_loader needs"),
        (dict(method_name="Init"), "declares no method named Initialize"),
    )
    for index, (override, expected) in enumerate(cases):
        problems = validate_pythonnet_runtime(
            _package(tmp_path / f"case{index}", runtime=python_runtime_image(**override)),
            {"found": False})
        assert any(expected in p for p in problems), (override, problems)


def test_gate_rejects_a_runtime_that_is_not_any_cpu(tmp_path, monkeypatch):
    """The real payload is IL-only/any-cpu; a 32-bit-required one would clash with a 64-bit exe."""
    _force_windows(monkeypatch)
    problems = validate_pythonnet_runtime(
        _package(tmp_path, runtime=python_runtime_image(il_only=False, requires_32bit=True)),
        {"found": False})
    assert any("IL-only/any-cpu" in p for p in problems), problems


def test_gate_rejects_a_runtime_no_netfx_can_host(tmp_path, monkeypatch):
    """.NETStandard 2.1 is not supported by any .NET Framework release — netfx could never load it."""
    _force_windows(monkeypatch)
    problems = validate_pythonnet_runtime(
        _package(tmp_path, runtime=python_runtime_image(target_framework=".NETStandard,Version=v2.1")),
        {"found": False})
    assert any("không bao giờ nạp được" in p for p in problems), problems


def test_gate_reports_the_required_netfx_release_it_derived(tmp_path, monkeypatch, capsys):
    _force_windows(monkeypatch)
    monkeypatch.setattr(bp, "dotnet_framework_release_of_build_machine", lambda: 533320)
    assert validate_pythonnet_runtime(_package(tmp_path), {"found": False}) == []
    logged = capsys.readouterr().out
    assert "461808" in logged and "Release=533320" in logged


def test_gate_fails_when_the_build_machine_netfx_is_too_old(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    monkeypatch.setattr(bp, "dotnet_framework_release_of_build_machine", lambda: 460798)   # 4.7.0
    problems = validate_pythonnet_runtime(_package(tmp_path), {"found": False})
    assert any("Release ≥ 461808" in p and "460798" in p for p in problems), problems


def test_gate_survives_an_unreadable_registry(monkeypatch):
    """An undetectable .NET Framework must warn nobody and block nothing — it is not a packaging defect."""
    monkeypatch.setattr(bp, "platform", types.SimpleNamespace(system=staticmethod(lambda: "Linux")))
    assert bp.dotnet_framework_release_of_build_machine() is None


def test_required_netfx_release_reads_a_netframework_target_too():
    image = dotnet_pe.ManagedImage(
        path="x", size=0, sha256="", machine="amd64", pe32_plus=True, is_dll=True, is_managed=True,
        cli_runtime_version="v4.0.30319", cli_flags=0, il_only=False, requires_32bit=False,
        strong_name_signed=False, exports=(), assembly=None, assembly_refs=(),
        target_frameworks=(".NETFramework,Version=v4.8",), type_names=(), methods={})
    assert bp.required_netfx_release(image)[0] == 528040


# ==================================================================== the build gate: the netfx host
def test_gate_rejects_a_clrloader_of_the_wrong_architecture(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    other = "x86" if _build_arch() == "amd64" else "amd64"
    problems = validate_pythonnet_runtime(
        _package(tmp_path, host=clr_loader_image(other)), {"found": False})
    assert any(f"thực chất là {other}" in p for p in problems), problems


def test_gate_rejects_a_clrloader_that_cannot_serve_the_cdef(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    problems = validate_pythonnet_runtime(
        _package(tmp_path, host=clr_loader_image(exports=("pyclr_initialize",))), {"found": False})
    joined = " ".join(problems)
    assert "không export" in joined and "pyclr_get_function" in joined, problems


def test_gate_rejects_an_il_only_clrloader(tmp_path, monkeypatch):
    """cffi dlopen()s the host, so it must be mixed-mode; an IL-only image has no native exports at all."""
    _force_windows(monkeypatch)
    problems = validate_pythonnet_runtime(
        _package(tmp_path, host=clr_loader_image(il_only=True, exports=())), {"found": False})
    assert any("mixed-mode" in p for p in problems), problems
    assert any("không export" in p for p in problems), problems


def test_gate_rejects_a_clrloader_that_is_not_a_pe(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    problems = validate_pythonnet_runtime(_package(tmp_path, host=b"NATIVE-HOST"), {"found": False})
    assert any("ClrLoader.dll" in p and "không đọc được" in p for p in problems), problems


def test_gate_still_accepts_both_architectures_because_the_hook_ships_both(tmp_path, monkeypatch):
    """hook-clr_loader collects amd64 AND x86; the superset must not be mistaken for a conflict."""
    _force_windows(monkeypatch)
    folder = _package(tmp_path, extra=[(_clrloader_rel("amd64"), clr_loader_image("amd64")),
                                       (_clrloader_rel("x86"), clr_loader_image("x86"))])
    assert validate_pythonnet_runtime(folder, {"found": False}) == []


# =============================================== duplicates: the stray copy must never silently win
def test_gate_proves_only_the_copy_pythonnet_will_actually_load(tmp_path, monkeypatch):
    """A stray second copy is reported, and the identity proof still runs on the one at the real path.

    This is the "no conflicting/duplicate runtime DLL wins unexpectedly" requirement: if the stray copy
    were the impostor, the gate must say so rather than quietly proving the good one.
    """
    _force_windows(monkeypatch)
    folder = _package(tmp_path, extra=[("_internal/Python.Runtime.dll", python_runtime_image())])
    problems = validate_pythonnet_runtime(folder, {"found": False})
    assert any("phải đúng 1" in p for p in problems), problems
    assert any("sai vị trí" in p and "_internal/Python.Runtime.dll" in p for p in problems), problems
    assert not any("danh tính assembly" in p for p in problems), "the real path holds a genuine runtime"


def test_gate_rejects_an_impostor_at_the_expected_path_even_when_a_good_copy_exists(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    folder = _package(tmp_path, runtime=python_runtime_image(assembly_name="Not.PythonNet"),
                      extra=[("_internal/spare/Python.Runtime.dll", python_runtime_image())])
    problems = validate_pythonnet_runtime(folder, {"found": False})
    assert any("danh tính assembly" in p for p in problems), problems


# =========================================================================== Mark-of-the-Web blocking
def test_mark_of_the_web_is_absent_for_an_ordinary_file(tmp_path):
    target = tmp_path / "Python.Runtime.dll"
    target.write_bytes(python_runtime_image())
    assert bp.mark_of_the_web(target) is None
    assert bp.mark_of_the_web(tmp_path / "absent.dll") is None


def test_gate_reports_a_blocked_runtime_dll(tmp_path, monkeypatch):
    """A blocked file exists at the right path with the right SHA256 — and still cannot be loaded."""
    _force_windows(monkeypatch)
    folder = _package(tmp_path)
    monkeypatch.setattr(bp, "mark_of_the_web",
                        lambda path: "[ZoneTransfer]\r\nZoneId=3" if path.name == "Python.Runtime.dll"
                        else None)
    problems = validate_pythonnet_runtime(folder, {"found": False})
    assert any("BỊ WINDOWS CHẶN" in p and "Unblock" in p for p in problems), problems


def test_gate_reports_a_blocked_netfx_host_too(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    folder = _package(tmp_path)
    monkeypatch.setattr(bp, "mark_of_the_web",
                        lambda path: "present" if path.name == "ClrLoader.dll" else None)
    problems = validate_pythonnet_runtime(folder, {"found": False})
    assert any("ClrLoader.dll" in p and "BỊ WINDOWS CHẶN" in p for p in problems), problems


def test_startup_unblock_only_runs_inside_a_frozen_windows_build(monkeypatch, tmp_path):
    """A development install gets its files from pip, which never attaches the stream — leave it alone."""
    calls = []
    runtime = tmp_path / "Python.Runtime.dll"
    runtime.write_bytes(python_runtime_image())      # the guard is is_file(), so the file must be real
    monkeypatch.setattr(desktop, "packaged_runtime_files", lambda: [runtime])
    monkeypatch.setattr(desktop, "_read_zone_identifier", lambda p: calls.append(p) or "present")
    for platform_name, packaged in (("win32", False), ("linux", True), ("win32", True)):
        monkeypatch.setattr(desktop.sys, "platform", platform_name)
        monkeypatch.setattr(desktop, "is_packaged", lambda: packaged)
        calls.clear()
        report = desktop.unblock_packaged_runtime()
        if platform_name == "win32" and packaged:
            assert report["blocked"] == ["Python.Runtime.dll"]
        else:
            assert report == {"checked": 0, "blocked": [], "unblocked": [], "failed": []}
            assert calls == [], "must not read anything outside a frozen Windows build"


def test_startup_unblock_clears_the_stream_and_reports_it(monkeypatch, tmp_path):
    runtime = tmp_path / "Python.Runtime.dll"
    runtime.write_bytes(python_runtime_image())
    removed = []
    monkeypatch.setattr(desktop.sys, "platform", "win32")
    monkeypatch.setattr(desktop, "is_packaged", lambda: True)
    monkeypatch.setattr(desktop, "packaged_runtime_files", lambda: [runtime])
    monkeypatch.setattr(desktop, "_read_zone_identifier", lambda p: "[ZoneTransfer]\r\nZoneId=3")
    monkeypatch.setattr(desktop, "_remove_zone_identifier", lambda p: removed.append(p) or True)
    report = desktop.unblock_packaged_runtime()
    assert report["checked"] == 1
    assert report["blocked"] == ["Python.Runtime.dll"]
    assert report["unblocked"] == ["Python.Runtime.dll"]
    assert report["failed"] == []
    assert removed == [runtime]


def test_startup_unblock_degrades_to_a_report_when_the_stream_cannot_be_removed(monkeypatch, tmp_path):
    """Read-only or locked files must never be the reason the app does not start."""
    runtime = tmp_path / "Python.Runtime.dll"
    runtime.write_bytes(python_runtime_image())
    monkeypatch.setattr(desktop.sys, "platform", "win32")
    monkeypatch.setattr(desktop, "is_packaged", lambda: True)
    monkeypatch.setattr(desktop, "packaged_runtime_files", lambda: [runtime])
    monkeypatch.setattr(desktop, "_read_zone_identifier", lambda p: "present")
    monkeypatch.setattr(desktop, "_remove_zone_identifier", lambda p: False)
    report = desktop.unblock_packaged_runtime()
    assert report["blocked"] == ["Python.Runtime.dll"]
    assert report["unblocked"] == []
    assert report["failed"] == ["Python.Runtime.dll"]


def test_startup_unblock_leaves_an_unblocked_package_alone(monkeypatch, tmp_path):
    runtime = tmp_path / "Python.Runtime.dll"
    runtime.write_bytes(python_runtime_image())
    monkeypatch.setattr(desktop.sys, "platform", "win32")
    monkeypatch.setattr(desktop, "is_packaged", lambda: True)
    monkeypatch.setattr(desktop, "packaged_runtime_files", lambda: [runtime])
    monkeypatch.setattr(desktop, "_read_zone_identifier", lambda p: None)
    report = desktop.unblock_packaged_runtime(remediate=True)
    assert report == {"checked": 1, "blocked": [], "unblocked": [], "failed": []}


def test_packaged_runtime_files_names_both_files_the_frozen_app_loads(monkeypatch, tmp_path):
    """pythonnet computes one path, clr_loader.ffi.load_netfx() computes the other; both must be covered."""
    bundle = tmp_path / "_internal"
    dll = bundle / "pythonnet" / "runtime" / "Python.Runtime.dll"
    monkeypatch.setattr(desktop, "pythonnet_runtime_dll", lambda: dll)
    monkeypatch.setattr(desktop.sys, "maxsize", 2 ** 63 - 1)
    files = desktop.packaged_runtime_files()
    assert files == [dll, bundle / "clr_loader" / "ffi" / "dlls" / "amd64" / "ClrLoader.dll"]
    monkeypatch.setattr(desktop, "pythonnet_runtime_dll", lambda: None)
    assert desktop.packaged_runtime_files() == []


# ================================================================== diagnostics must not lie about CLR
def _fake_pythonnet(monkeypatch, domain, domain_name=None, with_name=False):
    """A stand-in for ``pythonnet._RUNTIME`` shaped like clr_loader 0.2.10's ``NetFx``.

    ``NetFx`` stores BOTH the handle (``_domain``) and what it was asked to create (``_domain_name``);
    ``with_name`` reproduces the named-domain request that turns a NULL handle into a genuine failure.
    """
    if domain == "no-runtime":
        runtime = None
    else:
        attrs = {"_domain": domain}
        if with_name:
            attrs["_domain_name"] = domain_name
        runtime = types.SimpleNamespace(**attrs)
    monkeypatch.setitem(sys.modules, "pythonnet", types.SimpleNamespace(_RUNTIME=runtime))


def test_appdomain_report_reads_the_handle_instead_of_trusting_runtime_info(monkeypatch):
    """NetFx.info() hardcodes initialized=True, so get_runtime_info() cannot answer this question.

    PROMPT-030S-R corrected the VALUES reported from the handle: the NULL/unnamed combination is
    clr-loader's ROOT/CURRENT-domain success sentinel — never "the CLR itself would not start" — while a
    non-NULL handle still names the created domain and only 'no-runtime-selected' marks a real startup gap.
    """
    null = object()
    monkeypatch.setattr(desktop.sys, "platform", "win32")
    monkeypatch.setitem(sys.modules, "clr_loader", types.ModuleType("clr_loader"))
    monkeypatch.setitem(sys.modules, "clr_loader.ffi", types.SimpleNamespace(
        ffi=types.SimpleNamespace(NULL=null)))
    _fake_pythonnet(monkeypatch, 0x1234)
    assert desktop.appdomain_report() == "created(named AppDomain)"
    _fake_pythonnet(monkeypatch, 0x1234, with_name=True, domain_name="TNP-ReportExtractor")
    assert desktop.appdomain_report() == "created(named AppDomain 'TNP-ReportExtractor')"
    _fake_pythonnet(monkeypatch, null)
    report = desktop.appdomain_report()
    assert report.startswith("root("), report
    assert "AppDomain.CurrentDomain" in report
    for forbidden in ("failed", "would not start", "NULL("):
        assert forbidden not in report, f"root-domain NULL must not be reported as {forbidden!r}"
    _fake_pythonnet(monkeypatch, null, with_name=True, domain_name="TNP-ReportExtractor")
    failure = desktop.appdomain_report()
    assert failure.startswith("NULL(") and "named domain" in failure, \
        "a NAMED request whose handle came back zero is a real anomaly and stays reported as one"
    _fake_pythonnet(monkeypatch, "no-runtime")
    assert desktop.appdomain_report() == "no-runtime-selected"
    monkeypatch.setitem(sys.modules, "pythonnet", types.SimpleNamespace(
        _RUNTIME=types.SimpleNamespace()))                 # a coreclr/mono runtime has no _domain
    assert desktop.appdomain_report().startswith("not-a-netfx-runtime")
    monkeypatch.setattr(desktop.sys, "platform", "linux")
    assert desktop.appdomain_report() == "not-applicable"


def test_appdomain_report_falls_back_when_cffi_is_unavailable(monkeypatch):
    monkeypatch.setattr(desktop.sys, "platform", "win32")
    monkeypatch.setitem(sys.modules, "clr_loader", types.ModuleType("clr_loader"))
    monkeypatch.delitem(sys.modules, "clr_loader.ffi", raising=False)
    sys.modules.pop("clr_loader.ffi", None)
    _fake_pythonnet(monkeypatch, 0)
    assert desktop.appdomain_report().startswith("root(")


def test_appdomain_report_never_raises_when_pythonnet_is_absent(monkeypatch):
    monkeypatch.setattr(desktop.sys, "platform", "win32")
    # sys.modules[name] is None makes `import name` raise ImportError — no import hook needed.
    monkeypatch.setitem(sys.modules, "pythonnet", None)
    assert desktop.appdomain_report() == "unknown"


def test_diagnostics_report_the_block_and_the_domain(monkeypatch, tmp_path):
    """The log line a support engineer reads must carry the fields the hint tells them to look at."""
    runtime = tmp_path / "Python.Runtime.dll"
    runtime.write_bytes(python_runtime_image())
    monkeypatch.setattr(desktop.sys, "platform", "win32")
    monkeypatch.setattr(desktop, "is_packaged", lambda: True)
    monkeypatch.setattr(desktop, "packaged_runtime_files", lambda: [runtime])
    monkeypatch.setattr(desktop, "_read_zone_identifier", lambda p: "present")
    monkeypatch.setattr(desktop, "_remove_zone_identifier", lambda p: True)
    monkeypatch.setattr(desktop, "appdomain_report", lambda: "created")
    info = desktop.runtime_diagnostics(probe_dotnet=False)
    assert info["blocked_runtime"] == "Python.Runtime.dll"
    assert info["unblocked_runtime"] == "Python.Runtime.dll"
    assert info["runtime_files_checked"] == "1"
    assert "dotnet_appdomain" not in info, "the domain is only meaningful after a load attempt"
    assert info["dotnet_framework"].startswith("Release=") or "undetected" in info["dotnet_framework"] \
        or info["dotnet_framework"] == "not-applicable"


def test_hint_no_longer_asserts_an_app_domain_was_created():
    """PROMPT-028R's wording stated as fact something clr_loader cannot know; that misdirected the search."""
    hint = desktop.explain_clr_failure(
        "RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize from X")
    assert hint
    assert "clr_loader created a .NET Framework app domain but" not in hint
    assert "0x80131515" in hint, "must name the HRESULT a blocked assembly produces"
    assert "Zone.Identifier" in hint
    for field in ("blocked_runtime", "dotnet_appdomain", "dotnet_framework",
                  "python_runtime_dll_exists", "--validate-only"):
        assert field in hint, f"the hint must point at {field}"


def test_hint_does_not_misread_the_root_domain_sentinel_as_a_CLR_failure():
    """PROMPT-030's other half: the hint told readers dotnet_appdomain=NULL meant the CLR would not
    start / .NET Framework too old.  clr-loader 0.2.10 uses NULL to MEAN the root/current AppDomain,
    so the hint must teach the real discriminator instead of repeating that false inference."""
    hint = desktop.explain_clr_failure(
        "RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize from X")
    assert "the CLR itself would not start, i.e. .NET Framework older than 4.7.2" not in hint
    assert "root(AppDomain.CurrentDomain" in hint, "must name the normal default value"
    assert "NOT a startup failure" in hint
    assert "no-runtime-selected" in hint, "must point at the value that DOES mark a startup failure"
    assert "says NOTHING about the .NET Framework version" in hint, \
        "must stop inferring the framework age from the domain handle alone"


def test_hint_still_covers_the_older_signatures():
    assert "4.7.2" in desktop.explain_clr_failure("Could not load file or assembly 'netstandard, Version=2.0.0.0'")
    assert desktop.explain_clr_failure("System.BadImageFormatException: bad")
    assert desktop.explain_clr_failure("Could not find a suitable hostfxr library in X")
    assert desktop.explain_clr_failure("Python.Runtime.dll not found")
    assert desktop.explain_clr_failure("ValueError: unrelated") == ""


# ============================================================================== the loadability probe
def test_probe_source_is_valid_python_and_targets_the_packaged_paths():
    """The probe must exercise the PACKAGE, not the build venv — otherwise it proves nothing new."""
    source = bp.PROBE_PACKAGED_RUNTIME
    compile(source, "<probe>", "exec")
    assert '"_internal" / "clr_loader" / "ffi" / "dlls"' in source
    assert '"_internal" / "pythonnet" / "runtime" / "Python.Runtime.dll"' in source
    assert "site-packages" not in source and "importlib.util.find_spec" not in source


def test_probe_declares_the_same_native_interface_clr_loader_declares():
    """A drifted cdef would make the probe pass while the real app still fails."""
    spec = importlib.util.find_spec("clr_loader")
    if spec is None:
        pytest.skip("clr_loader is not installed on this host")
    cdef = (Path(spec.origin).parent / "ffi" / "netfx.py").read_text(encoding="utf-8")
    for name in bp.PYCLR_EXPORTS:
        assert name in bp.PROBE_PACKAGED_RUNTIME, f"the probe must call {name}"
        assert name in cdef
    assert "pyclr_create_appdomain" in cdef and "pyclr_get_function" in cdef


def test_probe_checks_the_domain_handle_that_clr_loader_never_reads():
    """NetFx.__init__ stores whatever pyclr_create_appdomain returned; the probe must report a NULL —
    as PROMPT-030S-R established, to NAME the root/current-domain sentinel the unnamed request is meant
    to produce, not to treat that sentinel as a failure (see the accept/reject tests below)."""
    assert '"appdomain"] = "NULL"' in bp.PROBE_PACKAGED_RUNTIME
    assert "domain == ffi.NULL" in bp.PROBE_PACKAGED_RUNTIME
    assert '"Python.Runtime.Loader"' in bp.PROBE_PACKAGED_RUNTIME and '"Initialize"' in bp.PROBE_PACKAGED_RUNTIME
    # The probe deliberately asks for the UNNAMED domain — the production pythonnet path (get_netfx()).
    assert "pyclr_create_appdomain(ffi.NULL, ffi.NULL)" in bp.PROBE_PACKAGED_RUNTIME


def test_probe_is_skipped_outside_windows_and_says_so(tmp_path, monkeypatch, capsys):
    """Never report a pass that was not performed."""
    monkeypatch.setattr(bp, "platform", types.SimpleNamespace(system=staticmethod(lambda: "Linux")))
    assert bp.probe_packaged_runtime(tmp_path, Path(sys.executable), _build_arch()) == []
    assert "BỎ QUA" in capsys.readouterr().out


def test_probe_turns_a_null_resolution_into_a_build_failure(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    payload = {"arch": "amd64", "appdomain": "created", "resolved": False}

    def fake_run(cmd, **kwargs):
        return types.SimpleNamespace(returncode=0, stdout=__import__("json").dumps(payload), stderr="")

    monkeypatch.setattr(bp.subprocess, "run", fake_run)
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("pyclr_get_function trả về NULL" in p for p in problems), problems


#: PROMPT-030S-R: the reproduction.  The verified healthy Windows evidence (root-domain NULL, real
#: resolution, Initialize 0) must be accepted — the old gate failed it, and the old test below encoded
#: that wrong expectation.  A NULL appdomain on the probe's unnamed/default request is clr-loader's
#: ROOT/CURRENT-domain sentinel (ClrLoader.cs answers index 0 = AppDomain.CurrentDomain), not a failure.
#:
#: PROMPT-030T — a successful INITIALIZE is only half of the lifecycle pythonnet 3.0.5 actually runs.
#: ``pythonnet/__init__.py::load()`` registers ``atexit(unload)`` the moment ``Initialize`` returns 0, and
#: ``unload()`` is ``Loader.Shutdown(b"full_shutdown")`` (``src/runtime/Loader.cs:40-60`` →
#: ``PythonEngine.Shutdown()``) BEFORE ``_RUNTIME.shutdown()``; clr-loader's own ``atexit(_release)`` then
#: calls ``pyclr_finalize()`` (``clr_loader/netfx.py:68,74``).  Skipping that half leaves
#: ``AppDomain.CurrentDomain.ProcessExit += OnProcessExit`` subscribed and the managed Finalizer queue
#: full, so the CLR tears the runtime down out of order and ``Py.cs:41``'s ``~GILState()`` throws
#: ``GIL must always be released…`` from the finalizer thread — an unhandled exception on .NET Framework,
#: i.e. a nonzero subprocess exit AFTER healthy-looking JSON was already printed.  A PASS therefore has to
#: prove shutdown and finalization too.
HEALTHY_ROOT_DOMAIN_EVIDENCE = {"arch": "amd64", "appdomain": "NULL", "resolved": True,
                                "initialize_rc": 0, "shutdown_resolved": True, "shutdown_rc": 0,
                                "finalized": True, "stage": "pythonnet-shutdown"}

#: What the probe emitted BEFORE PROMPT-030T: initialization proven, shutdown/finalization never performed.
#: On Windows that is the payload of a subprocess which then dies, so it must now FAIL CLOSED.
INITIALIZE_ONLY_EVIDENCE = {"arch": "amd64", "appdomain": "NULL", "resolved": True, "initialize_rc": 0}


def _feed_probe(monkeypatch, payload_text, returncode=0, stderr=""):
    _force_windows(monkeypatch)
    monkeypatch.setattr(bp.subprocess, "run", lambda *a, **k: types.SimpleNamespace(
        returncode=returncode, stdout=payload_text, stderr=stderr))


def test_probe_accepts_the_root_appdomain_null_sentinel(tmp_path, monkeypatch, capsys):
    """The reproduced Windows defect: appdomain=NULL + resolved=True + initialize_rc=0 must PASS."""
    _feed_probe(monkeypatch, json.dumps(HEALTHY_ROOT_DOMAIN_EVIDENCE))
    assert bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64") == []
    printed = capsys.readouterr().out
    assert "probe appdomain" in printed and "NULL" in printed, "the sentinel stays visible in the build log"
    assert "root/current AppDomain" in printed, "…and must be DESCRIBED as the root/current domain, not a failure"


def test_probe_turns_null_appdomain_plus_failed_resolution_into_a_build_failure(tmp_path, monkeypatch):
    """NULL domain is fine; NULL domain AND unresolved function is still the real PROMPT-030 crash."""
    _feed_probe(monkeypatch, json.dumps({"arch": "amd64", "appdomain": "NULL", "resolved": False}))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("pyclr_get_function trả về NULL" in p for p in problems), problems


@pytest.mark.parametrize("mutate, expected", [
    pytest.param(lambda p: p.pop("resolved"), "không xuất ra bằng chứng phân giải",
                 id="resolved-missing-fails-closed"),
    pytest.param(lambda p: p.pop("initialize_rc"), "chưa có kết quả hợp lệ",
                 id="initialize_rc-missing-fails-closed"),
    pytest.param(lambda p: p.update(initialize_rc=False), "chưa có kết quả hợp lệ",
                 id="initialize_rc-JSON-false-is-not-zero"),
    pytest.param(lambda p: p.update(initialize_rc=True), "chưa có kết quả hợp lệ",
                 id="initialize_rc-JSON-true-is-not-zero"),
    pytest.param(lambda p: p.update(arch="x86"), "kiến trúc", id="arch-echo-mismatch-fails"),
    pytest.param(lambda p: p.update(appdomain="weird"), "AppDomain hợp lệ",
                 id="malformed-appdomain-field-fails"),
    pytest.param(lambda p: p.update(error="OSError: cannot load library"), "thất bại",
                 id="probe-error-fails"),
])
def test_probe_fails_closed_on_incomplete_or_malformed_evidence(tmp_path, monkeypatch, mutate, expected):
    payload = dict(HEALTHY_ROOT_DOMAIN_EVIDENCE)
    mutate(payload)
    _feed_probe(monkeypatch, json.dumps(payload))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any(expected in p for p in problems), problems


def test_probe_fails_closed_on_nonzero_exit_with_parseable_json(tmp_path, monkeypatch):
    """The exact healthy payload with exit code 1: a completed-looking JSON line must NOT rescue it."""
    _feed_probe(monkeypatch, json.dumps(HEALTHY_ROOT_DOMAIN_EVIDENCE), returncode=1)
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("mã lỗi 1" in p for p in problems), problems


@pytest.mark.parametrize("stdout", [
    pytest.param("not json at all", id="non-JSON"),
    pytest.param("42", id="JSON-int-instead-of-object"),
    pytest.param("null", id="JSON-null-instead-of-object"),
    pytest.param("", id="empty-stdout"),
])
def test_probe_fails_closed_on_unparseable_output(tmp_path, monkeypatch, stdout):
    """A gate that cannot read the probe's evidence must stop the build, never pass on ignorance."""
    _feed_probe(monkeypatch, stdout)
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert problems and any("không trả về JSON hợp lệ" in p for p in problems), problems
    assert not any("Traceback" in p for p in problems), "the validator itself must not raise"


def test_gate_accepts_the_healthy_root_domain_probe_end_to_end(tmp_path, monkeypatch):
    """The Windows acceptance shape, decision-tested on any host: every structural proof plus the real
    root-domain probe evidence (NULL sentinel, resolved True, initialize_rc 0, shutdown resolved,
    shutdown_rc 0, finalized) yields NO problems — while each structural check below the probe stays
    independently enforced by the tests above."""
    _feed_probe(monkeypatch, json.dumps(HEALTHY_ROOT_DOMAIN_EVIDENCE))
    problems = validate_pythonnet_runtime(_package(tmp_path), {"found": False}, Path(sys.executable))
    assert problems == [], problems


def test_gate_rejects_an_initialize_only_probe_end_to_end(tmp_path, monkeypatch):
    """The same package, same exit code, but a probe that stopped after Initialize: the gate must stop
    the build even though every structural proof and every initialization field looks healthy."""
    _feed_probe(monkeypatch, json.dumps(INITIALIZE_ONLY_EVIDENCE))
    problems = validate_pythonnet_runtime(_package(tmp_path), {"found": False}, Path(sys.executable))
    assert problems and any("Shutdown" in p for p in problems), problems


def test_probe_accepts_a_runtime_that_resolves(tmp_path, monkeypatch):
    """The named-domain shape of success (a NON-NULL handle stays valid evidence where it applies)."""
    _feed_probe(monkeypatch, json.dumps({**HEALTHY_ROOT_DOMAIN_EVIDENCE, "appdomain": "created"}))
    assert bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64") == []


def test_probe_rejects_a_nonzero_initialize_result(tmp_path, monkeypatch):
    _feed_probe(monkeypatch, json.dumps({"arch": "amd64", "appdomain": "created",
                                         "resolved": True, "initialize_rc": 3}))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("Initialize trả về 3" in p for p in problems), problems


def test_probe_does_not_demand_shutdown_evidence_from_a_runtime_that_never_initialized(tmp_path, monkeypatch):
    """H — a failed Initialize must stay the single visible finding.  Shutting down (or demanding shutdown
    evidence for) a runtime that never started would bury the real cause under invented complaints."""
    _feed_probe(monkeypatch, json.dumps({"arch": "amd64", "appdomain": "NULL",
                                         "resolved": True, "initialize_rc": 1}))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert problems == ["Python.Runtime.Loader.Initialize trả về 1 (khác 0): pythonnet không khởi động "
                        "được trong gói đã build"], problems


def test_probe_reports_a_crash_instead_of_a_silent_pass(tmp_path, monkeypatch):
    """Exit 1 with NO JSON: still a named failure, and the native stderr excerpt survives into the report."""
    _feed_probe(monkeypatch, "", returncode=1, stderr="Traceback: OSError: cannot load library")
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("không trả về JSON hợp lệ" in p and "cannot load library" in p for p in problems), problems
    assert any("rc=1" in p for p in problems), "the reported exit code must not be hidden"


def test_probe_reports_its_own_exception(tmp_path, monkeypatch):
    """A native-load exception (dlopen refusing the packaged host) fails the gate by itself."""
    _feed_probe(monkeypatch, json.dumps({"error": "OSError: cannot load library 'ClrLoader.dll'"}))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("probe nạp runtime đã đóng gói thất bại" in p for p in problems), problems
    assert len(problems) == 1, "the exception is the finding; no invented secondary complaints"


# ================================================= the pythonnet lifecycle (PROMPT-030T)
def test_probe_fails_closed_when_the_shutdown_half_of_the_lifecycle_is_absent(tmp_path, monkeypatch):
    """D — THE reproduced Windows failure boundary, at decision level.

    This exact payload is what the pre-PROMPT-030T probe printed on Windows immediately before the
    subprocess died with ``GIL must always be released, and it must be released from the same thread that
    acquired it``.  Initialization evidence alone must NOT be a pass any more: the probe has to prove it
    ran ``Loader.Shutdown(b"full_shutdown")`` and finalized the CLR.
    """
    _feed_probe(monkeypatch, json.dumps(INITIALIZE_ONLY_EVIDENCE))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("Shutdown" in p for p in problems), problems
    assert problems, "initialization-only evidence must fail closed"


def test_probe_accepts_the_complete_root_domain_lifecycle(tmp_path, monkeypatch, capsys):
    """A — NULL root domain + Initialize 0 + Shutdown resolved, rc 0 + finalized, process rc 0 → PASS."""
    _feed_probe(monkeypatch, json.dumps(HEALTHY_ROOT_DOMAIN_EVIDENCE))
    assert bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64") == []
    printed = capsys.readouterr().out
    for key in ("appdomain", "resolved", "initialize_rc", "shutdown_resolved", "shutdown_rc", "finalized"):
        assert f"probe {key}" in printed, f"{key} must stay visible in the Windows build log"
    assert "root/current AppDomain" in printed, "030S-R semantics must survive"


def test_probe_fails_when_shutdown_cannot_be_resolved_after_successful_initialize(tmp_path, monkeypatch):
    """B — Initialize returned 0, so Shutdown was obligatory; an unresolvable Shutdown is a real defect."""
    _feed_probe(monkeypatch, json.dumps({**HEALTHY_ROOT_DOMAIN_EVIDENCE, "shutdown_resolved": False}))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("Shutdown" in p and "NULL" in p for p in problems), problems


@pytest.mark.parametrize("mutate, expected", [
    pytest.param(lambda p: p.pop("shutdown_resolved"), "bằng chứng",
                 id="shutdown_resolved-missing-fails-closed"),
    pytest.param(lambda p: p.update(shutdown_resolved="yes"), "bằng chứng",
                 id="shutdown_resolved-non-boolean-fails-closed"),
    pytest.param(lambda p: p.pop("shutdown_rc"), "chưa có kết quả hợp lệ",
                 id="shutdown_rc-missing-fails-closed"),
    pytest.param(lambda p: p.update(shutdown_rc=False), "chưa có kết quả hợp lệ",
                 id="shutdown_rc-JSON-false-is-not-zero"),
    pytest.param(lambda p: p.update(shutdown_rc=True), "chưa có kết quả hợp lệ",
                 id="shutdown_rc-JSON-true-is-not-zero"),
])
def test_probe_fails_closed_on_malformed_shutdown_evidence(tmp_path, monkeypatch, mutate, expected):
    """E — missing or malformed shutdown evidence is ignorance, and ignorance fails closed."""
    payload = dict(HEALTHY_ROOT_DOMAIN_EVIDENCE)
    mutate(payload)
    _feed_probe(monkeypatch, json.dumps(payload))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any(expected in p for p in problems), problems


def test_probe_rejects_a_nonzero_shutdown_result(tmp_path, monkeypatch):
    """C — Loader.Shutdown returning 1 means PythonEngine.Shutdown() threw; the runtime is not clean."""
    _feed_probe(monkeypatch, json.dumps({**HEALTHY_ROOT_DOMAIN_EVIDENCE, "shutdown_rc": 1}))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("Shutdown trả về 1" in p for p in problems), problems


@pytest.mark.parametrize("payload, expected", [
    pytest.param({**HEALTHY_ROOT_DOMAIN_EVIDENCE, "finalized": False,
                  "finalize_error": "OSError: pyclr_finalize crashed"},
                 "pyclr_finalize", id="finalization-raised"),
    pytest.param({**HEALTHY_ROOT_DOMAIN_EVIDENCE, "finalized": False},
                 "pyclr_finalize", id="finalization-reported-false"),
    pytest.param({k: v for k, v in HEALTHY_ROOT_DOMAIN_EVIDENCE.items() if k != "finalized"},
                 "pyclr_finalize", id="finalization-evidence-missing"),
    pytest.param({**HEALTHY_ROOT_DOMAIN_EVIDENCE, "finalized": 1},
                 "pyclr_finalize", id="finalized-JSON-1-is-not-true"),
])
def test_probe_fails_when_clr_loader_finalization_did_not_succeed(tmp_path, monkeypatch, payload, expected):
    """F — clr-loader registers atexit(_release) → pyclr_finalize(); skipping or failing it fails the gate."""
    _feed_probe(monkeypatch, json.dumps(payload))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any(expected in p for p in problems), problems
    assert not any("Shutdown" in p for p in problems), \
        "shutdown succeeded here; the finding must stay on the finalization stage"


def test_probe_fails_closed_on_nonzero_exit_even_with_a_complete_healthy_lifecycle(tmp_path, monkeypatch):
    """G — the exit-code gate is NOT weakened by the new positive evidence.

    This is the exact PROMPT-030T regression: a probe that printed a perfect lifecycle report and then
    died must still fail.  Nothing in the shutdown/finalization evidence may rescue a nonzero exit.
    """
    _feed_probe(monkeypatch, json.dumps(HEALTHY_ROOT_DOMAIN_EVIDENCE), returncode=1,
                stderr="System.InvalidOperationException: GIL must always be released, "
                       "and it must be released from the same thread that acquired it.")
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("mã lỗi 1" in p for p in problems), problems
    assert any("GIL must always be released" in p for p in problems), "the .NET cause must stay visible"


def test_probe_invents_no_lifecycle_evidence_after_an_exception_before_initialize(tmp_path, monkeypatch):
    """I — a dlopen failure happens before any CLR/Python.Runtime state exists: the report must not claim
    shutdown or finalization success for state the probe never created."""
    _feed_probe(monkeypatch, json.dumps({"error": "OSError: cannot load library 'ClrLoader.dll'",
                                         "stage": "start"}))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert len(problems) == 1, problems
    assert all("Shutdown" not in p and "pyclr_finalize" not in p for p in problems), problems


def test_probe_source_models_the_verified_shutdown_and_finalize_lifecycle():
    """L — the probe body must contain the shutdown half, in the upstream order, with structured cleanup.

    Pinned evidence: ``pythonnet/__init__.py:160-167`` (``Loader.Shutdown(b"full_shutdown")`` then
    ``_RUNTIME.shutdown()``) and ``clr_loader/netfx.py:74`` (``pyclr_finalize()`` in ``atexit``).
    """
    source = bp.PROBE_PACKAGED_RUNTIME
    compile(source, "<probe>", "exec")
    assert 'b"Python.Runtime.Loader", b"Shutdown"' in source, "the probe must resolve Loader.Shutdown"
    assert 'b"full_shutdown"' in source, "…and call it with pythonnet's own full_shutdown command"
    assert "fw.pyclr_finalize()" in source, "…and finish with clr-loader's pyclr_finalize()"
    assert source.index('b"Python.Runtime.Loader", b"Initialize"') < source.index(
        'b"Python.Runtime.Loader", b"Shutdown"') < source.index("fw.pyclr_finalize()"), \
        "lifecycle order must be Initialize → Shutdown → finalize (atexit is LIFO upstream)"
    assert "finally:" in source, "cleanup must be structured, not skipped when the native chain throws"
    assert '"shutdown_resolved"' in source and '"shutdown_rc"' in source and '"finalized"' in source, \
        "the validator must receive positive evidence for every stage it relies on"


def test_probe_source_only_shuts_down_a_runtime_that_initialized():
    """L/H — ``PythonEngine.Shutdown()`` on a runtime that never initialized is not the upstream contract,
    and the finalize call must be gated on ``pyclr_initialize()`` having actually run."""
    source = bp.PROBE_PACKAGED_RUNTIME
    assert 'if out["initialize_rc"] == 0:' in source, \
        "Shutdown must be nested inside the successful-initialize branch"
    assert source.index('if out["initialize_rc"] == 0:') < source.index(
        'b"Python.Runtime.Loader", b"Shutdown"'), "…so no shutdown is attempted after a failed Initialize"
    assert 'if stage != "start":' in source, \
        "pyclr_finalize must only run once pyclr_initialize succeeded, and exactly once"
    assert source.index('if stage != "start":') < source.index("fw.pyclr_finalize()")
    assert "pyclr_create_appdomain(ffi.NULL, ffi.NULL)" in source, \
        "030S-R: the probe still requests the unnamed root domain — never a named one to dodge the sentinel"
    assert 'out["appdomain"] = "NULL" if domain == ffi.NULL else "created"' in source


def test_probe_source_never_lets_a_cleanup_error_hide_the_original_failure():
    """6 — the first exception is written to ``error`` once, in the except block; cleanup records its own
    failure under a different key instead of overwriting the finding."""
    source = bp.PROBE_PACKAGED_RUNTIME
    assert 'out["error"] = f"{type(exc).__name__}: {exc}"[:400]' in source
    assert source.count('out["error"] =') == 1, "only the original failure may claim the error key"
    assert '"finalize_error"' in source, "a finalize failure is reported separately, not swallowed"
    assert 'out["finalized"] = False' in source, "…and failed cleanup is never reported as successful"


# ---------------------------------------------------------------- executing the real probe source
#: A minimal stand-in for the ``cffi`` module so ``PROBE_PACKAGED_RUNTIME`` itself — the production
#: string, unmodified — can be executed on any host and its control flow observed.  This does NOT replace
#: the real Windows probe (which still dlopens the packaged ClrLoader.dll); it only lets Linux prove the
#: probe's stage/cleanup logic instead of merely grepping it.  Behaviour is selected by TNP_PROBE_FAKE.
#: The fake entry points return 99 when the probe passes the wrong buffer/size, so a drifted argument
#: shows up as a nonzero rc rather than silently passing.
_CFFI_STUB = r"""
import json, os, sys, types as _types

_cfg = json.loads(os.environ.get("TNP_PROBE_FAKE", "{}"))
_calls = []


def _boom(what):
    raise OSError(_cfg.get(what, "injected failure"))


class _FakeFFI:
    NULL = None

    def cdef(self, text):
        _calls.append(("cdef", text.count("pyclr_")))

    def dlopen(self, path):
        if _cfg.get("dlopen_raises"):
            _boom("dlopen_raises")
        return _FakeLib(path)

    def from_buffer(self, kind, data):
        return data

    def cast(self, kind, value):
        return value


class _FakeLib:
    def __init__(self, path):
        self.path = path
        self.finalized = 0

    def pyclr_initialize(self):
        _calls.append(("pyclr_initialize",))
        if _cfg.get("initialize_raises_native"):
            _boom("initialize_raises_native")

    def pyclr_create_appdomain(self, name, config):
        _calls.append(("pyclr_create_appdomain", name, config))
        return None if _cfg.get("root_domain", True) else object()

    def pyclr_get_function(self, domain, assembly, class_name, function):
        _calls.append(("pyclr_get_function", assembly, class_name, function))
        if function == b"Shutdown" and not _cfg.get("shutdown_resolvable", True):
            return None
        return _EntryPoint(function)

    def pyclr_close_appdomain(self, domain):
        _calls.append(("pyclr_close_appdomain", domain))

    def pyclr_finalize(self):
        self.finalized += 1
        _calls.append(("pyclr_finalize", self.finalized))
        if self.finalized > 1:
            raise OSError("pyclr_finalize called twice")
        if _cfg.get("finalize_raises"):
            _boom("finalize_raises")


class _EntryPoint:
    def __init__(self, function):
        self.function = function

    def __call__(self, buffer, size):
        _calls.append(("call", self.function, bytes(buffer), size))
        if self.function == b"Initialize" and _cfg.get("initialize_call_raises"):
            _boom("initialize_call_raises")
        if self.function == b"Initialize":
            return 99 if (bytes(buffer) != b"" or size != 0) else int(_cfg.get("initialize_rc", 0))
        if self.function == b"Shutdown":
            return 99 if (bytes(buffer) != b"full_shutdown" or size != 13) \
                else int(_cfg.get("shutdown_rc", 0))
        return 99


_mod = _types.ModuleType("cffi")
_mod.FFI = _FakeFFI
sys.modules["cffi"] = _mod
"""


def _run_probe_source(monkeypatch, tmp_path, fake_cfg):
    """Execute bp.PROBE_PACKAGED_RUNTIME for real and return (returncode, parsed-json-or-None, calls)."""
    monkeypatch.setenv("TNP_PROBE_FAKE", json.dumps(fake_cfg))
    proc = subprocess.run(
        [sys.executable, "-c", _CFFI_STUB + bp.PROBE_PACKAGED_RUNTIME, str(tmp_path), _build_arch()],
        capture_output=True, text=True, encoding="utf-8", errors="replace")
    try:
        out = json.loads(proc.stdout.strip().splitlines()[-1])
    except (ValueError, IndexError):
        out = None
    return proc.returncode, out, proc.stderr


def test_probe_source_runs_the_complete_lifecycle_and_exits_zero(tmp_path, monkeypatch):
    """A — executed, not grepped: the production probe string emits evidence for every stage and exits 0."""
    rc, out, err = _run_probe_source(monkeypatch, tmp_path, {})
    assert rc == 0, err
    assert out == {**out, "arch": _build_arch(), "appdomain": "NULL", "resolved": True,
                   "initialize_rc": 0, "shutdown_resolved": True, "shutdown_rc": 0,
                   "finalized": True, "stage": "pythonnet-shutdown"}, out
    assert out["host"].replace("\\", "/").endswith(
        f"_internal/clr_loader/ffi/dlls/{_build_arch()}/ClrLoader.dll"), out["host"]
    assert out["assembly"].replace("\\", "/").endswith(
        "_internal/pythonnet/runtime/Python.Runtime.dll"), out["assembly"]


def test_probe_source_never_shuts_down_a_runtime_whose_initialize_failed(tmp_path, monkeypatch):
    """H — executed: Initialize rc != 0 means no Loader.Shutdown call at all, the original rc survives,
    and pyclr_finalize still runs because the CLR itself WAS started."""
    rc, out, err = _run_probe_source(monkeypatch, tmp_path, {"initialize_rc": 1})
    assert rc == 0, err
    assert out["initialize_rc"] == 1 and "error" not in out
    assert "shutdown_resolved" not in out and "shutdown_rc" not in out, \
        "no shutdown may be attempted after a failed Initialize"
    assert out["finalized"] is True and out["stage"] == "appdomain", out


def test_probe_source_reports_a_failed_shutdown_and_still_finalizes(tmp_path, monkeypatch):
    """C — executed: Loader.Shutdown returning 1 is reported, the stage stops before shutdown, and the
    CLR is still finalized so the probe does not abandon native state it created."""
    rc, out, err = _run_probe_source(monkeypatch, tmp_path, {"shutdown_rc": 1})
    assert rc == 0, err
    assert out["shutdown_resolved"] is True and out["shutdown_rc"] == 1
    assert out["finalized"] is True and out["stage"] == "pythonnet-initialized", out


def test_probe_source_reports_an_unresolvable_shutdown(tmp_path, monkeypatch):
    """B — executed: a NULL functor for Loader.Shutdown is recorded as shutdown_resolved False."""
    rc, out, err = _run_probe_source(monkeypatch, tmp_path, {"shutdown_resolvable": False})
    assert rc == 0, err
    assert out["shutdown_resolved"] is False and "shutdown_rc" not in out
    assert out["finalized"] is True and out["stage"] == "pythonnet-initialized", out


def test_probe_source_invents_no_lifecycle_evidence_when_dlopen_fails(tmp_path, monkeypatch):
    """I — executed: a dlopen failure happens before any CLR state exists, so the probe reports the
    exception, the stage it died at, and NO shutdown/finalization success."""
    rc, out, err = _run_probe_source(monkeypatch, tmp_path, {"dlopen_raises": "cannot load ClrLoader.dll"})
    assert rc == 0, err
    assert out["error"].startswith("OSError: cannot load ClrLoader.dll"), out
    assert out["stage"] == "start" and out["error_stage"] == "start", out
    assert "finalized" not in out and "shutdown_resolved" not in out, out


def test_probe_source_keeps_the_original_failure_when_finalize_also_fails(tmp_path, monkeypatch):
    """6 — executed: a cleanup exception must be reported under its own key and must not replace the
    original finding, and failed cleanup is reported as False, never as success."""
    rc, out, err = _run_probe_source(monkeypatch, tmp_path,
                                     {"initialize_call_raises": "Loader.Initialize exploded",
                                      "finalize_raises": "pyclr_finalize exploded"})
    assert rc == 0, err
    assert out["error"] == "OSError: Loader.Initialize exploded", out
    assert out["error_stage"] == "appdomain", out
    assert out["finalized"] is False, out
    assert out["finalize_error"] == "OSError: pyclr_finalize exploded", out


def test_probe_source_does_not_finalize_a_clr_it_never_started(tmp_path, monkeypatch):
    """6 — executed: ``pyclr_initialize`` itself failing means there is nothing to finalize, so the probe
    must not claim cleanup it never attempted (and must not call it either)."""
    rc, out, err = _run_probe_source(monkeypatch, tmp_path,
                                     {"initialize_raises_native": "pyclr_initialize exploded"})
    assert rc == 0, err
    assert out["error"] == "OSError: pyclr_initialize exploded", out
    assert out["stage"] == "start" and "finalized" not in out, out


def test_probe_source_finalizes_exactly_once(tmp_path, monkeypatch):
    """6 — executed: no double-finalize.  The stub raises on a second ``pyclr_finalize()``, so a healthy
    ``finalized is True`` can only mean the finally block ran the call exactly once."""
    _, out, err = _run_probe_source(monkeypatch, tmp_path, {})
    assert out["finalized"] is True and "finalize_error" not in out, (out, err)
    _, out2, err = _run_probe_source(monkeypatch, tmp_path, {"shutdown_rc": 1})
    assert out2["finalized"] is True and out2["stage"] == "pythonnet-initialized", (out2, err)


def test_probe_source_sends_pythonnets_own_entry_point_arguments(tmp_path, monkeypatch):
    """The buffers must match clr_loader.types.ClrFunction.__call__ exactly (``ffi.from_buffer`` +
    ``len(buf)``): ``b""``/0 for Initialize and ``b"full_shutdown"``/13 for Shutdown.  A wrong argument
    makes the fake entry point return 99, which fails the healthy-path assertions."""
    _, out, err = _run_probe_source(monkeypatch, tmp_path, {})
    assert out["initialize_rc"] == 0 and out["shutdown_rc"] == 0, (out, err)
    assert 99 not in (out["initialize_rc"], out["shutdown_rc"])


def test_probe_reports_a_timeout_as_a_failure(tmp_path, monkeypatch):
    _force_windows(monkeypatch)

    def _hang(*args, **kwargs):
        raise bp.subprocess.TimeoutExpired(cmd="probe", timeout=240)

    monkeypatch.setattr(bp.subprocess, "run", _hang)
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("không chạy được probe" in p for p in problems), problems


def test_gate_runs_the_probe_only_when_a_build_interpreter_is_known(tmp_path, monkeypatch):
    """The layout gate stays usable standalone; the probe needs an interpreter to run in."""
    _force_windows(monkeypatch)
    seen = []
    monkeypatch.setattr(bp, "probe_packaged_runtime", lambda *a, **k: seen.append(a) or [])
    validate_pythonnet_runtime(_package(tmp_path), {"found": False})
    assert seen == [], "no interpreter → no probe, and no claim that one ran"
    validate_pythonnet_runtime(_package(tmp_path / "b"), {"found": False}, Path(sys.executable))
    assert len(seen) == 1


# ============================================================================== re-proving a shipped folder
def test_validate_only_reuses_the_same_gate_and_reports_a_real_cause(tmp_path, monkeypatch, capsys):
    """The H:\\ report is post-build; this structural case must not send its fake host to Windows."""
    _force_windows(monkeypatch)
    folder = _package(tmp_path, runtime=python_runtime_image(assembly_name="Wrong.Payload"))
    monkeypatch.setattr(bp, "source_pythonnet_dll", lambda py: {"found": False})
    monkeypatch.setattr(bp, "probe_packaged_runtime", lambda *args, **kwargs: [])
    assert bp.validate_existing_package(folder, Path(sys.executable)) == 1
    out = capsys.readouterr().out
    assert "KHÔNG thể khởi động" in out and "Wrong.Payload" in out


@pytest.mark.skipif(sys.platform != "win32",
                    reason="Windows-only: exercises real CFFI/CLR DLL loading; Linux cannot claim native loadability")
def test_validate_only_passes_a_healthy_folder(tmp_path, capsys):
    """The Windows live-probe case uses installed, genuinely loadable runtime payloads.

    PROMPT-030T — the probe must now prove the WHOLE lifecycle (Initialize → Shutdown → pyclr_finalize)
    and the subprocess must exit 0, with no ``GIL must always be released`` finalizer exception.
    """
    py = _validate_only_interpreter()
    folder = _package_from_installed_runtime(tmp_path, py)
    assert bp.validate_existing_package(folder, py) == 0
    output = capsys.readouterr().out
    assert "OK: Python.Runtime.Loader.Initialize phân giải được" in output
    assert "probe appdomain" in output and "probe resolved" in output and "probe initialize_rc" in output
    assert "probe shutdown_resolved" in output and "probe shutdown_rc" in output, \
        "the Windows log must carry the pythonnet unload() evidence"
    assert "probe finalized" in output, "…and the clr-loader pyclr_finalize() evidence"
    assert "root/current AppDomain" in output, "030S-R root-domain semantics must stay visible"
    assert "GIL must always be released" not in output, \
        "an unhandled Py.cs ~GILState() finalizer exception means the lifecycle is still incomplete"


@pytest.mark.parametrize(
    ("probe_problems", "expected_status"),
    [([], 0), (["injected live-probe failure"], 1)],
)
def test_validate_only_obeys_injected_probe_result_for_structural_fixture(
        tmp_path, monkeypatch, capsys, probe_problems, expected_status):
    """Exercise validate-only's decision contract; the injected result makes no native-loadability claim."""
    _force_windows(monkeypatch)
    folder = _package(tmp_path)  # PE/metadata fixture only; never passed to the real Windows loader.
    monkeypatch.setattr(bp, "source_pythonnet_dll", lambda py: {"found": False})
    calls = []

    def injected_probe(probe_folder, py, arch):
        calls.append((probe_folder, py, arch))
        return probe_problems

    monkeypatch.setattr(bp, "probe_packaged_runtime", injected_probe)
    assert bp.validate_existing_package(folder, Path(sys.executable)) == expected_status
    assert len(calls) == 1 and calls[0][0] == folder
    output = capsys.readouterr().out
    if probe_problems:
        assert "injected live-probe failure" in output


def test_validate_only_refuses_a_folder_that_is_not_there(tmp_path):
    with pytest.raises(SystemExit):
        bp.validate_existing_package(tmp_path / "absent", Path(sys.executable))


def test_builder_cli_exposes_validate_only():
    """The Windows acceptance steps depend on this flag existing."""
    assert "--validate-only" in (ROOT / "tools" / "build_portable.py").read_text(encoding="utf-8")
