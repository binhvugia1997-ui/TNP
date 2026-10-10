"""PROMPT-030 — the packaged .NET runtime must be PROVEN LOADABLE, not merely present.

The real Windows Portable died before its first window::

    RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize from:
    H:\\ReportExtractor_v1.3.4_Portable\\_internal\\pythonnet\\runtime\\Python.Runtime.dll

That string is the *only* thing ``clr_loader/netfx.py`` reports: its cdef declares
``pyclr_initialize`` / ``pyclr_create_appdomain`` / ``pyclr_get_function`` / ``pyclr_close_appdomain`` /
``pyclr_finalize`` and **no error accessor**, so whatever the CLR actually threw is discarded, and
``NetFx.__init__`` stores whatever ``pyclr_create_appdomain`` returned without checking it for NULL while
``NetFx.info()`` hardcodes ``initialized=True``.  PROMPT-028R's gate proved presence, uniqueness, location
and SHA256 — and the build passed it, so the exe shipped and crashed anyway.

These tests pin what PROMPT-030 adds:

* the packaged payload's **managed identity** — assembly name, IL-only/any-cpu, target framework, and a
  ``Python.Runtime.Loader.Initialize`` that is ``static int32 (native int, int32)``, which is precisely the
  ``entry_point`` typedef clr_loader turns into a delegate;
* the packaged netfx host's **architecture, mixed-mode-ness and five ``pyclr_*`` exports**;
* a **Mark-of-the-Web block**, the one condition that appears AFTER the build (a folder extracted from
  ``release/*.zip`` or copied from another PC): ``LoadLibrary`` ignores ``Zone.Identifier`` so the native
  host loads and the app domain is created, while ``Assembly.LoadFrom`` refuses the managed assembly with
  0x80131515 — a DLL that exists at the right path with the right SHA256 and still cannot initialize;
* a **real loadability probe** that performs the failing call against the packaged bytes and checks the
  app-domain handle for NULL, which ``clr_loader`` itself never does;
* startup diagnostics that report the discriminating facts instead of asserting an unprovable one.

Nothing here needs Windows, a .NET installation, PyInstaller or a built artifact: the managed images are
synthesized from ECMA-335 by ``tests/dotnet_image_fixtures.py``, and the real shipped DLLs are asserted
against the same reader whenever ``pythonnet`` happens to be installed.
"""
from __future__ import annotations

import importlib.util
import struct
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
def _fake_pythonnet(monkeypatch, domain):
    runtime = None if domain == "no-runtime" else types.SimpleNamespace(_domain=domain)
    monkeypatch.setitem(sys.modules, "pythonnet", types.SimpleNamespace(_RUNTIME=runtime))


def test_appdomain_report_reads_the_handle_instead_of_trusting_runtime_info(monkeypatch):
    """NetFx.info() hardcodes initialized=True, so get_runtime_info() cannot answer this question."""
    null = object()
    monkeypatch.setattr(desktop.sys, "platform", "win32")
    monkeypatch.setitem(sys.modules, "clr_loader", types.ModuleType("clr_loader"))
    monkeypatch.setitem(sys.modules, "clr_loader.ffi", types.SimpleNamespace(
        ffi=types.SimpleNamespace(NULL=null)))
    _fake_pythonnet(monkeypatch, 0x1234)
    assert desktop.appdomain_report() == "created"
    _fake_pythonnet(monkeypatch, null)
    assert desktop.appdomain_report().startswith("NULL(")
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
    assert desktop.appdomain_report().startswith("NULL(")


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


def test_probe_checks_the_domain_handle_that_clr_loader_never_checks():
    """NetFx.__init__ stores whatever pyclr_create_appdomain returned; the probe must notice a NULL."""
    assert '"appdomain"] = "NULL"' in bp.PROBE_PACKAGED_RUNTIME
    assert "domain == ffi.NULL" in bp.PROBE_PACKAGED_RUNTIME
    assert '"Python.Runtime.Loader"' in bp.PROBE_PACKAGED_RUNTIME and '"Initialize"' in bp.PROBE_PACKAGED_RUNTIME


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


def test_probe_turns_a_null_appdomain_into_a_build_failure(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    payload = {"arch": "amd64", "appdomain": "NULL", "resolved": False}
    monkeypatch.setattr(bp.subprocess, "run", lambda *a, **k: types.SimpleNamespace(
        returncode=0, stdout=__import__("json").dumps(payload), stderr=""))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("pyclr_create_appdomain trả về NULL" in p for p in problems), problems


def test_probe_accepts_a_runtime_that_resolves(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    payload = {"arch": "amd64", "appdomain": "created", "resolved": True, "initialize_rc": 0}
    monkeypatch.setattr(bp.subprocess, "run", lambda *a, **k: types.SimpleNamespace(
        returncode=0, stdout=__import__("json").dumps(payload), stderr=""))
    assert bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64") == []


def test_probe_rejects_a_nonzero_initialize_result(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    payload = {"arch": "amd64", "appdomain": "created", "resolved": True, "initialize_rc": 3}
    monkeypatch.setattr(bp.subprocess, "run", lambda *a, **k: types.SimpleNamespace(
        returncode=0, stdout=__import__("json").dumps(payload), stderr=""))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("Initialize trả về 3" in p for p in problems), problems


def test_probe_reports_a_crash_instead_of_a_silent_pass(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    monkeypatch.setattr(bp.subprocess, "run", lambda *a, **k: types.SimpleNamespace(
        returncode=1, stdout="", stderr="Traceback: OSError: cannot load library"))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("không trả về JSON" in p and "cannot load library" in p for p in problems), problems


def test_probe_reports_its_own_exception(tmp_path, monkeypatch):
    _force_windows(monkeypatch)
    payload = {"error": "OSError: cannot load library 'ClrLoader.dll'"}
    monkeypatch.setattr(bp.subprocess, "run", lambda *a, **k: types.SimpleNamespace(
        returncode=0, stdout=__import__("json").dumps(payload), stderr=""))
    problems = bp.probe_packaged_runtime(tmp_path, Path(sys.executable), "amd64")
    assert any("probe nạp runtime đã đóng gói thất bại" in p for p in problems), problems


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
    """The Windows live-probe case uses installed, genuinely loadable runtime payloads."""
    py = _validate_only_interpreter()
    folder = _package_from_installed_runtime(tmp_path, py)
    assert bp.validate_existing_package(folder, py) == 0
    output = capsys.readouterr().out
    assert "OK: Python.Runtime.Loader.Initialize phân giải được" in output
    assert "probe appdomain" in output and "probe resolved" in output and "probe initialize_rc" in output


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
