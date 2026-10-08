"""Automatic publish to the LAN update folder (tools/publish_update.py + build_portable step 12) – build tooling only."""
import hashlib
import json
import os
import re
import sys
from pathlib import Path

import pytest

import app

ROOT = Path(app.__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import publish_update as pu  # noqa: E402
import build_portable as bp  # noqa: E402

BUILD_SRC = (ROOT / "tools" / "build_portable.py").read_text(encoding="utf-8")
BAT = (ROOT / "build_portable.bat").read_text(encoding="utf-8")


def _release(tmp_path, version="1.2.1", build=12, payload=b"ZIP" * 5000, name="release"):
    rel = tmp_path / name
    rel.mkdir(parents=True, exist_ok=True)
    zip_path = rel / f"ReportExtractor_{version}.zip"
    zip_path.write_bytes(payload)
    (rel / "version.json").write_text(json.dumps({"version": version, "build": build, "package": zip_path.name,
                                                  "sha256": hashlib.sha256(payload).hexdigest(),
                                                  "size": len(payload)}), encoding="utf-8")
    return rel, zip_path


def _manifest(dest):
    return json.loads((dest / "version.json").read_text(encoding="utf-8"))


class Recorder:
    """Records console output AND the folder state at each [PUBLISH] line (proves ZIP-before-manifest ordering)."""
    def __init__(self, dest):
        self.dest, self.lines, self.snapshots = dest, [], []

    def __call__(self, msg):
        self.lines.append(msg)
        self.snapshots.append((msg, sorted(p.name for p in self.dest.glob("*")) if self.dest.exists() else []))


# ------------------------------------------------------------------ success / ordering / sha / cleanup
def test_successful_publish_and_console_output(tmp_path):
    rel, zip_src = _release(tmp_path)
    dest = tmp_path / "ReportExtractor_Update"
    rec = Recorder(dest)
    res = pu.publish_release(rel, str(dest), say=rec, lan=r"\\192.168.103.11\ReportExtractor_Update")
    assert res.ok and res.package == zip_src.name and res.error == ""
    assert (dest / zip_src.name).read_bytes() == zip_src.read_bytes()
    assert _manifest(dest) == json.loads((rel / "version.json").read_text(encoding="utf-8"))
    text = "\n".join(rec.lines)
    for expected in (f"[PUBLISH] Update folder: {dest}", "[PUBLISH] Copying package...", "[PUBLISH] Verifying SHA256...",
                     "[PUBLISH] Package OK", "[PUBLISH] Publishing version.json LAST...", "[PUBLISH] SUCCESS",
                     f"Local update folder:\n    {dest}", r"LAN update path:" + "\n    " + r"\\192.168.103.11\ReportExtractor_Update"):
        assert expected in text
    assert not list(dest.glob("*.tmp"))                                  # temp files cleaned


def test_zip_is_published_before_manifest(tmp_path):
    rel, zip_src = _release(tmp_path)
    dest = tmp_path / "upd"
    rec = Recorder(dest)
    assert pu.publish_release(rel, str(dest), say=rec).ok
    at = {msg: names for msg, names in rec.snapshots}
    assert "version.json" not in at["[PUBLISH] Package OK"] and zip_src.name in at["[PUBLISH] Package OK"]
    assert "version.json" in at["[PUBLISH] SUCCESS"]
    # the manifest only ever appears after the final ZIP name exists (never a dangling manifest)
    seen_zip = False
    for _msg, names in rec.snapshots:
        seen_zip = seen_zip or zip_src.name in names
        if "version.json" in names:
            assert seen_zip


def test_sha256_verification_of_destination(tmp_path, monkeypatch):
    rel, zip_src = _release(tmp_path)
    dest = tmp_path / "upd"
    # previous valid release present
    old_payload = b"OLD" * 100
    (dest).mkdir()
    (dest / "ReportExtractor_1.0.9.zip").write_bytes(old_payload)
    old_manifest = {"version": "1.0.9", "build": 9, "package": "ReportExtractor_1.0.9.zip",
                    "sha256": hashlib.sha256(old_payload).hexdigest()}
    (dest / "version.json").write_text(json.dumps(old_manifest), encoding="utf-8")
    real_copy = pu.shutil.copyfile

    def corrupt_copy(src, dst, *a, **k):                                  # same size, different bytes (bit rot on LAN)
        real_copy(src, dst)
        data = bytearray(Path(dst).read_bytes())
        data[10] ^= 0xFF
        Path(dst).write_bytes(bytes(data))
    monkeypatch.setattr(pu.shutil, "copyfile", corrupt_copy)
    rec = Recorder(dest)
    res = pu.publish_release(rel, str(dest), say=rec)
    assert not res.ok and "SHA256" in res.error and "[PUBLISH] FAILED" in "\n".join(rec.lines)
    assert _manifest(dest) == old_manifest and (dest / "ReportExtractor_1.0.9.zip").read_bytes() == old_payload
    assert not (dest / zip_src.name).exists() and not list(dest.glob("*.tmp"))


def test_source_sha_mismatch_refuses_to_publish(tmp_path):
    rel, zip_src = _release(tmp_path)
    zip_src.write_bytes(b"tampered")
    dest = tmp_path / "upd"
    res = pu.publish_release(rel, str(dest), say=lambda m: None)
    assert not res.ok and "SHA256" in res.error and not (dest / "version.json").exists()


def test_old_zip_cleanup_after_publish_and_unrelated_files_preserved(tmp_path):
    rel, zip_src = _release(tmp_path)
    dest = tmp_path / "upd"
    dest.mkdir()
    for old in ("ReportExtractor_1.0.4.zip", "ReportExtractor_1.1.0-beta.old.zip"):
        (dest / old).write_bytes(b"old")
    (dest / "ReportExtractor_1.0.9.zip.tmp").write_bytes(b"stale tmp")
    keep = {"README_UPDATE.txt": b"notes", "Other_Tool_2.0.zip": b"other", "SHA256SUMS.txt": b"sums"}
    for n, b in keep.items():
        (dest / n).write_bytes(b)
    (dest / "subfolder").mkdir()
    (dest / "subfolder" / "ReportExtractor_0.9.zip").write_bytes(b"nested")
    rec = Recorder(dest)
    res = pu.publish_release(rel, str(dest), say=rec)
    assert res.ok and sorted(res.removed) == ["ReportExtractor_1.0.4.zip", "ReportExtractor_1.1.0-beta.old.zip"]
    assert all((dest / n).read_bytes() == b for n, b in keep.items())
    assert (dest / "subfolder" / "ReportExtractor_0.9.zip").exists()        # never recursive
    assert not (dest / "ReportExtractor_1.0.9.zip.tmp").exists()
    assert sorted(p.name for p in dest.glob("ReportExtractor_*.zip")) == [zip_src.name]
    text = "\n".join(rec.lines)
    assert "[PUBLISH] Removing obsolete package: ReportExtractor_1.0.4.zip" in text
    # cleanup happens only AFTER SUCCESS-critical steps: the removal lines come after "Publishing version.json LAST..."
    idx = [i for i, l in enumerate(rec.lines) if "Removing obsolete" in l]
    assert min(idx) > rec.lines.index("[PUBLISH] Publishing version.json LAST...")


def test_republish_same_version_keeps_single_package(tmp_path):
    rel, zip_src = _release(tmp_path)
    dest = tmp_path / "upd"
    assert pu.publish_release(rel, str(dest), say=lambda m: None).ok
    rel2, _ = _release(tmp_path, payload=b"NEW" * 7000, name="release2")
    assert pu.publish_release(rel2, str(dest), say=lambda m: None).ok
    assert (dest / zip_src.name).read_bytes() == b"NEW" * 7000 and _manifest(dest)["size"] == 21000
    assert [p.name for p in dest.glob("ReportExtractor_*.zip")] == [zip_src.name]


# ------------------------------------------------------------------ failure modes keep the previous release valid
def _prev_release(dest):
    dest.mkdir(parents=True, exist_ok=True)
    payload = b"PREV" * 50
    (dest / "ReportExtractor_1.0.9.zip").write_bytes(payload)
    m = {"version": "1.0.9", "build": 9, "package": "ReportExtractor_1.0.9.zip",
         "sha256": hashlib.sha256(payload).hexdigest()}
    (dest / "version.json").write_text(json.dumps(m), encoding="utf-8")
    return m, payload


def test_copy_failure_preserves_previous_manifest_and_package(tmp_path, monkeypatch):
    rel, zip_src = _release(tmp_path)
    dest = tmp_path / "upd"
    prev, payload = _prev_release(dest)
    monkeypatch.setattr(pu.shutil, "copyfile", lambda *a, **k: (_ for _ in ()).throw(OSError(5, "LAN dropped")))
    rec = Recorder(dest)
    res = pu.publish_release(rel, str(dest), say=rec)
    assert not res.ok and "LAN dropped" in res.error
    assert _manifest(dest) == prev and (dest / "ReportExtractor_1.0.9.zip").read_bytes() == payload
    assert not (dest / zip_src.name).exists() and not list(dest.glob("*.tmp"))
    assert "[PUBLISH] FAILED" in "\n".join(rec.lines) and "[PUBLISH] SUCCESS" not in rec.lines


def test_manifest_replacement_failure_does_not_corrupt_previous_manifest(tmp_path, monkeypatch):
    rel, zip_src = _release(tmp_path)
    dest = tmp_path / "upd"
    prev, payload = _prev_release(dest)
    real_replace = os.replace

    def flaky_replace(src, dst):
        if str(dst).endswith("version.json"):
            raise PermissionError(13, "manifest locked")
        return real_replace(src, dst)
    monkeypatch.setattr(pu.os, "replace", flaky_replace)
    res = pu.publish_release(rel, str(dest), say=lambda m: None)
    assert not res.ok and "manifest locked" in res.error
    assert _manifest(dest) == prev                                        # old manifest byte-for-byte valid
    assert (dest / prev["package"]).read_bytes() == payload               # and still points at an existing ZIP
    assert not (dest / zip_src.name).exists()                             # no unreferenced new ZIP left behind
    assert not list(dest.glob("*.tmp"))


def test_manifest_failure_without_previous_manifest_leaves_no_dangling_zip(tmp_path, monkeypatch):
    rel, zip_src = _release(tmp_path)
    dest = tmp_path / "upd"
    monkeypatch.setattr(pu.os, "replace", lambda s, d: (_ for _ in ()).throw(OSError("disk full"))
                        if str(d).endswith("version.json") else os.rename(s, d))
    res = pu.publish_release(rel, str(dest), say=lambda m: None)
    assert not res.ok and not (dest / "version.json").exists() and not (dest / zip_src.name).exists()


def test_missing_release_inputs_fail_cleanly(tmp_path):
    dest = tmp_path / "upd"
    assert not pu.publish_release(tmp_path / "nothing", str(dest), say=lambda m: None).ok
    rel, zip_src = _release(tmp_path)
    zip_src.unlink()
    assert not pu.publish_release(rel, str(dest), say=lambda m: None).ok
    assert not dest.exists() or not (dest / "version.json").exists()


# ------------------------------------------------------------------ build integration: only after a successful build
def test_publish_is_the_last_build_step_and_never_runs_after_failures():
    main_src = BUILD_SRC[BUILD_SRC.index("def main("):BUILD_SRC.index("def publish_step(")]
    order = [m.group(1) for m in re.finditer(r"step\((\d+),", main_src)]
    assert order == [str(i) for i in range(1, 13)] and 'step(12, "Xuất bản' in main_src
    publish_pos = main_src.index("publish_step(")
    for gate in ("pyflakes", 'run([py, "-m", "pytest", "-q"])', "PyInstaller", "validate_artifact(folder)",
                 "write_version_manifest(", "zipfile.ZipFile(zip_path"):
        assert main_src.index(gate) < publish_pos                         # every gate precedes publishing
    # fail() exits the process (code 1) -> publishing is unreachable after any failed gate
    assert "sys.exit(code)" in BUILD_SRC[BUILD_SRC.index("def fail("):BUILD_SRC.index("def run(")]
    assert "--no-publish" in main_src and "--publish-dir" in main_src


def test_publish_step_reports_failure_without_masking_build_success(tmp_path, capsys):
    rel, zip_src = _release(tmp_path)
    assert bp.publish_step(rel, str(tmp_path / "upd")) == 0
    out = capsys.readouterr().out
    assert "[PUBLISH] SUCCESS" in out
    zip_src.write_bytes(b"broken")                                        # release no longer matches its manifest
    assert bp.publish_step(rel, str(tmp_path / "upd2")) == bp.PUBLISH_FAILED_RC == 3
    out = capsys.readouterr().out
    assert "PUBLISH FAILED" in out and "BUILD THÀNH CÔNG" in out
    assert 'if "%RC%"=="3"' in BAT and "PUBLISH FAILED" in BAT


def test_no_publish_after_failed_test_gate(tmp_path, monkeypatch):
    """Simulate the build driver: pytest failing -> fail() -> SystemExit before publish_release is ever called."""
    calls = []
    monkeypatch.setattr(pu, "publish_release", lambda *a, **k: calls.append(1))
    monkeypatch.setattr(bp, "run", lambda cmd, cwd=None: (_ for _ in ()).throw(SystemExit(1))
                        if "pytest" in " ".join(map(str, cmd)) else None)
    monkeypatch.setattr(bp.subprocess, "run", lambda *a, **k: type("R", (), {"stdout": "", "stderr": ""})())
    monkeypatch.setattr(bp, "check_python", lambda py: None)
    monkeypatch.setattr(sys, "argv", ["build_portable.py", "--no-venv", "--allow-non-windows"])
    with pytest.raises(SystemExit) as ex:
        bp.main()
    assert ex.value.code == 1 and calls == []


def test_destination_is_centralised_and_configurable(monkeypatch):
    assert pu.DEFAULT_UPDATE_FOLDER == r"D:\ReportExtractor_Update"
    assert pu.DEFAULT_LAN_PATH == r"\\192.168.103.11\ReportExtractor_Update"
    assert "ReportExtractor_Update" not in BUILD_SRC                       # only publish_update.py knows the path
    assert "192.168.103.11" not in BUILD_SRC and "192.168.103.11" not in (ROOT / "app" / "updater.py").read_text(encoding="utf-8")
    monkeypatch.setenv("RE_UPDATE_FOLDER", r"E:\Share\Update")
    monkeypatch.setenv("RE_UPDATE_LAN_PATH", r"\\BUILD-PC\Update")
    assert pu.update_folder() == r"E:\Share\Update" and pu.lan_path() == r"\\BUILD-PC\Update"
    monkeypatch.delenv("RE_UPDATE_FOLDER")
    monkeypatch.delenv("RE_UPDATE_LAN_PATH")
    assert pu.update_folder() == pu.DEFAULT_UPDATE_FOLDER and pu.lan_path() == pu.DEFAULT_LAN_PATH


def test_build_010_identity():
    assert app.__version__ == "1.3.3" and app.BUILD_NUMBER == 16 and app.BUILD_LABEL == "Build 016"
