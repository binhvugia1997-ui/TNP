"""PROMPT-005 – offline/LAN self-update: manifest, numeric build comparison, package safety, staging, apply/rollback,
GUI controls.  No Git, no Internet, no Python required on the client."""
import hashlib
import json
import os
import subprocess
import sys
import threading
import zipfile
from pathlib import Path

import pytest

import app
import app.updater as up
from app.config import AppConfig
from app.gui_controller import GuiController

sys.path.insert(0, str(Path(app.__file__).resolve().parent.parent / "tools"))
import build_portable as bp  # noqa: E402

V_NEW, B_NEW = "1.0.5", 5
V_NEWER, B_NEWER = "1.3.6", 19           # newer than the 1.3.5 / Build 018 release app


# ---------------------------------------------------------------------------- helpers
def _version_txt(version, build):
    return f"Report Extractor\nversion={version}\nbuild={up.format_build(build)}\nbuilt=x\ngit=unavailable\n"


def make_release_zip(path: Path, version=V_NEW, build=B_NEW, top=True, with_exe=True, with_version=True,
                     extra=None) -> Path:
    """A release ZIP shaped like build_portable.py's output (top folder optional)."""
    prefix = f"ReportExtractor_v{version}_Portable/" if top else ""
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        if with_exe:
            zf.writestr(prefix + up.APP_EXE, b"MZ-new-" + version.encode())
        zf.writestr(prefix + "_internal/app/core.py", f"# core {version}".encode())
        zf.writestr(prefix + "_internal/pptx/templates/default.pptx", b"PK-template")
        if with_version:
            zf.writestr(prefix + "VERSION.txt", _version_txt(version, build))
        zf.writestr(prefix + "README.txt", f"readme {version}")
        zf.writestr(prefix + "Output/.keep", b"")
        zf.writestr(prefix + "logs/.keep", b"")
        zf.writestr(prefix + "config/.keep", b"")
        for name, data in (extra or {}).items():
            zf.writestr(prefix + name, data)
    return path


def make_update_folder(folder: Path, version=V_NEW, build=B_NEW, sha=True, manifest=True, package=True, **zip_kw):
    folder.mkdir(parents=True, exist_ok=True)
    pkg = f"ReportExtractor_{version}.zip"
    if package:
        make_release_zip(folder / pkg, version, build, **zip_kw)
    data = {"version": version, "build": build, "package": pkg}
    if sha and package:
        data["sha256"] = up.sha256_file(folder / pkg)
    if manifest:
        (folder / "version.json").write_text(json.dumps(data), encoding="utf-8")
    return folder


def make_portable(root: Path, version="1.0.4", build=4) -> Path:
    """Current installation with user data that must survive an update."""
    (root / "_internal" / "app").mkdir(parents=True)
    (root / up.APP_EXE).write_bytes(b"MZ-old")
    (root / "_internal" / "app" / "core.py").write_text("# core old", encoding="utf-8")
    (root / "_internal" / "old_only.dll").write_bytes(b"old")
    (root / "VERSION.txt").write_text(_version_txt(version, build), encoding="utf-8")
    (root / "README.txt").write_text("readme old", encoding="utf-8")
    (root / "config").mkdir()
    (root / "config" / "config.json").write_text('{"update_path": "X:\\\\Update"}', encoding="utf-8")
    (root / "logs").mkdir()
    (root / "logs" / "app.log").write_text("log line", encoding="utf-8")
    (root / "Output").mkdir()
    (root / "Output" / "Kiem_chung_09_2026.xlsx").write_bytes(b"xlsx")
    (root / "Output" / "backup").mkdir()
    (root / "Output" / "backup" / "k_backup.xlsx").write_bytes(b"bak")
    (root / "Kiem_chung.xlsx").write_bytes(b"master")
    return root


def _ctl(tmp_path, update_path=""):
    return GuiController(AppConfig(update_path=update_path), config_path=tmp_path / "cfg" / "config.json")


# ---------------------------------------------------------------------------- version identity (§21/§23)
def test_numeric_build_display_and_constants():
    assert app.__version__ == "1.3.5" and app.BUILD_NUMBER == 18
    assert app.BUILD_ID == "018" and app.BUILD_LABEL == "Build 018" and app.APP_TITLE == "Report Extractor v1.3.5"
    assert app.format_build(4) == "004" and app.format_build(5) == "005" and app.format_build(15) == "015"
    assert app.format_build(1234) == "1234"
    assert up.version_label() == "1.3.5 — Build 018" and up.version_label("1.0.5", 5) == "1.0.5 — Build 005"
    assert "git" not in app.BUILD_ID.lower() and app.BUILD_ID.isdigit()


def test_build_script_uses_the_same_constants():
    assert bp._version() == ("1.3.5", "018") and bp.build_number() == 18
    assert bp.package_name("1.0.5") == "ReportExtractor_1.0.5.zip"


# ---------------------------------------------------------------------------- check (§3, §6, §15-18, §22)
def test_no_update_path_configured(tmp_path):
    r = up.check_for_update("")
    assert r.status == "no_path" and r.message == "Chưa cấu hình đường dẫn cập nhật."
    c = _ctl(tmp_path)
    assert c.update_status_text() == "Chưa cấu hình đường dẫn cập nhật." and not c.check_update_async(startup=True)


def test_local_update_folder_update_available(tmp_path):
    folder = make_update_folder(tmp_path / "Update")
    r = up.check_for_update(str(folder), current_build=4)
    assert r.status == "available" and r.available and r.message == "Có phiên bản mới: 1.0.5 — Build 005"
    assert r.info.build == 5 and Path(r.package_path) == folder / "ReportExtractor_1.0.5.zip"


def test_unc_network_path_is_handled_without_exception():
    for unc in (r"\\SERVER\ReportExtractor\Update", r"\\192.168.1.10\ReportExtractor\Update"):
        r = up.check_for_update(unc, current_build=4)
        assert r.status in ("inaccessible", "no_permission", "no_manifest") and r.update_path == unc


def test_inaccessible_update_path(tmp_path):
    r = up.check_for_update(str(tmp_path / "does_not_exist"))
    assert r.status == "inaccessible" and r.message == "Không truy cập được thư mục cập nhật."
    f = tmp_path / "file_not_folder"
    f.write_text("x")
    assert up.check_for_update(str(f)).status == "inaccessible"


def test_no_permission_message(tmp_path, monkeypatch):
    def deny(_p):
        raise PermissionError(13, "denied")
    monkeypatch.setattr(up.os, "listdir", deny)
    r = up.check_for_update(str(tmp_path))
    assert r.status == "no_permission" and r.message == "Không có quyền truy cập đường dẫn cập nhật."


def test_missing_version_json(tmp_path):
    folder = make_update_folder(tmp_path / "Update", manifest=False)
    r = up.check_for_update(str(folder))
    assert r.status == "no_manifest" and "version.json" in r.message


def test_invalid_json(tmp_path):
    folder = tmp_path / "Update"
    folder.mkdir()
    (folder / "version.json").write_text("{not json", encoding="utf-8")
    r = up.check_for_update(str(folder))
    assert r.status == "invalid_manifest" and r.message.startswith("version.json không hợp lệ")


@pytest.mark.parametrize("data", [
    {"version": "1.0.5", "package": "x.zip"},                       # no build
    {"build": 5, "package": "x.zip"},                               # no version
    {"version": "1.0.5", "build": 5},                               # no package
    {"version": "1.0.5", "build": "five", "package": "x.zip"},      # build not numeric
    {"version": "1.0.5", "build": True, "package": "x.zip"},        # bool is not an int here
    {"version": "1.0.5", "build": 5, "package": "../x.zip"},        # path traversal
    {"version": "1.0.5", "build": 5, "package": "x.zip", "sha256": "zz"},
    [1, 2],
])
def test_missing_or_invalid_required_metadata(data):
    with pytest.raises(up.ManifestError):
        up.parse_manifest(json.dumps(data))


def test_build_as_numeric_string_is_accepted():
    info = up.parse_manifest('{"version": "1.0.5", "build": "005", "package": "p.zip"}')
    assert info.build == 5 and info.label() == "1.0.5 — Build 005"


def test_numeric_build_comparison_not_text():
    assert up.compare_builds(4, 5) == "available"
    assert up.compare_builds(5, 5) == "latest"
    assert up.compare_builds(6, 5) == "older"
    assert up.compare_builds(9, 10) == "available"           # "10" < "9" as text – must be numeric


def test_current_004_remote_005_update_available(tmp_path):
    folder = make_update_folder(tmp_path / "U")
    assert up.check_for_update(str(folder), current_build=4).status == "available"


def test_current_005_remote_005_latest(tmp_path):
    folder = make_update_folder(tmp_path / "U")
    r = up.check_for_update(str(folder), current_build=5)
    assert r.status == "latest" and r.message == "Đã là phiên bản mới nhất." and not r.available


def test_current_006_remote_005_no_downgrade(tmp_path):
    folder = make_update_folder(tmp_path / "U")
    r = up.check_for_update(str(folder), current_build=6)
    assert r.status == "older" and not r.available
    assert r.message == "Phiên bản trên đường dẫn cập nhật cũ hơn phiên bản hiện tại."


def test_package_missing(tmp_path):
    folder = make_update_folder(tmp_path / "U", package=False)
    r = up.check_for_update(str(folder), current_build=4)
    assert r.status == "package_missing" and "ReportExtractor_1.0.5.zip" in r.message and not r.available


def test_check_is_read_only(tmp_path):
    folder = make_update_folder(tmp_path / "U")
    before = {p: p.read_bytes() for p in folder.rglob("*") if p.is_file()}
    portable = make_portable(tmp_path / "app")
    snap = sorted(str(p.relative_to(portable)) for p in portable.rglob("*"))
    up.check_for_update(str(folder), current_build=4)
    assert {p: p.read_bytes() for p in folder.rglob("*") if p.is_file()} == before
    assert sorted(str(p.relative_to(portable)) for p in portable.rglob("*")) == snap


# ---------------------------------------------------------------------------- package safety (§11)
def _info(folder, **kw):
    d = json.loads((folder / "version.json").read_text(encoding="utf-8"))
    d.update(kw)
    return up.parse_manifest(json.dumps(d))


def test_valid_zip_passes(tmp_path):
    folder = make_update_folder(tmp_path / "U")
    ok, msg = up.verify_package(folder / "ReportExtractor_1.0.5.zip", _info(folder))
    assert ok and msg == "OK"


def test_valid_zip_without_top_folder_passes(tmp_path):
    folder = make_update_folder(tmp_path / "U", top=False)
    assert up.verify_package(folder / "ReportExtractor_1.0.5.zip", _info(folder))[0]


def test_corrupt_zip_rejected(tmp_path):
    folder = make_update_folder(tmp_path / "U", sha=False)
    z = folder / "ReportExtractor_1.0.5.zip"
    z.write_bytes(b"this is not a zip at all")
    ok, msg = up.verify_package(z, _info(folder))
    assert not ok and msg.startswith("Gói cập nhật không hợp lệ")
    big = make_release_zip(tmp_path / "c.zip", extra={"_internal/big.bin": bytes(range(256)) * 64})
    data = bytearray(big.read_bytes())
    data = data[: len(data) - 200]                                          # truncated copy (interrupted transfer)
    (tmp_path / "c.zip").write_bytes(bytes(data))
    assert not up.verify_package(tmp_path / "c.zip", _info(folder))[0]
    empty = tmp_path / "empty.zip"
    empty.write_bytes(b"")
    assert not up.verify_package(empty, _info(folder))[0]
    assert not up.verify_package(tmp_path / "missing.zip", _info(folder))[0]


def test_required_exe_missing_from_zip_rejected(tmp_path):
    folder = make_update_folder(tmp_path / "U", with_exe=False)
    ok, msg = up.verify_package(folder / "ReportExtractor_1.0.5.zip", _info(folder))
    assert not ok and up.APP_EXE in msg


def test_version_metadata_must_match_requested_update(tmp_path):
    folder = make_update_folder(tmp_path / "U")                                  # zip says 1.0.5 / 005
    z = folder / "ReportExtractor_1.0.5.zip"
    assert not up.verify_package(z, _info(folder, version="1.0.6"))[0]          # manifest says 1.0.6
    assert not up.verify_package(z, _info(folder, build=6))[0]                  # manifest says build 006
    folder2 = make_update_folder(tmp_path / "U2", with_version=False)
    ok, msg = up.verify_package(folder2 / "ReportExtractor_1.0.5.zip", _info(folder2))
    assert not ok and "VERSION.txt" in msg


def test_sha256_valid(tmp_path):
    folder = make_update_folder(tmp_path / "U", sha=True)
    info = _info(folder)
    assert info.sha256 == hashlib.sha256((folder / info.package).read_bytes()).hexdigest()
    assert up.verify_package(folder / info.package, info)[0]


def test_sha256_mismatch_aborts(tmp_path):
    folder = make_update_folder(tmp_path / "U", sha=True)
    info = _info(folder, sha256="0" * 64)
    ok, msg = up.verify_package(folder / info.package, info)
    assert not ok and msg == "Gói cập nhật không hợp lệ hoặc bị thay đổi."
    portable = make_portable(tmp_path / "app")
    check = up.check_for_update(str(folder), current_build=4)
    check.info = info
    with pytest.raises(RuntimeError, match="Gói cập nhật không hợp lệ hoặc bị thay đổi."):
        up.stage_update(check, portable)
    assert not (portable / up.STAGING_DIR).exists() and (portable / up.APP_EXE).read_bytes() == b"MZ-old"


# ---------------------------------------------------------------------------- staging + updater launch (§9, §10)
def test_package_copied_to_local_staging_before_installation(tmp_path):
    folder = make_update_folder(tmp_path / "Update")
    portable = make_portable(tmp_path / "app")
    check = up.check_for_update(str(folder), current_build=4)
    staged = up.stage_update(check, portable)
    staging = portable / up.STAGING_DIR
    assert Path(staged.zip_path) == staging / "ReportExtractor_1.0.5.zip" and Path(staged.zip_path).exists()
    assert Path(staged.source_root).is_relative_to(staging) and (Path(staged.source_root) / up.APP_EXE).exists()
    m = json.loads(Path(staged.manifest_path).read_text(encoding="utf-8"))
    assert m["old_pid"] == os.getpid() and m["target_root"] == str(portable) and m["build"] == 5
    # nothing installed yet
    assert (portable / up.APP_EXE).read_bytes() == b"MZ-old"
    # updater runs from the STAGED copy, never from the share and never the running exe
    cmd = up.updater_command(staged, packaged=True)
    assert cmd[0] == str(Path(staged.source_root) / up.APP_EXE) and cmd[1:] == ["--apply-update", staged.manifest_path]
    dev = up.updater_command(staged, packaged=False)
    assert dev[0] == sys.executable and dev[1].endswith("run.py") and dev[2] == "--apply-update"
    spawned = []
    up.launch_updater(staged, spawn=lambda cmd, **kw: spawned.append((cmd, kw)))
    assert spawned and spawned[0][1]["cwd"] == staged.source_root
    log = (portable / "logs" / up.UPDATE_LOG).read_text(encoding="utf-8")
    assert "UPDATE_STAGE start" in log and "verify=ok" in log and "UPDATE_LAUNCH" in log


def test_zip_slip_rejected(tmp_path):
    folder = make_update_folder(tmp_path / "U", sha=False, extra={"../../evil.txt": b"x"})
    portable = make_portable(tmp_path / "app")
    check = up.check_for_update(str(folder), current_build=4)
    with pytest.raises(RuntimeError):
        up.stage_update(check, portable)
    assert not (tmp_path / "evil.txt").exists()


# ---------------------------------------------------------------------------- apply (§9, §12, §13, §14)
def _staged(tmp_path, **folder_kw):
    folder = make_update_folder(tmp_path / "Update", **folder_kw)
    portable = make_portable(tmp_path / "app")
    check = up.check_for_update(str(folder), current_build=4)
    return portable, up.stage_update(check, portable)


def test_updater_waits_for_main_process_exit(tmp_path):
    portable, staged = _staged(tmp_path)
    alive_calls = {"n": 0}

    def alive(pid):
        alive_calls["n"] += 1
        return alive_calls["n"] <= 3                 # alive for three polls, then gone
    slept = []
    rc = up.apply_update(Path(staged.manifest_path), alive=alive, sleep=slept.append, spawn=lambda *a, **k: None)
    assert rc == 0 and alive_calls["n"] >= 4 and len(slept) == 3
    assert (portable / up.APP_EXE).read_bytes() == b"MZ-new-1.0.5"
    # process never exits -> nothing replaced
    portable2, staged2 = _staged(tmp_path / "second")
    rc2 = up.apply_update(Path(staged2.manifest_path), alive=lambda pid: True, sleep=lambda s: None,
                          spawn=lambda *a, **k: None, wait_timeout=0.0)
    assert rc2 == 1 and (portable2 / up.APP_EXE).read_bytes() == b"MZ-old"
    assert not (portable2 / up.BACKUP_DIR / up.APP_EXE).exists()


def test_application_files_replaced_and_user_data_preserved(tmp_path):
    portable, staged = _staged(tmp_path)
    spawned = []
    rc = up.apply_update(Path(staged.manifest_path), alive=lambda pid: False, sleep=lambda s: None,
                         spawn=lambda cmd, **kw: spawned.append((cmd, kw)))
    assert rc == 0
    # application files replaced
    assert (portable / up.APP_EXE).read_bytes() == b"MZ-new-1.0.5"
    assert (portable / "_internal" / "app" / "core.py").read_text(encoding="utf-8") == "# core 1.0.5"
    assert not (portable / "_internal" / "old_only.dll").exists()                 # stale runtime file gone
    assert "build=005" in (portable / "VERSION.txt").read_text(encoding="utf-8")
    assert (portable / "README.txt").read_text(encoding="utf-8") == "readme 1.0.5"
    # config.json preserved
    assert (portable / "config" / "config.json").read_text(encoding="utf-8") == '{"update_path": "X:\\\\Update"}'
    # Output preserved (incl. result workbook + Excel backups)
    assert (portable / "Output" / "Kiem_chung_09_2026.xlsx").read_bytes() == b"xlsx"
    assert (portable / "Output" / "backup" / "k_backup.xlsx").read_bytes() == b"bak"
    assert not (portable / "Output" / ".keep").exists()                            # package's empty dirs not merged
    # logs preserved (+ update log/result written there)
    assert (portable / "logs" / "app.log").read_text(encoding="utf-8") == "log line"
    assert (portable / "Kiem_chung.xlsx").read_bytes() == b"master"                # root-level user Excel untouched
    # previous runtime kept for recovery
    assert (portable / up.BACKUP_DIR / up.APP_EXE).read_bytes() == b"MZ-old"
    # restart of the NEW application
    assert spawned and spawned[0][0][0] == str(portable / up.APP_EXE) and spawned[0][1]["cwd"] == str(portable)
    res = json.loads((portable / "logs" / up.RESULT_FILE).read_text(encoding="utf-8"))
    assert res["status"] == "ok" and res["version"] == "1.0.5" and res["build"] == 5
    log = (portable / "logs" / up.UPDATE_LOG).read_text(encoding="utf-8")
    for key in ("UPDATE_APPLY start", "old process exited", "UPDATE_REPLACE ok", "UPDATE_RESTART ok"):
        assert key in log


def test_failed_install_triggers_rollback(tmp_path, monkeypatch):
    portable, staged = _staged(tmp_path)
    orig = up._copy
    calls = {"n": 0}

    def flaky_copy(src, dst):
        calls["n"] += 1
        if calls["n"] == 3:                                  # fail in the middle of the replacement
            raise OSError(5, "disk I/O error")
        orig(src, dst)
    monkeypatch.setattr(up, "_copy", flaky_copy)
    spawned = []
    rc = up.apply_update(Path(staged.manifest_path), alive=lambda pid: False, sleep=lambda s: None,
                         spawn=lambda cmd, **kw: spawned.append(cmd))
    assert rc == 1 and not spawned                           # no restart of a half-updated app
    assert (portable / up.APP_EXE).read_bytes() == b"MZ-old"
    assert (portable / "_internal" / "app" / "core.py").read_text(encoding="utf-8") == "# core old"
    assert (portable / "_internal" / "old_only.dll").exists()
    assert (portable / "README.txt").read_text(encoding="utf-8") == "readme old"
    assert "build=004" in (portable / "VERSION.txt").read_text(encoding="utf-8")
    assert (portable / "config" / "config.json").exists() and (portable / "Output" / "Kiem_chung_09_2026.xlsx").exists()
    res = json.loads((portable / "logs" / up.RESULT_FILE).read_text(encoding="utf-8"))
    assert res["status"] == "rolled_back"
    log = (portable / "logs" / up.UPDATE_LOG).read_text(encoding="utf-8")
    assert "UPDATE_REPLACE failed" in log and "UPDATE_ROLLBACK result=ok" in log


def test_successful_install_restarts_application_and_notice_shown_once(tmp_path, monkeypatch):
    portable, staged = _staged(tmp_path)
    spawned = []
    assert up.apply_update(Path(staged.manifest_path), alive=lambda pid: False, sleep=lambda s: None,
                           spawn=lambda cmd, **kw: spawned.append(cmd)) == 0
    assert spawned == [[str(portable / up.APP_EXE)]]
    # the restarted GUI shows the one-shot notice, then it is gone
    import app.gui_controller as gc
    monkeypatch.setattr(gc, "portable_root", lambda: portable)
    c = _ctl(tmp_path)
    assert c.consume_update_notice() == "Cập nhật thành công lên phiên bản 1.0.5."
    assert c.consume_update_notice() == "" and not (portable / "logs" / up.RESULT_FILE).exists()
    assert not (portable / up.STAGING_DIR).exists()          # staging cleaned after a successful restart


def test_apply_update_cli_entry_point(tmp_path, monkeypatch):
    portable, staged = _staged(tmp_path)
    monkeypatch.setattr(up, "process_alive", lambda pid: False)
    monkeypatch.setattr(up.subprocess, "Popen", lambda *a, **k: None)
    from app.main import main
    assert main(["--apply-update", staged.manifest_path]) == 0
    assert (portable / up.APP_EXE).read_bytes() == b"MZ-new-1.0.5"
    assert main(["--apply-update", str(tmp_path / "nope.json")]) == 2


def test_preserved_entries_rules():
    for name in ("config", "Config", "logs", "Output", "backup", "config.json", "update_staging", "update_backup",
                 "Kiem_chung.xlsx", "report.pptx", "history.json", "app.log"):
        assert up.is_preserved(name), name
    for name in ("ReportExtractor.exe", "_internal", "VERSION.txt", "README.txt", "FIRST_RUN.txt",
                 "Install_Ollama_Optional.bat", "SHA256SUMS.txt"):
        assert not up.is_preserved(name), name


# ---------------------------------------------------------------------------- controller / GUI (§4, §5, §7, §8, §16)
def test_update_path_persisted_in_config(tmp_path):
    c = _ctl(tmp_path)
    c.set_update_path(r'"\\SERVER\ReportExtractor\Update"')
    c.save_settings()
    data = json.loads((tmp_path / "cfg" / "config.json").read_text(encoding="utf-8"))
    assert data["update_path"] == r"\\SERVER\ReportExtractor\Update"
    again = GuiController(AppConfig.load(tmp_path / "cfg" / "config.json"), config_path=tmp_path / "cfg" / "config.json")
    assert again.update_path == r"\\SERVER\ReportExtractor\Update"
    assert "password" not in json.dumps(data).lower() and "user" not in data


def test_controller_check_async_and_status_texts(tmp_path):
    folder = make_update_folder(tmp_path / "Update", version=V_NEWER, build=B_NEWER)
    c = _ctl(tmp_path, str(folder))
    assert c.current_version_text() == "Phiên bản hiện tại: 1.3.5 — Build 018"
    assert c.update_status_text() == "Chưa kiểm tra cập nhật."
    assert c.check_update_async(startup=True)
    c._update_thread.join(10)
    assert c.update_status_text() == "Có phiên bản mới: 1.3.6 — Build 019" and c.update_available()
    # offline share: a status line, never an exception, processing stays available
    c.set_update_path(str(tmp_path / "offline_share"))
    assert c.update_check is None and c.check_update_async()
    c._update_thread.join(10)
    assert c.update_status_text() == "Không truy cập được thư mục cập nhật." and not c.update_available()
    assert c.state == "idle" and c.can_start() is not None


def test_startup_check_is_non_blocking(tmp_path, monkeypatch):
    gate = threading.Event()

    def slow(path, *a, **k):
        gate.wait(5)
        return up.UpdateCheck("inaccessible", up.MSG["inaccessible"], path)
    monkeypatch.setattr(up, "check_for_update", slow)
    c = _ctl(tmp_path, r"\\SLOW\share")
    import time
    t0 = time.monotonic()
    assert c.check_update_async(startup=True)
    assert time.monotonic() - t0 < 1.0 and c.update_busy and c.update_status_text() == "Đang kiểm tra cập nhật…"
    gate.set()
    c._update_thread.join(10)
    # PROMPT-009: automatic start-up failure is soft (status kept internally, no alarming text)
    assert c.update_check.status == "inaccessible"
    assert c.update_status_text() == "Không thể kiểm tra cập nhật tự động — chương trình vẫn hoạt động bình thường."
    c.check_update()                                                   # manual check keeps the explicit message
    assert c.update_status_text() == "Không truy cập được thư mục cập nhật."


def test_install_update_requires_confirmation_and_idle_state(tmp_path, monkeypatch):
    folder = make_update_folder(tmp_path / "Update", version=V_NEWER, build=B_NEWER)
    portable = make_portable(tmp_path / "app")
    import app.gui_controller as gc
    monkeypatch.setattr(gc, "portable_root", lambda: portable)
    c = _ctl(tmp_path, str(folder))
    assert c.install_update() == (False, "Không có bản cập nhật để cài.")      # never without a positive check
    c.check_update()
    spawned = []
    ok, msg = c.install_update(spawn=lambda cmd, **kw: spawned.append(cmd))
    assert ok and msg.startswith("Đang cài đặt 1.3.6 — Build 019")
    assert spawned and "--apply-update" in spawned[0] and spawned[0][-1].endswith("update.json")
    assert (portable / up.APP_EXE).read_bytes() == b"MZ-old"                    # files untouched until the updater runs
    # downgrade / latest never install
    c2 = _ctl(tmp_path / "x", str(folder))
    c2.update_check = up.check_for_update(str(folder), current_build=B_NEWER + 1)
    assert c2.install_update()[0] is False


def test_gui_update_card_controls(monkeypatch, tmp_path):
    from tests.test_gui_redesign import _make_app
    folder = make_update_folder(tmp_path / "Update", version=V_NEWER, build=B_NEWER)
    gui, a, reg, _ = _make_app(monkeypatch, tmp_path)
    btn_texts = {w.cfg.get("text") for w in reg["widgets"] if type(w).__name__ == "Button"}
    assert {"Chọn...", "Kiểm tra cập nhật", "Cập nhật ngay"} <= btn_texts
    labels = {w.cfg.get("text") for w in reg["widgets"] if type(w).__name__ == "Label"}
    assert "Đường dẫn cập nhật:" in labels and "Phiên bản hiện tại: 1.3.5 — Build 018" in labels
    assert a.lbl_build.cfg["text"] == "Build 018"
    assert a.btn_install_update.cfg["state"] == "disabled"
    # startup check scheduled asynchronously (after), not inline
    assert any(ms == gui.STARTUP_UPDATE_DELAY_MS for ms, _ in a.root.scheduled) and 2000 <= gui.STARTUP_UPDATE_DELAY_MS <= 5000
    # user enters a path and checks
    a.var_update_path.set(str(folder))
    a.check_update()
    a.ctl._update_thread.join(10)
    a._poll()
    assert a.lbl_update.cfg["text"] == "Có phiên bản mới: 1.3.6 — Build 019"
    assert a.btn_install_update.cfg["state"] == "normal"
    assert a.ctl.update_path == str(folder) and a.ctl.cfg.update_path == str(folder)     # persisted via save_settings
    from app.runtime_paths import config_file, portable_root
    test_root = Path(os.environ["TNP_TEST_RUNTIME_ROOT"]).resolve()
    assert portable_root() == test_root and config_file() == test_root / "config" / "config.json"
    assert config_file().is_file() and config_file().is_relative_to(tmp_path)
    assert AppConfig.load().update_path == str(folder)                              # saved under disposable pytest state
    # "Để sau" (dialog closed without confirming) -> nothing happens
    a.install_update()
    assert a.btn_install_update.cfg["state"] == "normal" and a.lbl_update.cfg["text"].startswith("Có phiên bản mới")
    # latest / older render as status text only
    a.ctl.update_check = up.check_for_update(str(folder), current_build=B_NEWER)
    a._render_update()
    assert a.lbl_update.cfg["text"] == "Đã là phiên bản mới nhất." and a.btn_install_update.cfg["state"] == "disabled"
    a.ctl.update_check = up.check_for_update(str(folder), current_build=B_NEWER + 1)
    a._render_update()
    assert a.lbl_update.cfg["text"] == "Phiên bản trên đường dẫn cập nhật cũ hơn phiên bản hiện tại."


def test_no_git_dependency_in_updater():
    src = Path(up.__file__).read_text(encoding="utf-8")
    assert "git" not in src.lower().replace("digit", "") or "rev-parse" not in src
    assert "subprocess.run" not in src                                             # only Popen for the updater/restart
    assert "pip" not in src and "urllib" not in src and "requests" not in src       # no Internet, no pip


# ---------------------------------------------------------------------------- release generation (§20)
def test_build_script_generates_version_json_with_exact_sha256(tmp_path):
    release = tmp_path / "release"
    release.mkdir()
    z = make_release_zip(release / bp.package_name("1.0.5"), "1.0.5", 5)
    out = bp.write_version_manifest(release, "1.0.5", 5, z, built="2026-10-02 10:00")
    data = json.loads(out.read_text(encoding="utf-8"))
    assert out.name == "version.json" and data["package"] == "ReportExtractor_1.0.5.zip" and data["build"] == 5
    assert data["version"] == "1.0.5" and data["sha256"] == hashlib.sha256(z.read_bytes()).hexdigest()
    # the generated manifest is exactly what the client understands
    info = up.parse_manifest(out.read_text(encoding="utf-8"))
    assert up.verify_package(z, info) == (True, "OK")
    assert up.check_for_update(str(release), current_build=4).status == "available"
    src = Path(bp.__file__).read_text(encoding="utf-8")
    assert "write_version_manifest(release" in src and "SAU CÙNG" in src
    assert subprocess is not None
