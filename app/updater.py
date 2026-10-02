"""Offline / LAN self-update for the Windows Portable build (PROMPT-005).

Update source = a Windows-accessible folder (``D:\\ReportExtractor_Update``, ``\\\\SERVER\\share\\Update`` …) that
contains ``version.json`` + the release ZIP.  No Git, no Internet, no Python on the client, no credentials: the
current Windows session's access rights are used as they are.

Flow (``ReportExtractor.exe`` never overwrites itself):

    check_for_update()      read version.json, numeric build comparison, package present      (read-only)
    stage_update()          copy ZIP -> <portable>/update_staging, verify (zip/exe/metadata/sha256), extract
    updater_command()       launch the *staged new* ReportExtractor.exe --apply-update <manifest>, then exit
    apply_update()          (runs in the staged copy) wait for the old process, move current runtime files to
                            <portable>/update_backup, copy the new release in, restore on failure, restart app

Persistent user data (config/, logs/, Output/, backup/, config.json, Excel files, update_* dirs …) is never
replaced or deleted.  Everything is logged to <portable>/logs/update.log (no credentials).
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import shutil
import subprocess
import sys
import time
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional, Tuple

from . import BUILD_NUMBER, __version__, format_build

LOG = logging.getLogger("report_extractor.update")

APP_EXE = "ReportExtractor.exe"
MANIFEST_NAME = "version.json"
STAGING_DIR = "update_staging"
BACKUP_DIR = "update_backup"
UPDATE_LOG = "update.log"
RESULT_FILE = "update_result.json"            # written to <portable>/logs by the updater, shown once by the GUI
STAGED_MANIFEST = "update.json"
# Entries of the portable folder that belong to the USER, never to a release package -> never replaced/deleted.
PRESERVED_ENTRIES = frozenset(e.lower() for e in (
    "config", "logs", "output", "backup", "config.json", STAGING_DIR, BACKUP_DIR, "history.json",
    "learning_data",                                   # PROMPT-006 image labels / thumbnails / local model
))
PRESERVED_SUFFIXES = (".xlsx", ".xlsm", ".xls", ".pptx", ".ppt", ".log", ".json")   # root-level user files

MSG = {
    "no_path": "Chưa cấu hình đường dẫn cập nhật.",
    "inaccessible": "Không truy cập được thư mục cập nhật.",
    "no_permission": "Không có quyền truy cập đường dẫn cập nhật.",
    "no_manifest": "Không tìm thấy version.json trong thư mục cập nhật.",
    "invalid_manifest": "version.json không hợp lệ",
    "package_missing": "Không tìm thấy gói cập nhật",
    "latest": "Đã là phiên bản mới nhất.",
    "older": "Phiên bản trên đường dẫn cập nhật cũ hơn phiên bản hiện tại.",
    "available": "Có phiên bản mới",
    "bad_package": "Gói cập nhật không hợp lệ hoặc bị thay đổi.",
}


def version_label(version: str = __version__, build: int = BUILD_NUMBER) -> str:
    """'1.0.4 — Build 004'."""
    return f"{version} — Build {format_build(build)}"


# ---------------------------------------------------------------------------- manifest
class ManifestError(ValueError):
    """version.json missing/invalid – message is user-facing Vietnamese."""


@dataclass
class UpdateInfo:
    version: str
    build: int
    package: str
    sha256: str = ""

    def label(self) -> str:
        return version_label(self.version, self.build)


def parse_manifest(text: str) -> UpdateInfo:
    try:
        data = json.loads(text)
    except Exception as e:  # noqa: BLE001
        raise ManifestError(f"{MSG['invalid_manifest']} (JSON lỗi: {type(e).__name__})") from e
    if not isinstance(data, dict):
        raise ManifestError(f"{MSG['invalid_manifest']} (không phải đối tượng JSON)")
    missing = [k for k in ("version", "build", "package") if k not in data]
    if missing:
        raise ManifestError(f"{MSG['invalid_manifest']} (thiếu {', '.join(missing)})")
    build = data["build"]
    if isinstance(build, bool) or not isinstance(build, (int, str)) or (isinstance(build, str) and not build.strip().isdigit()):
        raise ManifestError(f"{MSG['invalid_manifest']} (build phải là số nguyên)")
    build = int(build)
    version = str(data["version"]).strip()
    package = str(data["package"]).strip()
    if not version or not package or Path(package).name != package:
        raise ManifestError(f"{MSG['invalid_manifest']} (version/package không hợp lệ)")
    sha = str(data.get("sha256") or "").strip().lower()
    if sha and (len(sha) != 64 or any(c not in "0123456789abcdef" for c in sha)):
        raise ManifestError(f"{MSG['invalid_manifest']} (sha256 không hợp lệ)")
    return UpdateInfo(version=version, build=build, package=package, sha256=sha)


def compare_builds(current: int, remote: int) -> str:
    """Numeric ordering only: 'available' (remote > current), 'latest' (==), 'older' (remote < current)."""
    current, remote = int(current), int(remote)
    if remote > current:
        return "available"
    if remote == current:
        return "latest"
    return "older"


# ---------------------------------------------------------------------------- check (read-only)
@dataclass
class UpdateCheck:
    status: str                       # no_path | inaccessible | no_permission | no_manifest | invalid_manifest |
    message: str                      # package_missing | available | latest | older
    update_path: str = ""
    info: Optional[UpdateInfo] = None
    package_path: str = ""
    checked_at: str = ""

    @property
    def available(self) -> bool:
        return self.status == "available"


def _folder_state(path: Path) -> str:
    """'ok' | 'no_permission' | 'inaccessible' – never raises (UNC shares offline, denied, not a directory …)."""
    try:
        if path.is_dir():
            os.listdir(path)                       # an unreadable share says "exists" but refuses listing
            return "ok"
        return "inaccessible"
    except PermissionError:
        return "no_permission"
    except OSError as e:
        if getattr(e, "winerror", None) in (5, 1326, 1327, 1331, 1385):      # access denied / logon failures
            return "no_permission"
        return "inaccessible"


def check_for_update(update_path: str, current_build: int = BUILD_NUMBER,
                     current_version: str = __version__) -> UpdateCheck:
    """Read-only check of ``update_path`` (local or UNC).  Never modifies the application, never raises."""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    raw = (update_path or "").strip().strip('"')
    if not raw:
        return UpdateCheck("no_path", MSG["no_path"], "", checked_at=now)
    folder = Path(raw)
    LOG.info("UPDATE_CHECK current=%s build=%s path=%s", current_version, format_build(current_build), folder)
    state = _folder_state(folder)
    if state != "ok":
        LOG.warning("UPDATE_CHECK result=%s path=%s", state, folder)
        return UpdateCheck(state, MSG[state], raw, checked_at=now)
    manifest = folder / MANIFEST_NAME
    try:
        if not manifest.is_file():
            LOG.warning("UPDATE_CHECK result=no_manifest path=%s", manifest)
            return UpdateCheck("no_manifest", MSG["no_manifest"], raw, checked_at=now)
        info = parse_manifest(manifest.read_text(encoding="utf-8-sig"))
    except PermissionError:
        return UpdateCheck("no_permission", MSG["no_permission"], raw, checked_at=now)
    except ManifestError as e:
        LOG.warning("UPDATE_CHECK result=invalid_manifest reason=%s", e)
        return UpdateCheck("invalid_manifest", str(e), raw, checked_at=now)
    except OSError as e:
        LOG.warning("UPDATE_CHECK result=inaccessible reason=%s", e)
        return UpdateCheck("inaccessible", MSG["inaccessible"], raw, checked_at=now)
    pkg = folder / info.package
    try:
        pkg_ok = pkg.is_file()
    except OSError:
        pkg_ok = False
    LOG.info("UPDATE_CHECK remote=%s build=%s package=%s present=%s", info.version, format_build(info.build),
             info.package, pkg_ok)
    if not pkg_ok:
        return UpdateCheck("package_missing", f"{MSG['package_missing']} {info.package}.", raw, info, checked_at=now)
    rel = compare_builds(current_build, info.build)
    if rel == "available":
        msg = f"{MSG['available']}: {info.label()}"
    else:
        msg = MSG[rel]
    LOG.info("UPDATE_CHECK result=%s", rel)
    return UpdateCheck(rel, msg, raw, info, str(pkg), checked_at=now)


# ---------------------------------------------------------------------------- package safety
def sha256_file(path: Path, chunk: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(chunk), b""):
            h.update(block)
    return h.hexdigest()


def _exe_entry(names: List[str]) -> Optional[str]:
    """ZIP member that is the application executable: ``ReportExtractor.exe`` at the root or one folder deep."""
    for n in names:
        parts = n.replace("\\", "/").split("/")
        if parts[-1] == APP_EXE and len(parts) <= 2:
            return n
    return None


def verify_package(zip_path: Path, info: UpdateInfo) -> Tuple[bool, str]:
    """Package safety gate: exists, non-empty, opens, intact, contains the EXE, VERSION.txt matches, SHA256 matches."""
    try:
        if not zip_path.is_file() or zip_path.stat().st_size == 0:
            return False, f"{MSG['bad_package']} (gói trống hoặc không tồn tại)"
        if info.sha256:
            actual = sha256_file(zip_path)
            if actual != info.sha256:
                LOG.error("UPDATE_HASH result=mismatch expected=%s actual=%s", info.sha256, actual)
                return False, MSG["bad_package"]
            LOG.info("UPDATE_HASH result=ok sha256=%s", actual)
        else:
            LOG.info("UPDATE_HASH result=skipped (version.json không có sha256)")
        if not zipfile.is_zipfile(zip_path):
            return False, f"{MSG['bad_package']} (không phải file ZIP)"
        with zipfile.ZipFile(zip_path) as zf:
            if zf.testzip() is not None:
                return False, f"{MSG['bad_package']} (ZIP hỏng)"
            names = zf.namelist()
            exe = _exe_entry(names)
            if not exe:
                return False, f"{MSG['bad_package']} (không có {APP_EXE})"
            prefix = exe[: -len(APP_EXE)]
            ver_name = prefix + "VERSION.txt"
            if ver_name not in names:
                return False, f"{MSG['bad_package']} (thiếu VERSION.txt)"
            text = zf.read(ver_name).decode("utf-8-sig", errors="replace")
            if f"version={info.version}" not in text:
                return False, f"{MSG['bad_package']} (VERSION.txt không khớp phiên bản {info.version})"
            if f"build={format_build(info.build)}" not in text:
                return False, f"{MSG['bad_package']} (VERSION.txt không khớp Build {format_build(info.build)})"
    except (OSError, zipfile.BadZipFile) as e:
        return False, f"{MSG['bad_package']} ({type(e).__name__})"
    return True, "OK"


# ---------------------------------------------------------------------------- staging
@dataclass
class StagedUpdate:
    staging_dir: str
    zip_path: str
    source_root: str                  # extracted folder that contains ReportExtractor.exe
    target_root: str                  # portable folder to update
    manifest_path: str                # staging/update.json consumed by --apply-update
    info: UpdateInfo


def _clear_dir(p: Path) -> None:
    if p.exists():
        shutil.rmtree(p, ignore_errors=True)
    p.mkdir(parents=True, exist_ok=True)


def stage_update(check: UpdateCheck, portable_root: Path) -> StagedUpdate:
    """network/local source -> local staging -> validation -> extraction.  Raises RuntimeError (Vietnamese)."""
    if not check.available or not check.info or not check.package_path:
        raise RuntimeError("Không có bản cập nhật để cài.")
    info = check.info
    portable_root = Path(portable_root)
    staging = portable_root / STAGING_DIR
    _log_file(portable_root, f"UPDATE_STAGE start current={version_label()} remote={info.label()} "
                             f"source={check.package_path}")
    _clear_dir(staging)
    local_zip = staging / info.package
    try:
        shutil.copyfile(check.package_path, local_zip)
    except PermissionError as e:
        raise RuntimeError(MSG["no_permission"]) from e
    except OSError as e:
        raise RuntimeError(f"Không sao chép được gói cập nhật về máy ({e})") from e
    ok, msg = verify_package(local_zip, info)
    _log_file(portable_root, f"UPDATE_STAGE verify={'ok' if ok else 'failed'} detail={msg}")
    if not ok:
        shutil.rmtree(staging, ignore_errors=True)
        raise RuntimeError(msg)
    new_dir = staging / "new"
    with zipfile.ZipFile(local_zip) as zf:
        for member in zf.infolist():                       # zip-slip guard
            dest = (new_dir / member.filename).resolve()
            if new_dir.resolve() not in dest.parents and dest != new_dir.resolve():
                raise RuntimeError(f"{MSG['bad_package']} (đường dẫn không an toàn trong ZIP)")
        zf.extractall(new_dir)
    source_root = new_dir if (new_dir / APP_EXE).exists() else next(
        (d for d in new_dir.iterdir() if d.is_dir() and (d / APP_EXE).exists()), None)
    if source_root is None:
        raise RuntimeError(f"{MSG['bad_package']} (không có {APP_EXE} sau khi giải nén)")
    manifest = staging / STAGED_MANIFEST
    payload = {"version": info.version, "build": info.build, "package": info.package, "sha256": info.sha256,
               "source_root": str(source_root), "target_root": str(portable_root), "old_pid": os.getpid(),
               "staged_at": datetime.now().isoformat(timespec="seconds"),
               "from_version": __version__, "from_build": BUILD_NUMBER}
    manifest.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    _log_file(portable_root, f"UPDATE_STAGE ok source_root={source_root} manifest={manifest}")
    return StagedUpdate(str(staging), str(local_zip), str(source_root), str(portable_root), str(manifest), info)


def updater_command(staged: StagedUpdate, packaged: Optional[bool] = None) -> List[str]:
    """Command that runs the updater OUT of the staged copy (never the running executable)."""
    packaged = bool(getattr(sys, "frozen", False)) if packaged is None else packaged
    if packaged:
        return [str(Path(staged.source_root) / APP_EXE), "--apply-update", staged.manifest_path]
    run_py = Path(__file__).resolve().parent.parent / "run.py"
    return [sys.executable, str(run_py), "--apply-update", staged.manifest_path]


def launch_updater(staged: StagedUpdate, spawn: Callable = subprocess.Popen) -> List[str]:
    cmd = updater_command(staged)
    kwargs = {"cwd": staged.source_root, "close_fds": True}
    if os.name == "nt":
        kwargs["creationflags"] = 0x00000008 | 0x00000200        # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    spawn(cmd, **kwargs)
    _log_file(Path(staged.target_root), f"UPDATE_LAUNCH updater={' '.join(cmd)}")
    return cmd


# ---------------------------------------------------------------------------- apply (runs in the staged copy)
def process_alive(pid: int) -> bool:
    if not pid or pid <= 0:
        return False
    if os.name == "nt":
        import ctypes
        k32 = ctypes.windll.kernel32                           # type: ignore[attr-defined]
        h = k32.OpenProcess(0x1000, False, int(pid))          # PROCESS_QUERY_LIMITED_INFORMATION
        if not h:
            return False
        try:
            code = ctypes.c_ulong()
            if k32.GetExitCodeProcess(h, ctypes.byref(code)):
                return code.value == 259                        # STILL_ACTIVE
            return False
        finally:
            k32.CloseHandle(h)
    try:
        os.kill(int(pid), 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False


def wait_for_exit(pid: int, timeout: float = 90.0, alive: Callable[[int], bool] = process_alive,
                  sleep: Callable[[float], None] = time.sleep, interval: float = 0.5) -> bool:
    deadline = time.monotonic() + timeout
    while alive(pid):
        if time.monotonic() >= deadline:
            return False
        sleep(interval)
    return True


def is_preserved(name: str) -> bool:
    low = name.lower()
    return low in PRESERVED_ENTRIES or (("." in low) and low.endswith(PRESERVED_SUFFIXES))


def release_entries(source_root: Path) -> List[str]:
    """Top-level names of the new release that may replace the current runtime (user data never included)."""
    return sorted(p.name for p in Path(source_root).iterdir() if not is_preserved(p.name))


def _move(src: Path, dst: Path) -> None:
    if dst.exists():
        shutil.rmtree(dst, ignore_errors=True) if dst.is_dir() else dst.unlink()
    shutil.move(str(src), str(dst))


def _copy(src: Path, dst: Path) -> None:
    if src.is_dir():
        shutil.copytree(src, dst)
    else:
        shutil.copy2(src, dst)


def apply_update(manifest_path: Path, alive: Optional[Callable[[int], bool]] = None,
                 sleep: Callable[[float], None] = time.sleep, spawn: Optional[Callable] = subprocess.Popen,
                 wait_timeout: float = 90.0, restart: bool = True) -> int:
    """Replace the runtime files of ``target_root`` with ``source_root`` (both from the staged manifest).

    0 = installed (+ restarted), 1 = failed and rolled back (previous application restored), 2 = cannot start.
    """
    alive = alive or process_alive
    manifest_path = Path(manifest_path)
    try:
        m = json.loads(manifest_path.read_text(encoding="utf-8"))
        source, target = Path(m["source_root"]), Path(m["target_root"])
        info = UpdateInfo(str(m["version"]), int(m["build"]), str(m["package"]), str(m.get("sha256") or ""))
    except Exception as e:  # noqa: BLE001
        LOG.error("UPDATE_APPLY cannot read manifest %s: %s", manifest_path, e)
        return 2
    log = lambda msg: _log_file(target, msg)  # noqa: E731
    log(f"UPDATE_APPLY start target={target} source={source} to={info.label()} pid_old={m.get('old_pid')}")
    if not (source / APP_EXE).exists() and not (source / "run.py").exists() and not (source / "_internal").exists():
        log("UPDATE_APPLY abort: source_root không chứa bản phát hành")
        _write_result(target, "failed", info, "Thư mục staging không hợp lệ")
        return 2
    if not wait_for_exit(int(m.get("old_pid") or 0), wait_timeout, alive, sleep):
        log("UPDATE_APPLY abort: ứng dụng cũ vẫn đang chạy")
        _write_result(target, "failed", info, "Ứng dụng cũ chưa thoát – không thay thế file")
        return 1
    log("UPDATE_APPLY old process exited")
    backup = target / BACKUP_DIR
    _clear_dir(backup)
    entries = release_entries(source)
    moved: List[str] = []
    copied: List[str] = []
    try:
        for name in entries:
            cur = target / name
            if cur.exists():
                _move(cur, backup / name)
                moved.append(name)
            _copy(source / name, cur)
            copied.append(name)
        log(f"UPDATE_REPLACE ok entries={entries} backup={backup}")
    except Exception as e:  # noqa: BLE001
        log(f"UPDATE_REPLACE failed error={type(e).__name__}: {e} – rollback")
        rb_ok = True
        for name in copied:
            try:
                p = target / name
                if p.exists():
                    shutil.rmtree(p, ignore_errors=True) if p.is_dir() else p.unlink()
            except Exception as e2:  # noqa: BLE001
                rb_ok = False
                log(f"UPDATE_ROLLBACK cannot remove {name}: {e2}")
        for name in moved:
            try:
                _move(backup / name, target / name)
            except Exception as e2:  # noqa: BLE001
                rb_ok = False
                log(f"UPDATE_ROLLBACK cannot restore {name}: {e2}")
        log(f"UPDATE_ROLLBACK result={'ok' if rb_ok else 'INCOMPLETE'}")
        _write_result(target, "rolled_back" if rb_ok else "rollback_failed", info, str(e))
        return 1
    _write_result(target, "ok", info, "")
    if restart and spawn is not None:
        try:
            exe = target / APP_EXE
            cmd = [str(exe)] if exe.exists() else [sys.executable, str(target / "run.py")]
            kwargs = {"cwd": str(target), "close_fds": True}
            if os.name == "nt":
                kwargs["creationflags"] = 0x00000008 | 0x00000200
            spawn(cmd, **kwargs)
            log(f"UPDATE_RESTART ok cmd={cmd}")
        except Exception as e:  # noqa: BLE001
            log(f"UPDATE_RESTART failed {type(e).__name__}: {e}")
    return 0


def _write_result(target: Path, status: str, info: UpdateInfo, detail: str) -> None:
    try:
        (target / "logs").mkdir(parents=True, exist_ok=True)
        (target / "logs" / RESULT_FILE).write_text(json.dumps(
            {"status": status, "version": info.version, "build": info.build, "detail": detail,
             "at": datetime.now().isoformat(timespec="seconds")}, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def consume_result(portable_root: Path) -> Optional[dict]:
    """Read + delete the one-shot result written by the updater (GUI shows 'Cập nhật thành công …' once)."""
    p = Path(portable_root) / "logs" / RESULT_FILE
    if not p.exists():
        return None
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        data = None
    try:
        p.unlink()
    except OSError:
        pass
    return data


def cleanup_staging(portable_root: Path) -> None:
    """Best-effort removal of update_staging/ (the updater exe may still be locked right after a restart)."""
    shutil.rmtree(Path(portable_root) / STAGING_DIR, ignore_errors=True)


def _log_file(portable_root: Path, message: str) -> None:
    LOG.info(message)
    try:
        d = Path(portable_root) / "logs"
        d.mkdir(parents=True, exist_ok=True)
        with open(d / UPDATE_LOG, "a", encoding="utf-8") as fh:
            fh.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
    except OSError:
        pass


def staged_to_dict(s: StagedUpdate) -> dict:
    d = asdict(s)
    d["info"] = asdict(s.info)
    return d

