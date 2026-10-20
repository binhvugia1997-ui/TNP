"""Validated offline/LAN self-update for the Windows Portable build.

The configured folder is read-only update input (``version.json`` and the release ZIP). Installation is
explicit and only available to a verified frozen Portable runtime. The live installation is never used as
an extraction workspace: the archive is copied to a unique local stage, hashed, securely extracted, and
validated before the separate staged EXE is launched in ``--apply-update`` helper mode.

The helper waits for the old process, takes the inter-process update lock, creates and verifies a complete
application-only backup, replaces only the documented package-owned root entries, validates the result,
and restarts the new EXE. Any failed replacement/restart attempts a bounded rollback while leaving runtime
and user data outside the package-owned allowlist untouched.
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import shutil
import stat
import subprocess
import sys
import time
import unicodedata
import uuid
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath, PureWindowsPath
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from . import BUILD_NUMBER, __version__, format_build
from .runtime_paths import is_packaged

LOG = logging.getLogger("report_extractor.update")

APP_EXE = "ReportExtractor.exe"
MANIFEST_NAME = "version.json"
STAGING_DIR = "update_staging"
BACKUP_DIR = "update_backup"
LOCK_NAME = "update_install.lock"
UPDATE_LOG = "update.log"
RESULT_FILE = "update_result.json"
STAGED_MANIFEST = "update.json"
MAX_PACKAGE_BYTES = 8 * 1024 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 16 * 1024 * 1024 * 1024
MAX_MEMBER_BYTES = 8 * 1024 * 1024 * 1024
MAX_ARCHIVE_MEMBERS = 50_000
MAX_COMPRESSION_RATIO = 1_000
MAX_MANIFEST_BYTES = 64 * 1024
MAX_BACKUP_MANIFEST_BYTES = 32 * 1024 * 1024
MAX_VERSION_FILE_BYTES = 16 * 1024
MAX_BACKUPS = 2
STAGE_NAME_RE = re.compile(r"^stage-(\d{3,})-([0-9a-f]{32})$")
BACKUP_NAME_RE = re.compile(r"^backup-(\d{3,})-([0-9a-f]{32})$")
VERSION_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9.+_-]{0,63}$")
PACKAGE_RE = re.compile(r"^ReportExtractor_[A-Za-z0-9][A-Za-z0-9.+_-]*\.zip$", re.IGNORECASE)

# Exact Portable onedir entries produced by ReportExtractor.spec + tools/build_portable.py. Replacement is
# allowlist-based: unknown root entries are never moved or deleted, even if they share a package suffix.
PACKAGE_OWNED_TOP_LEVEL = frozenset({
    APP_EXE, "_internal", "frontend", "VERSION.txt", "README.txt", "FIRST_RUN.txt",
    "Install_Ollama_Optional.bat", "SHA256SUMS.txt",
})
_PACKAGE_OWNED_BY_CASE = {name.casefold(): name for name in PACKAGE_OWNED_TOP_LEVEL}
# These empty marker directories are emitted by the builder but are user/runtime-owned at installation time.
PACKAGE_SCAFFOLD_DIRS = frozenset({"config", "logs", "output"})
PRESERVED_ENTRIES = frozenset(name.casefold() for name in (
    "config", "logs", "Output", "backup", "config.json", STAGING_DIR, BACKUP_DIR, LOCK_NAME,
    "history.json", "learning_data", "manual_fields.json",
))
PRESERVED_SUFFIXES = (".xlsx", ".xlsm", ".xls", ".pptx", ".ppt", ".log", ".json")

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
    "development_only": "Chỉ có thể cài đặt cập nhật từ ứng dụng Portable đã đóng gói; môi trường phát triển chỉ có thể kiểm tra.",
    "busy": "Không thể cập nhật khi một tác vụ đang chạy. Hãy đợi tác vụ hoàn tất rồi thử lại.",
    "already_installing": "Đang cập nhật – vui lòng đợi.",
    "lock_failed": "Không thể khóa an toàn thư mục Portable; hãy kiểm tra quyền truy cập hiện tại.",
}


def version_label(version: str = __version__, build: int = BUILD_NUMBER) -> str:
    return f"{version} — Build {format_build(build)}"


class ManifestError(ValueError):
    """version.json missing or invalid – message is safe to return to the user."""


def _safe_manifest_package(package: str, version: str) -> bool:
    if not package or "\x00" in package or "/" in package or "\\" in package or ":" in package:
        return False
    if Path(package).name != package or PurePosixPath(package).name != package or PureWindowsPath(package).name != package:
        return False
    if PureWindowsPath(package).is_absolute() or PureWindowsPath(package).drive:
        return False
    if package in (".", "..") or not PACKAGE_RE.fullmatch(package):
        return False
    return package.casefold() == f"ReportExtractor_{version}.zip".casefold()


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    build: int
    package: str
    sha256: str
    size: Optional[int] = None

    def label(self) -> str:
        return version_label(self.version, self.build)


def parse_manifest(text: str) -> UpdateInfo:
    try:
        data = json.loads(text)
    except Exception as exc:  # noqa: BLE001
        raise ManifestError(f"{MSG['invalid_manifest']} (JSON lỗi: {type(exc).__name__})") from exc
    if not isinstance(data, dict):
        raise ManifestError(f"{MSG['invalid_manifest']} (không phải đối tượng JSON)")
    required = ("version", "build", "package", "sha256")
    missing = [key for key in required if key not in data]
    if missing:
        raise ManifestError(f"{MSG['invalid_manifest']} (thiếu {', '.join(missing)})")

    version_value = data["version"]
    if not isinstance(version_value, str):
        raise ManifestError(f"{MSG['invalid_manifest']} (version không hợp lệ)")
    version = version_value.strip()
    if not VERSION_RE.fullmatch(version):
        raise ManifestError(f"{MSG['invalid_manifest']} (version không hợp lệ)")

    build_value = data["build"]
    if isinstance(build_value, bool) or not isinstance(build_value, (int, str)):
        raise ManifestError(f"{MSG['invalid_manifest']} (build phải là số nguyên)")
    if isinstance(build_value, str) and not build_value.strip().isdigit():
        raise ManifestError(f"{MSG['invalid_manifest']} (build phải là số nguyên)")
    try:
        build = int(build_value)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ManifestError(f"{MSG['invalid_manifest']} (build phải là số nguyên)") from exc
    if build < 1:
        raise ManifestError(f"{MSG['invalid_manifest']} (build phải lớn hơn 0)")

    package_value = data["package"]
    if not isinstance(package_value, str):
        raise ManifestError(f"{MSG['invalid_manifest']} (package không hợp lệ)")
    package = package_value.strip()
    if not _safe_manifest_package(package, version):
        raise ManifestError(f"{MSG['invalid_manifest']} (tên package không an toàn)")

    sha_value = data["sha256"]
    if not isinstance(sha_value, str):
        raise ManifestError(f"{MSG['invalid_manifest']} (sha256 không hợp lệ)")
    sha = sha_value.strip().lower()
    if len(sha) != 64 or any(ch not in "0123456789abcdef" for ch in sha):
        raise ManifestError(f"{MSG['invalid_manifest']} (sha256 không hợp lệ)")

    size: Optional[int] = None
    if "size" in data:
        size_value = data["size"]
        if isinstance(size_value, bool) or not isinstance(size_value, (int, str)):
            raise ManifestError(f"{MSG['invalid_manifest']} (size không hợp lệ)")
        if isinstance(size_value, str) and not size_value.strip().isdigit():
            raise ManifestError(f"{MSG['invalid_manifest']} (size không hợp lệ)")
        try:
            size = int(size_value)
        except (TypeError, ValueError, OverflowError) as exc:
            raise ManifestError(f"{MSG['invalid_manifest']} (size không hợp lệ)") from exc
        if size <= 0 or size > MAX_PACKAGE_BYTES:
            raise ManifestError(f"{MSG['invalid_manifest']} (size ngoài giới hạn cho phép)")

    return UpdateInfo(version=version, build=build, package=package, sha256=sha, size=size)


def compare_builds(current: int, remote: int) -> str:
    """Numeric ordering only: available (remote > current), latest (=), older (<)."""
    current, remote = int(current), int(remote)
    if remote > current:
        return "available"
    if remote == current:
        return "latest"
    return "older"


def _is_reparse_or_link(path: Path) -> bool:
    try:
        info = Path(path).lstat()
    except OSError:
        return False
    if stat.S_ISLNK(info.st_mode):
        return True
    attributes = getattr(info, "st_file_attributes", 0)
    reparse = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)
    return bool(attributes & reparse)


def _regular_file(path: Path) -> bool:
    try:
        mode = Path(path).lstat().st_mode
    except OSError:
        return False
    return stat.S_ISREG(mode) and not _is_reparse_or_link(Path(path))


def _directory(path: Path) -> bool:
    try:
        mode = Path(path).lstat().st_mode
    except OSError:
        return False
    return stat.S_ISDIR(mode) and not _is_reparse_or_link(Path(path))


@dataclass
class UpdateCheck:
    status: str
    message: str
    update_path: str = ""
    info: Optional[UpdateInfo] = None
    package_path: str = ""
    checked_at: str = ""

    @property
    def available(self) -> bool:
        return self.status == "available"


def _folder_state(path: Path) -> str:
    try:
        if path.is_dir():
            os.listdir(path)
            return "ok"
        return "inaccessible"
    except PermissionError:
        return "no_permission"
    except OSError as exc:
        if getattr(exc, "winerror", None) in (5, 1326, 1327, 1331, 1385):
            return "no_permission"
        return "inaccessible"


def _safe_child_file(folder: Path, name: str) -> Optional[Path]:
    """Return an ordinary, direct child file. Never follow a package symlink or a path from the manifest."""
    try:
        root = Path(folder).resolve(strict=True)
        candidate = Path(folder) / name
        if candidate.parent.resolve(strict=True) != root or not _regular_file(candidate):
            return None
        resolved = candidate.resolve(strict=True)
        if resolved.parent != root:
            return None
        return candidate
    except OSError:
        return None


def check_for_update(update_path: str, current_build: int = BUILD_NUMBER,
                     current_version: str = __version__) -> UpdateCheck:
    """Read-only configured-folder check; it never copies or extracts the release package."""
    LOG.info("UPDATE_CHECK_START")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    raw = (update_path or "").strip().strip('"')
    if not raw:
        result = UpdateCheck("no_path", MSG["no_path"], "", checked_at=now)
        LOG.info("UPDATE_CHECK_RESULT status=no_path")
        return result

    folder = Path(raw)
    folder_state = _folder_state(folder)
    if folder_state != "ok":
        result = UpdateCheck(folder_state, MSG[folder_state], raw, checked_at=now)
        LOG.info("UPDATE_CHECK_RESULT status=%s", folder_state)
        return result

    manifest = folder / MANIFEST_NAME
    if not _regular_file(manifest):
        result = UpdateCheck("no_manifest", MSG["no_manifest"], raw, checked_at=now)
        LOG.info("UPDATE_CHECK_RESULT status=no_manifest")
        return result
    try:
        if manifest.stat().st_size > MAX_MANIFEST_BYTES:
            raise ManifestError(f"{MSG['invalid_manifest']} (kích thước vượt giới hạn)")
        with manifest.open("rb") as handle:
            manifest_bytes = handle.read(MAX_MANIFEST_BYTES + 1)
        if len(manifest_bytes) > MAX_MANIFEST_BYTES:
            raise ManifestError(f"{MSG['invalid_manifest']} (kích thước vượt giới hạn)")
        try:
            manifest_text = manifest_bytes.decode("utf-8-sig")
        except UnicodeError as exc:
            raise ManifestError(f"{MSG['invalid_manifest']} (mã hóa không hợp lệ)") from exc
        info = parse_manifest(manifest_text)
    except ManifestError as exc:
        LOG.warning("UPDATE_CHECK_RESULT status=invalid_manifest reason=%s", exc)
        return UpdateCheck("invalid_manifest", str(exc), raw, checked_at=now)
    except PermissionError:
        LOG.warning("UPDATE_CHECK_RESULT status=no_permission")
        return UpdateCheck("no_permission", MSG["no_permission"], raw, checked_at=now)
    except OSError:
        LOG.warning("UPDATE_CHECK_RESULT status=inaccessible")
        return UpdateCheck("inaccessible", MSG["inaccessible"], raw, checked_at=now)

    package_path = _safe_child_file(folder, info.package)
    if package_path is None:
        result = UpdateCheck("package_missing", f"{MSG['package_missing']} {info.package}.", raw,
                             info, checked_at=now)
        LOG.info("UPDATE_CHECK_RESULT status=package_missing")
        return result
    try:
        size = package_path.stat().st_size
    except OSError:
        size = 0
    if size <= 0 or size > MAX_PACKAGE_BYTES or (info.size is not None and size != info.size):
        result = UpdateCheck("invalid_manifest", f"{MSG['invalid_manifest']} (kích thước package không khớp)",
                             raw, info, checked_at=now)
        LOG.warning("UPDATE_CHECK_RESULT status=package_size_mismatch")
        return result

    relation = compare_builds(current_build, info.build)
    message = f"{MSG['available']}: {info.label()}" if relation == "available" else MSG[relation]
    LOG.info("UPDATE_CHECK_RESULT status=%s current_build=%s remote_build=%s", relation,
             format_build(current_build), format_build(info.build))
    return UpdateCheck(relation, message, raw, info, str(package_path), now)


def sha256_file(path: Path, chunk: int = 1 << 20) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(chunk), b""):
            digest.update(block)
    return digest.hexdigest()


@dataclass(frozen=True)
class _ZipMember:
    info: zipfile.ZipInfo
    parts: Tuple[str, ...]
    is_dir: bool


def _win_component_is_reserved(component: str) -> bool:
    base = component.split(".", 1)[0].rstrip(" .").casefold()
    if base in {"con", "prn", "aux", "nul"}:
        return True
    return bool(re.fullmatch(r"(?:com|lpt)[1-9]", base))


def _parse_zip_member(info: zipfile.ZipInfo) -> _ZipMember:
    raw = info.filename
    if not raw or "\x00" in raw or "\\" in raw:
        raise ValueError("unsafe archive name")
    is_dir = raw.endswith("/")
    name = raw[:-1] if is_dir else raw
    if not name or name.startswith("/") or name.startswith("//") or re.match(r"^[A-Za-z]:", name):
        raise ValueError("absolute or drive-qualified archive name")
    if PureWindowsPath(name).drive or PureWindowsPath(name).is_absolute() or PurePosixPath(name).is_absolute():
        raise ValueError("absolute archive name")
    raw_parts = name.split("/")
    if any(part in ("", ".", "..") for part in raw_parts):
        raise ValueError("unsafe archive component")
    parts: List[str] = []
    for part in raw_parts:
        normalized = unicodedata.normalize("NFC", part)
        if (normalized != part or part.endswith((".", " "))
                or len(part.encode("utf-16-le")) // 2 > 255
                or any(ch in part for ch in '<>:"|?*')):
            raise ValueError("non-canonical Windows archive component")
        if _win_component_is_reserved(part):
            raise ValueError("reserved Windows archive component")
        parts.append(normalized)

    mode = (info.external_attr >> 16) & 0xFFFF
    kind = stat.S_IFMT(mode)
    if kind == stat.S_IFLNK:
        raise ValueError("archive symlink is not allowed")
    if kind not in (0, stat.S_IFREG, stat.S_IFDIR):
        raise ValueError("unsupported archive file type")
    dos_attributes = info.external_attr & 0xFFFF
    if dos_attributes & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400):
        raise ValueError("archive reparse point is not allowed")
    if is_dir and kind == stat.S_IFREG:
        raise ValueError("directory metadata mismatch")
    if not is_dir and kind == stat.S_IFDIR:
        raise ValueError("file metadata mismatch")
    if info.flag_bits & 0x1:
        raise ValueError("encrypted archive member")
    if info.file_size < 0 or info.file_size > MAX_MEMBER_BYTES:
        raise ValueError("archive member exceeds size limit")
    if info.compress_size < 0:
        raise ValueError("invalid compressed size")
    if info.file_size and info.compress_size == 0:
        raise ValueError("invalid compressed size")
    if info.compress_size and info.file_size / info.compress_size > MAX_COMPRESSION_RATIO:
        raise ValueError("archive compression ratio exceeds limit")
    if is_dir and info.file_size != 0:
        raise ValueError("directory contains data")
    return _ZipMember(info=info, parts=tuple(parts), is_dir=is_dir)


def _archive_plan(zf: zipfile.ZipFile) -> Tuple[List[_ZipMember], Tuple[str, ...]]:
    infos = zf.infolist()
    if not infos or len(infos) > MAX_ARCHIVE_MEMBERS:
        raise ValueError("empty archive or too many members")
    plan: List[_ZipMember] = []
    canonical: Dict[str, _ZipMember] = {}
    expanded = 0
    exe_entries: List[_ZipMember] = []
    for info in infos:
        member = _parse_zip_member(info)
        key = "/".join(part.casefold() for part in member.parts)
        if key in canonical:
            raise ValueError("duplicate or case-colliding archive member")
        canonical[key] = member
        expanded += info.file_size
        if expanded > MAX_UNCOMPRESSED_BYTES:
            raise ValueError("archive exceeds expanded size limit")
        if member.parts[-1].casefold() == APP_EXE.casefold() and not member.is_dir:
            exe_entries.append(member)
        plan.append(member)
    if len(exe_entries) != 1 or len(exe_entries[0].parts) not in (1, 2):
        raise ValueError("archive must contain exactly one root Portable executable")
    prefix = exe_entries[0].parts[:-1]
    if prefix:
        for member in plan:
            if member.parts[0].casefold() != prefix[0].casefold():
                raise ValueError("archive contains entries outside its Portable root")
            if len(member.parts) == 1 and not member.is_dir:
                raise ValueError("archive contains a file beside its Portable root")
    elif any(len(member.parts) > 1 and member.parts[0].casefold() in _PACKAGE_OWNED_BY_CASE
             for member in plan):
        # root-level packages are valid; this branch intentionally does not disallow descendants.
        pass

    root_component_count = len(prefix)
    relative_members = []
    for member in plan:
        relative = member.parts[root_component_count:]
        if not relative:
            continue
        first = relative[0].casefold()
        if first in PACKAGE_SCAFFOLD_DIRS:
            # The builder emits only empty .keep markers in these runtime-owned folders.
            if len(relative) == 1 and member.is_dir:
                continue
            if len(relative) == 2 and relative[1] == ".keep" and not member.is_dir and member.info.file_size == 0:
                continue
            raise ValueError("archive contains runtime/user data")
        if first not in _PACKAGE_OWNED_BY_CASE:
            raise ValueError("archive contains an unexpected Portable root entry")
        relative_members.append(member)
    return plan, prefix


def _version_fields(text: str) -> Tuple[str, int]:
    fields: Dict[str, str] = {}
    for line in text.lstrip("\ufeff").splitlines():
        if "=" in line:
            key, value = line.split("=", 1)
            fields[key.strip().casefold()] = value.strip()
    version, build = fields.get("version", ""), fields.get("build", "")
    if not version or not build.isdigit():
        raise ValueError("VERSION metadata is missing version/build")
    return version, int(build)


def _read_version_fields(path: Path) -> Tuple[str, int]:
    path = Path(path)
    if not _regular_file(path) or path.stat().st_size > MAX_VERSION_FILE_BYTES:
        raise ValueError("VERSION metadata is not a bounded ordinary file")
    with path.open("rb") as handle:
        raw = handle.read(MAX_VERSION_FILE_BYTES + 1)
    if len(raw) > MAX_VERSION_FILE_BYTES:
        raise ValueError("VERSION metadata exceeds the allowed size")
    return _version_fields(raw.decode("utf-8-sig"))


def validate_staged_portable(root: Path, info: UpdateInfo, *, windows_runtime: Optional[bool] = None) -> Tuple[bool, str]:
    """Validate the actual PyInstaller onedir + React bundle layout emitted by the Portable builder."""
    root = Path(root)
    try:
        if not _directory(root):
            return False, "thư mục staged không hợp lệ"
        for item in root.rglob("*"):
            if _is_reparse_or_link(item) or not (_directory(item) or _regular_file(item)):
                return False, "cây Portable chứa file link/reparse hoặc kiểu file không an toàn"
        exe = root / APP_EXE
        internal = root / "_internal"
        version_file = root / "VERSION.txt"
        if not _regular_file(exe) or exe.stat().st_size <= 0:
            return False, f"thiếu hoặc hỏng {APP_EXE}"
        with exe.open("rb") as handle:
            if handle.read(2) != b"MZ":
                return False, f"{APP_EXE} không phải Windows executable"
        if not _directory(internal) or not any(internal.iterdir()):
            return False, "thiếu runtime _internal của Portable"
        if not _regular_file(version_file):
            return False, "thiếu VERSION.txt"
        version, build = _read_version_fields(version_file)
        if version != info.version or build != info.build:
            return False, "VERSION.txt không khớp manifest version/build"

        for base in (root / "frontend" / "dist", internal / "frontend" / "dist"):
            if not _regular_file(base / "index.html"):
                return False, "thiếu frontend/dist/index.html"
            assets = base / "assets"
            if not _directory(assets):
                return False, "thiếu frontend/dist/assets"
            files = [p for p in assets.iterdir() if _regular_file(p)]
            if not any(p.suffix.casefold() == ".js" for p in files):
                return False, "thiếu JavaScript frontend bundle"
            if not any(p.suffix.casefold() == ".css" for p in files):
                return False, "thiếu CSS frontend bundle"
        for name in ("api.js", "lib/dom_json.js", "state.js"):
            if not _regular_file(internal / "webview" / "js" / name):
                return False, "thiếu pywebview JS bridge runtime"

        use_windows = sys.platform == "win32" if windows_runtime is None else windows_runtime
        if use_windows:
            python_runtime = internal / "pythonnet" / "runtime" / "Python.Runtime.dll"
            clr_loader = internal / "clr_loader" / "ffi" / "dlls"
            webview_core = internal / "webview" / "lib" / "Microsoft.Web.WebView2.Core.dll"
            webview_forms = internal / "webview" / "lib" / "Microsoft.Web.WebView2.WinForms.dll"
            loader = internal / "webview" / "lib" / "runtimes" / "win-x64" / "native" / "WebView2Loader.dll"
            clr_dlls = list(clr_loader.glob("*/ClrLoader.dll")) if _directory(clr_loader) else []
            if not _regular_file(python_runtime) or not any(_regular_file(p) for p in clr_dlls):
                return False, "thiếu .NET/pythonnet runtime bắt buộc"
            if not all(_regular_file(p) for p in (webview_core, webview_forms, loader)):
                return False, "thiếu WebView2 runtime bắt buộc"
        return True, "OK"
    except (OSError, ValueError, UnicodeError) as exc:
        return False, f"layout/version metadata không hợp lệ ({type(exc).__name__})"


def _verify_zip(zip_path: Path, info: UpdateInfo) -> Tuple[bool, str, Tuple[str, ...]]:
    try:
        if not _regular_file(zip_path):
            return False, f"{MSG['bad_package']} (package không phải file thường)", ()
        size = zip_path.stat().st_size
        if size <= 0 or size > MAX_PACKAGE_BYTES or (info.size is not None and size != info.size):
            return False, f"{MSG['bad_package']} (kích thước package không hợp lệ)", ()
        if not info.sha256 or sha256_file(zip_path) != info.sha256:
            LOG.error("UPDATE_VERIFY_FAILED stage=UPDATE_VERIFY reason=sha256_mismatch")
            return False, MSG["bad_package"], ()
        if not zipfile.is_zipfile(zip_path):
            return False, f"{MSG['bad_package']} (không phải file ZIP)", ()
        with zipfile.ZipFile(zip_path) as zf:
            plan, prefix = _archive_plan(zf)
            bad_member = zf.testzip()
            if bad_member is not None:
                return False, f"{MSG['bad_package']} (ZIP hỏng)", ()
            relative: Dict[str, _ZipMember] = {}
            for member in plan:
                rel_parts = member.parts[len(prefix):]
                if rel_parts:
                    relative["/".join(rel_parts)] = member
            exe = relative.get(APP_EXE)
            version_entry = relative.get("VERSION.txt")
            internal = relative.get("_internal")
            frontend = relative.get("frontend")
            if exe is None or exe.is_dir or version_entry is None or version_entry.is_dir:
                return False, f"{MSG['bad_package']} (thiếu {APP_EXE} hoặc VERSION.txt)", ()
            if internal is None or not internal.is_dir or frontend is None or not frontend.is_dir:
                return False, f"{MSG['bad_package']} (thiếu runtime _internal/frontend)", ()
            try:
                if version_entry.info.file_size > 16 * 1024:
                    return False, f"{MSG['bad_package']} (VERSION.txt vượt giới hạn kích thước)", ()
                with zf.open(version_entry.info, "r") as version_stream:
                    version_text = version_stream.read(16 * 1024 + 1).decode("utf-8-sig")
                if len(version_text) > 16 * 1024:
                    return False, f"{MSG['bad_package']} (VERSION.txt vượt giới hạn kích thước)", ()
                found_version, found_build = _version_fields(version_text)
            except (KeyError, UnicodeError, ValueError):
                return False, f"{MSG['bad_package']} (VERSION.txt không hợp lệ)", ()
            if found_version != info.version or found_build != info.build:
                return False, f"{MSG['bad_package']} (VERSION.txt không khớp version/build)", ()
            required = (
                "frontend/dist/index.html", "_internal/frontend/dist/index.html",
                "_internal/webview/js/api.js", "_internal/webview/js/lib/dom_json.js", "_internal/webview/js/state.js",
            )
            if any(name not in relative or relative[name].is_dir for name in required):
                return False, f"{MSG['bad_package']} (thiếu bundle frontend/bridge)", ()
            for base in ("frontend/dist/assets", "_internal/frontend/dist/assets"):
                if not any(key.startswith(base + "/") and key.casefold().endswith(".js") and not entry.is_dir
                           for key, entry in relative.items()):
                    return False, f"{MSG['bad_package']} (thiếu JavaScript frontend bundle)", ()
                if not any(key.startswith(base + "/") and key.casefold().endswith(".css") and not entry.is_dir
                           for key, entry in relative.items()):
                    return False, f"{MSG['bad_package']} (thiếu CSS frontend bundle)", ()
            with zf.open(exe.info, "r") as executable_stream:
                exe_bytes = executable_stream.read(2)
            if exe_bytes != b"MZ":
                return False, f"{MSG['bad_package']} ({APP_EXE} không phải Windows executable)", ()
            return True, "OK", prefix
    except (OSError, zipfile.BadZipFile, ValueError, RuntimeError) as exc:
        return False, f"{MSG['bad_package']} ({type(exc).__name__})", ()


def verify_package(zip_path: Path, info: UpdateInfo) -> Tuple[bool, str]:
    """Verify local staged ZIP checksum, size, archive paths, structure, and VERSION identity."""
    ok, message, _prefix = _verify_zip(Path(zip_path), info)
    return ok, message


class UpdateInstallLock:
    """Cross-process exclusive lock; the old app holds it through handoff, then the helper acquires it."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self._file = None
        self.acquired = False

    def acquire(self, timeout: float = 0.0, *, sleep: Callable[[float], None] = time.sleep,
                clock: Callable[[], float] = time.monotonic, write_owner: bool = True) -> bool:
        if self.acquired:
            return True
        path = self.path
        path.parent.mkdir(parents=True, exist_ok=True)
        if _is_reparse_or_link(path):
            raise OSError("unsafe update lock path")
        flags = os.O_RDWR | os.O_CREAT
        flags |= getattr(os, "O_BINARY", 0)
        flags |= getattr(os, "O_NOFOLLOW", 0)
        fd = os.open(path, flags, 0o600)
        try:
            opened = os.fstat(fd)
            if not stat.S_ISREG(opened.st_mode) or opened.st_nlink != 1:
                raise OSError("unsafe update lock file")
            if _is_reparse_or_link(path):
                raise OSError("unsafe update lock path")
            self._file = os.fdopen(fd, "r+b", buffering=0)
            fd = -1
        finally:
            if fd >= 0:
                os.close(fd)
        self._file.seek(0, os.SEEK_END)
        if self._file.tell() == 0:
            self._file.write(b"\0")
            self._file.flush()
        try:
            bounded_timeout = min(90.0, max(0.0, float(timeout)))
        except (TypeError, ValueError, OverflowError) as exc:
            self._file.close()
            self._file = None
            raise OSError("invalid update lock timeout") from exc
        deadline = clock() + bounded_timeout
        waited = 0.0
        while True:
            try:
                self._file.seek(0)
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(self._file.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(self._file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except (BlockingIOError, OSError):
                remaining = min(deadline - clock(), bounded_timeout - waited)
                if remaining <= 0:
                    self._file.close()
                    self._file = None
                    return False
                delay = min(0.1, remaining)
                try:
                    sleep(delay)
                except Exception:  # noqa: BLE001
                    self._file.close()
                    self._file = None
                    return False
                waited += delay
                continue

            self.acquired = True
            if not write_owner:
                return True
            try:
                self._file.seek(0)
                self._file.truncate()
                self._file.write(json.dumps({"pid": os.getpid(), "at": datetime.now(timezone.utc).isoformat()},
                                            separators=(",", ":")).encode("utf-8"))
                self._file.flush()
                try:
                    os.fsync(self._file.fileno())
                except OSError:
                    pass
                return True
            except OSError:
                self.release()
                raise

    def release(self) -> None:
        if not self._file:
            return
        try:
            if self.acquired:
                self._file.seek(0)
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(self._file.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(self._file.fileno(), fcntl.LOCK_UN)
        except OSError:
            LOG.warning("UPDATE_LOCK_RELEASE_FAILED")
        finally:
            self.acquired = False
            self._file.close()
            self._file = None

    def __enter__(self):
        if not self.acquire():
            raise RuntimeError(MSG["already_installing"])
        return self

    def __exit__(self, *_exc):
        self.release()


def acquire_update_lock(portable_root: Path, timeout: float = 0.0, *, write_owner: bool = True
                        ) -> Optional[UpdateInstallLock]:
    lock = UpdateInstallLock(Path(portable_root) / LOCK_NAME)
    return lock if lock.acquire(timeout=timeout, write_owner=write_owner) else None


def update_lock_problem(portable_root: Path) -> str:
    """Probe the shared lock without owner writes, distinguishing contention from access/path failure."""
    try:
        lock = acquire_update_lock(portable_root, timeout=0, write_owner=False)
    except OSError:
        return MSG["lock_failed"]
    if lock is None:
        return MSG["already_installing"]
    lock.release()
    return ""


def update_lock_available(portable_root: Path) -> bool:
    """Non-destructively probe the shared scan/process/install lock for UI capability state."""
    return not update_lock_problem(portable_root)


def portable_root_layout_matches(portable_root: Path, *, version: str, build: int) -> bool:
    """Check an installed/staged Portable root without assuming it is this process's executable directory."""
    try:
        root = Path(portable_root).resolve(strict=True)
        executable = root / APP_EXE
        internal = root / "_internal"
        version_file = root / "VERSION.txt"
        if not _regular_file(executable) or executable.stat().st_size <= 0 or not _directory(internal):
            return False
        with executable.open("rb") as handle:
            if handle.read(2) != b"MZ":
                return False
        if not _regular_file(version_file):
            return False
        found_version, found_build = _read_version_fields(version_file)
        required = (
            root / "frontend" / "dist" / "index.html",
            internal / "frontend" / "dist" / "index.html",
            internal / "webview" / "js" / "api.js",
        )
        return (found_version == version and found_build == int(build)
                and all(_regular_file(path) for path in required))
    except (OSError, ValueError, UnicodeError):
        return False


def _is_local_portable_root(root: Path) -> bool:
    """Reject UNC and mapped-network Windows installs: sibling staging must be on local storage."""
    if os.name != "nt":
        return True
    try:
        windows_path = PureWindowsPath(str(root))
        drive = windows_path.drive
        raw_path = str(root).replace("/", "\\")
        if raw_path.casefold().startswith("\\\\?\\unc\\"):
            return False
        if not drive or (drive.startswith("\\\\") and not drive.startswith("\\\\?\\")):
            return False
        if drive.startswith("\\\\?\\"):
            drive_letter = drive[4:]
            if len(drive_letter) != 2 or not drive_letter[0].isalpha() or drive_letter[1] != ":":
                return False
        drive_root = drive + "\\"
        import ctypes
        return int(ctypes.windll.kernel32.GetDriveTypeW(drive_root)) != 4  # DRIVE_REMOTE
    except Exception:  # noqa: BLE001 - inability to prove local storage fails closed
        return False


def production_install_capability(portable_root: Path, *, version: str = __version__, build: int = BUILD_NUMBER
                                 ) -> Tuple[bool, str]:
    """Fail-closed production gate pinned to a frozen EXE, local disk, and exact Portable layout/root."""
    if not is_packaged():
        return False, MSG["development_only"]
    try:
        resolved_root = Path(portable_root).resolve(strict=True)
        executable = Path(sys.executable).resolve(strict=True)
        if executable.name.casefold() != APP_EXE.casefold() or executable.parent != resolved_root:
            return False, "Không xác định được thư mục Portable đang chạy an toàn."
        if not _is_local_portable_root(resolved_root):
            return False, "Chỉ có thể cập nhật Portable đang chạy từ ổ đĩa cục bộ; không thay thế bản chạy từ mạng."
        if not portable_root_layout_matches(resolved_root, version=version, build=build):
            return False, "Portable hiện tại không có layout/version metadata hợp lệ; không thể cài đặt an toàn."
        return True, ""
    except (OSError, ValueError, UnicodeError):
        return False, "Không xác định được thư mục Portable đang chạy an toàn."


@dataclass
class StagedUpdate:
    staging_dir: str
    zip_path: str
    source_root: str
    target_root: str
    manifest_path: str
    info: UpdateInfo
    stage_id: str = ""


STAGE_PREPARING, STAGE_COPYING, STAGE_VERIFYING, STAGE_EXTRACTING, STAGE_VALIDATING, STAGE_READY = (
    "PREPARING", "COPYING", "VERIFYING", "EXTRACTING", "VALIDATING", "READY")
COPY_CHUNK = 1 << 20


def copy_with_progress(src: Path, dst: Path, progress: Optional[Callable[[int, int, float], None]] = None,
                       chunk: int = COPY_CHUNK) -> int:
    src, dst = Path(src), Path(dst)
    if not _regular_file(src):
        raise OSError("update package source is not a regular file")
    total = src.stat().st_size
    if total <= 0 or total > MAX_PACKAGE_BYTES:
        raise OSError("update package size is outside the allowed range")
    copied = 0
    if progress:
        progress(0, total, 0.0)
    with src.open("rb") as source, dst.open("xb") as target:
        while True:
            block = source.read(chunk)
            if not block:
                break
            target.write(block)
            copied += len(block)
            if copied > total or copied > MAX_PACKAGE_BYTES:
                raise OSError("update package changed during copy")
            if progress:
                progress(copied, total, copied * 100.0 / total)
        target.flush()
        try:
            os.fsync(target.fileno())
        except OSError:
            pass
    if copied != total or src.stat().st_size != total or dst.stat().st_size != total:
        raise OSError("copied byte count does not match the source")
    if progress:
        progress(copied, total, 100.0)
    return copied


def _safe_managed_dir(path: Path, *, create: bool = False) -> Path:
    path = Path(path)
    if path.exists() or _is_reparse_or_link(path):
        if not _directory(path):
            raise RuntimeError("unsafe managed update directory")
    elif create:
        path.mkdir(parents=True, exist_ok=False)
    else:
        raise RuntimeError("managed update directory is missing")
    return path.resolve(strict=True)


def _remove_tree_owned(path: Path) -> None:
    path = Path(path)
    if not path.exists() and not _is_reparse_or_link(path):
        return
    if _is_reparse_or_link(path):
        raise RuntimeError("refusing to remove a link/reparse point")
    if not _directory(path):
        path.unlink()
        return
    for child in path.iterdir():
        if _is_reparse_or_link(child):
            raise RuntimeError("refusing to remove a tree containing a link/reparse point")
        if _directory(child):
            _remove_tree_owned(child)
        else:
            child.unlink()
    path.rmdir()


def _clean_stage_dirs_under_lock(portable_root: Path, *, keep: Optional[str] = None,
                                 older_than_seconds: float = 0.0) -> int:
    parent = Path(portable_root) / STAGING_DIR
    if not parent.exists():
        return 0
    if not _directory(parent):
        raise RuntimeError("unsafe update staging root")
    now = time.time()
    removed = 0
    for child in parent.iterdir():
        if child.name == keep or not STAGE_NAME_RE.fullmatch(child.name) or not _directory(child):
            continue
        try:
            if now - child.stat().st_mtime < older_than_seconds:
                continue
            _remove_tree_owned(child)
            removed += 1
        except OSError:
            LOG.warning("UPDATE_STAGE_CLEANUP_FAILED")
    return removed


def _relative_source_root(new_dir: Path, prefix: Sequence[str]) -> Path:
    return Path(new_dir).joinpath(*prefix) if prefix else Path(new_dir)


def _extract_archive(zip_path: Path, destination: Path, info: UpdateInfo,
                     progress: Optional[Callable[[], None]] = None) -> Tuple[Path, int]:
    ok, message, prefix = _verify_zip(zip_path, info)
    if not ok:
        raise RuntimeError(message)
    destination = Path(destination)
    if destination.exists():
        raise RuntimeError("extraction destination already exists")
    destination.mkdir(parents=True, exist_ok=False)
    root = destination.resolve(strict=True)
    extracted = 0
    with zipfile.ZipFile(zip_path) as zf:
        plan, inspected_prefix = _archive_plan(zf)
        if inspected_prefix != prefix:
            raise RuntimeError("archive changed during extraction")
        for member in plan:
            package_relative = member.parts[len(prefix):]
            if not package_relative:
                continue
            first = package_relative[0].casefold()
            if first in PACKAGE_SCAFFOLD_DIRS:
                continue
            # Preserve the publisher's wrapper directory inside local staging; source_root points to this directory.
            relative = member.parts
            target = destination.joinpath(*relative)
            normalized = target.resolve(strict=False)
            if normalized != root and root not in normalized.parents:
                raise RuntimeError("unsafe archive extraction path")
            current = root
            for component in relative[:-1]:
                current = current / component
                if current.exists() or _is_reparse_or_link(current):
                    if not _directory(current):
                        raise RuntimeError("unsafe archive parent path")
                else:
                    current.mkdir()
            if member.is_dir:
                if target.exists() or _is_reparse_or_link(target):
                    if not _directory(target):
                        raise RuntimeError("archive directory collides with a file")
                else:
                    target.mkdir()
                continue
            if target.exists() or _is_reparse_or_link(target):
                raise RuntimeError("duplicate extraction target")
            written = 0
            with zf.open(member.info, "r") as source:
                fd = os.open(target, os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_BINARY", 0), 0o600)
                with os.fdopen(fd, "wb") as output:
                    while True:
                        block = source.read(COPY_CHUNK)
                        if not block:
                            break
                        written += len(block)
                        if written > member.info.file_size or written > MAX_MEMBER_BYTES:
                            raise RuntimeError("archive member expanded beyond declared size")
                        output.write(block)
                    output.flush()
                    try:
                        os.fsync(output.fileno())
                    except OSError:
                        pass
            if written != member.info.file_size:
                raise RuntimeError("archive member size mismatch")
            extracted += 1
            if progress:
                progress()
    source_root = _relative_source_root(destination, prefix)
    return source_root, extracted


def _require_production_root(root: Path) -> None:
    allowed, reason = production_install_capability(root)
    if not allowed:
        raise RuntimeError(reason)


def stage_update(check: UpdateCheck, portable_root: Path,
                 progress: Optional[Callable[[str, int, int, float], None]] = None,
                 *, lock: Optional[UpdateInstallLock] = None) -> StagedUpdate:
    """Copy configured source to unique local stage, verify, extract safely, and validate before handoff."""
    portable_root = Path(portable_root)
    _require_production_root(portable_root)
    if not check.available or not check.info or not check.package_path or not check.update_path:
        raise RuntimeError("Không có bản cập nhật mới hợp lệ để cài.")
    if check.info.build <= BUILD_NUMBER:
        raise RuntimeError("Bản cập nhật không mới hơn phiên bản đang chạy; không cho phép cài lại hoặc hạ cấp.")
    if lock is None or not lock.acquired:
        raise RuntimeError(MSG["already_installing"])

    emit = progress or (lambda *_args: None)
    parent = portable_root / STAGING_DIR
    if parent.exists() or _is_reparse_or_link(parent):
        _safe_managed_dir(parent)
    else:
        parent.mkdir(parents=True, exist_ok=False)
    _clean_stage_dirs_under_lock(portable_root, older_than_seconds=0)

    # Re-read the authoritative manifest immediately before copying. A changed target must be checked again.
    latest = check_for_update(check.update_path, current_build=BUILD_NUMBER, current_version=__version__)
    if not latest.available or latest.info != check.info or not latest.package_path:
        raise RuntimeError("version.json đã thay đổi kể từ lần kiểm tra; hãy kiểm tra lại cập nhật.")
    stage_id = f"stage-{check.info.build:03d}-{uuid.uuid4().hex}"
    stage_dir = parent / stage_id
    stage_dir.mkdir(mode=0o700)
    local_zip = stage_dir / check.info.package
    new_dir = stage_dir / "new"
    manifest = stage_dir / STAGED_MANIFEST
    emit(STAGE_PREPARING, 0, 0, 0.0)
    _log_file(portable_root, f"UPDATE_STAGE_START build={format_build(check.info.build)}")
    try:
        copied = copy_with_progress(Path(latest.package_path), local_zip,
                                    lambda done, total, pct: emit(STAGE_COPYING, done, total, pct))
        if check.info.size is not None and copied != check.info.size:
            raise RuntimeError("Kích thước gói cập nhật không khớp version.json.")
        _log_file(portable_root, f"UPDATE_COPY_DONE bytes={copied}")
        emit(STAGE_VERIFYING, copied, copied, 100.0)
        ok, message = verify_package(local_zip, check.info)
        if not ok:
            raise RuntimeError(message)
        _log_file(portable_root, "UPDATE_VERIFY_OK")
        emit(STAGE_EXTRACTING, 0, 0, 0.0)
        source_root, member_count = _extract_archive(local_zip, new_dir, check.info)
        _log_file(portable_root, f"UPDATE_EXTRACT_OK members={member_count}")
        emit(STAGE_VALIDATING, copied, copied, 100.0)
        valid, reason = validate_staged_portable(source_root, check.info)
        if not valid:
            raise RuntimeError(f"Gói cập nhật không đạt kiểm tra Portable: {reason}.")
        _log_file(portable_root, "UPDATE_STAGE_VALIDATE_OK")

        payload = {
            "schema": 1,
            "version": check.info.version,
            "build": check.info.build,
            "package": check.info.package,
            "sha256": check.info.sha256,
            "size": check.info.size,
            "source_root": str(source_root.resolve(strict=True)),
            "target_root": str(portable_root.resolve(strict=True)),
            "stage_id": stage_id,
            "old_pid": os.getpid(),
            "from_version": __version__,
            "from_build": BUILD_NUMBER,
            "staged_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        }
        with manifest.open("x", encoding="utf-8", newline="\n") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        emit(STAGE_READY, copied, copied, 100.0)
        return StagedUpdate(str(stage_dir), str(local_zip), str(source_root), str(portable_root),
                            str(manifest), check.info, stage_id)
    except Exception as exc:
        try:
            _remove_tree_owned(stage_dir)
        except OSError:
            LOG.warning("UPDATE_STAGE_CLEANUP_FAILED")
        if isinstance(exc, RuntimeError):
            raise
        raise RuntimeError(f"Cập nhật thất bại ở bước chuẩn bị ({type(exc).__name__}).") from exc


def updater_command(staged: StagedUpdate, packaged: Optional[bool] = None) -> List[str]:
    """Structured helper argument vector. Source-mode helper launch is forbidden."""
    if packaged is False or not is_packaged():
        raise RuntimeError(MSG["development_only"])
    allowed, reason = production_install_capability(Path(staged.target_root))
    if not allowed:
        raise RuntimeError(reason)
    source = Path(staged.source_root).resolve(strict=True)
    manifest = Path(staged.manifest_path).resolve(strict=True)
    if not _regular_file(source / APP_EXE) or manifest.parent.resolve() != Path(staged.staging_dir).resolve():
        raise RuntimeError("Thư mục helper đã staging không hợp lệ.")
    return [str(source / APP_EXE), "--apply-update", str(manifest)]


def launch_updater(staged: StagedUpdate, spawn: Callable = subprocess.Popen) -> List[str]:
    cmd = updater_command(staged, packaged=True)
    kwargs = {"cwd": staged.source_root, "close_fds": True, "shell": False}
    if os.name == "nt":
        kwargs["creationflags"] = 0x00000008 | 0x00000200  # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    spawn(cmd, **kwargs)
    _log_file(Path(staged.target_root), f"UPDATE_HELPER_STARTED build={format_build(staged.info.build)}")
    return cmd


def process_alive(pid: int) -> bool:
    if isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
        return False
    if os.name == "nt":
        import ctypes
        k32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
        handle = k32.OpenProcess(0x1000, False, int(pid))  # PROCESS_QUERY_LIMITED_INFORMATION
        if not handle:
            error = k32.GetLastError()
            return error != 87  # only ERROR_INVALID_PARAMETER proves the PID no longer exists
        try:
            code = ctypes.c_ulong()
            if k32.GetExitCodeProcess(handle, ctypes.byref(code)):
                return code.value == 259  # STILL_ACTIVE
            return True
        finally:
            k32.CloseHandle(handle)
    try:
        os.kill(int(pid), 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return True  # an unknown process-query error is not evidence that the old app exited


def wait_for_exit(pid: int, timeout: float = 90.0, alive: Callable[[int], bool] = process_alive,
                  sleep: Callable[[float], None] = time.sleep, interval: float = 0.5,
                  clock: Callable[[], float] = time.monotonic) -> bool:
    if isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
        return False
    try:
        bounded_timeout = min(90.0, max(0.0, float(timeout)))
    except (TypeError, ValueError, OverflowError):
        return False
    deadline = clock() + bounded_timeout
    slept = 0.0
    try:
        poll_interval = min(0.5, max(0.01, float(interval)))
    except (TypeError, ValueError, OverflowError):
        return False
    while True:
        try:
            if not alive(pid):
                return True
        except Exception:  # noqa: BLE001 - process query failure fails closed without replacing files
            return False
        remaining = min(deadline - clock(), bounded_timeout - slept)
        if remaining <= 0:
            return False
        delay = min(poll_interval, remaining)
        try:
            sleep(delay)
        except Exception:  # noqa: BLE001 - sleep failure must not allow replacement
            return False
        slept += delay


def is_preserved(name: str) -> bool:
    low = name.casefold()
    return low in PRESERVED_ENTRIES or ("." in low and low.endswith(PRESERVED_SUFFIXES))


def release_entries(source_root: Path) -> List[str]:
    """Exact app-owned roots present in the validated staged release; unknown entries are never installed."""
    source_root = Path(source_root)
    if not _directory(source_root):
        raise RuntimeError("Thư mục release staged không hợp lệ.")
    names: List[str] = []
    for entry in source_root.iterdir():
        canonical = _PACKAGE_OWNED_BY_CASE.get(entry.name.casefold())
        if canonical is None:
            if entry.name.casefold() in PACKAGE_SCAFFOLD_DIRS and _directory(entry):
                continue
            raise RuntimeError("Gói chứa thành phần ngoài layout Portable được hỗ trợ.")
        if canonical != entry.name or _is_reparse_or_link(entry):
            raise RuntimeError("Gói chứa tên hoặc kiểu file Portable không an toàn.")
        names.append(canonical)
    required = {APP_EXE, "_internal", "frontend", "VERSION.txt"}
    if not required.issubset(names):
        raise RuntimeError("Gói Portable thiếu file ứng dụng bắt buộc.")
    return sorted(names)


def _tree_signature(path: Path) -> Tuple[Tuple[str, str, int], ...]:
    path = Path(path)
    if _is_reparse_or_link(path):
        raise OSError("backup tree contains a link/reparse point")
    if _regular_file(path):
        return (("", sha256_file(path), path.stat().st_size),)
    if not _directory(path):
        raise OSError("backup tree has an unsupported root type")
    signature: List[Tuple[str, str, int]] = []
    for item in sorted(path.rglob("*"), key=lambda candidate: candidate.relative_to(path).as_posix().casefold()):
        rel = item.relative_to(path).as_posix()
        if _is_reparse_or_link(item):
            raise OSError("backup tree contains a link/reparse point")
        if _directory(item):
            signature.append((rel + "/", "DIR", 0))
        elif _regular_file(item):
            signature.append((rel, sha256_file(item), item.stat().st_size))
        else:
            raise OSError("backup tree contains an unsupported file type")
    return tuple(signature)


def _copy(src: Path, dst: Path) -> None:
    src, dst = Path(src), Path(dst)
    if _is_reparse_or_link(src):
        raise OSError("refusing to copy a link/reparse point")
    if _directory(src):
        _tree_signature(src)  # refuse nested symlinks/junctions before copytree can follow them
        shutil.copytree(src, dst, symlinks=False)
        _tree_signature(dst)
    elif _regular_file(src):
        shutil.copy2(src, dst)
    else:
        raise OSError("refusing to copy unsupported package file")


def _remove_package_entry(path: Path) -> None:
    if not path.exists() and not _is_reparse_or_link(path):
        return
    if _is_reparse_or_link(path):
        raise OSError("package-owned target is a link/reparse point")
    if _directory(path):
        _remove_tree_owned(path)
    elif _regular_file(path):
        path.unlink()
    else:
        raise OSError("package-owned target has an unsupported file type")


def _new_backup_dir(target: Path, from_build: int) -> Path:
    backup_root = target / BACKUP_DIR
    if backup_root.exists() or _is_reparse_or_link(backup_root):
        _safe_managed_dir(backup_root)
    else:
        backup_root.mkdir(mode=0o700)
    backup = backup_root / f"backup-{from_build:03d}-{uuid.uuid4().hex}"
    backup.mkdir(mode=0o700)
    return backup


def _create_verified_backup(target: Path, backup: Path, *, from_version: str, from_build: int) -> List[str]:
    entries: List[str] = []
    signatures: Dict[str, List[List[object]]] = {}
    try:
        for name in sorted(PACKAGE_OWNED_TOP_LEVEL):
            source = Path(target) / name
            if not source.exists() and not _is_reparse_or_link(source):
                continue
            if _is_reparse_or_link(source) or not (_directory(source) or _regular_file(source)):
                raise OSError("current package-owned entry is unsafe")
            destination = Path(backup) / name
            _copy(source, destination)
            source_signature = _tree_signature(source)
            backup_signature = _tree_signature(destination)
            if source_signature != backup_signature:
                raise OSError("backup verification mismatch")
            entries.append(name)
            signatures[name] = [list(item) for item in backup_signature]
        metadata = {"schema": 1, "entries": entries, "signatures": signatures,
                    "version": from_version, "build": int(from_build),
                    "created_at": datetime.now(timezone.utc).isoformat(timespec="seconds")}
        with (Path(backup) / "backup.json").open("x", encoding="utf-8") as handle:
            json.dump(metadata, handle, ensure_ascii=False, indent=2)
            handle.flush()
            os.fsync(handle.fileno())
        return entries
    except Exception:
        try:
            _remove_tree_owned(backup)
        except OSError:
            LOG.error("UPDATE_BACKUP_CLEANUP_FAILED")
        raise


def _read_backup_entries(backup: Path) -> Tuple[List[str], Dict[str, Tuple[Tuple[str, str, int], ...]]]:
    manifest = Path(backup) / "backup.json"
    if not _regular_file(manifest) or manifest.stat().st_size > MAX_BACKUP_MANIFEST_BYTES:
        raise OSError("backup manifest is missing or oversized")
    with manifest.open("rb") as handle:
        raw_manifest = handle.read(MAX_BACKUP_MANIFEST_BYTES + 1)
    if len(raw_manifest) > MAX_BACKUP_MANIFEST_BYTES:
        raise OSError("backup manifest exceeds the size limit")
    data = json.loads(raw_manifest.decode("utf-8"))
    entries = data.get("entries") if isinstance(data, dict) and data.get("schema") == 1 else None
    if not isinstance(entries, list) or any(not isinstance(n, str) or n not in PACKAGE_OWNED_TOP_LEVEL for n in entries):
        raise OSError("backup manifest is invalid")
    if len(entries) != len(set(entries)):
        raise OSError("backup manifest has duplicate entries")
    raw_signatures = data.get("signatures")
    if not isinstance(raw_signatures, dict) or set(raw_signatures) != set(entries):
        raise OSError("backup signatures are missing")
    signatures: Dict[str, Tuple[Tuple[str, str, int], ...]] = {}
    for name in entries:
        rows = raw_signatures.get(name)
        if not isinstance(rows, list):
            raise OSError("backup signature is invalid")
        parsed: List[Tuple[str, str, int]] = []
        for row in rows:
            if (not isinstance(row, list) or len(row) != 3 or not isinstance(row[0], str)
                    or not isinstance(row[1], str) or isinstance(row[2], bool)
                    or not isinstance(row[2], int) or row[2] < 0):
                raise OSError("backup signature is invalid")
            if row[1] != "DIR" and (len(row[1]) != 64 or any(c not in "0123456789abcdef" for c in row[1])):
                raise OSError("backup signature hash is invalid")
            parsed.append((row[0], row[1], row[2]))
        signatures[name] = tuple(parsed)
    return entries, signatures


def _rollback(target: Path, backup: Path, info: UpdateInfo,
              log: Callable[[str], None]) -> bool:
    log("UPDATE_ROLLBACK_START")
    restored = True
    try:
        entries, expected_signatures = _read_backup_entries(backup)
    except Exception as exc:  # noqa: BLE001
        log(f"UPDATE_ROLLBACK_FAILED stage=backup_manifest error={type(exc).__name__}")
        return False
    for name in sorted(PACKAGE_OWNED_TOP_LEVEL):
        try:
            _remove_package_entry(Path(target) / name)
        except Exception as exc:  # noqa: BLE001
            restored = False
            log(f"UPDATE_ROLLBACK_FAILED stage=remove entry={name} error={type(exc).__name__}")
    for name in entries:
        try:
            source = Path(backup) / name
            if not source.exists() or _is_reparse_or_link(source):
                raise OSError("backup entry missing or unsafe")
            if _tree_signature(source) != expected_signatures[name]:
                raise OSError("backup bytes no longer match the verified snapshot")
            _copy(source, Path(target) / name)
            if _tree_signature(source) != _tree_signature(Path(target) / name):
                raise OSError("restored bytes mismatch")
        except Exception as exc:  # noqa: BLE001
            restored = False
            log(f"UPDATE_ROLLBACK_FAILED stage=restore entry={name} error={type(exc).__name__}")
    if restored:
        log("UPDATE_ROLLBACK_OK")
        _write_result(target, "rolled_back", info, "Đã khôi phục file ứng dụng trước cập nhật.",
                      helper_pid=os.getpid())
    else:
        log("UPDATE_ROLLBACK_FAILED stage=restore")
        _write_result(target, "rollback_failed", info, "Không thể khôi phục hoàn toàn file ứng dụng.",
                      helper_pid=os.getpid())
    return restored


def _prune_backups(target: Path, keep: int = MAX_BACKUPS) -> None:
    root = Path(target) / BACKUP_DIR
    if not _directory(root):
        return
    backups = [p for p in root.iterdir() if BACKUP_NAME_RE.fullmatch(p.name) and _directory(p)]
    backups.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    for old in backups[max(1, keep):]:
        try:
            _remove_tree_owned(old)
        except OSError:
            LOG.warning("UPDATE_BACKUP_PRUNE_FAILED")


def _manifest_path_contract(manifest_path: Path, data: dict) -> Tuple[Path, Path, Path, UpdateInfo]:
    manifest_path = Path(manifest_path)
    if manifest_path.name != STAGED_MANIFEST or not _regular_file(manifest_path):
        raise ValueError("helper manifest is not an ordinary staged file")
    stage = manifest_path.parent.resolve(strict=True)
    if not STAGE_NAME_RE.fullmatch(stage.name) or not _directory(stage):
        raise ValueError("invalid stage identity")
    parent = stage.parent.resolve(strict=True)
    if parent.name.casefold() != STAGING_DIR.casefold() or not _directory(parent):
        raise ValueError("invalid staging parent")
    target = parent.parent.resolve(strict=True)
    if not _is_local_portable_root(target):
        raise ValueError("target Portable root is not on local storage")
    if data.get("schema") != 1 or data.get("stage_id") != stage.name:
        raise ValueError("invalid update manifest schema/stage")
    if Path(str(data.get("target_root", ""))).resolve(strict=True) != target:
        raise ValueError("target root is not the owner of staging")
    version = data.get("version")
    build = data.get("build")
    package = data.get("package")
    sha = data.get("sha256")
    size = data.get("size")
    if not isinstance(version, str) or isinstance(build, bool) or not isinstance(build, int):
        raise ValueError("invalid staged version/build")
    if size is not None and (isinstance(size, bool) or not isinstance(size, int)):
        raise ValueError("invalid staged size")
    info = parse_manifest(json.dumps({"version": version, "build": build, "package": package,
                                      "sha256": sha, **({"size": size} if size is not None else {})}))
    source = Path(str(data.get("source_root", ""))).resolve(strict=True)
    if source != (stage / "new").resolve(strict=True) and (stage / "new").resolve(strict=True) not in source.parents:
        raise ValueError("source root is outside local staging")
    package_path = stage / info.package
    if not _regular_file(package_path) or package_path.resolve(strict=True).parent != stage:
        raise ValueError("staged package is missing or outside stage")
    old_pid, from_build, from_version = data.get("old_pid"), data.get("from_build"), data.get("from_version")
    if (isinstance(old_pid, bool) or not isinstance(old_pid, int) or old_pid <= 0
            or isinstance(from_build, bool) or not isinstance(from_build, int) or from_build <= 0
            or not isinstance(from_version, str) or not VERSION_RE.fullmatch(from_version)):
        raise ValueError("invalid source runtime identity")
    if info.build <= from_build:
        raise ValueError("normal update cannot reinstall or downgrade")
    return stage, target, source, info


def _restart_command(target: Path) -> Tuple[List[str], dict]:
    executable = Path(target) / APP_EXE
    if not _regular_file(executable):
        raise OSError("restart executable is missing")
    kwargs = {"cwd": str(target), "close_fds": True, "shell": False}
    if os.name == "nt":
        kwargs["creationflags"] = 0x00000008 | 0x00000200
    return [str(executable)], kwargs


def _restart_process_ok(process, *, grace: float, sleep: Callable[[float], None],
                        clock: Callable[[], float]) -> bool:
    if process is None or not callable(getattr(process, "poll", None)):
        return True  # test adapters may record argv only; real Popen always has poll().
    try:
        bounded_grace = min(5.0, max(0.0, float(grace)))
    except (TypeError, ValueError, OverflowError):
        return False
    deadline = clock() + bounded_grace
    slept = 0.0
    while True:
        try:
            if process.poll() is not None:
                return False
        except Exception:  # noqa: BLE001
            return False
        remaining = min(deadline - clock(), bounded_grace - slept)
        if remaining <= 0:
            return True
        delay = min(0.1, remaining)
        try:
            sleep(delay)
        except Exception:  # noqa: BLE001
            return False
        slept += delay


def _restart_restored_application(target: Path, info: UpdateInfo, log: Callable[[str], None], spawn: Callable,
                                  sleep: Callable[[float], None], clock: Callable[[], float],
                                  grace: float) -> bool:
    try:
        command, kwargs = _restart_command(target)
        process = spawn(command, **kwargs)
        if not _restart_process_ok(process, grace=grace, sleep=sleep, clock=clock):
            raise OSError("restored application exited during startup grace period")
        log("UPDATE_ROLLBACK_RESTART")
        return True
    except Exception as exc:  # noqa: BLE001
        log(f"UPDATE_ROLLBACK_RESTART_FAILED error={type(exc).__name__}")
        _write_result(target, "rolled_back", info,
                      "File ứng dụng cũ đã được khôi phục nhưng không thể tự khởi chạy; hãy mở ứng dụng thủ công.",
                      helper_pid=os.getpid())
        return False


def apply_update(manifest_path: Path, alive: Optional[Callable[[int], bool]] = None,
                 sleep: Callable[[float], None] = time.sleep, spawn: Optional[Callable] = subprocess.Popen,
                 wait_timeout: float = 90.0, restart: bool = True, restart_grace: float = 1.0,
                 clock: Callable[[], float] = time.monotonic) -> int:
    """Apply a previously validated staged Portable release in the separate frozen helper process.

    Return 0 on installed/restarted, 1 on a safe failure/rollback, and 2 on an invalid helper/failed rollback.
    """
    manifest_path = Path(manifest_path)
    if not is_packaged():
        LOG.error("UPDATE_HELPER_REJECTED reason=development_runtime")
        return 2
    target: Optional[Path] = None
    info: Optional[UpdateInfo] = None
    try:
        if not _regular_file(manifest_path) or manifest_path.stat().st_size > MAX_MANIFEST_BYTES:
            raise ValueError("staged helper manifest is not a bounded ordinary file")
        with manifest_path.open("rb") as handle:
            manifest_bytes = handle.read(MAX_MANIFEST_BYTES + 1)
        if len(manifest_bytes) > MAX_MANIFEST_BYTES:
            raise ValueError("staged helper manifest exceeds the size limit")
        data = json.loads(manifest_bytes.decode("utf-8-sig"))
        if not isinstance(data, dict):
            raise ValueError("invalid manifest object")
        stage, target, source, info = _manifest_path_contract(manifest_path, data)
        executable_root = Path(sys.executable).resolve(strict=True).parent
        if executable_root != source:
            raise ValueError("helper is not running from the staged release root")
        if int(data["old_pid"]) == os.getpid():
            raise ValueError("helper PID must differ from original app PID")
        if not _regular_file(target / APP_EXE) or not _directory(target / "_internal"):
            raise ValueError("target Portable layout is missing")
        current_version, current_build = _read_version_fields(target / "VERSION.txt")
        if current_version != data["from_version"] or current_build != int(data["from_build"]):
            raise ValueError("target version changed since check/stage")
        if not portable_root_layout_matches(target, version=current_version, build=current_build):
            raise ValueError("target is not a verified production Portable root")
        valid, reason = validate_staged_portable(source, info)
        if not valid:
            raise ValueError(f"staged application validation failed: {reason}")
        package_path = stage / info.package
        ok, message = verify_package(package_path, info)
        if not ok:
            raise ValueError(f"staged ZIP verification failed: {message}")
    except Exception as exc:  # noqa: BLE001
        LOG.error("UPDATE_HELPER_REJECTED stage=UPDATE_STAGE_VALIDATE error=%s", type(exc).__name__)
        if target is not None and info is not None:
            _write_result(target, "failed", info, "Gói hoặc thông tin cài đặt không hợp lệ.", helper_pid=os.getpid())
        return 2

    assert target is not None and info is not None
    log = lambda message: _log_file(target, message)  # noqa: E731
    helper_lock = UpdateInstallLock(target / LOCK_NAME)
    try:
        if not helper_lock.acquire(timeout=wait_timeout, sleep=sleep, clock=clock):
            log("UPDATE_HELPER_LOCK_FAILED")
            _write_result(target, "failed", info, "Một tiến trình cập nhật khác đang chạy.", helper_pid=os.getpid())
            return 1
    except OSError as exc:
        log(f"UPDATE_HELPER_LOCK_FAILED error={type(exc).__name__}")
        _write_result(target, "failed", info, "Không thể khóa an toàn tiến trình cập nhật.", helper_pid=os.getpid())
        return 1

    backup: Optional[Path] = None
    try:
        # The application holds this same process lock through helper handoff, so obtaining it proves the old
        # application has released the filesystem transaction. The explicit PID wait remains a bounded check.
        log(f"UPDATE_WAIT_EXIT build={format_build(info.build)}")
        original_pid = int(data["old_pid"])
        if not wait_for_exit(original_pid, wait_timeout, alive or process_alive, sleep, clock=clock):
            log("UPDATE_WAIT_EXIT_FAILED reason=timeout")
            _write_result(target, "failed", info, "Ứng dụng cũ chưa thoát trong thời gian cho phép; chưa thay file.",
                          helper_pid=os.getpid())
            return 1
        log("UPDATE_WAIT_EXIT_OK")

        # Recheck the target identity too: another helper may have completed while this helper waited for the lock.
        locked_version, locked_build = _read_version_fields(target / "VERSION.txt")
        if (locked_version != data["from_version"] or locked_build != int(data["from_build"])
                or info.build <= locked_build
                or not portable_root_layout_matches(target, version=locked_version, build=locked_build)):
            raise RuntimeError("target Portable version changed since staging")

        # Recheck all stage data after obtaining the shared process lock.
        valid, reason = validate_staged_portable(source, info)
        if not valid:
            raise RuntimeError(f"stage validation failed: {reason}")
        ok, message = verify_package(stage / info.package, info)
        if not ok:
            raise RuntimeError(f"stage hash/package validation failed: {message}")
        names = release_entries(source)

        payload = stage / "install-ready"
        if payload.exists():
            _remove_tree_owned(payload)
        payload.mkdir(mode=0o700)
        for name in names:
            _copy(source / name, payload / name)
        valid, reason = validate_staged_portable(payload, info)
        if not valid:
            raise RuntimeError(f"install-ready validation failed: {reason}")

        backup = _new_backup_dir(target, int(data["from_build"]))
        backup_entries = _create_verified_backup(target, backup, from_version=data["from_version"],
                                                  from_build=int(data["from_build"]))
        log(f"UPDATE_BACKUP_OK entries={len(backup_entries)}")
        _write_result(target, "installing", info, "", helper_pid=os.getpid())

        for name in sorted(PACKAGE_OWNED_TOP_LEVEL):
            current = target / name
            if current.exists() or _is_reparse_or_link(current):
                _remove_package_entry(current)
            staged_entry = payload / name
            if staged_entry.exists():
                os.replace(staged_entry, current)
        valid, reason = validate_staged_portable(target, info)
        if not valid:
            raise RuntimeError(f"installed layout validation failed: {reason}")
        log("UPDATE_INSTALL_OK")

        if not restart or spawn is None:
            _write_result(target, "ok", info, "", helper_pid=os.getpid())
            _prune_backups(target)
            return 0

        _write_result(target, "restarting", info, "", helper_pid=os.getpid())
        try:
            cmd, kwargs = _restart_command(target)
            new_process = spawn(cmd, **kwargs)
            log("UPDATE_RESTART")
            if not _restart_process_ok(new_process, grace=restart_grace, sleep=sleep, clock=clock):
                raise OSError("new executable exited during startup grace period")
        except Exception as exc:  # noqa: BLE001
            log(f"UPDATE_RESTART_FAILED error={type(exc).__name__}")
            if not _rollback(target, backup, info, log):
                return 2
            if spawn is not None:
                _restart_restored_application(target, info, log, spawn, sleep, clock, restart_grace)
            return 1

        _write_result(target, "ok", info, "", helper_pid=os.getpid())
        _prune_backups(target)
        return 0
    except Exception as exc:  # noqa: BLE001
        log(f"UPDATE_INSTALL_FAILED stage=UPDATE_INSTALL error={type(exc).__name__}")
        if backup is not None and backup.exists():
            if not _rollback(target, backup, info, log):
                return 2
            if spawn is not None:
                _restart_restored_application(target, info, log, spawn, sleep, clock, restart_grace)
            return 1
        # Backup or payload preparation failed: target has not yet been changed.
        _write_result(target, "failed", info, "Sao lưu hoặc chuẩn bị gói thất bại; phiên bản hiện tại chưa bị thay đổi.",
                      helper_pid=os.getpid())
        return 1
    finally:
        helper_lock.release()


def _write_result(target: Path, status: str, info: UpdateInfo, detail: str, *, helper_pid: int = 0) -> None:
    try:
        logs = Path(target) / "logs"
        if logs.exists() or _is_reparse_or_link(logs):
            if not _directory(logs):
                return
        else:
            logs.mkdir(parents=True, exist_ok=True)
        destination = logs / RESULT_FILE
        temp = logs / (RESULT_FILE + ".tmp")
        if _is_reparse_or_link(destination) or _is_reparse_or_link(temp):
            return
        payload = {"status": status, "version": info.version, "build": info.build,
                   "detail": detail[:500], "helper_pid": int(helper_pid),
                   "at": datetime.now(timezone.utc).isoformat(timespec="seconds")}
        with temp.open("w", encoding="utf-8", newline="\n") as handle:
            json.dump(payload, handle, ensure_ascii=False)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp, destination)
    except OSError:
        LOG.warning("UPDATE_RESULT_WRITE_FAILED status=%s", status)


def read_update_result(portable_root: Path) -> Optional[dict]:
    path = Path(portable_root) / "logs" / RESULT_FILE
    if not _regular_file(path):
        return None
    try:
        if path.stat().st_size > 16 * 1024:
            return None
        data = json.loads(path.read_text(encoding="utf-8"))
        if (not isinstance(data, dict) or data.get("status") not in {
                "installing", "restarting", "ok", "rolled_back", "rollback_failed", "failed"}
                or not isinstance(data.get("version"), str)
                or not VERSION_RE.fullmatch(data["version"])
                or isinstance(data.get("build"), bool) or not isinstance(data.get("build"), int)
                or data["build"] < 1
                or isinstance(data.get("helper_pid", 0), bool)
                or not isinstance(data.get("helper_pid", 0), int) or data.get("helper_pid", 0) < 0
                or not isinstance(data.get("detail", ""), str)):
            return None
        data["detail"] = data.get("detail", "")[:500]
        return data
    except (OSError, ValueError, UnicodeError):
        return None


def consume_result(portable_root: Path) -> Optional[dict]:
    """Read and delete terminal one-shot updater result; never consume a pending handoff state."""
    path = Path(portable_root) / "logs" / RESULT_FILE
    data = read_update_result(portable_root)
    if not data or data.get("status") in ("installing", "restarting"):
        return None
    helper_pid = int(data.get("helper_pid") or 0)
    if helper_pid and process_alive(helper_pid):
        return data  # keep it visible until the helper exits; also keeps staging alive during helper execution.
    try:
        path.unlink()
    except OSError:
        pass
    return data


def cleanup_staged_update(staged: StagedUpdate) -> None:
    """Remove one caller-owned stage after a failed helper launch while its process lock is still held."""
    stage = Path(staged.staging_dir)
    target = Path(staged.target_root).resolve(strict=True)
    if (stage.name != staged.stage_id or not STAGE_NAME_RE.fullmatch(stage.name)
            or stage.parent.resolve(strict=True) != (target / STAGING_DIR).resolve(strict=True)
            or not _directory(stage)):
        return
    _remove_tree_owned(stage)


def cleanup_staging(portable_root: Path) -> None:
    """Remove only updater-owned stage directories, and only while no process holds the update lock."""
    root = Path(portable_root)
    try:
        lock = acquire_update_lock(root, timeout=0)
    except OSError:
        return
    if lock is None:
        return
    try:
        _clean_stage_dirs_under_lock(root, older_than_seconds=0)
    except (OSError, RuntimeError):
        LOG.warning("UPDATE_STAGE_CLEANUP_FAILED")
    finally:
        lock.release()


def _log_file(portable_root: Path, message: str) -> None:
    """Append bounded, secret-free stage diagnostics without recording configured network/local paths."""
    LOG.info(message)
    try:
        logs = Path(portable_root) / "logs"
        if logs.exists() or _is_reparse_or_link(logs):
            if not _directory(logs):
                return
        else:
            logs.mkdir(parents=True, exist_ok=True)
        path = logs / UPDATE_LOG
        if _is_reparse_or_link(path):
            return
        # Keep diagnostics bounded (last 2 MiB) without following a user-controlled link.
        if _regular_file(path) and path.stat().st_size > 2 * 1024 * 1024:
            with path.open("rb") as source:
                source.seek(-1 * 1024 * 1024, os.SEEK_END)
                tail = source.read()
            with path.open("wb") as output:
                output.write(tail)
        with path.open("a", encoding="utf-8", newline="\n") as handle:
            handle.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {message[:500]}\n")
    except OSError:
        pass


def staged_to_dict(staged: StagedUpdate) -> dict:
    data = asdict(staged)
    data["info"] = asdict(staged.info)
    return data
