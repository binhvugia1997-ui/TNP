"""Publish a built release to the LAN update folder (build/release tooling only – never used at runtime).

Called by tools/build_portable.py as the LAST step, only after pyflakes, pytest, PyInstaller, artifact validation,
ZIP and version.json all succeeded.  Safe publish order (mandatory – work PCs auto-check version.json):

    1. copy  release/ReportExtractor_<version>.zip  ->  <update>/ReportExtractor_<version>.zip.tmp
    2. verify size + SHA256 of the destination against release/version.json
    3. rename .tmp -> final ZIP name (same volume -> atomic)
    4. write   <update>/version.json.tmp, then os.replace() -> version.json   (manifest LAST)
    5. remove obsolete ReportExtractor_*.zip (other releases only; unrelated files are never touched)

Any failure before step 4 leaves the previous ZIP + manifest intact; a failure in step 4 keeps the previous
manifest (os.replace is atomic) and the new ZIP is removed again so nothing dangles.

The destination is centralised here (env overrides for other build PCs):
    RE_UPDATE_FOLDER   local folder shared to the LAN     (default D:\\ReportExtractor_Update)
    RE_UPDATE_LAN_PATH UNC path the work PCs configure    (default \\\\192.168.103.11\\ReportExtractor_Update)
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, List, Optional

from console_safe import configure_console, safe_text

DEFAULT_UPDATE_FOLDER = r"D:\ReportExtractor_Update"
DEFAULT_LAN_PATH = r"\\192.168.103.11\ReportExtractor_Update"
MANIFEST_NAME = "version.json"
PACKAGE_GLOB = "ReportExtractor_*.zip"
TAG = "[PUBLISH]"


class PublishError(RuntimeError):
    pass


def update_folder() -> str:
    return (os.environ.get("RE_UPDATE_FOLDER") or DEFAULT_UPDATE_FOLDER).strip() or DEFAULT_UPDATE_FOLDER


def lan_path() -> str:
    return (os.environ.get("RE_UPDATE_LAN_PATH") or DEFAULT_LAN_PATH).strip() or DEFAULT_LAN_PATH


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with Path(path).open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


@dataclass
class PublishResult:
    ok: bool
    folder: Path
    package: str = ""
    removed: List[str] = field(default_factory=list)
    error: str = ""


def _log(say: Callable[[str], None], msg: str) -> None:
    text = f"{TAG} {msg}"
    try:
        say(text)
    except UnicodeEncodeError:  # external callbacks may be strict legacy-codepage streams
        say(safe_text(text, encoding="ascii"))


def _cleanup(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except OSError:
        pass


def publish_release(release_dir: Path, folder: Optional[str] = None, *, say: Callable[[str], None] = print,
                    lan: Optional[str] = None) -> PublishResult:
    """Publish release_dir/{<package>.zip, version.json} into ``folder`` with the safe ordering above."""
    release_dir = Path(release_dir)
    dest = Path(folder or update_folder())
    lan = lan or lan_path()
    manifest_src = release_dir / MANIFEST_NAME
    tmp_zip = tmp_manifest = None
    try:
        # ---- 0. source validation (never publish a half-built release) ---------------------------------
        if not manifest_src.is_file():
            raise PublishError(f"thiếu {manifest_src}")
        info = json.loads(manifest_src.read_text(encoding="utf-8-sig"))
        pkg_name = str(info.get("package") or "")
        expected_sha = str(info.get("sha256") or "").lower()
        if not pkg_name or not expected_sha:
            raise PublishError("version.json thiếu package/sha256")
        zip_src = release_dir / pkg_name
        if not zip_src.is_file():
            raise PublishError(f"thiếu gói {zip_src}")
        if sha256_file(zip_src) != expected_sha:
            raise PublishError("SHA256 của ZIP nguồn không khớp version.json")
        src_size = zip_src.stat().st_size

        _log(say, f"Update folder: {dest}")
        dest.mkdir(parents=True, exist_ok=True)
        final_zip = dest / pkg_name
        tmp_zip = dest / (pkg_name + ".tmp")
        manifest_dst = dest / MANIFEST_NAME
        tmp_manifest = dest / (MANIFEST_NAME + ".tmp")

        # ---- 1-3. package FIRST: copy to .tmp, verify, rename ----------------------------------------
        _log(say, "Copying package...")
        _cleanup(tmp_zip)
        shutil.copyfile(zip_src, tmp_zip)
        dst_size = tmp_zip.stat().st_size
        if dst_size != src_size:
            raise PublishError(f"kích thước đích {dst_size} != nguồn {src_size}")
        _log(say, "Verifying SHA256...")
        if sha256_file(tmp_zip) != expected_sha:
            raise PublishError("SHA256 của gói đã copy không khớp")
        os.replace(tmp_zip, final_zip)
        tmp_zip = None
        _log(say, "Package OK")

        # ---- 4. manifest LAST (atomic replace) -------------------------------------------------------
        _log(say, "Publishing version.json LAST...")
        _cleanup(tmp_manifest)
        try:
            tmp_manifest.write_text(manifest_src.read_text(encoding="utf-8"), encoding="utf-8")
            os.replace(tmp_manifest, manifest_dst)
        except Exception:
            _cleanup(tmp_manifest)
            # previous manifest (if any) is intact; do not leave a ZIP the old manifest does not reference
            if manifest_dst.is_file():
                try:
                    old_pkg = json.loads(manifest_dst.read_text(encoding="utf-8-sig")).get("package")
                except Exception:  # noqa: BLE001
                    old_pkg = None
                if old_pkg != pkg_name:
                    _cleanup(final_zip)
            else:
                _cleanup(final_zip)
            raise
        tmp_manifest = None

        # ---- 5. obsolete packages only AFTER the new release is live ------------------------------------
        removed: List[str] = []
        for old in sorted(dest.glob(PACKAGE_GLOB)):
            if old.name == pkg_name or not old.is_file():
                continue
            _log(say, f"Removing obsolete package: {old.name}")
            try:
                old.unlink()
                removed.append(old.name)
            except OSError as e:
                _log(say, f"(không xoá được {old.name}: {e})")
        for stale in sorted(dest.glob(PACKAGE_GLOB + ".tmp")):
            _cleanup(stale)
        _log(say, "SUCCESS")
        success_text = f"\nLocal update folder:\n    {dest}\n\nLAN update path:\n    {lan}\n"
        try:
            say(success_text)
        except UnicodeEncodeError:
            say(safe_text(success_text, encoding="ascii"))
        return PublishResult(True, dest, pkg_name, removed)
    except Exception as e:  # noqa: BLE001 – build stays successful; publishing is reported separately
        if tmp_zip is not None:
            _cleanup(tmp_zip)
        if tmp_manifest is not None:
            _cleanup(tmp_manifest)
        msg = f"{type(e).__name__}: {e}" if not isinstance(e, PublishError) else str(e)
        _log(say, f"FAILED – {msg}")
        _log(say, "Gói/manifest cũ trong thư mục cập nhật được giữ nguyên (nếu có).")
        return PublishResult(False, dest, error=msg)


def main(argv: Optional[List[str]] = None) -> int:
    """Manual re-publish of an already built release: python tools\\publish_update.py [--release DIR] [--dest DIR]"""
    configure_console()
    import argparse
    ap = argparse.ArgumentParser(description="Publish release/ to the LAN update folder (ZIP first, version.json last)")
    ap.add_argument("--release", default=str(Path(__file__).resolve().parent.parent / "release"))
    ap.add_argument("--dest", default=None, help=f"mặc định {update_folder()}")
    args = ap.parse_args(argv)
    return 0 if publish_release(Path(args.release), args.dest).ok else 3


if __name__ == "__main__":
    raise SystemExit(main())
