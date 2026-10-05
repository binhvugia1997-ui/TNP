"""PROMPT-012 – ONE-CLICK BUILD + PUBLISH FROM GITHUB (``BUILD_AND_PUBLISH.bat`` → this module).

    safe Git fetch / pull  →  tools/build_portable.py (pyflakes, pytest, PyInstaller, validate, ZIP, SHA256,
    version.json)  →  publish to the LAN update folder (ZIP first, version.json LAST)  →  work PCs update.

Safety rules of the Git step (``git_sync``):

* the checkout must be a Git work tree with ``git`` available and an upstream branch;
* tracked files must be clean (untracked / ignored build artefacts are fine) – never stash or discard user edits;
* only fast-forward pulls: local commits not on GitHub (ahead) or a diverged branch STOP the run – nothing is
  rebased, reset or force-pulled;
* ``--branch NAME`` (or ``RE_BUILD_BRANCH``) asserts the branch that may be published from;
* the build that is about to be published is checked against the manifest already in the update folder: the same
  or an older build number is refused (``--force`` to republish deliberately).

Everything is logged to ``logs/build_and_publish_<timestamp>.log`` in addition to the console.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

from console_safe import child_env, configure_console, safe_text  # noqa: E402
import publish_update  # noqa: E402

RC_OK, RC_GIT, RC_BUILD, RC_PUBLISH, RC_REFUSED = 0, 2, 1, 3, 4


class SyncError(RuntimeError):
    """A safety check failed; the message says what the user must do – nothing has been changed."""


@dataclass
class SyncResult:
    branch: str = ""
    upstream: str = ""
    before: str = ""
    after: str = ""
    behind: int = 0
    ahead: int = 0
    pulled: bool = False
    notes: List[str] = field(default_factory=list)

    @property
    def revision(self) -> str:
        return self.after[:12]


def _git(args: List[str], cwd: Path, check: bool = True, timeout: float = 120.0) -> str:
    try:
        cp = subprocess.run(["git", *args], cwd=str(cwd), capture_output=True, text=True, encoding="utf-8",
                            errors="replace", timeout=timeout, env=child_env())
    except FileNotFoundError as e:
        raise SyncError("Không tìm thấy lệnh git. Cài Git for Windows (https://git-scm.com) rồi chạy lại.") from e
    except subprocess.TimeoutExpired as e:
        raise SyncError(f"git {' '.join(args)} quá thời gian ({timeout:.0f}s) – kiểm tra mạng/GitHub.") from e
    if check and cp.returncode != 0:
        raise SyncError(f"git {' '.join(args)} thất bại:\n{(cp.stderr or cp.stdout).strip()}")
    return cp.stdout.strip()


def git_sync(root: Path = ROOT, branch: Optional[str] = None, say: Callable[[str], None] = print,
             remote: str = "origin") -> SyncResult:
    """Fetch + fast-forward pull with the safety rules above.  Raises :class:`SyncError` when it must not proceed."""
    res = SyncResult()
    if _git(["rev-parse", "--is-inside-work-tree"], root, check=False) != "true":
        raise SyncError(f"{root} không phải là bản sao Git (git clone) – không thể đồng bộ từ GitHub.")
    res.branch = _git(["rev-parse", "--abbrev-ref", "HEAD"], root)
    if res.branch == "HEAD":
        raise SyncError("Đang ở trạng thái detached HEAD – checkout nhánh phát hành trước (git checkout <nhánh>).")
    want = branch or os.environ.get("RE_BUILD_BRANCH", "").strip() or None
    if want and res.branch != want:
        raise SyncError(f"Đang ở nhánh '{res.branch}', nhánh phát hành yêu cầu là '{want}'. "
                        f"Chạy: git checkout {want}")
    dirty = _git(["status", "--porcelain", "--untracked-files=no"], root)
    if dirty:
        raise SyncError("Thư mục làm việc có thay đổi chưa commit – không kéo code đè lên:\n" + dirty
                        + "\n\nCommit/push hoặc hoàn tác (git checkout -- <file>) rồi chạy lại.")
    res.before = _git(["rev-parse", "HEAD"], root)
    say(f"   Nhánh: {res.branch}   HEAD: {res.before[:12]}")
    say(f"   git fetch {remote} --prune ...")
    _git(["fetch", remote, "--prune"], root, timeout=300.0)
    upstream = _git(["rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}"], root, check=False)
    if not upstream or upstream == "@{u}":
        cand = f"{remote}/{res.branch}"
        if _git(["rev-parse", "--verify", "--quiet", cand], root, check=False):
            upstream = cand
            res.notes.append(f"nhánh chưa đặt upstream – dùng {cand}")
        else:
            raise SyncError(f"Nhánh '{res.branch}' không có trên {remote} – không có gì để kéo về.")
    res.upstream = upstream
    counts = _git(["rev-list", "--left-right", "--count", f"HEAD...{upstream}"], root)
    res.ahead, res.behind = (int(x) for x in counts.split())
    if res.ahead and res.behind:
        raise SyncError(f"Nhánh cục bộ và {upstream} đã phân nhánh (ahead {res.ahead}, behind {res.behind}). "
                        "Không tự động merge/rebase – xử lý thủ công rồi chạy lại.")
    if res.ahead:
        raise SyncError(f"Có {res.ahead} commit cục bộ chưa đẩy lên GitHub – không build gói từ mã chưa công bố. "
                        f"Chạy: git push {remote} {res.branch}  (hoặc bỏ commit đó) rồi chạy lại.")
    if res.behind:
        say(f"   Có {res.behind} commit mới trên {upstream} – git pull --ff-only ...")
        _git(["pull", "--ff-only", remote, res.branch], root, timeout=300.0)
        res.pulled = True
    else:
        say(f"   Đã mới nhất so với {upstream}.")
    res.after = _git(["rev-parse", "HEAD"], root)
    if res.pulled:
        say(f"   HEAD mới: {res.after[:12]}")
        log = _git(["log", "--oneline", f"{res.before}..{res.after}"], root, check=False)
        for ln in log.splitlines()[:20]:
            say("     " + ln)
    return res


# ---------------------------------------------------------------------------- published-build guard
def read_published_manifest(folder: str) -> Optional[dict]:
    p = Path(folder) / "version.json"
    if not p.is_file():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return None


def check_not_already_published(version: str, build: int, folder: str, force: bool = False) -> str:
    """Refuse to republish a build number that is already (or newer) in the update folder."""
    m = read_published_manifest(folder)
    if m is None:
        return f"Thư mục cập nhật chưa có version.json – sẽ xuất bản {version} / Build {build:03d}."
    try:
        pub_build = int(m.get("build", 0))
    except (TypeError, ValueError):
        pub_build = 0
    pub_ver = str(m.get("version", "?"))
    if pub_build >= build and not force:
        raise SyncError(f"Thư mục cập nhật đã có {pub_ver} / Build {pub_build:03d} (>= Build {build:03d} sắp build). "
                        "Tăng BUILD_NUMBER trong app/__init__.py rồi push, hoặc chạy với --force để xuất bản lại.")
    if pub_build >= build:
        return f"--force: xuất bản lại Build {build:03d} dù thư mục đang có Build {pub_build:03d}."
    return f"Đang phát hành: {pub_ver} / Build {pub_build:03d}  →  mới: {version} / Build {build:03d}."


# ---------------------------------------------------------------------------- orchestration
class Tee:
    def __init__(self, log_path: Path):
        log_path.parent.mkdir(parents=True, exist_ok=True)
        self.fh = open(log_path, "a", encoding="utf-8")
        self.path = log_path

    def __call__(self, msg: str = "") -> None:
        print(safe_text(msg, stream=sys.stdout), flush=True)
        self.fh.write(msg + "\n")
        self.fh.flush()

    def close(self) -> None:
        try:
            self.fh.close()
        except OSError:
            pass


def build_command(py: str, args: argparse.Namespace) -> List[str]:
    cmd = [sys.executable, str(ROOT / "tools" / "build_portable.py"), "--python", py]
    if args.skip_tests:
        cmd.append("--skip-tests")
    if args.no_publish:
        cmd.append("--no-publish")
    if args.publish_dir:
        cmd += ["--publish-dir", args.publish_dir]
    if args.allow_non_windows:
        cmd.append("--allow-non-windows")
    return cmd


def run_streamed(cmd: List[str], say: Callable[[str], None], cwd: Path = ROOT) -> int:
    proc = subprocess.Popen(cmd, cwd=str(cwd), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                            encoding="utf-8", errors="replace", bufsize=1, env=child_env())
    assert proc.stdout is not None
    for line in proc.stdout:
        say(line.rstrip("\n"))
    return proc.wait()


def main(argv: Optional[List[str]] = None) -> int:
    configure_console()
    ap = argparse.ArgumentParser(description="ONE-CLICK: git pull → test/build/validate → ZIP+version.json → publish LAN")
    ap.add_argument("--python", default=sys.executable, help="Python dùng để tạo .venv-build")
    ap.add_argument("--branch", default=None, help="nhánh phát hành bắt buộc (mặc định: nhánh hiện tại / RE_BUILD_BRANCH)")
    ap.add_argument("--no-sync", action="store_true", help="bỏ qua git fetch/pull (build đúng mã đang có)")
    ap.add_argument("--skip-tests", action="store_true")
    ap.add_argument("--no-publish", action="store_true", help="chỉ build, không xuất bản")
    ap.add_argument("--publish-dir", default=None, help=f"thư mục cập nhật (mặc định {publish_update.update_folder()})")
    ap.add_argument("--force", action="store_true", help="cho phép xuất bản lại cùng số Build")
    ap.add_argument("--allow-non-windows", action="store_true")
    ap.add_argument("--dry-run", action="store_true", help="chỉ kiểm tra Git + phiên bản, không build")
    args = ap.parse_args(argv)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    say = Tee(ROOT / "logs" / f"build_and_publish_{stamp}.log")
    try:
        say("=" * 70)
        say(f" Report Extractor – BUILD + PUBLISH ({stamp})   log: {say.path}")
        say("=" * 70)
        say("[1/4] Đồng bộ mã nguồn từ GitHub")
        if args.no_sync:
            say("   (bỏ qua theo --no-sync)")
        else:
            try:
                sync = git_sync(ROOT, args.branch, say)
            except SyncError as e:
                say(f"DỪNG: {e}")
                return RC_GIT
            for n in sync.notes:
                say(f"   lưu ý: {n}")
            say(f"   GIT_SYNC branch={sync.branch} revision={sync.revision} pulled={sync.pulled}")

        say("[2/4] Phiên bản sẽ build")
        import importlib
        import build_portable as bp
        importlib.reload(bp)                       # the pull may have changed app/__init__.py
        version, build_id = bp._version()
        build = bp.build_number()
        say(f"   {version} / Build {build_id}  →  {bp.package_name(version)}")
        folder = args.publish_dir or publish_update.update_folder()
        if not args.no_publish:
            try:
                say("   " + check_not_already_published(version, build, folder, args.force))
            except SyncError as e:
                say(f"DỪNG: {e}")
                return RC_REFUSED
            say(f"   Thư mục xuất bản: {folder}")
        if args.dry_run:
            say("   (--dry-run: dừng trước khi build)")
            return RC_OK

        say("[3/4] Test / build / validate / ZIP / SHA256 / version.json (tools/build_portable.py)")
        rc = run_streamed(build_command(args.python, args), say)
        if rc == bp.PUBLISH_FAILED_RC:
            say("[4/4] BUILD THÀNH CÔNG nhưng XUẤT BẢN THẤT BẠI – gói nằm trong release\\; "
                "chạy lại: python tools\\publish_update.py")
            return RC_PUBLISH
        if rc != 0:
            say(f"[4/4] BUILD THẤT BẠI (mã {rc}) – thư mục cập nhật KHÔNG bị thay đổi.")
            return RC_BUILD
        if args.no_publish:
            say("[4/4] Build xong (không xuất bản theo --no-publish).")
        else:
            say(f"[4/4] ĐÃ XUẤT BẢN {version} / Build {build_id} vào {folder} – các máy làm việc sẽ tự phát hiện.")
        return RC_OK
    finally:
        say.close()


if __name__ == "__main__":
    sys.exit(main())
