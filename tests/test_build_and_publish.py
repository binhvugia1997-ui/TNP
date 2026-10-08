"""PROMPT-012 – ONE-CLICK BUILD + PUBLISH: safe Git sync, published-build guard, orchestration, BUILD_AND_PUBLISH.bat."""
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

import app
import build_and_publish as bap
from build_and_publish import (RC_BUILD, RC_GIT, RC_OK, RC_PUBLISH, RC_REFUSED, SyncError, build_command,
                               check_not_already_published, git_sync)

ROOT = Path(__file__).resolve().parent.parent


# ------------------------------------------------------------------ git fixtures (real git, temp repos)
def _git(cwd: Path, *args: str) -> str:
    env = dict(os.environ, GIT_AUTHOR_NAME="t", GIT_AUTHOR_EMAIL="t@x", GIT_COMMITTER_NAME="t", GIT_COMMITTER_EMAIL="t@x")
    cp = subprocess.run(["git", *args], cwd=str(cwd), capture_output=True, text=True, env=env)
    assert cp.returncode == 0, cp.stderr
    return cp.stdout.strip()


def _commit(repo: Path, name: str, text: str, msg: str = "c") -> str:
    (repo / name).write_text(text, encoding="utf-8")
    _git(repo, "add", name)
    _git(repo, "commit", "-q", "-m", msg)
    return _git(repo, "rev-parse", "HEAD")


@pytest.fixture
def repos(tmp_path):
    """bare 'github' remote + 'dev' clone (the agent) + 'build' clone (the build PC), branch main."""
    remote = tmp_path / "remote.git"
    _git(tmp_path, "init", "-q", "--bare", "-b", "main", str(remote))
    dev = tmp_path / "dev"
    _git(tmp_path, "clone", "-q", str(remote), str(dev))
    _git(dev, "checkout", "-q", "-b", "main")
    _commit(dev, "a.txt", "1", "first")
    _git(dev, "push", "-q", "-u", "origin", "main")
    build = tmp_path / "build"
    _git(tmp_path, "clone", "-q", str(remote), str(build))
    return dev, build, remote


def _quiet(_msg: str = "") -> None:
    pass


def test_sync_up_to_date(repos):
    dev, build, _ = repos
    r = git_sync(build, say=_quiet)
    assert r.branch == "main" and r.upstream == "origin/main" and not r.pulled and r.before == r.after
    assert r.ahead == 0 and r.behind == 0 and r.revision == r.after[:12]


def test_sync_pulls_fast_forward(repos):
    dev, build, _ = repos
    new = _commit(dev, "a.txt", "2", "second")
    _git(dev, "push", "-q")
    lines = []
    r = git_sync(build, say=lines.append)
    assert r.pulled and r.behind == 1 and r.after == new and (build / "a.txt").read_text() == "2"
    assert any("pull --ff-only" in ln for ln in lines) and any("second" in ln for ln in lines)


def test_sync_refuses_dirty_tracked_files(repos):
    dev, build, _ = repos
    (build / "a.txt").write_text("edited locally", encoding="utf-8")
    with pytest.raises(SyncError, match="chưa commit"):
        git_sync(build, say=_quiet)
    assert (build / "a.txt").read_text() == "edited locally"          # nothing discarded


def test_sync_allows_untracked_and_ignored_artifacts(repos):
    dev, build, _ = repos
    (build / "dist").mkdir()
    (build / "dist" / "x.exe").write_bytes(b"MZ")
    (build / "scratch.txt").write_text("untracked")
    assert git_sync(build, say=_quiet).branch == "main"


def test_sync_refuses_local_unpushed_commits(repos):
    dev, build, _ = repos
    _commit(build, "b.txt", "local", "local only")
    with pytest.raises(SyncError, match="chưa đẩy lên GitHub"):
        git_sync(build, say=_quiet)


def test_sync_refuses_diverged(repos):
    dev, build, _ = repos
    _commit(build, "b.txt", "local", "local only")
    _commit(dev, "c.txt", "remote", "remote only")
    _git(dev, "push", "-q")
    head = _git(build, "rev-parse", "HEAD")
    with pytest.raises(SyncError, match="phân nhánh"):
        git_sync(build, say=_quiet)
    assert _git(build, "rev-parse", "HEAD") == head                   # no merge / rebase / reset happened


def test_sync_refuses_wrong_branch(repos, monkeypatch):
    dev, build, _ = repos
    _git(build, "checkout", "-q", "-b", "feature")
    with pytest.raises(SyncError, match="nhánh phát hành yêu cầu là 'main'"):
        git_sync(build, branch="main", say=_quiet)
    monkeypatch.setenv("RE_BUILD_BRANCH", "main")
    with pytest.raises(SyncError, match="git checkout main"):
        git_sync(build, say=_quiet)
    monkeypatch.delenv("RE_BUILD_BRANCH")
    with pytest.raises(SyncError, match="không có trên origin"):      # feature never pushed
        git_sync(build, say=_quiet)


def test_sync_refuses_detached_head_and_non_repo(repos, tmp_path):
    dev, build, _ = repos
    _git(build, "checkout", "-q", "--detach")
    with pytest.raises(SyncError, match="detached HEAD"):
        git_sync(build, say=_quiet)
    plain = tmp_path / "plain"
    plain.mkdir()
    with pytest.raises(SyncError, match="không phải là bản sao Git"):
        git_sync(plain, say=_quiet)


def test_sync_without_upstream_uses_origin_branch(repos):
    dev, build, _ = repos
    _git(build, "branch", "--unset-upstream")
    r = git_sync(build, say=_quiet)
    assert r.upstream == "origin/main" and any("upstream" in n for n in r.notes)


def test_sync_git_missing(monkeypatch, tmp_path):
    def boom(*a, **k):
        raise FileNotFoundError("git")
    monkeypatch.setattr(bap.subprocess, "run", boom)
    with pytest.raises(SyncError, match="Không tìm thấy lệnh git"):
        git_sync(tmp_path, say=_quiet)


# ------------------------------------------------------------------ published-build guard
def test_published_guard(tmp_path):
    folder = tmp_path / "Update"
    assert "chưa có version.json" in check_not_already_published("1.2.1", 12, str(folder))
    folder.mkdir()
    (folder / "version.json").write_text(json.dumps({"version": "1.2.0", "build": 11}), encoding="utf-8")
    assert "Build 011  →  mới: 1.2.1 / Build 012" in check_not_already_published("1.2.1", 12, str(folder))
    (folder / "version.json").write_text(json.dumps({"version": "1.2.1", "build": 12}), encoding="utf-8")
    with pytest.raises(SyncError, match="đã có 1.2.1 / Build 012"):
        check_not_already_published("1.2.1", 12, str(folder))
    (folder / "version.json").write_text(json.dumps({"version": "9.9", "build": 99}), encoding="utf-8")
    with pytest.raises(SyncError, match=">= Build 012"):
        check_not_already_published("1.2.1", 12, str(folder))
    assert "--force" in check_not_already_published("1.2.1", 12, str(folder), force=True)
    (folder / "version.json").write_text("not json", encoding="utf-8")
    assert "chưa có version.json" in check_not_already_published("1.2.1", 12, str(folder))   # unreadable = none


# ------------------------------------------------------------------ orchestration (build step mocked)
def _run_main(monkeypatch, tmp_path, argv, build_rc=0, sync=None):
    monkeypatch.setattr(bap, "ROOT", tmp_path)
    calls = {}
    if sync is None:
        monkeypatch.setattr(bap, "git_sync", lambda root, branch, say: bap.SyncResult(branch="main", before="a" * 40,
                                                                                      after="a" * 40))
    else:
        monkeypatch.setattr(bap, "git_sync", sync)

    def fake_run(cmd, say, cwd=None):
        calls["cmd"] = cmd
        say("   (fake build output)")
        return build_rc
    monkeypatch.setattr(bap, "run_streamed", fake_run)
    rc = bap.main(argv)
    logs = sorted((tmp_path / "logs").glob("build_and_publish_*.log"))
    return rc, calls, logs[-1].read_text(encoding="utf-8") if logs else ""


def test_main_full_success(monkeypatch, tmp_path):
    rc, calls, log = _run_main(monkeypatch, tmp_path, ["--publish-dir", str(tmp_path / "Update"), "--python", "PY"])
    assert rc == RC_OK
    assert calls["cmd"][1:] == [str(tmp_path / "tools" / "build_portable.py"), "--python", "PY",
                                "--publish-dir", str(tmp_path / "Update")]
    assert f"ĐÃ XUẤT BẢN {app.__version__} / Build {app.BUILD_ID}" in log and "GIT_SYNC branch=main" in log
    assert f"ReportExtractor_{app.__version__}.zip" in log


def test_main_stops_on_git_error_before_build(monkeypatch, tmp_path):
    def bad(root, branch, say):
        raise SyncError("Có 1 commit cục bộ chưa đẩy lên GitHub")
    rc, calls, log = _run_main(monkeypatch, tmp_path, ["--publish-dir", str(tmp_path / "U")], sync=bad)
    assert rc == RC_GIT and "cmd" not in calls and "DỪNG: Có 1 commit cục bộ" in log


def test_main_refuses_already_published(monkeypatch, tmp_path):
    u = tmp_path / "U"
    u.mkdir()
    (u / "version.json").write_text(json.dumps({"version": app.__version__, "build": app.BUILD_NUMBER}), encoding="utf-8")
    rc, calls, log = _run_main(monkeypatch, tmp_path, ["--publish-dir", str(u)])
    assert rc == RC_REFUSED and "cmd" not in calls and "--force" in log
    rc, calls, log = _run_main(monkeypatch, tmp_path, ["--publish-dir", str(u), "--force"])
    assert rc == RC_OK and "cmd" in calls


def test_main_build_failure_and_publish_failure_codes(monkeypatch, tmp_path):
    rc, _, log = _run_main(monkeypatch, tmp_path, ["--publish-dir", str(tmp_path / "U")], build_rc=1)
    assert rc == RC_BUILD and "KHÔNG bị thay đổi" in log
    rc, _, log = _run_main(monkeypatch, tmp_path, ["--publish-dir", str(tmp_path / "U")], build_rc=3)
    assert rc == RC_PUBLISH and "publish_update.py" in log


def test_main_dry_run_and_no_sync_and_no_publish(monkeypatch, tmp_path):
    rc, calls, log = _run_main(monkeypatch, tmp_path, ["--dry-run", "--no-sync", "--publish-dir", str(tmp_path / "U")])
    assert rc == RC_OK and "cmd" not in calls and "--no-sync" in log and "--dry-run" in log
    rc, calls, log = _run_main(monkeypatch, tmp_path, ["--no-publish", "--skip-tests"])
    assert rc == RC_OK and "--no-publish" in calls["cmd"] and "--skip-tests" in calls["cmd"]
    assert "không xuất bản" in log


def test_build_command_flags():
    ns = bap.argparse.Namespace(skip_tests=True, no_publish=False, publish_dir=r"D:\U", allow_non_windows=True)
    cmd = build_command("PYEXE", ns)
    assert cmd[0] == sys.executable and cmd[2:] == ["--python", "PYEXE", "--skip-tests", "--publish-dir", r"D:\U",
                                                    "--allow-non-windows"]


# ------------------------------------------------------------------ the deliverable
def test_build_and_publish_bat_exists_and_wires_everything():
    bat = ROOT / "BUILD_AND_PUBLISH.bat"
    assert bat.exists(), "BUILD_AND_PUBLISH.bat must exist at the repository root"
    text = bat.read_text(encoding="utf-8")
    assert "tools\\build_and_publish.py" in text and 'cd /d "%~dp0"' in text
    assert "where git" in text and "3.13 3.12 3.11 3.10" in text and "pause" in text
    for rc in ("0", "2", "3", "4"):
        assert f'"%RC%"=="{rc}"' in text
    assert "exit /b %RC%" in text
    src = (ROOT / "tools" / "build_and_publish.py").read_text(encoding="utf-8")
    import re
    git_calls = re.findall(r'_git\(\[([^\]]*)\]', src)        # every git invocation of the sync layer
    joined = " ".join(git_calls)
    for forbidden in ('"reset"', '"stash"', '"rebase"', '"merge"', '"checkout"', '"clean"', '"--force"', '"-f"'):
        assert forbidden not in joined, forbidden            # never discards or rewrites user work
    assert '"--ff-only"' in joined


def test_identity_build_017():
    assert app.__version__ == "1.3.4" and app.BUILD_NUMBER == 17 and app.BUILD_LABEL == "Build 017"
    import build_portable as bp
    assert bp.package_name(app.__version__) == "ReportExtractor_1.3.4.zip" and bp._version() == ("1.3.4", "017")
