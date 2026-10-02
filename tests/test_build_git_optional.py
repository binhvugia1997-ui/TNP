"""PROMPT-004E – the portable build must never require Git (revision is informational metadata only)."""
import subprocess
import sys
from pathlib import Path

import pytest

import app

sys.path.insert(0, str(Path(app.__file__).resolve().parent.parent / "tools"))
import build_portable as bp  # noqa: E402


class _R:
    def __init__(self, out="", code=0):
        self.stdout, self.returncode = out, code


# 1. git installed + valid repo -> short revision used normally
def test_git_available_returns_short_revision(monkeypatch):
    seen = {}

    def fake_run(cmd, **kw):
        seen["cmd"], seen["kw"] = cmd, kw
        return _R("abc1234\n")
    monkeypatch.setattr(bp.subprocess, "run", fake_run)
    assert bp.get_git_revision(Path("/repo")) == "abc1234"
    assert seen["cmd"] == ["git", "rev-parse", "--short", "HEAD"] and seen["kw"]["cwd"] == "/repo"
    assert seen["kw"]["check"] is True and seen["kw"]["timeout"] > 0


# 2. git executable missing (WinError 2 on Windows)
def test_git_executable_missing_falls_back(monkeypatch, capsys):
    def fake_run(cmd, **kw):
        raise FileNotFoundError(2, "The system cannot find the file specified")
    monkeypatch.setattr(bp.subprocess, "run", fake_run)
    assert bp.get_git_revision() == "unavailable" == bp.GIT_UNAVAILABLE
    assert "FileNotFoundError" in capsys.readouterr().out


# 3. git installed but the folder is not a repository (exit 128)
def test_not_a_repository_falls_back(monkeypatch):
    def fake_run(cmd, **kw):
        raise subprocess.CalledProcessError(128, cmd, stderr="fatal: not a git repository")
    monkeypatch.setattr(bp.subprocess, "run", fake_run)
    assert bp.get_git_revision() == "unavailable"


def test_real_git_outside_a_repository_falls_back(tmp_path):
    """Real subprocess: an empty temp dir is never a repo; works whether or not git is installed."""
    assert bp.get_git_revision(tmp_path) == "unavailable"


# 4. other git command failures: timeout, OSError, empty output
@pytest.mark.parametrize("exc", [subprocess.TimeoutExpired(["git"], 10), OSError(13, "Permission denied"),
                                 PermissionError("locked"), RuntimeError("weird")])
def test_git_command_failure_never_raises(monkeypatch, exc):
    def fake_run(cmd, **kw):
        raise exc
    monkeypatch.setattr(bp.subprocess, "run", fake_run)
    assert bp.get_git_revision() == "unavailable"


def test_empty_git_output_falls_back(monkeypatch):
    monkeypatch.setattr(bp.subprocess, "run", lambda cmd, **kw: _R("   \n"))
    assert bp.get_git_revision() == "unavailable"


# 5. fallback metadata written correctly (VERSION.txt / README.txt remain valid)
def test_fallback_metadata_written_correctly(tmp_path, monkeypatch):
    def fake_run(cmd, **kw):
        raise FileNotFoundError(2, "git")
    monkeypatch.setattr(bp.subprocess, "run", fake_run)
    folder = tmp_path / "ReportExtractor_v1.0.3_Portable"
    folder.mkdir()
    rev = bp.write_release_metadata(folder, "1.0.3", "PROMPT-003", folder.name, built="2026-10-02 10:00")
    assert rev == "unavailable"
    version_txt = (folder / "VERSION.txt").read_text(encoding="utf-8-sig")
    assert "git=unavailable" in version_txt and "Git revision: unavailable" in version_txt
    assert "version=1.0.3" in version_txt and "prompt=PROMPT-003" in version_txt and "built=2026-10-02 10:00" in version_txt
    readme = (folder / "README.txt").read_text(encoding="utf-8-sig")
    assert "Git: unavailable" in readme and "{git}" not in readme
    for f in ("FIRST_RUN.txt", "Install_Ollama_Optional.bat"):
        assert (folder / f).is_file()


def test_metadata_with_git_available_uses_revision(tmp_path, monkeypatch):
    monkeypatch.setattr(bp.subprocess, "run", lambda cmd, **kw: _R("deadbee\n"))
    folder = tmp_path / "P"
    folder.mkdir()
    assert bp.write_release_metadata(folder, "1.0.3", "PROMPT-003", "P") == "deadbee"
    assert "git=deadbee" in (folder / "VERSION.txt").read_text(encoding="utf-8-sig")


# 6. the build does not fail because Git metadata is unavailable: step 8 + step 9 validation succeed
def test_build_metadata_and_artifact_validation_succeed_without_git(tmp_path, monkeypatch):
    def fake_run(cmd, **kw):
        raise FileNotFoundError(2, "git")
    monkeypatch.setattr(bp.subprocess, "run", fake_run)
    folder = tmp_path / "ReportExtractor_v1.0.3_Portable"
    (folder / "_internal").mkdir(parents=True)
    (folder / "ReportExtractor.exe").write_bytes(b"MZ")
    for sub in ("Output", "logs", "config"):
        (folder / sub).mkdir()
    bp.write_release_metadata(folder, "1.0.3", "PROMPT-003", folder.name)       # must not raise
    assert bp.validate_artifact(folder) == []
    # main() no longer calls git directly: the only git invocation lives in the safe helper
    src = Path(bp.__file__).read_text(encoding="utf-8")
    body = src.split("def main(")[1]
    assert "rev-parse" not in body and src.count('"rev-parse"') == 1
    assert "write_release_metadata(folder, version, build_id, name)" in body
