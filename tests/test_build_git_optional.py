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
def test_git_available_returns_short_revision(monkeypatch, tmp_path):
    seen = {}

    def fake_run(cmd, **kw):
        seen["cmd"], seen["kw"] = cmd, kw
        return _R("abc1234\n")
    monkeypatch.setattr(bp.subprocess, "run", fake_run)
    assert bp.get_git_revision(tmp_path) == "abc1234"
    assert seen["cmd"] == ["git", "rev-parse", "--short", "HEAD"]
    assert Path(seen["kw"]["cwd"]).resolve() == tmp_path.resolve()          # platform-independent cwd check
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
    rev = bp.write_release_metadata(folder, "1.0.4", "004", folder.name, built="2026-10-02 10:00")
    assert rev == "unavailable"
    version_txt = (folder / "VERSION.txt").read_text(encoding="utf-8-sig")
    assert "git=unavailable" in version_txt and "Git revision: unavailable" in version_txt
    assert "version=1.0.4" in version_txt and "build=004" in version_txt and "prompt=" not in version_txt and "built=2026-10-02 10:00" in version_txt
    readme = (folder / "README.txt").read_text(encoding="utf-8-sig")
    assert "Git (nội bộ): unavailable" in readme and "{git}" not in readme
    for f in ("FIRST_RUN.txt", "Install_Ollama_Optional.bat"):
        assert (folder / f).is_file()


def test_metadata_with_git_available_uses_revision(tmp_path, monkeypatch):
    monkeypatch.setattr(bp.subprocess, "run", lambda cmd, **kw: _R("deadbee\n"))
    folder = tmp_path / "P"
    folder.mkdir()
    assert bp.write_release_metadata(folder, "1.0.4", "004", "P") == "deadbee"
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
    bp.write_release_metadata(folder, "1.0.4", "004", folder.name)       # must not raise
    assert bp.validate_artifact(folder) == []
    # main() no longer calls git directly: the only git invocation lives in the safe helper
    src = Path(bp.__file__).read_text(encoding="utf-8")
    body = src.split("def main(")[1]
    assert "rev-parse" not in body and src.count('"rev-parse"') == 1
    assert "write_release_metadata(folder, version, build_id, name)" in body


# ---------------------------------------------------------------- PROMPT-004F: validator vs dependency runtime resources
def _portable_skeleton(tmp_path):
    folder = tmp_path / "ReportExtractor_v1.0.3_Portable"
    (folder / "_internal").mkdir(parents=True)
    (folder / "ReportExtractor.exe").write_bytes(b"MZ")
    for sub in ("Output", "logs", "config"):
        (folder / sub).mkdir()
    for f in ("README.txt", "FIRST_RUN.txt", "VERSION.txt", "Install_Ollama_Optional.bat"):
        (folder / f).write_text("x", encoding="utf-8")
    return folder


def test_f1_python_pptx_default_template_is_allowed(tmp_path):
    folder = _portable_skeleton(tmp_path)
    tpl = folder / "_internal" / "pptx" / "templates" / "default.pptx"
    tpl.parent.mkdir(parents=True)
    tpl.write_bytes(b"PK")
    assert bp.validate_artifact(folder) == []


def test_f2_windows_and_posix_representations_are_allowed():
    assert bp.is_allowed_dependency_resource("_internal/pptx/templates/default.pptx")
    assert bp.is_allowed_dependency_resource("_internal\\pptx\\templates\\default.pptx")
    assert bp.is_allowed_dependency_resource(Path("_internal", "pptx", "templates", "default.pptx"))
    assert bp.is_allowed_dependency_resource("_INTERNAL\\PPTX\\Templates\\Default.PPTX")      # case-insensitive FS


def test_f3_arbitrary_pptx_under_package_root_rejected(tmp_path):
    folder = _portable_skeleton(tmp_path)
    (folder / "default.pptx").write_bytes(b"PK")
    (folder / "(CTMS)_20506_260918080-VOC_A185_Rear.pptx").write_bytes(b"PK")
    problems = bp.validate_artifact(folder)
    assert any("default.pptx" in p for p in problems) and any("260918080-VOC" in p for p in problems)


def test_f4_report_pptx_under_real_data_rejected(tmp_path):
    folder = _portable_skeleton(tmp_path)
    for where in (folder / "real_data", folder / "_internal" / "real_data", folder / "_internal" / "September"):
        where.mkdir(parents=True)
        (where / "260925015_bao_cao.pptx").write_bytes(b"PK")
    problems = bp.validate_artifact(folder)
    assert len([p for p in problems if "260925015_bao_cao.pptx" in p]) == 3


def test_f5_kiem_chung_and_result_xlsx_rejected(tmp_path):
    folder = _portable_skeleton(tmp_path)
    (folder / "Kiem_chung.xlsx").write_bytes(b"PK")
    (folder / "Output" / "Kiem_chung_09_2026.xlsx").write_bytes(b"PK")
    (folder / "_internal" / "Kiem_chung.xlsx").write_bytes(b"PK")
    problems = bp.validate_artifact(folder)
    assert len([p for p in problems if "Kiem_chung" in p]) == 3


def test_f6_machine_config_logs_and_ollama_rejected(tmp_path):
    folder = _portable_skeleton(tmp_path)
    (folder / "config" / "config.json").write_text("{}", encoding="utf-8")
    (folder / "config.json").write_text("{}", encoding="utf-8")
    (folder / "_internal" / "ollama.exe").write_bytes(b"x")
    (folder / "_internal" / "qwen3-4b.gguf").write_bytes(b"x")
    problems = bp.validate_artifact(folder)
    assert any("config/config.json" in p for p in problems) and any("config.json của máy dev" in p for p in problems)
    assert any("ollama.exe" in p for p in problems) and any(".gguf" in p for p in problems)
    # log / cache / pyc dev artefacts stay forbidden as before
    (folder / "_internal" / "__pycache__").mkdir()
    (folder / "_internal" / "__pycache__" / "x.pyc").write_bytes(b"x")
    (folder / "tests").mkdir()
    problems = bp.validate_artifact(folder)
    assert any("__pycache__" in p for p in problems) and any("tests" in p for p in problems)


def test_f7_genuine_dev_user_data_next_to_the_allowed_template_still_caught(tmp_path):
    folder = _portable_skeleton(tmp_path)
    tdir = folder / "_internal" / "pptx" / "templates"
    tdir.mkdir(parents=True)
    (tdir / "default.pptx").write_bytes(b"PK")                 # allowed
    (tdir / "report.pptx").write_bytes(b"PK")                  # NOT allowed – same folder, different file
    (folder / "_internal" / "pptx" / "default.pptx").write_bytes(b"PK")   # NOT allowed – different folder
    (folder / "_internal" / "sample_data" / "x").mkdir(parents=True)
    problems = bp.validate_artifact(folder)
    assert not any(str(Path("_internal", "pptx", "templates", "default.pptx")) in p for p in problems)
    assert any("report.pptx" in p for p in problems)
    assert any(str(Path("_internal", "pptx", "default.pptx")) in p for p in problems)
    assert any("sample_data" in p for p in problems)


def test_f8_no_broad_internal_pptx_exemption(tmp_path):
    assert bp.ALLOWED_DEPENDENCY_RESOURCES == (("_internal", "pptx", "templates", "default.pptx"),)
    assert not bp.is_allowed_dependency_resource("_internal/anything.pptx")
    assert not bp.is_allowed_dependency_resource("_internal/pptx/templates/other.pptx")
    assert not bp.is_allowed_dependency_resource("_internal/pptx/templates/default.pptx/extra.pptx")
    assert not bp.is_allowed_dependency_resource("pptx/templates/default.pptx")
    folder = _portable_skeleton(tmp_path)
    n = 0
    for name in ("a.pptx", "b.ppt", "c.xlsx"):
        (folder / "_internal" / name).write_bytes(b"PK")
        n += 1
    assert len([p for p in bp.validate_artifact(folder) if "dữ liệu mẫu/dev" in p]) == n
