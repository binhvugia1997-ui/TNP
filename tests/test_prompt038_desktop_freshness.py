"""PROMPT-038: source mode must never silently serve a stale pre-build frontend bundle.

`frontend/dist` is gitignored/untracked, so merging newer `frontend/src` cannot update it; the
Windows integration acceptance after PROMPT-034/037 still showed the pre-merge UI (unbounded
report list, visible "Bỏ qua" rows) for exactly this reason. The freshness guard in
`app.desktop` compares the newest frontend source input against `dist/index.html` and `main()`
refuses to start (source mode only) with the rebuild command. These tests pin that contract.
"""
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import desktop


def _make_frontend(
    tmp_path: Path,
    *,
    dist_mtime: float,
    src_mtime: float,
    index_html_mtime: float | None = None,
) -> Path:
    """Build a tiny Vite-like tree and set every freshness input to an explicit timestamp."""
    frontend = tmp_path / "frontend"
    dist = frontend / "dist"
    assets = dist / "assets"
    assets.mkdir(parents=True)
    source = frontend / "src"
    source.mkdir()
    source_css = source / "index.css"
    source_css.write_text(".report-list-viewport{}", encoding="utf-8")
    source_index = frontend / "index.html"
    source_index.write_text("<html></html>", encoding="utf-8")
    (assets / "index-abc.js").write_text("js", encoding="utf-8")
    (assets / "index-def.css").write_text("css", encoding="utf-8")
    index = dist / "index.html"
    index.write_text('<script src="./assets/index-abc.js"></script>', encoding="utf-8")

    # Use explicit whole-second test timestamps, separated by >=100 seconds where ordering matters.
    # This avoids write-time drift and filesystem/sub-second precision differences on Windows/POSIX.
    source_index_mtime = src_mtime if index_html_mtime is None else index_html_mtime
    os.utime(source_css, (src_mtime, src_mtime))
    os.utime(source_index, (source_index_mtime, source_index_mtime))
    os.utime(index, (dist_mtime, dist_mtime))
    return dist


def test_fresh_dist_is_not_stale(tmp_path):
    now = int(time.time())
    dist = _make_frontend(tmp_path, dist_mtime=now, src_mtime=now - 100)
    report = desktop.frontend_dist_freshness(dist)
    assert report["stale"] is False
    assert report["reason"] == "fresh"
    assert report["source_mtime"] < report["dist_mtime"]


def test_source_newer_than_dist_is_stale_and_reports_the_newest_source(tmp_path):
    now = int(time.time())
    dist = _make_frontend(tmp_path, dist_mtime=now - 100, src_mtime=now, index_html_mtime=now - 200)
    report = desktop.frontend_dist_freshness(dist)
    assert report["stale"] is True
    assert report["reason"] == "source-newer-than-dist"
    assert Path(report["newest_source"]).name == "index.css"
    assert report["source_mtime"] > report["dist_mtime"]


def test_index_html_remains_a_freshness_input(tmp_path):
    now = int(time.time())
    dist = _make_frontend(tmp_path, dist_mtime=now - 100, src_mtime=now - 200, index_html_mtime=now)
    report = desktop.frontend_dist_freshness(dist)
    assert report["stale"] is True
    assert report["reason"] == "source-newer-than-dist"
    assert Path(report["newest_source"]).name == "index.html"


def test_equal_source_mtimes_choose_the_same_canonical_path(tmp_path):
    now = int(time.time())
    dist = _make_frontend(tmp_path, dist_mtime=now - 100, src_mtime=now, index_html_mtime=now)
    report = desktop.frontend_dist_freshness(dist)
    assert report["stale"] is True
    # Equal source timestamps are diagnostic ties; normalized relative-path order chooses index.html first.
    assert Path(report["newest_source"]).name == "index.html"
    assert report["source_mtime"] > report["dist_mtime"]


def test_missing_dist_is_stale(tmp_path):
    report = desktop.frontend_dist_freshness(tmp_path / "frontend" / "dist")
    assert report["stale"] is True
    assert report["reason"] == "dist-missing"


def test_no_source_tree_cannot_be_judged_and_is_allowed(tmp_path):
    dist = tmp_path / "frontend" / "dist"
    dist.mkdir(parents=True)
    (dist / "index.html").write_text("<html></html>", encoding="utf-8")
    report = desktop.frontend_dist_freshness(dist)
    assert report["stale"] is False
    assert report["reason"] == "no-source-tree"


def test_asset_report_names_hashed_assets_with_digests(tmp_path):
    now = int(time.time())
    dist = _make_frontend(tmp_path, dist_mtime=now, src_mtime=now - 100)
    report = desktop.frontend_asset_report(dist)
    assert len(report["index_sha256"]) == 12
    assert any(name.startswith("index-abc.js#") for name in report["assets"])
    assert any(name.startswith("index-def.css#") for name in report["assets"])
