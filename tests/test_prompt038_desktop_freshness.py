"""PROMPT-038: source mode must never silently serve a stale pre-build frontend bundle.

`frontend/dist` is gitignored/untracked, so merging newer `frontend/src` cannot update it; the
Windows integration acceptance after PROMPT-034/037 still showed the pre-merge UI (unbounded
report list, visible "Bỏ qua" rows) for exactly this reason. The freshness guard in
`app.desktop` compares the newest frontend source input against `dist/index.html` and main()
refuses to start (source mode only) with the rebuild command. These tests pin that contract.
"""
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import desktop


def _make_frontend(tmp_path: Path, *, dist_mtime: float, src_mtime: float) -> Path:
    frontend = tmp_path / "frontend"
    dist = frontend / "dist"
    assets = dist / "assets"
    assets.mkdir(parents=True)
    (frontend / "src").mkdir()
    (frontend / "src" / "index.css").write_text(".report-list-viewport{}", encoding="utf-8")
    (frontend / "index.html").write_text("<html></html>", encoding="utf-8")
    (assets / "index-abc.js").write_text("js", encoding="utf-8")
    (assets / "index-def.css").write_text("css", encoding="utf-8")
    index = dist / "index.html"
    index.write_text('<script src="./assets/index-abc.js"></script>', encoding="utf-8")
    os.utime(frontend / "src" / "index.css", (src_mtime, src_mtime))
    os.utime(index, (dist_mtime, dist_mtime))
    return dist


def test_fresh_dist_is_not_stale(tmp_path):
    now = time.time()
    dist = _make_frontend(tmp_path, dist_mtime=now, src_mtime=now - 100)
    report = desktop.frontend_dist_freshness(dist)
    assert report["stale"] is False
    assert report["reason"] == "fresh"


def test_source_newer_than_dist_is_stale(tmp_path):
    now = time.time()
    dist = _make_frontend(tmp_path, dist_mtime=now - 100, src_mtime=now)
    report = desktop.frontend_dist_freshness(dist)
    assert report["stale"] is True
    assert report["reason"] == "source-newer-than-dist"
    assert report["newest_source"].endswith("index.css")


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
    now = time.time()
    dist = _make_frontend(tmp_path, dist_mtime=now, src_mtime=now - 100)
    report = desktop.frontend_asset_report(dist)
    assert len(report["index_sha256"]) == 12
    assert any(name.startswith("index-abc.js#") for name in report["assets"])
    assert any(name.startswith("index-def.css#") for name in report["assets"])
