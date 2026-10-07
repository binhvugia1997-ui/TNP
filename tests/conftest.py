import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

from make_samples import make_report, make_template, make_sample_tree  # noqa: E402


@pytest.fixture(autouse=True)
def _isolated_runtime_root(tmp_path, monkeypatch):
    """All default writable application state must stay inside this test's disposable directory.

    Covers config/report/output preferences, manual-field retry state, logs/history, learning data and updater state;
    individual tests may still inject a narrower path when that is the behavior under test.
    """
    import app.runtime_paths as rp
    monkeypatch.setenv(rp.TEST_RUNTIME_ROOT_ENV, str(tmp_path / "tnp-runtime"))
    yield


@pytest.fixture(autouse=True)
def _never_spawn_real_updater(monkeypatch):
    """Safety net: no test may launch the real external updater (it would apply a fake release onto the repository,
    which is the dev portable_root).  Tests that need a spawn pass their own recorder; a default spawn fails loudly."""
    import app.updater as up

    def _forbidden(cmd, **kwargs):
        raise AssertionError(f"test tried to spawn the real updater: {cmd}")
    monkeypatch.setattr(up.launch_updater, "__defaults__", (_forbidden,))
    yield


@pytest.fixture(scope="session")
def sample_tree(tmp_path_factory):
    root = tmp_path_factory.mktemp("sample")
    return make_sample_tree(root)


@pytest.fixture(scope="session")
def a185_report(sample_tree):
    return sample_tree["files"][0]


@pytest.fixture
def template(tmp_path):
    return make_template(tmp_path / "Verification.xlsx")


@pytest.fixture
def report_factory(tmp_path):
    def _mk(name="r.pptx", **kw):
        return make_report(tmp_path / name, **kw)
    return _mk
