import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

from make_samples import make_report, make_template, make_sample_tree  # noqa: E402


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
