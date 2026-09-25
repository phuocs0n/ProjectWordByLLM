import json
from pathlib import Path

import pytest

from wordreport.spec import ReportSpec

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture
def example_spec() -> ReportSpec:
    return ReportSpec.model_validate(json.loads((ROOT / "examples" / "mang-may-tinh-do-an.json").read_text(encoding="utf-8")))


@pytest.fixture(autouse=True)
def _skills_dir(monkeypatch):
    monkeypatch.setenv("WORDREPORT_SKILLS_DIR", str(ROOT / "skills"))
