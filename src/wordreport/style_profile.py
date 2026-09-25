"""Nạp style profile (YAML) - hỗ trợ kế thừa qua khoá `extends`."""

from __future__ import annotations

import copy
from pathlib import Path
from typing import Any

import yaml

PROFILE_DIR = Path(__file__).parent / "profiles"


def _deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    out = copy.deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = _deep_merge(out[key], value)
        else:
            out[key] = copy.deepcopy(value)
    return out


def _resolve_path(name_or_path: str) -> Path:
    candidate = Path(name_or_path)
    if candidate.suffix in {".yaml", ".yml"} and candidate.exists():
        return candidate
    builtin = PROFILE_DIR / f"{name_or_path}.yaml"
    if builtin.exists():
        return builtin
    raise FileNotFoundError(
        f"Không tìm thấy style profile '{name_or_path}'. Có sẵn: {', '.join(list_profiles())}"
    )


def load_profile(name_or_path: str = "hcmus-clc") -> dict[str, Any]:
    path = _resolve_path(name_or_path)
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    parent = data.pop("extends", None)
    if parent:
        data = _deep_merge(load_profile(parent), data)
    return data


def list_profiles() -> list[str]:
    return sorted(p.stem for p in PROFILE_DIR.glob("*.yaml"))


def describe_profiles() -> list[dict[str, str]]:
    return [{"name": n, "description": load_profile(n).get("description", "")} for n in list_profiles()]
