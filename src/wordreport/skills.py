"""Kho skill: mỗi skill là thư mục `skills/<name>/SKILL.md` (định dạng Agent Skills của Claude:
YAML frontmatter `name`, `description` + nội dung Markdown).

Agent áp dụng "progressive disclosure": system prompt chỉ chứa danh sách tên + mô tả, nội dung
đầy đủ chỉ được nạp khi model gọi tool `load_skill`. Cùng thư mục này có thể chép vào
`~/.claude/skills/` để Claude Code / Claude Desktop dùng trực tiếp.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

import yaml


@dataclass
class Skill:
    name: str
    description: str
    path: Path

    @property
    def body(self) -> str:
        return _split_frontmatter(self.path.read_text(encoding="utf-8"))[1]


def _split_frontmatter(text: str) -> tuple[dict, str]:
    if text.startswith("---"):
        _, fm, body = text.split("---", 2)
        return yaml.safe_load(fm) or {}, body.strip()
    return {}, text.strip()


def default_skills_dir() -> Path:
    env = os.environ.get("WORDREPORT_SKILLS_DIR")
    candidates = [Path(env)] if env else []
    candidates += [Path.cwd() / "skills", Path(__file__).resolve().parents[2] / "skills"]
    for candidate in candidates:
        if candidate.is_dir():
            return candidate
    raise FileNotFoundError("Không tìm thấy thư mục skills/ (đặt biến môi trường WORDREPORT_SKILLS_DIR).")


class SkillLibrary:
    def __init__(self, root: str | Path | None = None) -> None:
        self.root = Path(root) if root else default_skills_dir()
        self.skills: dict[str, Skill] = {}
        for skill_md in sorted(self.root.glob("*/SKILL.md")):
            meta, _ = _split_frontmatter(skill_md.read_text(encoding="utf-8"))
            name = meta.get("name") or skill_md.parent.name
            self.skills[name] = Skill(name=name, description=meta.get("description", ""), path=skill_md)

    def names(self) -> list[str]:
        return list(self.skills)

    def index(self) -> str:
        """Danh mục ngắn để nhúng vào system prompt."""
        return "\n".join(f"- {s.name}: {s.description}" for s in self.skills.values())

    def load(self, name: str) -> str:
        skill = self.skills.get(name)
        if skill is None:
            raise KeyError(f"Không có skill '{name}'. Có: {', '.join(self.names())}")
        return f"# Skill: {skill.name}\n\n{skill.body}"
