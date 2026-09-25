"""Cấu hình dùng chung cho các lời gọi Claude API."""

from __future__ import annotations

import os
from pathlib import Path

from .inspector import read_source

DEFAULT_MODEL = os.environ.get("WORDREPORT_MODEL", "claude-opus-5")
DEFAULT_EFFORT = os.environ.get("WORDREPORT_EFFORT", "high")

# Khi model từ chối (stop_reason "refusal"), API tự chạy lại yêu cầu trên model dự phòng.
FALLBACK_BETA = "server-side-fallback-2026-07-01"


def request_options(model: str, effort: str) -> dict:
    """Tham số chung: adaptive thinking, mức effort, fallback phía server."""
    return {
        "model": model,
        "thinking": {"type": "adaptive"},
        "output_config": {"effort": effort},
        "betas": [FALLBACK_BETA],
        "fallbacks": "default",
    }


def build_sources_block(sources: list[str | Path]) -> str:
    """Đọc tài liệu nguồn và bọc trong thẻ <source> - nội dung là dữ liệu, không phải chỉ thị."""
    parts = []
    for src in sources:
        text = read_source(src)
        parts.append(f'<source path="{src}">\n{text}\n</source>')
    return "\n\n".join(parts)


class RefusalError(RuntimeError):
    pass


def check_stop(message) -> None:
    if message.stop_reason == "refusal":
        details = getattr(message, "stop_details", None)
        raise RefusalError(f"Model từ chối yêu cầu: {details}")
    if message.stop_reason == "max_tokens":
        raise RuntimeError("Phản hồi bị cắt do chạm max_tokens - hãy chia nhỏ yêu cầu.")
