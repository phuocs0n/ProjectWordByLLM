"""Đánh số tiêu đề, bảng, hình một cách tất định.

Báo cáo gốc (PDF mẫu) bị lệch số thứ tự ("Bài 2" thiếu số, mục "4.1" nằm trong bài 2).
Renderer tự đánh số nên LLM chỉ cần cung cấp tiêu đề trơn.

Hai kiểu đánh số (khoá `heading_numbering` trong style profile):
  section: I. / 1. / 1.1.        (mục con đánh lại từ 1 trong mỗi chương)
  chapter: I. / 1.1. / 1.1.1.    (mục con mang số chương; chú thích "Hình 2.3", "Bảng 1.2")
"""

from __future__ import annotations

_ROMAN = [
    (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
    (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
]


def to_roman(n: int) -> str:
    out = []
    for value, symbol in _ROMAN:
        while n >= value:
            out.append(symbol)
            n -= value
    return "".join(out)


class HeadingNumberer:
    def __init__(self, scheme: str = "section") -> None:
        if scheme not in {"section", "chapter"}:
            raise ValueError(f"heading_numbering không hợp lệ: {scheme}")
        self.scheme = scheme
        self.counters = [0, 0, 0]

    @property
    def chapter(self) -> int:
        return self.counters[0]

    def next(self, level: int) -> str:
        idx = level - 1
        self.counters[idx] += 1
        for j in range(idx + 1, len(self.counters)):
            self.counters[j] = 0
        c1, c2, c3 = self.counters
        if level == 1:
            return f"{to_roman(c1)}."
        if self.scheme == "chapter":
            return f"{c1}.{c2}." if level == 2 else f"{c1}.{c2}.{c3}."
        return f"{c2}." if level == 2 else f"{c2}.{c3}."


class CaptionCounter:
    """Đếm hình/bảng. Kiểu chapter: số chương + thứ tự trong chương ("2.3"), đánh lại mỗi chương."""

    def __init__(self, scheme: str = "section") -> None:
        self.scheme = scheme
        self.chapter = 0
        self.counts: dict[str, int] = {}

    def new_chapter(self, chapter: int) -> None:
        if self.scheme == "chapter":
            self.chapter = chapter
            self.counts = {}

    def next(self, kind: str) -> tuple[str, str, bool]:
        """Trả về (tiền tố chương, số thứ tự, có phải hình/bảng đầu tiên của chương)."""
        first = kind not in self.counts
        self.counts[kind] = self.counts.get(kind, 0) + 1
        prefix = f"{self.chapter}." if self.scheme == "chapter" and self.chapter else ""
        return prefix, str(self.counts[kind]), first
