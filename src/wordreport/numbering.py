"""Đánh số tiêu đề, bảng, hình một cách tất định.

Báo cáo gốc (PDF mẫu) bị lệch số thứ tự ("Bài 2" thiếu số, mục "4.1" nằm trong bài 2).
Renderer tự đánh số nên LLM chỉ cần cung cấp tiêu đề trơn.
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
    """Level 1 -> "I.", level 2 -> "1.", level 3 -> "1.1." (theo quy ước báo cáo mẫu)."""

    def __init__(self) -> None:
        self.counters = [0, 0, 0]

    def next(self, level: int) -> str:
        idx = level - 1
        self.counters[idx] += 1
        for j in range(idx + 1, len(self.counters)):
            self.counters[j] = 0
        if level == 1:
            return f"{to_roman(self.counters[0])}."
        if level == 2:
            return f"{self.counters[1]}."
        return f"{self.counters[1]}.{self.counters[2]}."
