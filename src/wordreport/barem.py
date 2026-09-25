"""Chấm barem: báo cáo có đủ các phần bắt buộc của khung mẫu hay chưa.

Barem (profile `hcmus-clc`, trích từ báo cáo mẫu CLC HCMUS):
  Bìa (trường, loại báo cáo, học phần/đề tài, GVHD, thành viên, nơi – năm)
  LỜI MỞ ĐẦU -> MỤC LỤC
  I. Giới thiệu chung: 1. Thành viên nhóm, 2. Bảng phân công công việc
  II. Nội dung: các chương của báo cáo
  III. Tài liệu tham khảo
Renderer tự dựng khung; module này chỉ ra dữ liệu còn thiếu để khung không bị trống.
"""

from __future__ import annotations

import re

from .inspector import LintIssue
from .spec import FigurePlaceholderBlock, HeadingBlock, ImageBlock, ListBlock, NoteBlock, ParagraphBlock, ReportSpec, TableBlock
from .style_profile import load_profile

_REF = re.compile(r"\[\[([\w\-.:]+)\]\]")
_BUILTIN_LABELS = {"barem-members", "barem-assignments"}


def _texts(blocks) -> list[str]:
    out: list[str] = []
    for b in blocks:
        if isinstance(b, (ParagraphBlock, NoteBlock)):
            out.append(b.text)
        elif isinstance(b, ListBlock):
            out.extend(b.items)
        elif isinstance(b, TableBlock):
            out.append(b.caption)
            out.extend(cell for row in b.rows for cell in row)
        elif isinstance(b, (ImageBlock, FigurePlaceholderBlock)):
            out.append(b.caption)
    return out


def check_spec(spec: ReportSpec, profile: str | dict | None = None) -> list[LintIssue]:
    prof = load_profile(profile or spec.profile) if not isinstance(profile, dict) else profile
    issues: list[LintIssue] = []
    add = lambda sev, msg: issues.append(LintIssue(sev, "barem", msg))  # noqa: E731
    m = spec.meta
    barem = (prof.get("barem") or {}).get("enabled")

    # --- tham chiếu chéo: mọi [[nhãn]] phải trỏ tới block có label
    blocks = list(spec.front_matter) + list(spec.introduction) + list(spec.body)
    labels = {getattr(b, "label", "") for b in blocks} | (_BUILTIN_LABELS if barem else set())
    for text in _texts(blocks) + list(spec.preface):
        for name in _REF.findall(text):
            if name not in labels:
                add("error", f"Tham chiếu [[{name}]] không trỏ tới tiêu đề/bảng/hình nào (thiếu label).")

    if not barem:
        return issues

    # --- trang bìa
    if not (m.university or m.school):
        add("warning", "Bìa: thiếu tên đại học / trường.")
    if not (m.subject or m.topic):
        add("error", "Bìa: thiếu tên học phần hoặc đề tài.")
    if not m.instructors:
        add("warning", "Bìa: chưa có giảng viên hướng dẫn.")
    if not (m.city or m.year):
        add("info", "Bìa: chưa có nơi – năm thực hiện.")
    if not m.faculty:
        add("info", "Header trái (Khoa) đang trống.")
    if not m.class_code:
        add("info", "Header phải (mã lớp) đang trống.")

    # --- phần đầu
    if not spec.preface:
        add("warning", "Thiếu LỜI MỞ ĐẦU (spec.preface).")
    if not spec.include_toc:
        add("warning", "Barem yêu cầu MỤC LỤC (include_toc: true).")

    # --- I. Giới thiệu chung
    if not m.members:
        add("error", "I.1 Thành viên nhóm: chưa có thành viên (meta.members).")
    elif any(not mem.student_id for mem in m.members):
        add("warning", "I.1 Thành viên nhóm: có thành viên thiếu MSSV.")
    if not spec.assignments:
        add("warning", "I.2 Bảng phân công công việc: chưa có dữ liệu (spec.assignments) - bảng đang để trống cho nhóm điền.")
    else:
        names = {mem.name for mem in m.members}
        for a in spec.assignments:
            if names and a.member not in names:
                add("info", f"I.2 Phân công: '{a.member}' không có trong danh sách thành viên.")
            if not a.tasks or not a.completion:
                add("warning", f"I.2 Phân công: '{a.member}' thiếu nhiệm vụ hoặc mức độ hoàn thành.")

    # --- II. Nội dung
    chapters = [b for b in spec.body if isinstance(b, HeadingBlock) and b.level == 1]
    if not chapters:
        add("error", "II. Nội dung: cần ít nhất một chương (tiêu đề cấp 1 trong spec.body).")
    placeholders = [b for b in blocks if isinstance(b, FigurePlaceholderBlock)]
    if placeholders:
        add("info", f"Còn {len(placeholders)} khung giữ chỗ hình cần chèn ảnh thật.")

    # --- III. Tài liệu tham khảo
    if not spec.references:
        add("warning", "III. Tài liệu tham khảo: chưa có tài liệu nào.")
    return issues


def barem_outline_labels(profile: dict) -> list[str]:
    """Các tiêu đề bắt buộc phải xuất hiện trong file .docx (dùng cho lint file có sẵn)."""
    cfg = profile.get("barem") or {}
    if not cfg.get("enabled"):
        return []
    labels = profile["labels"]
    return [labels["preface"], labels["toc"], cfg["intro_title"], cfg["members_title"],
            cfg["assignments_title"], cfg["content_title"], labels["references"]]
