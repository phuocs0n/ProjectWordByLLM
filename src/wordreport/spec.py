"""ReportSpec - định dạng trung gian (IR) giữa LLM và bộ dựng Word.

LLM không bao giờ ghi trực tiếp XML của Word. Nó chỉ sinh (hoặc sửa từng bước qua MCP)
một ReportSpec: metadata + danh sách block phẳng. Renderer biến spec thành .docx một cách
tất định, nên mọi báo cáo đều đồng nhất định dạng theo style profile.

Danh sách block là phẳng (không đệ quy); cấu trúc chương/mục được suy ra từ `heading.level`.
Điều này giữ JSON Schema đơn giản để dùng được với structured outputs.
"""

from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, Field


class Member(BaseModel):
    name: str = Field(description="Họ và tên")
    student_id: str = Field(default="", description="MSSV / mã nhân viên")
    email: str = Field(default="")


class CoverImage(BaseModel):
    """Ảnh trang trí trôi nổi trên trang bìa (hoa văn góc, đường viền), định vị tuyệt đối theo trang."""

    path: str
    x_cm: float = Field(description="Khoảng cách từ mép trái trang")
    y_cm: float = Field(description="Khoảng cách từ mép trên trang")
    width_cm: float
    height_cm: float = Field(default=0, description="0 = giữ tỉ lệ ảnh")


class ReportMeta(BaseModel):
    """Thông tin trang bìa và header/footer."""

    university: str = Field(default="", description="Dòng 1 trang bìa, ví dụ: ĐẠI HỌC QUỐC GIA THÀNH PHỐ HỒ CHÍ MINH")
    school: str = Field(default="", description="Dòng 2 trang bìa, ví dụ: TRƯỜNG ĐẠI HỌC KHOA HỌC TỰ NHIÊN")
    faculty: str = Field(default="", description="Khoa - hiển thị ở header trái")
    report_type: str = Field(default="BÁO CÁO", description="Ví dụ: BÁO CÁO ĐỒ ÁN HỌC PHẦN")
    subject: str = Field(default="", description="Tên học phần / dự án, ví dụ: MẠNG MÁY TÍNH")
    topic: str = Field(default="", description="Đề tài (không bắt buộc)")
    subtitle: str = Field(default="", description="Dòng phụ dưới đề tài, ví dụ tên tiếng Anh")
    instructors: list[str] = Field(default_factory=list, description="Giảng viên hướng dẫn")
    members: list[Member] = Field(default_factory=list, description="Thành viên nhóm")
    class_code: str = Field(default="", description="Mã lớp - hiển thị ở header phải")
    city: str = Field(default="TP. Hồ Chí Minh")
    year: str = Field(default="")
    logo_path: str | None = Field(default=None, description="Đường dẫn logo trường (png/jpg), có thể bỏ trống")
    cover_images: list[CoverImage] = Field(default_factory=list, description="Ảnh trang trí trôi nổi trên bìa")


# ---------------------------------------------------------------------------
# Blocks
# ---------------------------------------------------------------------------
# Văn bản trong mọi block hỗ trợ inline markup tối giản:
#   **đậm**, *nghiêng*, `mã`, [chữ](https://link)


class HeadingBlock(BaseModel):
    type: Literal["heading"]
    level: Literal[1, 2, 3] = Field(description="1 = chương (I.), 2 = mục (1.), 3 = tiểu mục (1.1.)")
    text: str = Field(description="Tiêu đề KHÔNG kèm số thứ tự - renderer tự đánh số")
    numbered: bool = Field(default=True, description="False cho mục không đánh số: TÓM TẮT, MỞ ĐẦU, CHỮ VIẾT TẮT, BẢNG PHÂN CÔNG...")


class ParagraphBlock(BaseModel):
    type: Literal["paragraph"]
    text: str
    align: Literal["justify", "left", "center", "right"] = "justify"


class ListBlock(BaseModel):
    type: Literal["list"]
    items: list[str]
    style: Literal["bullet", "dash", "number", "roman", "alpha"] = "dash"
    level: Literal[0, 1, 2] = Field(default=0, description="Độ thụt lề của danh sách")


class TableBlock(BaseModel):
    type: Literal["table"]
    columns: list[str] = Field(description="Tiêu đề cột")
    rows: list[list[str]] = Field(description="Mỗi ô có thể chứa nhiều dòng phân tách bằng \\n; dòng bắt đầu '- ' thành gạch đầu dòng")
    caption: str = Field(default="", description="Chú thích bảng (renderer tự thêm 'Bảng N:')")
    col_widths: list[float] = Field(default_factory=list, description="Độ rộng cột theo tỉ lệ tương đối, ví dụ [1, 3, 5]")


class ImageBlock(BaseModel):
    type: Literal["image"]
    path: str
    caption: str = ""
    width_cm: float = Field(default=14.0)


class FigurePlaceholderBlock(BaseModel):
    """Khung giữ chỗ cho ảnh chụp màn hình mà LLM không có - người dùng chèn ảnh sau."""

    type: Literal["figure_placeholder"]
    caption: str
    description: str = Field(default="", description="Mô tả ảnh cần chụp/chèn")


class CodeBlock(BaseModel):
    type: Literal["code"]
    code: str = Field(description="Nội dung lệnh / mã nguồn, giữ nguyên xuống dòng")
    language: str = ""
    caption: str = ""


class NoteBlock(BaseModel):
    type: Literal["note"]
    text: str
    kind: Literal["note", "tip", "warning"] = "note"


class PageBreakBlock(BaseModel):
    type: Literal["page_break"]


class DividerBlock(BaseModel):
    """Dấu kết thúc phần kiểu '---o0o---'."""

    type: Literal["divider"]


Block = Annotated[
    Union[
        HeadingBlock,
        ParagraphBlock,
        ListBlock,
        TableBlock,
        ImageBlock,
        FigurePlaceholderBlock,
        CodeBlock,
        NoteBlock,
        PageBreakBlock,
        DividerBlock,
    ],
    Field(discriminator="type"),
]


class Reference(BaseModel):
    text: str = Field(description="Mô tả tài liệu: tác giả, tên, nguồn, năm")
    url: str = ""


class ReportSpec(BaseModel):
    profile: str = Field(default="hcmus-clc", description="Tên style profile")
    meta: ReportMeta = Field(default_factory=ReportMeta)
    preface: list[str] = Field(default_factory=list, description="Các đoạn của LỜI MỞ ĐẦU; để trống nếu không cần")
    preface_title: str = Field(default="", description="Đổi tiêu đề lời mở đầu, ví dụ 'LỜI NÓI ĐẦU'; trống = theo profile")
    front_matter: list[Block] = Field(default_factory=list, description="Phần đặt giữa lời mở đầu và MỤC LỤC, ví dụ TÓM TẮT")
    include_toc: bool = True
    list_of_figures: bool = Field(default=False, description="Thêm DANH MỤC HÌNH tự động sau mục lục")
    list_of_tables: bool = Field(default=False, description="Thêm DANH MỤC BẢNG tự động sau mục lục")
    body: list[Block] = Field(default_factory=list)
    references: list[Reference] = Field(default_factory=list)
    references_title: str = Field(default="", description="Đổi tiêu đề tài liệu tham khảo; trống = theo profile")
    references_numbered: bool | None = Field(default=None, description="Tiêu đề tài liệu tham khảo có đánh số chương hay không; null = theo profile")

    def outline(self) -> list[str]:
        """Dàn ý dạng chuỗi - dùng cho MCP get_outline và log."""
        from .numbering import HeadingNumberer

        numberer = HeadingNumberer()
        lines = []
        for i, block in enumerate(self.body):
            if isinstance(block, HeadingBlock):
                label = numberer.next(block.level) if block.numbered else "–"
                lines.append(f"[{i}] {'  ' * (block.level - 1)}{label} {block.text}")
            else:
                lines.append(f"[{i}] {'  ' * 3}<{block.type}>")
        return lines
