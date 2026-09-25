---
name: watermark-remover
description: Xoá nhãn công cụ/AI khỏi file Word đầu ra - metadata "Tạo bởi…", tên thư viện/ứng dụng, template, ngày tạo cũ, ảnh thu nhỏ template, đoạn "Generated with…". Dùng trước khi bàn giao file .docx, hoặc khi lint báo ai-label.
---

# Watermark remover

## Mặc định

Mọi file do `save_report` / `wordreport render` / `wordreport generate` tạo ra đã được làm sạch tự động
ở bước cuối (sau khi điền số trang mục lục). Chỉ cần gọi thủ công cho file .docx từ nguồn khác.

## Tool

- `remove_watermarks(path, output_path="", author="")` → `{changes, remaining}`.
  `output_path` trống = ghi đè; `author` đặt tác giả/người sửa cuối (mặc định giữ tên người thật,
  chỉ xoá khi đó là tên thư viện/AI).
- `lint_document(path)` báo rule `ai-label` nếu còn nhãn.
- Dòng lệnh: `watermark-remover a.docx b.docx`, `watermark-remover in.docx -o out.docx --author "Nguyễn A"`,
  `watermark-remover --check a.docx` (chỉ liệt kê).

## Những gì bị xoá

| Vị trí | Nội dung |
|---|---|
| `docProps/core.xml` | mô tả/ghi chú, từ khoá, danh mục; tác giả là tên thư viện/AI; người sửa cuối; ngày tạo/sửa (đặt lại = hiện tại); số lần sửa |
| `docProps/app.xml` | Application, AppVersion, Template, Company, Manager, TotalTime (Word tự điền lại khi lưu) |
| `docProps/thumbnail.jpeg` | ảnh thu nhỏ của template gốc |
| Thân bài, header, footer | đoạn văn CHỈ gồm nhãn trình tạo/AI: "Generated with Claude Code", "Tạo bởi AI", "AI-generated content"… |

## Không bị xoá

- Nội dung, bảng, hình, định dạng, field (mục lục, số trang).
- Câu có nội dung khác mà nhắc tới AI ("Mô hình AI được nhóm huấn luyện…") và dòng tác giả thật
  ("Written by John Smith", "Sinh viên thực hiện: …").
