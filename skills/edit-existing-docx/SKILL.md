---
name: edit-existing-docx
description: Chỉnh sửa hoặc chuẩn hoá định dạng một file .docx có sẵn (do người khác soạn) - đọc nội dung, dựng lại theo profile, hoặc sửa nhỏ trực tiếp. Dùng khi người dùng muốn "định dạng lại", "sửa lỗi trình bày" file Word đã có.
---

# Chỉnh sửa .docx có sẵn

## Cách 1 – Dựng lại theo profile (khuyến nghị cho "định dạng lại toàn bộ")

1. `read_document(path)` → Markdown (giữ tiêu đề, bảng).
2. `lint_document(path)` để liệt kê lỗi hiện có (dùng làm checklist).
3. Chuyển Markdown thành ReportSpec: `#` → heading (bỏ số thứ tự cũ), bảng → `table`, gạch đầu dòng → `list`,
   ảnh → `figure_placeholder` (hoặc `image` nếu người dùng cung cấp thư mục ảnh).
4. `create_report` + `add_blocks` → `save_report` ra file MỚI (không ghi đè bản gốc).
5. So sánh lint trước/sau và báo cáo cho người dùng.

## Cách 2 – Sửa nhỏ giữ nguyên bố cục

Khi chỉ cần thay chữ/sửa chính tả mà phải giữ nguyên định dạng gốc: dùng skill docx của Claude
(giải nén → sửa `word/document.xml` → nén lại) hoặc Word MCP bên thứ ba (xem README, mục MCP).
Luôn ghi ra file mới và nói rõ đã sửa gì.
