---
name: word-mcp
description: Dùng MCP Microsoft Word (Office-Word-MCP-Server, server "word") để sửa chi tiết một file .docx đã lưu - thay chữ, định dạng đoạn/ô bảng, độ rộng cột, gộp ô, chú thích cuối trang, bình luận, bảo vệ tài liệu - mà không làm mất định dạng của profile. Dùng sau save_report hoặc khi người dùng muốn chỉnh một file Word có sẵn.
---

# MCP Microsoft Word

Server `word` (Office-Word-MCP-Server) thao tác trực tiếp trên file .docx theo đường dẫn. Chạy bằng
`uvx --from office-word-mcp-server word_mcp_server` (đã khai báo trong `.mcp.json`).

## Phân vai

| Việc | Dùng |
|---|---|
| Dựng báo cáo mới, đổi cấu trúc chương mục, định dạng theo profile | `word-report` (`create_report`, `add_blocks`, `save_report`) |
| Sửa nhỏ trên file đã lưu, giữ nguyên bố cục | `word` (bảng dưới) |

Không dùng `create_document`/`add_heading`/`add_paragraph` của server `word` để soạn lại cả báo cáo: style
mặc định của nó không theo profile.

## Tool hay dùng

- Kiểm tra: `get_document_info(filename)`, `get_document_outline(filename)`, `find_text_in_document(filename, text_to_find)`.
- Chữ: `search_and_replace(filename, find_text, replace_text)`, `format_text(filename, paragraph_index, start_pos, end_pos, bold, italic, ...)`.
- Bảng (chỉ số từ 0, gồm cả bảng trang bìa): `merge_table_cells_vertical(...)`, `set_table_cell_alignment(...)`,
  `format_table_cell_text(...)`.
- KHÔNG dùng `set_table_column_widths` (bản 1.1.11 chèn `w:tcW` sai thứ tự schema, Word có thể báo file lỗi);
  đổi độ rộng cột bằng `col_widths` của block `table` rồi `save_report` lại.
- Chú thích cuối trang: `add_footnote_after_text(filename, search_text, footnote_text)`.
- Bàn giao: `protect_document(filename, password)` khi người dùng yêu cầu.

## Quy tắc

1. Luôn làm trên file kết quả của `save_report`; muốn giữ bản gốc thì `copy_document` trước.
2. Sau khi sửa, chạy `lint_document` và `remove_watermarks` của `word-report`.
3. Không gọi `convert_to_pdf` trừ khi người dùng yêu cầu PDF; mặc định chỉ bàn giao .docx.
4. Mục lục/danh mục hình/bảng do Microsoft Word cập nhật khi mở file (chọn **Yes**) hoặc qua Word COM trên Windows.
