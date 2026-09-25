---
name: lists
description: Chọn kiểu danh sách (gạch đầu dòng, chấm tròn, số, La Mã, chữ cái) với đánh số tự động khởi động lại. Dùng khi liệt kê yêu cầu, các bước thực hiện, ý chính.
---

# Danh sách

| Nội dung | `style` |
|---|---|
| Ý song song không thứ tự (mặc định báo cáo VN) | `dash` ( - ) |
| Ý nhấn mạnh / tóm tắt | `bullet` ( • ) |
| Yêu cầu đề bài, các bước tuần tự | `number` (1. 2. 3.) |
| Các bước con trong một thao tác | `roman` (i. ii. iii.) |
| Phương án lựa chọn | `alpha` (a. b. c.) |

- Mỗi `list` block là một danh sách mới, đánh số bắt đầu lại từ 1.
- Danh sách con: block `list` tiếp theo với `level: 1`.
- KHÔNG gõ "- " hoặc "1." ở đầu paragraph – lint báo `manual-bullet`.
- Mỗi mục là một câu/cụm hoàn chỉnh; các mục cùng dạng ngữ pháp; kết thúc bằng dấu chấm nếu là câu.
