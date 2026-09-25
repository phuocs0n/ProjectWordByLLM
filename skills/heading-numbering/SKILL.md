---
name: heading-numbering
description: Quy tắc tiêu đề nhiều cấp và đánh số tự động (I. / 1. / 1.1.) để mục lục đúng và không lệch số. Dùng khi thêm/sắp xếp lại chương mục hoặc khi lint báo heading-skip / heading-unnumbered.
---

# Tiêu đề & đánh số

- `heading.level`: 1 = chương (La Mã "I."), 2 = mục ("1."), 3 = tiểu mục ("1.1.").
- `text` KHÔNG chứa số thứ tự. Sai: `"2. Bảng phân công"`. Đúng: `"Bảng phân công"`.
- Không nhảy cấp (level 1 → level 3). Lint báo `heading-skip` nếu vi phạm.
- Tiêu đề ngắn gọn (≤ 12 từ), không kết thúc bằng dấu chấm/dấu hai chấm.
- Tiêu đề song song cùng cấp dùng cấu trúc ngữ pháp giống nhau ("Cấu hình DHCP", "Cấu hình DNS"...).
- Bản mẫu PDF gốc từng bị lệch số ("Bài 2" thiếu số, mục "4.1" nằm trong Bài 2) – đây chính là lỗi mà
  đánh số tự động loại bỏ.

Sắp xếp lại: `get_outline(doc_id)` để xem chỉ số block, rồi `move_block` / `update_block` / `delete_block`.
Số thứ tự và mục lục được tính lại khi `save_report`.
