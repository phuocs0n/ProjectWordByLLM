---
name: quality-check
description: Kiểm tra chất lượng báo cáo Word sau khi lưu (lint font, tiêu đề, chú thích, gạch đầu dòng gõ tay, chính tả, mục lục, số trang) và vòng lặp sửa lỗi. Dùng SAU MỖI lần save_report và khi người dùng muốn rà soát một file .docx có sẵn.
---

# Kiểm tra chất lượng

1. `lint_document(path, profile)` → danh sách issue `{severity, rule, message, location}`.
2. Sửa theo thứ tự: `error` → `warning` → `info` (info có thể chấp nhận nếu có lý do).
3. Với báo cáo đang soạn trong phiên: sửa block qua `update_block` rồi `save_report` lại.
   Với file .docx bên ngoài: xem skill `edit-existing-docx`.
4. Lặp tối đa 3 vòng; tóm tắt cho người dùng các lỗi còn lại và lý do.

| rule | Ý nghĩa | Cách sửa |
|---|---|---|
| no-toc | Thiếu mục lục tự động | `include_toc: true` |
| no-page-number | Footer không có số trang | Render lại bằng profile |
| heading-skip | Nhảy cấp tiêu đề | Chỉnh `level` |
| heading-unnumbered | Tiêu đề lẫn lộn có/không số | Để renderer đánh số |
| figure-without-caption / table-without-caption | Thiếu chú thích | Thêm `caption` |
| manual-bullet | Gạch đầu dòng gõ tay | Chuyển sang `list` |
| font-mismatch | Font lạ | Render lại theo profile |
| typo | Lỗi chính tả phổ biến | Sửa chữ |
| ai-label | Còn nhãn trình tạo/AI trong metadata hoặc nội dung | `remove_watermarks(path)` (skill `watermark-remover`) |
| empty-paragraphs, double-space, long-paragraph | Trình bày | Xoá / tách đoạn |

Checklist thủ công thêm: thông tin bìa đúng, không còn `figure_placeholder` khi nộp bản cuối,
số liệu khớp giữa bảng và văn bản.
