---
name: quality-check
description: Chấm barem và kiểm tra chất lượng báo cáo Word (đủ phần bắt buộc của barem, font, tiêu đề, chú thích, gạch đầu dòng gõ tay, chính tả, mục lục, số trang) và vòng lặp sửa lỗi. Dùng SAU MỖI lần save_report và khi người dùng muốn rà soát một file .docx có sẵn.
---

# Kiểm tra chất lượng

1. Trước khi lưu: `check_barem(doc_id)` → phần còn thiếu so với barem (thành viên, phân công, lời mở đầu,
   GVHD, chương nội dung, tài liệu tham khảo, tham chiếu `[[label]]` hỏng). `save_report` cũng trả về mục `barem`.
2. Sau khi lưu: `lint_document(path, profile)` → danh sách issue `{severity, rule, message, location}`.
2. Sửa theo thứ tự: `error` → `warning` → `info` (info có thể chấp nhận nếu có lý do).
3. Với báo cáo đang soạn trong phiên: sửa block qua `update_block` rồi `save_report` lại.
   Với file .docx bên ngoài: xem skill `edit-existing-docx`.
4. Lặp tối đa 3 vòng; tóm tắt cho người dùng các lỗi còn lại và lý do.

| rule | Ý nghĩa | Cách sửa |
|---|---|---|
| barem (spec) | Thiếu dữ liệu cho một phần của barem | Bổ sung `meta`, `preface`, `assignments`, `body`, `references` |
| barem-missing (docx) | File thiếu tiêu đề bắt buộc (LỜI MỞ ĐẦU, MỤC LỤC, Giới thiệu chung, Thành viên nhóm, Bảng phân công, Nội dung, Tài liệu tham khảo) | Dựng lại theo barem (skill `report-structure-vn`) |
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
