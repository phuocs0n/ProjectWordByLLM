---
name: report-structure-vn
description: Dàn ý chuẩn cho báo cáo đồ án/học phần/doanh nghiệp tiếng Việt (bìa, lời mở đầu, mục lục, chương, tài liệu tham khảo). Dùng ĐẦU TIÊN khi bắt đầu soạn một báo cáo mới từ ghi chú hoặc đề bài.
---

# Cấu trúc báo cáo tiếng Việt

## Khung chuẩn (ánh xạ sang ReportSpec)

| Phần báo cáo | Trường trong spec | Ghi chú |
|---|---|---|
| Trang bìa | `meta` | Xem skill `cover-page` |
| Lời mở đầu | `preface` (list đoạn) | 2–4 đoạn: bối cảnh → vấn đề → phạm vi báo cáo |
| Mục lục | `include_toc: true` | Tự sinh, xem `table-of-contents` |
| I. Giới thiệu chung | `heading` level 1 | Thành viên, phân công (bảng) |
| II. Nội dung | `heading` level 1 | Mỗi bài/nhiệm vụ là level 2; "Đề bài", "Bài làm", "Kết quả" là level 3 |
| III. Kết luận (khuyến nghị) | `heading` level 1 | Kết quả đạt được, hạn chế, hướng phát triển |
| Tài liệu tham khảo | `references` | Renderer tự đặt thành chương cuối có số |

## Quy tắc

1. Tối đa 3 cấp tiêu đề. Không gõ số thứ tự vào tiêu đề – renderer tự đánh "I.", "1.", "1.1.".
2. Mỗi chương (level 1) tự bắt đầu trang mới; không chèn `page_break` thủ công trước chương.
3. Mỗi tiêu đề phải có ít nhất một khối nội dung ngay sau (tránh tiêu đề "treo").
4. Bài thực hành kỹ thuật: *Đề bài* (tóm tắt yêu cầu bằng `list` kiểu `number`) → *Bài làm* (từng bước, lệnh trong `code`, ảnh chụp bằng `image`/`figure_placeholder`) → *Kết quả kiểm thử*.
5. Kết thúc phần lời mở đầu/mục lục/chương cuối có thể thêm `divider` ("---o0o---") như bản mẫu.

## Quy trình với MCP `word-report`

1. `create_report(profile, meta)` → nhận `doc_id`.
2. `set_preface(doc_id, paragraphs)`.
3. `add_blocks(doc_id, blocks)` theo từng chương (gửi theo lô 10–30 block để tiết kiệm lượt gọi).
4. `get_outline(doc_id)` để rà soát dàn ý.
5. `save_report(doc_id, output_path)` → sau đó `lint_document` (skill `quality-check`).
