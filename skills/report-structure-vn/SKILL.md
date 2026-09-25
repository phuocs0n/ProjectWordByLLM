---
name: report-structure-vn
description: BAREM CHUẨN của dự án - khung báo cáo bắt buộc (bìa, LỜI MỞ ĐẦU, MỤC LỤC, I. Giới thiệu chung với Thành viên nhóm + Bảng phân công, II. Nội dung, III. Tài liệu tham khảo) và cách chuyển MỌI báo cáo/ghi chú vào khung đó. Dùng ĐẦU TIÊN mỗi khi soạn mới hoặc định dạng lại một báo cáo.
---

# Barem báo cáo (khung bắt buộc)

Mọi báo cáo đầu ra phải giống báo cáo mẫu của dự án (`examples/output/mang-may-tinh-do-an.docx`),
bất kể tài liệu nguồn trình bày thế nào. **Không** sao chép bố cục, màu, font, kiểu đánh số của file
nguồn — chỉ lấy NỘI DUNG rồi đổ vào khung dưới đây. Profile mặc định `hcmus-clc` là barem.

| # | Phần | Lấy từ | Renderer làm gì |
|---|---|---|---|
| 1 | Trang bìa (khung viền) | `meta`: university, school, report_type, subject, topic, instructors, members, city, year, logo_path | Dựng bìa chuẩn |
| 2 | LỜI MỞ ĐẦU | `preface` (2–4 đoạn) | Trang riêng + `---o0o---` |
| 3 | MỤC LỤC | `include_toc: true` | Field TOC, Word điền số trang |
| 4 | **I. Giới thiệu chung** | | Tự tạo tiêu đề |
| 4.1 | 1. Thành viên nhóm | `meta.members` (name, student_id, email) | Tự tạo bảng STT/MSSV/Họ và tên/(Email) |
| 4.2 | 2. Bảng phân công công việc | `assignments` [{member, tasks[], completion}] | Tự tạo bảng; thiếu dữ liệu thì để trống cho nhóm điền |
| 4.3 | 3., 4., … | `introduction` (tiêu đề cấp 1 → "3.", cấp 2 → "3.1.") | Tóm tắt đề tài, lý do chọn đề tài, chữ viết tắt… |
| 5 | **II. Nội dung** | `body` (tiêu đề cấp 1 = chương → "1.", cấp 2 → "1.1.", cấp 3 → "1.1.1.") | Tự tạo tiêu đề, hạ cấp các chương |
| 6 | **III. Tài liệu tham khảo** | `references` | Chương cuối, danh sách 1. 2. 3. |

Header: `faculty` bên trái, `class_code` bên phải; footer: số trang giữa. Chữ Times New Roman 14pt,
tiêu đề xanh #2F5496, bảng có hàng tiêu đề xám, chú thích "Bảng N:" trên bảng, "Hình N:" dưới hình.

## Chuyển một báo cáo có sẵn vào barem

| Trong file nguồn | Đưa vào |
|---|---|
| Lời nói đầu / Lời mở đầu / Lời cảm ơn | `preface` |
| Tóm tắt, Mở đầu (lý do, mục tiêu, bố cục), Chữ viết tắt, Khái niệm chung | `introduction` (mỗi phần một tiêu đề cấp 1) |
| Các chương / bài (Chương 1…, I…, Bài 1…) | `body`, mỗi chương một tiêu đề cấp 1, viết thường kiểu câu, bỏ số cũ |
| Tiểu mục 1.1, 2.5.1… | tiêu đề cấp 2, cấp 3 tương ứng |
| Bảng thành viên / MSSV trên bìa | `meta.members` |
| Bảng phân công | `assignments` |
| Danh mục tài liệu | `references` |
| Mục lục, danh mục hình/bảng, số trang, bìa gốc | Bỏ — renderer tự tạo theo barem |

- Câu dẫn "Bảng 4.1", "Hình 4.3", "Chương IV" trong nguồn phải đổi sang tham chiếu chéo: gắn `label`
  cho bảng/hình/tiêu đề và viết `[[label]]` trong văn bản (renderer thay bằng "Bảng 9", "Hình 19", "4").
- Chương đầu của nguồn tên "Giới thiệu chung" → đổi thành phần `introduction` để không trùng mục I.
- Không bịa: thiếu GVHD, mã lớp, email, phân công thì để trống; `check_barem` sẽ liệt kê cho người dùng.

## Quy trình MCP

1. `create_report(profile="hcmus-clc", meta={... members ...})`
2. `set_preface`, `set_assignments`
3. `add_blocks(section="introduction", blocks=[...])` (nếu có)
4. `add_blocks(blocks=[...])` cho từng chương nội dung
5. `get_outline` (xem khung thật sẽ dựng) → `check_barem` → sửa phần thiếu
6. `save_report` → `lint_document` (skill `quality-check`)
