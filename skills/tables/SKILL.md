---
name: tables
description: Tạo bảng chuyên nghiệp (tiêu đề cột tô nền, chú thích "Bảng N", độ rộng cột, gạch đầu dòng trong ô) cho danh sách thành viên, phân công, bảng IP, số liệu, so sánh. Dùng khi dữ liệu có cấu trúc hàng–cột.
---

# Bảng

```json
{"type": "table", "caption": "Phân công công việc",
 "columns": ["STT", "Người phụ trách", "Nhiệm vụ", "Mức độ hoàn thành"],
 "col_widths": [1, 3, 4, 2.4],
 "rows": [["1", "Nguyễn Minh An", "- Làm bài 1.\n- Viết báo cáo bài 1.", "100%"]]}
```

## Quy tắc

- Luôn có `caption` → renderer thêm "Bảng N:" (field SEQ) phía TRÊN bảng.
- Hàng tiêu đề tự tô nền, in đậm, lặp lại khi bảng sang trang; hàng không bị tách đôi.
- `col_widths` là tỉ lệ tương đối; bỏ trống thì tự tính theo độ dài nội dung.
- Ô nhiều dòng: dùng `\n`; dòng bắt đầu `"- "` thành gạch đầu dòng trong ô.
- Cột ngắn (STT, mã, %, số) tự căn giữa; cột văn bản dài căn trái.
- Link/email trong ô: `[text](mailto:...)`.
- Khi nào KHÔNG dùng bảng: nội dung chỉ 1 cột → dùng `list`; văn bản dài → paragraph.
- Bảng quá 7 cột → cân nhắc tách bảng hoặc profile khổ ngang.

## Mẫu hay dùng

- Thành viên nhóm: `STT | MSSV | Họ và tên | Email`.
- Phân công: `STT | Người phụ trách | Nhiệm vụ được giao | Mức độ hoàn thành`.
- Địa chỉ IP: `Thiết bị | Cổng | Địa chỉ IP | Subnet | Gateway`.
- Kết quả thực nghiệm: `Thử nghiệm | Kỳ vọng | Kết quả | Đạt/Không đạt`.
