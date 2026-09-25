---
name: cover-page
description: Điền thông tin trang bìa (trường, khoa, loại báo cáo, học phần, đề tài, giảng viên, thành viên, mã lớp, nơi – năm, logo) và header/footer. Dùng khi tạo báo cáo mới hoặc người dùng yêu cầu sửa trang bìa.
---

# Trang bìa & header/footer

Trang bìa được dựng hoàn toàn từ `meta` – KHÔNG dùng heading/paragraph để vẽ bìa.

```json
{
  "university": "Đại học Quốc gia Thành phố Hồ Chí Minh",
  "school": "Trường Đại học Khoa học Tự nhiên",
  "faculty": "Khoa Công nghệ Thông tin",
  "report_type": "Báo cáo đồ án học phần",
  "subject": "Mạng máy tính",
  "topic": "Thiết kế, cấu hình mô hình mạng logic",
  "instructors": ["ThS. Nguyễn Văn A"],
  "members": [{"name": "Nguyễn Minh An", "student_id": "22127102", "email": "..."}],
  "class_code": "22CLC03",
  "city": "TP. Hồ Chí Minh",
  "year": "2024",
  "logo_path": "assets/logo.png"
}
```

## Quy tắc

- Renderer tự VIẾT HOA tên trường, loại báo cáo, học phần, đề tài – hãy nhập dạng câu bình thường.
- `faculty` hiển thị ở header trái, `class_code` ở header phải (profile `hcmus-clc`). Doanh nghiệp: dùng `faculty` = tên phòng ban, `class_code` = mã dự án/số hiệu văn bản.
- Không bịa tên giảng viên, MSSV, email. Thiếu thông tin → để chuỗi rỗng và báo người dùng bổ sung.
- `year` mặc định là năm hiện tại nếu người dùng không nói khác. `city` viết đúng "TP. Hồ Chí Minh" (lỗi hay gặp: "Hồ Chí Mình").
- Logo: chỉ đặt `logo_path` khi người dùng cung cấp file; đường dẫn tương đối tính theo thư mục file spec.
- Trang bìa có khung viền khi profile bật `page.cover_border`.

Sửa bìa trên tài liệu đang soạn: `set_meta(doc_id, {...})` (chỉ gửi trường cần đổi).
