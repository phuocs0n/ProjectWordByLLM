---
name: figures-captions
description: Chèn hình ảnh/ảnh chụp màn hình/biểu đồ kèm chú thích "Hình N" tự đánh số, hoặc khung giữ chỗ khi chưa có ảnh. Dùng khi báo cáo cần sơ đồ, ảnh chụp cấu hình, biểu đồ.
---

# Hình ảnh & chú thích

- Có file ảnh: `{"type": "image", "path": "img/topo.png", "caption": "Sơ đồ mạng bài 1", "width_cm": 14}`.
  Ảnh căn giữa, không vượt quá bề rộng vùng in; chú thích "Hình N:" nằm DƯỚI ảnh.
- Chưa có ảnh (LLM không tự chụp được màn hình): dùng
  `{"type": "figure_placeholder", "caption": "...", "description": "Ảnh chụp tab Services → DHCP sau khi lưu"}`.
  Mô tả phải đủ cụ thể để người dùng biết cần chụp gì. Khi có ảnh, `update_block` sang `image`.
- Đường dẫn ảnh không tồn tại → renderer tự đổi thành khung giữ chỗ (không làm hỏng file).
- Mọi hình đều phải có chú thích ngắn (≤ 15 từ), không kết thúc bằng dấu chấm.
- Trước mỗi hình nên có một câu dẫn ("Kết quả được thể hiện ở hình dưới:").
- Ảnh chụp chữ (terminal, cấu hình CLI) → ưu tiên `code` block (skill `code-cli-blocks`) để người đọc sao chép được.
- Biểu đồ từ Python: lưu PNG dpi ≥ 150 bằng matplotlib, rồi chèn bằng `image` (skill `data-analysis-report`).
