---
name: references
description: Định dạng danh mục tài liệu tham khảo (sách, bài báo, website, video) có đánh số và hyperlink. Dùng khi báo cáo trích dẫn nguồn.
---

# Tài liệu tham khảo

`add_reference(doc_id, text, url)` hoặc trường `references` trong spec. Renderer tự thêm chương
"Tài liệu tham khảo" có số, danh sách đánh số [1], [2]... và URL là hyperlink.

Mẫu `text` (kiểu IEEE rút gọn):
- Sách: `J. F. Kurose, K. W. Ross, "Computer Networking: A Top-Down Approach", 8th ed., Pearson, 2021.`
- Website: `Computer Networking Notes, "How to configure DHCP Relay Agent on Cisco routers", truy cập 10/2024.`
- Video: `Tên kênh, "Tiêu đề video", YouTube, năm.` (không để link trần không có mô tả như bản mẫu gốc).

Quy tắc: chỉ liệt kê nguồn người dùng cung cấp hoặc thực sự đã dùng; không bịa DOI/URL.
Trong thân bài trích dẫn bằng số: "…theo hướng dẫn của Cisco [2]."
