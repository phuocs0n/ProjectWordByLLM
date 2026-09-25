---
name: code-cli-blocks
description: Trình bày lệnh CLI (Cisco IOS, Linux, PowerShell), đoạn mã nguồn, file cấu hình trong khung code có nền và font đơn cách. Dùng khi báo cáo có lệnh cấu hình, mã nguồn, log.
---

# Khối lệnh & mã nguồn

```json
{"type": "code", "caption": "Cấu hình định tuyến tĩnh trên Router 1", "language": "cisco",
 "code": "Router> enable\nRouter# configure terminal\nRouter(config)# ip route 172.2.5.0 255.255.255.0 172.2.4.1"}
```

- Giữ nguyên xuống dòng bằng `\n`; không thêm số dòng thủ công.
- Giữ prompt (`Router#`, `$`, `PS>`) khi minh hoạ phiên làm việc; bỏ prompt nếu là file cấu hình/mã nguồn.
- Lệnh/tên file ngắn trong câu dùng inline: `` `show ip route` ``.
- Cú pháp tổng quát dùng placeholder dạng `<mạng đích>`.
- Không dán ảnh chụp chỉ để hiển thị lệnh; nếu cần ảnh kết quả thì dùng thêm hình.
- Khối > 40 dòng → tách nhỏ theo bước hoặc đưa vào phụ lục.
