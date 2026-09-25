# Ghi chú thô để thử `wordreport generate`

Báo cáo đồ án học phần Mạng máy tính, đề tài: thiết kế, cấu hình mô hình mạng logic.
Trường ĐH Khoa học Tự nhiên – ĐHQG TP.HCM, Khoa CNTT, lớp 22CLC03, năm 2024.
GVHD: ThS. Nguyễn Văn A. Nhóm: Nguyễn Minh An (22127102) làm bài 1, Lê Thu Bình (22127117) làm bài 2 + tạo khuôn báo cáo. Cả hai hoàn thành 100%.

Bài 1: 3 router nối cáp chéo, switch nối cáp thẳng. X = 02.
- R1 f0/0 172.2.1.254/24, f0/1 172.2.2.1/24; R2 f0/0 172.2.3.254/24 f0/1 172.2.2.2/24 f1/0 172.2.4.2/24 (phải gắn module NM-2FE2W); R3 f0/0 172.2.5.254/24 f0/1 172.2.4.1/24
- DHCP server cấp 172.2.1.1-100, gw 172.2.1.254, max users 100
- static route trên R1: ip route 172.2.3.0 255.255.255.0 172.2.2.2 ; 172.2.4.0 qua 172.2.2.2 ; 172.2.5.0 qua 172.2.4.1
- ping PC3 -> PC5 ok (có ảnh chụp simulation)

Bài 2: web server www.congtyhanhoa.vn (index.html + thanhvien.html), DNS bản ghi A, RIP trên 3 router,
DHCP relay: ip helper-address trên cổng LAN của R1, R3. Truy cập web từ PC3 (192.2.3.0/24) và PC5 (192.2.5.0/24) thành công.

Tham khảo: https://www.computernetworkingnotes.com/ccna-study-guide/how-to-configure-dhcp-relay-agent-on-cisco-routers.html
