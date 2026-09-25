# Phân tích báo cáo mẫu → yêu cầu thiết kế công cụ

Nguồn: *Báo cáo đồ án học phần Mạng máy tính – CLC HCMUS* (PDF 27 trang, bản tải từ Studocu).
Số đo lấy tự động bằng PyMuPDF (vị trí, font, cỡ chữ, màu của từng span và hình vẽ).

## 1. Cấu trúc tài liệu

| Trang | Phần | Nhận xét |
|---|---|---|
| 1 | Trang bìa của Studocu | Không thuộc báo cáo |
| 2 | Bìa: ĐHQG, trường, logo, tiêu đề, GVHD, nhóm, nơi – năm | Có khung viền trang |
| 3 | Bìa phụ: lặp lại tiêu đề + "Đề tài" | Trùng lặp với trang 2 |
| 4 | LỜI MỞ ĐẦU + `---o0o---` | Tiêu đề 20pt đậm căn giữa |
| 5 | MỤC LỤC | Có dấu chấm dẫn và số trang |
| 6 | I. Giới thiệu chung: bảng thành viên, bảng phân công | Tiêu đề xanh |
| 7–25 | II. Nội dung: Bài 1, Bài 2 (đề bài → bài làm, ~30 ảnh chụp màn hình) | |
| 26 | III. Tài liệu tham khảo | 4 URL trần |
| 27 | Trang trống | Thừa |

→ Công cụ mô hình hoá thành **ReportSpec**: `meta` (bìa) + `preface` + `include_toc` + `body` (block phẳng)
+ `references`; renderer dựng 3 section: bìa / phần đầu / nội dung.

## 2. Quy chuẩn định dạng đo được (→ `profiles/hcmus-clc.yaml`)

| Thuộc tính | Giá trị đo | Trong profile |
|---|---|---|
| Khổ giấy | 612 × 792 pt (Letter) | `page.size: letter` |
| Lề trái/phải | văn bản bắt đầu x = 72 pt, kết thúc ≈ 543 pt | 2,54 cm |
| Font thân bài | Times New Roman 13,9 pt, màu đen | TNR 14 pt |
| Khoảng cách dòng | 16,1 pt cho cỡ 14 (≈ single) | `line_spacing: 1.15` |
| Tiêu đề chương/mục | TNR đậm 13,9 pt, màu `#2F5496` | `headings.color` |
| Tiêu đề phụ | màu `#1F3763` | `level3.color` |
| LỜI MỞ ĐẦU / MỤC LỤC | TNR đậm 19,9 pt, căn giữa | `headings.front` |
| Header | "Khoa Công nghệ Thông Tin" trái, "22CLC03" phải, 12 pt, đường kẻ dưới | `header` |
| Footer | số trang 12 pt căn giữa, đường kẻ trên | `footer` |
| Bảng | viền đen, hàng tiêu đề tô `#D0CECE`, chữ đậm căn giữa | `table.header_fill` |
| Liên kết | `#0563C1` gạch chân | `hyperlink_color` |
| Dấu kết phần | `---o0o---` 12 pt | `divider_text` |

## 3. Lỗi trình bày phát hiện được (→ tính năng tự động hoá / luật lint)

| Lỗi trong bản mẫu | Cách công cụ xử lý |
|---|---|
| Đánh số tiêu đề không nhất quán: "Bài 2" thiếu "2.", mục con của Bài 2 đánh "4.1/4.2" | Tiêu đề không chứa số; renderer tự đánh `I. / 1. / 1.1.`; lint `heading-unnumbered`, `heading-skip` |
| Thụt lề tiêu đề cùng cấp khác nhau ("1. Thành viên nhóm" x=90, "2. Bảng phân công" x=72) | Thụt lề nằm trong style Heading, không định dạng trực tiếp |
| ~30 ảnh chụp không có chú thích, không thể lập danh mục hình | `image`/`figure_placeholder` luôn kèm "Hình N" (field SEQ); lint `figure-without-caption` |
| Lệnh Cisco (`ip route`, `ip helper-address`) chỉ có dạng ảnh chụp | Block `code` font đơn cách, sao chép được |
| Danh sách gõ tay ("- ", "i.", "ii.") và câu bị giãn chữ khi căn đều | Block `list` dùng numbering thật của Word; lint `manual-bullet` |
| Lỗi chính tả: "Hồ Chí Mình" (2 lần), "chúng tỏ", "hoạt đồng", "đảm bào", "thự hiện", "tên mien" | Từ điển lỗi phổ biến trong lint `typo`, skill `academic-writing-vn` |
| Văn nói: "vô Physical", "bấm lệnh" | Skill `academic-writing-vn` |
| Tài liệu tham khảo là URL trần, không mô tả | Block `references` có mô tả + hyperlink, skill `references` |
| Bìa lặp 2 lần, trang trống cuối | Bìa dựng một lần từ `meta`; không chèn trang trống |
| Mục lục phải cập nhật tay | Field TOC thật + tự điền số trang (Word COM hoặc LibreOffice) |

## 4. Kết quả tái dựng

`examples/mang-may-tinh-do-an.json` là bản tái dựng nội dung báo cáo mẫu (tên người đã thay bằng tên giả).
`wordreport render examples/mang-may-tinh-do-an.json --pdf` tạo file 9 trang: bìa có khung, lời mở đầu,
mục lục có số trang, 3 bảng + 4 hình có chú thích, 2 khối lệnh, 2 hộp ghi chú, tài liệu tham khảo có link;
lint 0 lỗi/0 cảnh báo; qua kiểm tra XSD của OOXML.
