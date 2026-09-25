---
name: table-of-contents
description: Tạo và cập nhật MỤC LỤC tự động (field TOC) kèm số trang. Dùng khi người dùng hỏi về mục lục, số trang mục lục sai/trống, hoặc sau khi lưu báo cáo.
---

# Mục lục

- Bật bằng `include_toc: true` (mặc định). Mục lục gồm LỜI MỞ ĐẦU, MỤC LỤC, mọi heading cấp 1–3 và Tài liệu tham khảo.
- Mục lục là field `TOC \o "1-3" \h \z \u` thật của Word: bấm vào mục sẽ nhảy tới tiêu đề, Word cập nhật được.
- Số trang được điền khi `save_report(..., update_toc=true)`:
  - Windows có Microsoft Word + `pywin32`: Word tự cập nhật (chính xác tuyệt đối).
  - Không có cả hai: file đã bật `updateFields`, Word hỏi cập nhật khi mở → chọn **Yes**.
- Hướng dẫn người dùng thủ công: click vào mục lục → **F9** → *Update entire table*.
- Không bao giờ gõ tay mục lục bằng paragraph + dấu chấm.
