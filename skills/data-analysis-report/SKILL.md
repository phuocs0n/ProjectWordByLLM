---
name: data-analysis-report
description: Biến kết quả phân tích dữ liệu bằng Python (pandas, matplotlib) thành báo cáo Word - bảng thống kê, biểu đồ, nhận định, khuyến nghị. Dùng khi người dùng đưa notebook/CSV/Excel hoặc số liệu cần báo cáo.
---

# Báo cáo phân tích dữ liệu

## Dàn ý

1. Tóm tắt điều hành (3–5 gạch đầu dòng: phát hiện chính + con số).
2. Dữ liệu & phương pháp (nguồn, giai đoạn, số bản ghi, làm sạch, công cụ).
3. Kết quả – mỗi câu hỏi phân tích là một mục level 2: biểu đồ → bảng số liệu → nhận định.
4. Khuyến nghị (danh sách `number`, mỗi ý gắn với một phát hiện).
5. Phụ lục: mã nguồn chính (`code`), định nghĩa chỉ số.

## Từ pandas sang block

```python
df_summary = df.groupby("khu_vuc")["doanh_thu"].agg(["sum", "mean"]).round(1).reset_index()
block = {"type": "table", "caption": "Doanh thu theo khu vực",
         "columns": ["Khu vực", "Tổng (tỷ đồng)", "Trung bình"],
         "rows": df_summary.astype(str).values.tolist()}
```

Biểu đồ: `fig.savefig("out/doanh-thu.png", dpi=200, bbox_inches="tight")` →
`{"type": "image", "path": "out/doanh-thu.png", "caption": "Doanh thu theo tháng", "width_cm": 15}`.

## Quy tắc

- Số trong bảng làm tròn nhất quán, có đơn vị ở tiêu đề cột.
- Mỗi biểu đồ đi kèm ≥1 câu nhận định định lượng ("tăng 18% so với Q1").
- Nhận định chỉ dựa trên số liệu thật đã tính; không suy diễn nhân quả khi chỉ có tương quan.
- Bảng > 15 dòng → đưa top-N vào thân bài, bảng đầy đủ vào phụ lục.
