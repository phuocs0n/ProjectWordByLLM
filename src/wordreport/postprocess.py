"""Hậu xử lý bằng Microsoft Word: điền số trang cho MỤC LỤC / DANH MỤC HÌNH / DANH MỤC BẢNG và xuất PDF.

python-docx không dàn trang nên không biết tiêu đề nằm ở trang nào. Chỉ Microsoft Word dàn trang
chính xác, vì vậy công cụ dùng Word:
  - Windows có Microsoft Word + pywin32: điều khiển Word qua COM để cập nhật mọi field rồi lưu.
  - Nền tảng khác: file đã bật cờ `updateFields`; khi mở bằng Word, chọn **Yes** để Word tự cập nhật.
Công cụ không dùng LibreOffice.
"""

from __future__ import annotations

import sys
from contextlib import contextmanager
from pathlib import Path

from .watermark_remover import remove_watermarks

WD_FORMAT_PDF = 17


def word_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # type: ignore[import-not-found]  # noqa: F401
    except ImportError:
        return False
    return True


@contextmanager
def _word():
    import win32com.client  # type: ignore[import-not-found]

    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        yield app
    finally:
        app.Quit()


def update_toc(docx_path: str | Path) -> str:
    """Cho Microsoft Word cập nhật mục lục, danh mục hình/bảng và mọi field. Trả về phương pháp đã dùng."""
    if not word_available():
        return "chờ Word: mở file bằng Microsoft Word và chọn Yes để cập nhật mục lục"
    with _word() as app:
        doc = app.Documents.Open(str(Path(docx_path).resolve()))
        try:
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(i).Update()
            for i in range(1, doc.TablesOfFigures.Count + 1):
                doc.TablesOfFigures(i).Update()
            doc.Fields.Update()
            doc.Save()
        finally:
            doc.Close()
    return "microsoft-word"


def export_pdf(docx_path: str | Path, out_dir: str | Path | None = None) -> Path:
    """Xuất PDF bằng Microsoft Word (chỉ khi được yêu cầu rõ ràng)."""
    if not word_available():
        raise RuntimeError("Xuất PDF cần Microsoft Word trên Windows (pip install pywin32).")
    docx_path = Path(docx_path).resolve()
    pdf = Path(out_dir or docx_path.parent).resolve() / (docx_path.stem + ".pdf")
    with _word() as app:
        doc = app.Documents.Open(str(docx_path))
        try:
            doc.SaveAs2(str(pdf), FileFormat=WD_FORMAT_PDF)
        finally:
            doc.Close(False)
    return pdf


def finalize(docx_path: str | Path, toc: bool = True, author: str | None = None) -> dict[str, object]:
    """Bước cuối cho mọi file xuất ra: cập nhật mục lục bằng Word (nếu có) rồi xoá nhãn trình tạo/AI.

    Làm sạch chạy SAU cùng vì Word có thể ghi lại metadata khi lưu.
    """
    result: dict[str, object] = {}
    if toc:
        try:
            result["toc"] = update_toc(docx_path)
        except Exception as err:  # bước phụ, không làm hỏng việc xuất file
            result["toc"] = f"lỗi cập nhật mục lục: {err}"
    result["watermarks_removed"] = remove_watermarks(docx_path, author=author)
    return result
