"""Hậu xử lý: điền số trang cho MỤC LỤC và xuất PDF.

python-docx không có bộ dàn trang nên không biết tiêu đề nằm ở trang nào. Chiến lược:
  1. Windows + Microsoft Word (pywin32): dùng COM để Word tự cập nhật mọi field - chính xác nhất.
  2. Có LibreOffice: xuất PDF tạm, dò trang chứa từng tiêu đề, ghi số trang vào kết quả
     đệm của field TOC (Liberation Serif cùng metric với Times New Roman nên khớp với Word).
  3. Không có gì: giữ nguyên - file đã bật cờ updateFields, Word sẽ đề nghị cập nhật khi mở.
"""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document

from .watermark_remover import remove_watermarks

TOC_STYLES = {"toc 1", "toc 2", "toc 3"}


def find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    if sys.platform == "win32":
        for candidate in (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ):
            if Path(candidate).exists():
                return candidate
    return None


def export_pdf(docx_path: str | Path, out_dir: str | Path | None = None) -> Path:
    """Xuất PDF bằng LibreOffice (headless)."""
    docx_path = Path(docx_path).resolve()
    out_dir = Path(out_dir or docx_path.parent).resolve()
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("Không tìm thấy LibreOffice (soffice) để xuất PDF.")
    with tempfile.TemporaryDirectory() as profile:
        subprocess.run(
            [soffice, f"-env:UserInstallation={Path(profile).as_uri()}", "--headless", "--norestore",
             "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)],
            check=True, capture_output=True, timeout=300,
        )
    pdf = out_dir / (docx_path.stem + ".pdf")
    if not pdf.exists():
        raise RuntimeError(f"LibreOffice không tạo được {pdf}")
    return pdf


def pdf_pages_text(pdf_path: Path) -> list[str]:
    """Văn bản theo từng trang - ưu tiên pdftotext (poppler), sau đó PyMuPDF / pypdf."""
    if shutil.which("pdftotext"):
        out = subprocess.run(["pdftotext", "-layout", str(pdf_path), "-"], check=True,
                             capture_output=True, text=True, encoding="utf-8").stdout
        pages = out.split("\f")
        return pages[:-1] if pages and not pages[-1].strip() else pages
    try:
        import pymupdf

        with pymupdf.open(str(pdf_path)) as doc:
            return [page.get_text() for page in doc]
    except ImportError:
        pass
    from pypdf import PdfReader

    return [page.extract_text() or "" for page in PdfReader(str(pdf_path)).pages]


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip().casefold()


def _toc_paragraphs(document) -> list:
    return [p for p in document.paragraphs if p.style is not None and p.style.name in TOC_STYLES]


def _entry_text(paragraph) -> str:
    return paragraph.text.split("\t")[0]


def locate_headings(entries: list[str], pages: list[str]) -> list[int | None]:
    """Trả về số trang (1-based) cho từng mục, dò tuần tự theo thứ tự tài liệu."""
    is_toc_page = [len(re.findall(r"\.{5,}", page)) >= 3 for page in pages]
    page_lines = [[_norm(line) for line in page.splitlines() if line.strip()] for page in pages]
    results: list[int | None] = []
    cursor = 0
    for raw in entries:
        entry = _norm(raw)
        found = None
        for idx in range(cursor, len(pages)):
            for line in page_lines[idx]:
                if is_toc_page[idx]:
                    # Trên trang mục lục chỉ nhận dòng trùng khớp tuyệt đối (chính tiêu đề "MỤC LỤC"),
                    # không nhận các dòng mục lục có dấu chấm dẫn.
                    matched = line == entry
                else:
                    matched = line == entry or line.startswith(entry) or (len(line) >= 12 and entry.startswith(line))
                if matched:
                    found = idx
                    break
            if found is not None:
                break
        if found is not None:
            cursor = found
            results.append(found + 1)
        else:
            results.append(None)
    return results


def _update_with_word(docx_path: Path) -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        return False
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(docx_path.resolve()))
        for i in range(1, doc.TablesOfContents.Count + 1):
            doc.TablesOfContents(i).Update()
        doc.Fields.Update()
        doc.Save()
        doc.Close()
    finally:
        word.Quit()
    return True


def update_toc(docx_path: str | Path) -> str:
    """Điền số trang cho mục lục. Trả về tên phương pháp đã dùng."""
    docx_path = Path(docx_path)
    if _update_with_word(docx_path):
        return "word-com"
    if not find_soffice():
        return "skipped (cần Microsoft Word hoặc LibreOffice; Word sẽ tự cập nhật khi mở file)"

    document = Document(str(docx_path))
    toc = _toc_paragraphs(document)
    if not toc:
        return "no-toc"
    with tempfile.TemporaryDirectory() as tmp:
        pdf = export_pdf(docx_path, tmp)
        pages = pdf_pages_text(pdf)
    numbers = locate_headings([_entry_text(p) for p in toc], pages)
    for paragraph, number in zip(toc, numbers):
        if number is None:
            continue
        for run in reversed(paragraph.runs):
            if run.text.endswith("\t"):
                run.text = f"{run.text}{number}"
                break
    document.save(str(docx_path))
    return "libreoffice"


def finalize(docx_path: str | Path, toc: bool = True, author: str | None = None) -> dict[str, object]:
    """Bước cuối cho mọi file xuất ra: điền số trang mục lục rồi xoá nhãn trình tạo/AI.

    Chỉ ghi ra .docx (PDF tạm dùng để dò số trang nằm trong thư mục tạm và bị xoá ngay).
    Làm sạch chạy SAU cùng vì Word COM có thể ghi lại metadata khi lưu.
    """
    result: dict[str, object] = {}
    if toc:
        try:
            result["toc"] = update_toc(docx_path)
        except Exception as err:  # bước phụ, không làm hỏng việc xuất file
            result["toc"] = f"lỗi cập nhật mục lục: {err}"
    result["watermarks_removed"] = remove_watermarks(docx_path, author=author)
    return result
