"""wordreport - tự động soạn thảo & định dạng báo cáo Word (.docx) bằng LLM."""

from .renderer import DocxRenderer, render_spec
from .spec import ReportSpec

__all__ = ["DocxRenderer", "ReportSpec", "render_spec"]
__version__ = "0.1.0"
