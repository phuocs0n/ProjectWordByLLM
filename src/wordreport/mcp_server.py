"""MCP server `word-report`: bộ công cụ Word cho LLM.

LLM (Claude Desktop, Claude Code, hoặc agent trong `wordreport.agent`) gọi các tool này để
dựng báo cáo từng bước. Mỗi báo cáo đang soạn là một phiên (doc_id) giữ một ReportSpec trong
bộ nhớ; `save_report` render spec ra .docx (kèm file .spec.json để chỉnh sửa lại về sau).

Chạy:  wordreport-mcp            (stdio - dùng cho Claude Desktop / Claude Code)
       wordreport mcp --http     (streamable HTTP tại http://127.0.0.1:8765/mcp)
"""

from __future__ import annotations

import argparse
import functools
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from pydantic import TypeAdapter

from .inspector import lint_report, outline as docx_outline, read_source
from .postprocess import export_pdf, update_toc
from .renderer import DocxRenderer
from .skills import SkillLibrary
from .spec import Block, Reference, ReportMeta, ReportSpec
from .style_profile import describe_profiles

INSTRUCTIONS = """\
Bộ công cụ soạn thảo & định dạng báo cáo Microsoft Word (.docx) chuyên nghiệp.
Quy trình: list_skills → load_skill('report-structure-vn') → create_report → set_preface →
add_blocks (theo lô, mỗi chương một lần) → get_outline → save_report → lint_document.
Không tự đánh số tiêu đề, không gõ tay gạch đầu dòng/mục lục: renderer lo phần định dạng.
"""

server = MCPServer("word-report", instructions=INSTRUCTIONS)

_block_adapter: TypeAdapter[Any] = TypeAdapter(Block)


def tool(fn):
    """Đăng ký tool; lỗi dự kiến (dữ liệu sai, thiếu file...) được trả về cho LLM với thông điệp rõ ràng
    để nó tự sửa, thay vì thông báo chung chung "Error executing tool"."""

    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except ToolError:
            raise
        except (ValueError, KeyError, IndexError, FileNotFoundError, RuntimeError) as err:
            raise ToolError(str(err)) from err

    return server.tool()(wrapper)


@dataclass
class Session:
    spec: ReportSpec
    base_dir: Path = field(default_factory=Path.cwd)


SESSIONS: dict[str, Session] = {}
_skills: SkillLibrary | None = None


def _library() -> SkillLibrary:
    global _skills
    if _skills is None:
        _skills = SkillLibrary()
    return _skills


def _session(doc_id: str) -> Session:
    try:
        return SESSIONS[doc_id]
    except KeyError:
        raise ValueError(f"doc_id '{doc_id}' không tồn tại. Gọi create_report hoặc load_spec trước.") from None


def _append(doc_id: str, blocks: list[dict[str, Any]], position: int | None = None) -> dict[str, Any]:
    session = _session(doc_id)
    parsed = [_block_adapter.validate_python(b) for b in blocks]
    if position is None:
        session.spec.body.extend(parsed)
        start = len(session.spec.body) - len(parsed)
    else:
        start = max(0, min(position, len(session.spec.body)))
        session.spec.body[start:start] = parsed
    return {"added": len(parsed), "first_index": start, "total_blocks": len(session.spec.body)}


# ---------------------------------------------------------------------------
# Skills & profiles
# ---------------------------------------------------------------------------

@tool
def list_skills() -> list[dict[str, str]]:
    """Liệt kê kho skill xử lý tác vụ Word (tên + khi nào dùng)."""
    return [{"name": s.name, "description": s.description} for s in _library().skills.values()]


@tool
def load_skill(name: str) -> str:
    """Nạp nội dung đầy đủ của một skill (quy tắc + ví dụ JSON block). Gọi trước khi làm tác vụ tương ứng."""
    return _library().load(name)


@server.resource("skill://{name}", mime_type="text/markdown")
def skill_resource(name: str) -> str:
    """Nội dung skill dưới dạng MCP resource."""
    return _library().load(name)


@tool
def list_profiles() -> list[dict[str, str]]:
    """Liệt kê style profile (bộ quy chuẩn định dạng: khổ giấy, lề, font, màu tiêu đề, header/footer)."""
    return describe_profiles()


# ---------------------------------------------------------------------------
# Soạn báo cáo
# ---------------------------------------------------------------------------

@tool
def create_report(profile: str = "hcmus-clc", meta: dict[str, Any] | None = None,
                  include_toc: bool = True, base_dir: str = "") -> dict[str, Any]:
    """Tạo phiên báo cáo mới.

    Args:
        profile: tên style profile (xem list_profiles).
        meta: thông tin trang bìa - university, school, faculty, report_type, subject, topic,
            instructors[], members[{name, student_id, email}], class_code, city, year, logo_path.
        include_toc: có tạo MỤC LỤC tự động hay không.
        base_dir: thư mục gốc để phân giải đường dẫn ảnh/logo tương đối.
    """
    spec = ReportSpec(profile=profile, meta=ReportMeta.model_validate(meta or {}), include_toc=include_toc)
    doc_id = uuid.uuid4().hex[:8]
    SESSIONS[doc_id] = Session(spec=spec, base_dir=Path(base_dir) if base_dir else Path.cwd())
    return {"doc_id": doc_id, "profile": profile}


@tool
def set_meta(doc_id: str, meta: dict[str, Any]) -> dict[str, Any]:
    """Cập nhật một phần thông tin trang bìa/header (chỉ gửi các trường cần đổi)."""
    session = _session(doc_id)
    merged = {**session.spec.meta.model_dump(), **meta}
    session.spec.meta = ReportMeta.model_validate(merged)
    return session.spec.meta.model_dump()


@tool
def set_preface(doc_id: str, paragraphs: list[str]) -> str:
    """Đặt nội dung LỜI MỞ ĐẦU (mỗi phần tử là một đoạn). Danh sách rỗng = bỏ lời mở đầu."""
    _session(doc_id).spec.preface = paragraphs
    return f"Đã đặt lời mở đầu ({len(paragraphs)} đoạn)."


@tool
def add_blocks(doc_id: str, blocks: list[dict[str, Any]], position: int | None = None) -> dict[str, Any]:
    """Thêm nhiều block một lần (cách hiệu quả nhất). Mỗi block có trường `type`:
    heading{level 1-3, text} | paragraph{text, align} | list{items[], style dash|bullet|number|roman|alpha, level}
    | table{columns[], rows[][], caption, col_widths[]} | image{path, caption, width_cm}
    | figure_placeholder{caption, description} | code{code, caption, language}
    | note{text, kind note|tip|warning} | page_break | divider.
    Văn bản hỗ trợ **đậm**, *nghiêng*, `mã`, [chữ](url). `position` = chèn tại chỉ số (mặc định: cuối).
    """
    return _append(doc_id, blocks, position)


@tool
def add_heading(doc_id: str, text: str, level: int = 1) -> dict[str, Any]:
    """Thêm tiêu đề cấp 1-3. KHÔNG kèm số thứ tự - renderer tự đánh số."""
    return _append(doc_id, [{"type": "heading", "level": level, "text": text}])


@tool
def add_paragraph(doc_id: str, text: str, align: str = "justify") -> dict[str, Any]:
    """Thêm đoạn văn (justify|left|center|right)."""
    return _append(doc_id, [{"type": "paragraph", "text": text, "align": align}])


@tool
def add_list(doc_id: str, items: list[str], style: str = "dash", level: int = 0) -> dict[str, Any]:
    """Thêm danh sách tự động: dash | bullet | number | roman | alpha."""
    return _append(doc_id, [{"type": "list", "items": items, "style": style, "level": level}])


@tool
def add_table(doc_id: str, columns: list[str], rows: list[list[str]], caption: str = "",
              col_widths: list[float] | None = None) -> dict[str, Any]:
    """Thêm bảng có hàng tiêu đề tô nền + chú thích 'Bảng N'. Ô nhiều dòng dùng \\n, dòng '- ' thành gạch đầu dòng."""
    return _append(doc_id, [{"type": "table", "columns": columns, "rows": rows, "caption": caption,
                             "col_widths": col_widths or []}])


@tool
def add_image(doc_id: str, path: str, caption: str = "", width_cm: float = 14.0) -> dict[str, Any]:
    """Chèn ảnh (png/jpg) căn giữa kèm chú thích 'Hình N'."""
    return _append(doc_id, [{"type": "image", "path": path, "caption": caption, "width_cm": width_cm}])


@tool
def add_figure_placeholder(doc_id: str, caption: str, description: str = "") -> dict[str, Any]:
    """Chèn khung giữ chỗ cho ảnh chụp màn hình chưa có (người dùng chèn ảnh sau)."""
    return _append(doc_id, [{"type": "figure_placeholder", "caption": caption, "description": description}])


@tool
def add_code_block(doc_id: str, code: str, caption: str = "", language: str = "") -> dict[str, Any]:
    """Chèn khối lệnh/mã nguồn (font đơn cách, nền xám, viền trái)."""
    return _append(doc_id, [{"type": "code", "code": code, "caption": caption, "language": language}])


@tool
def add_note(doc_id: str, text: str, kind: str = "note") -> dict[str, Any]:
    """Chèn hộp ghi chú: note | tip | warning."""
    return _append(doc_id, [{"type": "note", "text": text, "kind": kind}])


@tool
def add_reference(doc_id: str, text: str, url: str = "") -> str:
    """Thêm một tài liệu tham khảo (renderer tự tạo chương 'Tài liệu tham khảo')."""
    session = _session(doc_id)
    session.spec.references.append(Reference(text=text, url=url))
    return f"Đã thêm tài liệu tham khảo [{len(session.spec.references)}]."


@tool
def update_block(doc_id: str, index: int, block: dict[str, Any]) -> str:
    """Thay block tại chỉ số `index` (xem chỉ số bằng get_outline)."""
    session = _session(doc_id)
    if not 0 <= index < len(session.spec.body):
        raise ValueError(f"index {index} ngoài phạm vi 0..{len(session.spec.body) - 1}")
    session.spec.body[index] = _block_adapter.validate_python(block)
    return f"Đã cập nhật block {index}."


@tool
def delete_block(doc_id: str, index: int) -> str:
    """Xoá block tại chỉ số `index`."""
    session = _session(doc_id)
    removed = session.spec.body.pop(index)
    return f"Đã xoá block {index} ({removed.type})."


@tool
def move_block(doc_id: str, index: int, new_index: int) -> str:
    """Di chuyển block từ `index` sang `new_index`."""
    body = _session(doc_id).spec.body
    block = body.pop(index)
    body.insert(new_index, block)
    return f"Đã chuyển block {index} → {new_index}."


@tool
def get_outline(doc_id: str) -> str:
    """Dàn ý hiện tại kèm số thứ tự tiêu đề sẽ được đánh và chỉ số block."""
    spec = _session(doc_id).spec
    head = f"profile={spec.profile} | toc={spec.include_toc} | preface={len(spec.preface)} đoạn | refs={len(spec.references)}"
    return head + "\n" + "\n".join(spec.outline())


@tool
def get_spec(doc_id: str) -> str:
    """Toàn bộ ReportSpec dạng JSON."""
    return _session(doc_id).spec.model_dump_json(indent=2)


@tool
def load_spec(spec_path: str) -> dict[str, Any]:
    """Mở một file ReportSpec JSON (ví dụ *.spec.json đã lưu) thành phiên soạn thảo mới."""
    path = Path(spec_path)
    spec = ReportSpec.model_validate_json(path.read_text(encoding="utf-8"))
    doc_id = uuid.uuid4().hex[:8]
    SESSIONS[doc_id] = Session(spec=spec, base_dir=path.parent)
    return {"doc_id": doc_id, "blocks": len(spec.body)}


@tool
def save_report(doc_id: str, output_path: str, update_toc_pages: bool = True,
                pdf: bool = False) -> dict[str, Any]:
    """Render báo cáo ra .docx (và .spec.json cạnh bên), điền số trang mục lục, lint, tuỳ chọn xuất PDF."""
    session = _session(doc_id)
    out = Path(output_path)
    if out.suffix.lower() != ".docx":
        out = out.with_suffix(".docx")
    DocxRenderer(session.spec.profile, base_dir=session.base_dir).save(session.spec, out)
    spec_file = out.with_suffix(".spec.json")
    spec_file.write_text(session.spec.model_dump_json(indent=2), encoding="utf-8")
    result: dict[str, Any] = {"docx": str(out.resolve()), "spec": str(spec_file.resolve())}
    if update_toc_pages and session.spec.include_toc:
        try:
            result["toc"] = update_toc(out)
        except Exception as err:  # không làm hỏng việc lưu vì bước phụ
            result["toc"] = f"lỗi cập nhật mục lục: {err}"
    if pdf:
        try:
            result["pdf"] = str(export_pdf(out))
        except Exception as err:
            result["pdf"] = f"lỗi xuất PDF: {err}"
    lint = lint_report(out, session.spec.profile)
    result["lint"] = {k: lint[k] for k in ("errors", "warnings", "infos")}
    result["lint_issues"] = lint["issues"][:20]
    return result


@tool
def render_spec_file(spec_path: str, output_path: str) -> dict[str, Any]:
    """Render trực tiếp một file ReportSpec JSON ra .docx."""
    loaded = load_spec(spec_path)
    return save_report(loaded["doc_id"], output_path)


# ---------------------------------------------------------------------------
# Đọc & kiểm tra tài liệu có sẵn
# ---------------------------------------------------------------------------

@tool
def read_document(path: str) -> str:
    """Đọc .docx/.md/.txt (và PDF/PPTX/XLSX nếu có MarkItDown) thành Markdown."""
    return read_source(path)


@tool
def document_outline(path: str) -> list[dict[str, Any]]:
    """Danh sách tiêu đề (level, text) của một file .docx."""
    return docx_outline(path)


@tool
def lint_document(path: str, profile: str = "hcmus-clc") -> dict[str, Any]:
    """Kiểm tra chất lượng định dạng file .docx: font, tiêu đề, chú thích, gạch đầu dòng gõ tay, chính tả, mục lục, số trang."""
    return lint_report(path, profile)


@tool
def export_to_pdf(path: str) -> str:
    """Xuất .docx sang PDF bằng LibreOffice."""
    return str(export_pdf(path))


def main() -> None:
    parser = argparse.ArgumentParser(description="MCP server word-report")
    parser.add_argument("--http", action="store_true", help="Chạy streamable HTTP thay vì stdio")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()
    if args.http:
        server.run("streamable-http", host=args.host, port=args.port)
    else:
        server.run()


if __name__ == "__main__":
    main()

