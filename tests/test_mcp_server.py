import json

import anyio
from mcp import Client

from wordreport.mcp_server import server


def test_mcp_end_to_end(tmp_path):
    async def scenario():
        async with Client(server) as client:
            names = {t.name for t in (await client.list_tools()).tools}
            assert {"create_report", "add_blocks", "save_report", "lint_document", "load_skill", "remove_watermarks"} <= names
            assert not any("pdf" in n for n in names), "MCP chỉ xuất .docx"

            created = await client.call_tool("create_report", {"meta": {"subject": "Kiểm thử", "class_code": "X1"}})
            doc_id = created.structured_content["doc_id"]
            await client.call_tool("set_preface", {"doc_id": doc_id, "paragraphs": ["Mở đầu."]})
            added = await client.call_tool("add_blocks", {"doc_id": doc_id, "blocks": [
                {"type": "heading", "level": 1, "text": "Chương một"},
                {"type": "paragraph", "text": "Nội dung **quan trọng**."},
                {"type": "table", "columns": ["A", "B"], "rows": [["1", "2"]], "caption": "Bảng thử"},
            ]})
            assert added.structured_content["total_blocks"] == 3
            await client.call_tool("add_reference", {"doc_id": doc_id, "text": "Tài liệu A", "url": "https://example.com"})

            outline = await client.call_tool("get_outline", {"doc_id": doc_id})
            assert "I. Chương một" in outline.content[0].text

            bad = await client.call_tool("add_blocks", {"doc_id": doc_id, "blocks": [{"type": "khong-co"}]})
            assert bad.is_error and "khong-co" in bad.content[0].text

            out = tmp_path / "mcp.docx"
            saved = await client.call_tool("save_report", {"doc_id": doc_id, "output_path": str(out), "update_toc_pages": False})
            result = saved.structured_content
            assert out.exists() and (tmp_path / "mcp.spec.json").exists()
            assert not list(tmp_path.glob("*.pdf"))
            assert result["watermarks_removed"] is not None
            assert result["lint"]["errors"] == 0 and result["lint"]["warnings"] == 0

            reloaded = await client.call_tool("load_spec", {"spec_path": str(tmp_path / "mcp.spec.json")})
            assert reloaded.structured_content["blocks"] == 3

            md = await client.call_tool("read_document", {"path": str(out)})
            assert "Chương một" in md.content[0].text

            skill = await client.call_tool("load_skill", {"name": "tables"})
            assert "Bảng" in skill.content[0].text
            json.dumps(result)

    anyio.run(scenario)
