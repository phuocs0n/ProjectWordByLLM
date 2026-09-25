import json

import anyio
from mcp import Client

from wordreport.mcp_server import server


def test_mcp_end_to_end(tmp_path):
    async def scenario():
        async with Client(server) as client:
            names = {t.name for t in (await client.list_tools()).tools}
            assert {"create_report", "add_blocks", "save_report", "lint_document", "load_skill", "remove_watermarks",
                    "set_assignments", "check_barem"} <= names
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

            await client.call_tool("set_assignments", {"doc_id": doc_id, "assignments": [
                {"member": "An", "tasks": ["Viết báo cáo"], "completion": "100%"}]})
            intro = await client.call_tool("add_blocks", {"doc_id": doc_id, "section": "introduction", "blocks": [
                {"type": "heading", "level": 1, "text": "Tóm tắt"}]})
            assert intro.structured_content["section"] == "introduction"

            outline = (await client.call_tool("get_outline", {"doc_id": doc_id})).content[0].text
            for line in ("I. Giới thiệu chung", "1. Thành viên nhóm", "2. Bảng phân công công việc",
                         "3. Tóm tắt", "II. Nội dung", "1. Chương một", "III. Tài liệu tham khảo"):
                assert line in outline

            barem = (await client.call_tool("check_barem", {"doc_id": doc_id})).structured_content["result"]
            assert any("Thành viên nhóm" in i["message"] for i in barem)  # meta chưa có members

            bad = await client.call_tool("add_blocks", {"doc_id": doc_id, "blocks": [{"type": "khong-co"}]})
            assert bad.is_error and "khong-co" in bad.content[0].text

            out = tmp_path / "mcp.docx"
            saved = await client.call_tool("save_report", {"doc_id": doc_id, "output_path": str(out), "update_toc_pages": False})
            result = saved.structured_content
            assert out.exists()
            assert [f.name for f in tmp_path.iterdir()] == ["mcp.docx"], "chỉ xuất đúng một file .docx"
            assert result["watermarks_removed"] is not None
            assert result["lint"]["errors"] == 0 and result["lint"]["warnings"] == 0
            assert any(i["rule"] == "barem" for i in result["barem"])

            spec_json = (await client.call_tool("get_spec", {"doc_id": doc_id})).content[0].text
            spec_path = tmp_path / "mcp-spec.json"
            spec_path.write_text(spec_json, encoding="utf-8")
            reloaded = await client.call_tool("load_spec", {"spec_path": str(spec_path)})
            assert reloaded.structured_content["blocks"] == 3

            md = await client.call_tool("read_document", {"path": str(out)})
            assert "Chương một" in md.content[0].text

            skill = await client.call_tool("load_skill", {"name": "tables"})
            assert "Bảng" in skill.content[0].text
            json.dumps(result)

    anyio.run(scenario)
