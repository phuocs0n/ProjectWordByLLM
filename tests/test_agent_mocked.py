"""Chạy agent end-to-end với Claude API giả lập (không cần API key, không tốn phí).

Mock đóng vai model: gọi create_report → add_blocks → save_report → trả lời. Test kiểm tra
cả tham số gửi lên API (adaptive thinking, fallbacks, skill index trong system prompt)."""

import json
import re

import anthropic
import anyio
import httpx2
from anthropic import DefaultAsyncHttpxClient

from wordreport.agent import run_agent
from wordreport.llm_common import FALLBACK_BETA
from wordreport.mcp_server import server


def _message(content, stop_reason):
    return {
        "id": "msg_test", "type": "message", "role": "assistant", "model": "claude-opus-5",
        "content": content, "stop_reason": stop_reason, "stop_sequence": None,
        "usage": {"input_tokens": 10, "output_tokens": 10},
    }


def _tool_use(idx, name, args):
    return {"type": "tool_use", "id": f"toolu_{idx}", "name": name, "input": args}


def test_agent_builds_report_via_mcp(tmp_path):
    out = tmp_path / "agent.docx"
    requests = []

    def handler(request: httpx2.Request) -> httpx2.Response:
        body = json.loads(request.content)
        requests.append((request, body))
        step = sum(1 for m in body["messages"] if m["role"] == "assistant")
        if step == 0:
            content = [_tool_use(0, "create_report", {"meta": {"subject": "Mạng máy tính", "class_code": "22CLC03"}})]
            return httpx2.Response(200, json=_message(content, "tool_use"))
        if step == 1:
            last = json.dumps(body["messages"][-1], ensure_ascii=False)
            doc_id = re.search(r'doc_id\\?"\s*:\s*\\?"(\w+)', last).group(1)
            blocks = [
                {"type": "heading", "level": 1, "text": "Nội dung"},
                {"type": "paragraph", "text": "Cấu hình **DHCP** cho mạng 172.2.1.0/24."},
                {"type": "list", "style": "number", "items": ["Bật dịch vụ.", "Lưu cấu hình."]},
            ]
            content = [
                _tool_use(1, "add_blocks", {"doc_id": doc_id, "blocks": blocks}),
                _tool_use(2, "save_report", {"doc_id": doc_id, "output_path": str(out), "update_toc_pages": False}),
            ]
            return httpx2.Response(200, json=_message(content, "tool_use"))
        return httpx2.Response(200, json=_message([{"type": "text", "text": f"Đã lưu {out}"}], "end_turn"))

    client = anthropic.AsyncAnthropic(
        api_key="test-key", http_client=DefaultAsyncHttpxClient(transport=httpx2.MockTransport(handler))
    )
    events = []
    answer = anyio.run(lambda: run_agent(
        "Soạn báo cáo DHCP", out, client=client, word_server=server, on_event=events.append))

    assert out.exists()
    assert "Đã lưu" in answer
    assert "→ add_blocks" in events and "→ save_report" in events
    assert len(requests) == 3

    first_request, first = requests[0]
    assert first["model"] == "claude-opus-5"
    assert first["thinking"] == {"type": "adaptive"}
    assert first["fallbacks"] == "default"
    assert FALLBACK_BETA in first_request.headers.get("anthropic-beta", "")
    system = first["system"] if isinstance(first["system"], str) else first["system"][0]["text"]
    assert "report-structure-vn" in system and "quality-check" in system
    assert {t["name"] for t in first["tools"]} >= {"create_report", "add_blocks", "save_report", "load_skill"}

    # kết quả lỗi/thành công của tool được trả lại cho model trong cùng một user message
    tool_results = [b for b in requests[2][1]["messages"][-1]["content"] if b["type"] == "tool_result"]
    assert len(tool_results) == 2 and not any(r.get("is_error") for r in tool_results)
