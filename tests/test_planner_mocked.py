"""Chế độ plan với API giả lập trả về stream SSE chứa ReportSpec JSON."""

import json

import anthropic
import httpx2
from anthropic import DefaultHttpxClient

from wordreport.planner import plan_report
from wordreport.renderer import DocxRenderer


def _sse(events):
    return "".join(f"event: {e['type']}\ndata: {json.dumps(e, ensure_ascii=False)}\n\n" for e in events)


def test_plan_report_parses_structured_output(tmp_path):
    spec_json = json.dumps({
        "profile": "hcmus-clc",
        "meta": {"subject": "Phân tích dữ liệu", "members": [{"name": "Nguyễn A"}]},
        "preface": ["Mở đầu."],
        "include_toc": True,
        "body": [
            {"type": "heading", "level": 1, "text": "Kết quả"},
            {"type": "table", "columns": ["Khu vực", "Doanh thu"], "rows": [["Bắc", "12,5"]], "caption": "Doanh thu"},
            {"type": "page_break"},
        ],
        "references": [],
    }, ensure_ascii=False)
    seen = {}

    def handler(request: httpx2.Request) -> httpx2.Response:
        body = json.loads(request.content)
        seen["body"] = body
        events = [
            {"type": "message_start", "message": {"id": "msg_1", "type": "message", "role": "assistant",
             "model": body["model"], "content": [], "stop_reason": None, "stop_sequence": None,
             "usage": {"input_tokens": 5, "output_tokens": 1}}},
            {"type": "content_block_start", "index": 0, "content_block": {"type": "text", "text": ""}},
            {"type": "content_block_delta", "index": 0, "delta": {"type": "text_delta", "text": spec_json[:40]}},
            {"type": "content_block_delta", "index": 0, "delta": {"type": "text_delta", "text": spec_json[40:]}},
            {"type": "content_block_stop", "index": 0},
            {"type": "message_delta", "delta": {"stop_reason": "end_turn", "stop_sequence": None},
             "usage": {"output_tokens": 50}},
            {"type": "message_stop"},
        ]
        return httpx2.Response(200, text=_sse(events), headers={"content-type": "text/event-stream"})

    client = anthropic.Anthropic(api_key="k", http_client=DefaultHttpxClient(transport=httpx2.MockTransport(handler)))
    spec = plan_report("Báo cáo doanh thu", profile="nd30-a4", client=client)

    assert spec.profile == "nd30-a4"
    assert spec.body[1].type == "table" and spec.body[1].rows == [["Bắc", "12,5"]]
    body = seen["body"]
    assert body["output_config"]["effort"] == "high"
    assert body["output_config"]["format"]["type"] == "json_schema"
    assert "Bảng" in body["system"]  # nội dung skill `tables` được nhúng
    DocxRenderer(spec.profile).save(spec, tmp_path / "plan.docx")
