"""Giao diện dòng lệnh `wordreport`.

Ví dụ:
  wordreport render examples/mang-may-tinh-do-an.json -o out/bao-cao.docx
  wordreport clean bao-cao.docx        # watermark-remover: xoá nhãn công cụ/AI
  wordreport generate "Soạn báo cáo đồ án..." --source de-bai.pdf --notes ghi-chu.md -o out/bc.docx
  wordreport lint out/bao-cao.docx
  wordreport mcp                       # chạy MCP server (stdio)
  wordreport skills install            # chép kho skill vào ~/.claude/skills
"""

from __future__ import annotations

import argparse
import asyncio
import json
import shutil
import sys
from pathlib import Path

from .inspector import docx_to_markdown, lint_report, outline
from .postprocess import export_pdf, finalize
from .renderer import DocxRenderer
from .skills import SkillLibrary
from .spec import ReportSpec
from .style_profile import describe_profiles
from .watermark_remover import main as watermark_remover_main


def _finish(out: Path, spec: ReportSpec, pdf: bool, toc: bool) -> None:
    result = finalize(out, toc=toc and spec.include_toc, author=DocxRenderer.author(spec))
    if "toc" in result:
        print(f"Mục lục: {result['toc']}")
    print(f"Watermark-remover: đã xoá nhãn trình tạo/AI ({len(result['watermarks_removed'])} thay đổi)")
    lint = lint_report(out, spec.profile)
    print(f"Lint: {lint['errors']} lỗi, {lint['warnings']} cảnh báo, {lint['infos']} gợi ý")
    for issue in lint["issues"]:
        if issue["severity"] != "info":
            print(f"  [{issue['severity']}] {issue['rule']}: {issue['message']}")
    if pdf:
        print(f"PDF: {export_pdf(out)}")
    print(f"Đã tạo: {out}")


def cmd_render(args) -> None:
    spec_path = Path(args.spec)
    spec = ReportSpec.model_validate_json(spec_path.read_text(encoding="utf-8"))
    if args.profile:
        spec.profile = args.profile
    out = Path(args.output or spec_path.with_suffix(".docx"))
    DocxRenderer(spec.profile, base_dir=spec_path.parent).save(spec, out)
    _finish(out, spec, args.pdf, not args.no_toc_update)


def _request_text(args) -> str:
    parts = [args.request] if args.request else []
    if args.notes:
        parts.append(Path(args.notes).read_text(encoding="utf-8"))
    if not parts:
        sys.exit("Cần mô tả yêu cầu (tham số request) hoặc --notes.")
    return "\n\n".join(parts)


def cmd_generate(args) -> None:
    request = _request_text(args)
    out = Path(args.output)
    options = {"profile": args.profile, "sources": args.source}
    if args.model:
        options["model"] = args.model
    if args.effort:
        options["effort"] = args.effort

    if args.mode == "plan":
        from .planner import plan_report

        spec = plan_report(request, **options)
        out.parent.mkdir(parents=True, exist_ok=True)
        DocxRenderer(spec.profile, base_dir=Path.cwd()).save(spec, out)
        _finish(out, spec, args.pdf, True)
    else:
        from .agent import run_agent

        answer = asyncio.run(run_agent(request, out, use_markitdown=args.markitdown,
                                       use_word_mcp=not args.no_word_mcp, **options))
        print("\n" + answer)
        if args.pdf and out.exists():
            print(f"PDF: {export_pdf(out)}")


def cmd_inspect(args) -> None:
    if args.outline:
        for item in outline(args.docx):
            print(f"{'  ' * (item['level'] - 1)}{item['text']}")
    else:
        print(docx_to_markdown(args.docx))


def cmd_lint(args) -> None:
    report = lint_report(args.docx, args.profile)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return
    print(f"{report['file']}: {report['errors']} lỗi, {report['warnings']} cảnh báo, {report['infos']} gợi ý")
    for issue in report["issues"]:
        loc = f" ({issue['location']})" if issue["location"] else ""
        print(f"  [{issue['severity']}] {issue['rule']}{loc}: {issue['message']}")


def cmd_mcp(args) -> None:
    from .mcp_server import server

    if args.http:
        server.run("streamable-http", host=args.host, port=args.port)
    else:
        server.run()


def cmd_skills(args) -> None:
    library = SkillLibrary()
    if args.action == "list":
        for skill in library.skills.values():
            print(f"{skill.name:24} {skill.description}")
    elif args.action == "show":
        print(library.load(args.name))
    elif args.action == "install":
        dest = Path(args.dest).expanduser()
        dest.mkdir(parents=True, exist_ok=True)
        for skill in library.skills.values():
            target = dest / skill.name
            shutil.copytree(skill.path.parent, target, dirs_exist_ok=True)
        print(f"Đã chép {len(library.skills)} skill vào {dest}")


def cmd_clean(args) -> None:
    argv = list(args.files)
    if args.output:
        argv += ["-o", args.output]
    if args.author:
        argv += ["--author", args.author]
    if args.check:
        argv.append("--check")
    watermark_remover_main(argv)


def cmd_profiles(_args) -> None:
    for profile in describe_profiles():
        print(f"{profile['name']:12} {profile['description']}")


def cmd_pdf(args) -> None:
    print(export_pdf(args.docx))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="wordreport", description="Soạn & định dạng báo cáo Word bằng LLM")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("render", help="Render ReportSpec JSON ra .docx (không cần LLM)")
    p.add_argument("spec")
    p.add_argument("-o", "--output")
    p.add_argument("--profile", help="Ghi đè style profile trong spec")
    p.add_argument("--pdf", action="store_true", help="Xuất thêm PDF (mặc định chỉ xuất .docx; cần Microsoft Word trên Windows)")
    p.add_argument("--no-toc-update", action="store_true", help="Không điền số trang mục lục")
    p.set_defaults(func=cmd_render)

    p = sub.add_parser("generate", help="Dùng Claude soạn báo cáo từ yêu cầu/ghi chú/tài liệu nguồn")
    p.add_argument("request", nargs="?", default="", help="Mô tả yêu cầu báo cáo")
    p.add_argument("--notes", help="File ghi chú (.md/.txt) bổ sung vào yêu cầu")
    p.add_argument("--source", action="append", default=[], help="Tài liệu nguồn (.docx/.pdf/.pptx/.xlsx/.md), lặp lại được")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--mode", choices=["agent", "plan"], default="agent",
                   help="agent: Claude gọi tool MCP từng bước (mặc định); plan: 1 lượt gọi sinh ReportSpec")
    p.add_argument("--profile", default="hcmus-clc")
    p.add_argument("--model", help="Model Claude (mặc định claude-opus-5 hoặc biến WORDREPORT_MODEL)")
    p.add_argument("--effort", choices=["low", "medium", "high", "xhigh", "max"])
    p.add_argument("--markitdown", action="store_true", help="Kết nối thêm MCP markitdown (markitdown-mcp)")
    p.add_argument("--no-word-mcp", action="store_true",
                   help="Không kết nối MCP Microsoft Word (Office-Word-MCP-Server, mặc định bật)")
    p.add_argument("--pdf", action="store_true", help="Xuất thêm PDF (mặc định chỉ xuất .docx)")
    p.set_defaults(func=cmd_generate)

    p = sub.add_parser("inspect", help="Đọc .docx thành Markdown / dàn ý")
    p.add_argument("docx")
    p.add_argument("--outline", action="store_true")
    p.set_defaults(func=cmd_inspect)

    p = sub.add_parser("lint", help="Kiểm tra chất lượng định dạng .docx")
    p.add_argument("docx")
    p.add_argument("--profile", default="hcmus-clc")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_lint)

    p = sub.add_parser("mcp", help="Chạy MCP server word-report")
    p.add_argument("--http", action="store_true")
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=8765)
    p.set_defaults(func=cmd_mcp)

    p = sub.add_parser("skills", help="Quản lý kho skill")
    p.add_argument("action", choices=["list", "show", "install"])
    p.add_argument("name", nargs="?")
    p.add_argument("--dest", default="~/.claude/skills")
    p.set_defaults(func=cmd_skills)

    p = sub.add_parser("clean", help="watermark-remover: xoá nhãn công cụ/AI khỏi file .docx")
    p.add_argument("files", nargs="+")
    p.add_argument("-o", "--output")
    p.add_argument("--author", help="Đặt tên tác giả / người sửa cuối")
    p.add_argument("--check", action="store_true", help="Chỉ liệt kê nhãn còn sót")
    p.set_defaults(func=cmd_clean)

    p = sub.add_parser("profiles", help="Liệt kê style profile")
    p.set_defaults(func=cmd_profiles)

    p = sub.add_parser("pdf", help="Xuất .docx sang PDF")
    p.add_argument("docx")
    p.set_defaults(func=cmd_pdf)
    return parser


def main(argv: list[str] | None = None) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")  # console Windows
    args = build_parser().parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
