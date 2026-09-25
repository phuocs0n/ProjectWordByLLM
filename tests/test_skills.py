from pathlib import Path

from wordreport.skills import SkillLibrary

ROOT = Path(__file__).resolve().parents[1]


def test_every_skill_is_well_formed():
    library = SkillLibrary(ROOT / "skills")
    assert len(library.skills) >= 12
    for name, skill in library.skills.items():
        assert skill.path.parent.name == name, "tên skill phải trùng tên thư mục"
        assert 40 <= len(skill.description) <= 1024
        assert skill.body.startswith("#")
    assert "report-structure-vn" in library.index()
    assert library.load("tables").startswith("# Skill: tables")


def test_skills_reference_real_mcp_tools():
    """Tên tool nhắc trong skill phải tồn tại trong MCP server."""
    import re

    from wordreport import mcp_server

    tool_names = {name for name in dir(mcp_server) if not name.startswith("_")}
    library = SkillLibrary(ROOT / "skills")
    mentioned = set()
    for skill in library.skills.values():
        mentioned |= set(re.findall(r"`([a-z_]+)\(", skill.body))
    assert mentioned, "skill nên hướng dẫn cách gọi tool"
    external = {  # tool của MCP server markitdown và MCP Microsoft Word (Office-Word-MCP-Server)
        "convert_to_markdown", "get_document_info", "get_document_outline", "find_text_in_document",
        "search_and_replace", "format_text", "set_table_column_widths", "merge_table_cells_vertical",
        "set_table_cell_alignment", "format_table_cell_text", "add_footnote_after_text", "protect_document",
        "copy_document", "create_document", "add_heading", "add_paragraph", "convert_to_pdf",
    }
    assert mentioned - external <= tool_names, mentioned - external - tool_names
