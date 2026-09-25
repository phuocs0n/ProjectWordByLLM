"""Renderer tất định: ReportSpec + style profile -> tài liệu .docx.

Bố cục 3 section giống báo cáo mẫu:
  1. Trang bìa  - khung viền, không header/footer.
  2. Phần đầu   - LỜI MỞ ĐẦU, MỤC LỤC; footer có số trang.
  3. Nội dung   - header (Khoa | Mã lớp) có đường kẻ, footer số trang có đường kẻ.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from ..numbering import HeadingNumberer
from ..spec import (
    CodeBlock,
    DividerBlock,
    FigurePlaceholderBlock,
    HeadingBlock,
    ImageBlock,
    ListBlock,
    NoteBlock,
    PageBreakBlock,
    ParagraphBlock,
    ReportSpec,
    TableBlock,
)
from ..style_profile import load_profile
from . import ooxml

ALIGN = {
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}

PAGE_SIZES_CM = {"letter": (21.59, 27.94), "a4": (21.0, 29.7)}

FRONT_STYLE = "Front Heading"
CODE_STYLE = "Code Block"
CAPTION_STYLE = "Caption"

# **đậm**, *nghiêng*, `mã`, [chữ](url)
_INLINE = re.compile(r"(\*\*.+?\*\*|\*[^*]+?\*|`[^`]+?`|\[[^\]]+?\]\([^)]+?\))")

NOTE_LABELS = {"note": "Lưu ý: ", "tip": "Mẹo: ", "warning": "Cảnh báo: "}
NOTE_COLORS = {"note": ("2F5496", "EAF1FB"), "tip": ("2E7D32", "EDF7EE"), "warning": ("C00000", "FDECEA")}


class DocxRenderer:
    def __init__(self, profile: dict[str, Any] | str = "hcmus-clc", base_dir: str | Path = ".") -> None:
        self.p = load_profile(profile) if isinstance(profile, str) else profile
        self.base_dir = Path(base_dir)
        self.font = self.p["fonts"]["body"]
        self.body_size = self.p["body"]["size_pt"]
        self.figure_count = 0
        self.table_count = 0

    # ------------------------------------------------------------------ public

    def render(self, spec: ReportSpec):
        self.figure_count = 0
        self.table_count = 0
        doc = Document()
        self.doc = doc
        self._setup_styles()
        self.numbering = ooxml.NumberingManager(doc, self.font)

        cover = doc.sections[0]
        self._setup_page(cover)
        if self.p["page"].get("cover_border"):
            ooxml.page_border(cover)
        self._cover(spec)

        front = doc.add_section(WD_SECTION.NEW_PAGE)
        self._setup_page(front)
        self._header_footer(front, spec, with_header=False)
        toc_entries = self._toc_entries(spec)
        if spec.preface:
            self._front_heading(self.p["labels"]["preface"])
            for text in spec.preface:
                self._paragraph(text)
            self._divider()
        if spec.include_toc:
            if spec.preface:
                self._page_break()
            self._front_heading(self.p["labels"]["toc"])
            self._toc(toc_entries)
            self._divider()

        body = doc.add_section(WD_SECTION.NEW_PAGE)
        self._setup_page(body)
        self._header_footer(body, spec, with_header=True)
        self._body(spec)

        ooxml.enable_update_fields_on_open(doc)
        props = doc.core_properties
        props.title = " - ".join(x for x in (spec.meta.report_type, spec.meta.subject) if x)
        props.subject = spec.meta.topic
        props.author = ", ".join(m.name for m in spec.meta.members)
        props.comments = "Tạo bởi wordreport (LLM + MCP)"
        return doc

    def save(self, spec: ReportSpec, path: str | Path) -> Path:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.render(spec).save(str(path))
        return path

    # ------------------------------------------------------------------ setup

    @property
    def text_width_cm(self) -> float:
        page = self.p["page"]
        width, _ = PAGE_SIZES_CM[page["size"]]
        return width - page["margin_left_cm"] - page["margin_right_cm"]

    def _setup_page(self, section) -> None:
        page = self.p["page"]
        width, height = PAGE_SIZES_CM[page["size"]]
        section.page_width = Cm(width)
        section.page_height = Cm(height)
        section.top_margin = Cm(page["margin_top_cm"])
        section.bottom_margin = Cm(page["margin_bottom_cm"])
        section.left_margin = Cm(page["margin_left_cm"])
        section.right_margin = Cm(page["margin_right_cm"])
        section.header_distance = Cm(page["header_distance_cm"])
        section.footer_distance = Cm(page["footer_distance_cm"])

    def _setup_styles(self) -> None:
        styles = self.doc.styles
        body = self.p["body"]
        normal = styles["Normal"]
        ooxml.set_style_font(normal, self.font, self.body_size)
        pf = normal.paragraph_format
        pf.line_spacing = body["line_spacing"]
        pf.space_after = Pt(body["space_after_pt"])
        pf.space_before = Pt(0)

        head = self.p["headings"]
        for level in (1, 2, 3):
            cfg = {"color": head["color"], **head[f"level{level}"]}
            style = styles[f"Heading {level}"]
            ooxml.set_style_font(style, self.font, cfg["size_pt"])
            style.font.bold = cfg.get("bold", True)
            style.font.italic = cfg.get("italic", False)
            style.font.color.rgb = ooxml.hex_color(cfg["color"])
            spf = style.paragraph_format
            spf.space_before = Pt(cfg["space_before_pt"])
            spf.space_after = Pt(cfg["space_after_pt"])
            spf.left_indent = Cm(cfg.get("indent_cm", 0))
            spf.first_line_indent = Cm(0)
            spf.keep_with_next = True
            spf.line_spacing = 1.0
            spf.alignment = WD_ALIGN_PARAGRAPH.LEFT

        front_cfg = head["front"]
        front = styles.add_style(FRONT_STYLE, WD_STYLE_TYPE.PARAGRAPH)
        front.base_style = styles["Heading 1"]
        front.next_paragraph_style = normal
        ooxml.set_style_font(front, self.font, front_cfg["size_pt"])
        front.font.bold = True
        front.font.color.rgb = ooxml.hex_color(front_cfg["color"])
        front.paragraph_format.alignment = ALIGN[front_cfg["align"]]
        front.paragraph_format.space_before = Pt(0)
        front.paragraph_format.space_after = Pt(18)
        front.paragraph_format.left_indent = Cm(0)
        ooxml.set_outline_level(front, 0)  # để MỤC LỤC liệt kê được

        cap = self.p["caption"]
        caption = styles[CAPTION_STYLE] if CAPTION_STYLE in [s.name for s in styles] else styles.add_style(CAPTION_STYLE, WD_STYLE_TYPE.PARAGRAPH)
        caption.base_style = normal
        ooxml.set_style_font(caption, self.font, cap["size_pt"])
        caption.font.italic = cap["italic"]
        caption.font.bold = False
        caption.font.color.rgb = ooxml.hex_color("000000")
        caption.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption.paragraph_format.space_before = Pt(3)
        caption.paragraph_format.space_after = Pt(10)

        code_cfg = self.p["code"]
        code = styles.add_style(CODE_STYLE, WD_STYLE_TYPE.PARAGRAPH)
        code.base_style = normal
        ooxml.set_style_font(code, self.p["fonts"]["code"], code_cfg["size_pt"])
        code.paragraph_format.line_spacing = 1.0
        code.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        code.paragraph_format.space_after = Pt(8)
        code.paragraph_format.keep_together = True

        for level in (1, 2, 3):
            toc = styles.add_style(f"toc {level}", WD_STYLE_TYPE.PARAGRAPH)
            toc.base_style = normal
            toc.paragraph_format.left_indent = Cm(0.6 * (level - 1))
            toc.paragraph_format.space_after = Pt(3)
            toc.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            toc.paragraph_format.tab_stops.add_tab_stop(
                Cm(self.text_width_cm), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
            )
            toc.font.bold = level == 1

    # ------------------------------------------------------------------ primitives

    def _inline(self, paragraph, text: str, size_pt: float | None = None, bold: bool | None = None,
                italic: bool | None = None) -> None:
        size_pt = size_pt or self.body_size
        for part in _INLINE.split(text):
            if not part:
                continue
            if part.startswith("**") and part.endswith("**") and len(part) > 4:
                run = paragraph.add_run(part[2:-2])
                run.bold = True
                ooxml.set_run_font(run, self.font, size_pt)
            elif part.startswith("`") and part.endswith("`") and len(part) > 2:
                run = paragraph.add_run(part[1:-1])
                ooxml.set_run_font(run, self.p["fonts"]["code"], size_pt - 1)
            elif part.startswith("[") and "](" in part and part.endswith(")"):
                label, url = part[1:-1].split("](", 1)
                ooxml.add_hyperlink(paragraph, label, url, self.font, size_pt, self.p["hyperlink_color"])
                continue
            elif part.startswith("*") and part.endswith("*") and len(part) > 2:
                run = paragraph.add_run(part[1:-1])
                run.italic = True
                ooxml.set_run_font(run, self.font, size_pt)
            else:
                run = paragraph.add_run(part)
                ooxml.set_run_font(run, self.font, size_pt)
            if bold is not None:
                run.bold = bold or run.bold
            if italic is not None:
                run.italic = italic or run.italic

    def _paragraph(self, text: str, align: str | None = None, container=None):
        container = container or self.doc
        para = container.add_paragraph()
        para.alignment = ALIGN[align or self.p["body"]["align"]]
        indent = self.p["body"].get("first_line_indent_cm", 0)
        if indent:
            para.paragraph_format.first_line_indent = Cm(indent)
        self._inline(para, text)
        return para

    def _centered(self, text: str, size_pt: float, bold: bool = True, space_after: float = 0,
                  space_before: float = 0, italic: bool = False):
        para = self.doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_after = Pt(space_after)
        para.paragraph_format.space_before = Pt(space_before)
        para.paragraph_format.line_spacing = 1.0
        run = para.add_run(text)
        run.bold = bold
        run.italic = italic
        ooxml.set_run_font(run, self.font, size_pt)
        return para

    def _front_heading(self, text: str) -> None:
        self.doc.add_paragraph(text, style=FRONT_STYLE)

    def _divider(self) -> None:
        self._centered(self.p["divider_text"], 12, bold=False, space_before=6, space_after=6)

    def _page_break(self) -> None:
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    def _caption(self, label: str, number: int, text: str, container=None) -> None:
        container = container or self.doc
        para = container.add_paragraph(style=CAPTION_STYLE)
        size = self.p["caption"]["size_pt"]
        run = para.add_run(f"{label} ")
        run.bold = True
        ooxml.set_run_font(run, self.font, size)
        ooxml.add_field(para, f"SEQ {label} \\* ARABIC", str(number), font=self.font, size_pt=size, bold=True)
        sep = para.add_run(": " if text else "")
        sep.bold = True
        ooxml.set_run_font(sep, self.font, size)
        if text:
            self._inline(para, text, size_pt=size)

    # ------------------------------------------------------------------ cover

    def _cover(self, spec: ReportSpec) -> None:
        m = spec.meta
        labels = self.p["labels"]
        if m.university:
            self._centered(m.university.upper(), 14, space_after=4)
        if m.school:
            self._centered(m.school.upper(), 14, space_after=12)
        logo = self._resolve(m.logo_path) if m.logo_path else None
        if logo and logo.exists():
            para = self.doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = Pt(12)
            para.add_run().add_picture(str(logo), width=Cm(5))
            gap = 36
        else:
            gap = 120
        self._centered(m.report_type.upper(), 24, space_before=gap, space_after=0)
        if m.subject:
            self._centered(m.subject.upper(), 24, space_after=18)
        if m.topic:
            para = self._centered("", 16, space_after=24)
            run = para.add_run(f"{labels['topic']} ")
            run.bold = True
            ooxml.set_run_font(run, self.font, 16)
            run = para.add_run(m.topic.upper())
            run.bold = True
            ooxml.set_run_font(run, self.font, 16)

        rows: list[tuple[str, list[str]]] = []
        if m.instructors:
            rows.append((labels["instructors"], m.instructors))
        if m.members:
            rows.append((labels["members"], [
                f"{mem.name} – {mem.student_id}" if mem.student_id else mem.name for mem in m.members
            ]))
        if rows:
            table = self.doc.add_table(rows=len(rows), cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            ooxml.table_no_borders(table)
            widths = (Cm(6.5), Cm(8.5))
            for r, (label, values) in enumerate(rows):
                left, right = table.rows[r].cells
                left.width, right.width = widths
                lp = left.paragraphs[0]
                lp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                lp.paragraph_format.space_after = Pt(12)
                run = lp.add_run(label)
                run.bold = True
                ooxml.set_run_font(run, self.font, 14)
                for i, value in enumerate(values):
                    vp = right.paragraphs[0] if i == 0 else right.add_paragraph()
                    vp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    vp.paragraph_format.space_after = Pt(12 if i == len(values) - 1 else 0)
                    ooxml.set_run_font(vp.add_run(value), self.font, 14)

        place = " – ".join(x for x in (m.city, m.year) if x)
        if place:
            # Đẩy dòng "TP. ... – năm" xuống gần đáy trang bìa
            lines_used = len(m.instructors) + len(m.members) + (2 if m.topic else 0)
            if logo and logo.exists():
                lines_used += 8
            self._centered(place, 14, space_before=max(36, 300 - 17 * lines_used))

    def _resolve(self, path: str) -> Path:
        candidate = Path(path)
        return candidate if candidate.is_absolute() else self.base_dir / candidate

    # ------------------------------------------------------------------ header / footer

    def _header_footer(self, section, spec: ReportSpec, with_header: bool) -> None:
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        header_cfg, footer_cfg = self.p["header"], self.p["footer"]
        values = spec.meta.model_dump()

        hp = section.header.paragraphs[0]
        # Style "Header" mặc định có tab giữa 8.25cm làm lệch chữ bên phải -> dùng Normal
        hp.style = self.doc.styles["Normal"]
        hp.paragraph_format.space_after = Pt(0)
        if with_header:
            left = header_cfg["left"].format(**values)
            right = header_cfg["right"].format(**values)
            hp.paragraph_format.tab_stops.add_tab_stop(Cm(self.text_width_cm), WD_TAB_ALIGNMENT.RIGHT)
            ooxml.set_run_font(hp.add_run(f"{left}\t{right}"), self.font, header_cfg["size_pt"])
            if header_cfg.get("border_bottom"):
                ooxml.paragraph_border(hp, {"bottom": {"sz": "6", "color": "000000"}})

        fp = section.footer.paragraphs[0]
        fp.style = self.doc.styles["Normal"]
        fp.paragraph_format.space_after = Pt(0)
        fp.alignment = ALIGN[footer_cfg.get("page_number", "center")]
        ooxml.add_field(fp, "PAGE", "1", font=self.font, size_pt=footer_cfg["size_pt"])
        if footer_cfg.get("border_top"):
            ooxml.paragraph_border(fp, {"top": {"sz": "6", "color": "000000"}})

    # ------------------------------------------------------------------ TOC

    def _toc_entries(self, spec: ReportSpec) -> list[tuple[int, str]]:
        entries: list[tuple[int, str]] = []
        labels = self.p["labels"]
        if spec.preface:
            entries.append((1, labels["preface"]))
        if spec.include_toc:
            entries.append((1, labels["toc"]))
        numberer = HeadingNumberer()
        for block in spec.body:
            if isinstance(block, HeadingBlock):
                entries.append((block.level, f"{numberer.next(block.level)} {block.text}"))
        if spec.references:
            title = labels["references"]
            if self.p.get("references_numbered"):
                title = f"{numberer.next(1)} {title}"
            entries.append((1, title))
        return entries

    def _toc(self, entries: list[tuple[int, str]]) -> None:
        """Field TOC có sẵn kết quả đệm (không số trang). postprocess.update_toc() hoặc
        Word (F9) sẽ điền số trang thật."""
        first = self.doc.add_paragraph(style="toc 1")

        # begin + instr + separate nằm trong đoạn đầu tiên, end nằm trong đoạn cuối
        def fld(paragraph, kind: str):
            run = paragraph.add_run()
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
            run._element.append(el)

        fld(first, "begin")
        instr = first.add_run()
        instr_el = OxmlElement("w:instrText")
        instr_el.set(qn("xml:space"), "preserve")
        instr_el.text = ' TOC \\o "1-3" \\h \\z \\u '
        instr._element.append(instr_el)
        fld(first, "separate")
        paragraphs = []
        for i, (level, text) in enumerate(entries):
            para = first if i == 0 else self.doc.add_paragraph(style=f"toc {level}")
            if i == 0 and level != 1:
                para.style = self.doc.styles[f"toc {level}"]
            ooxml.set_run_font(para.add_run(f"{text}\t"), self.font, self.body_size)
            paragraphs.append(para)
        if not paragraphs:
            paragraphs.append(first)
        fld(paragraphs[-1], "end")

    # ------------------------------------------------------------------ body

    def _body(self, spec: ReportSpec) -> None:
        numberer = HeadingNumberer()
        chapter_break = self.p.get("chapter_page_break", True)
        seen_chapter = False
        for block in spec.body:
            if isinstance(block, HeadingBlock):
                if block.level == 1 and seen_chapter and chapter_break:
                    self._page_break()
                if block.level == 1:
                    seen_chapter = True
                label = numberer.next(block.level)
                self.doc.add_paragraph(f"{label} {block.text}", style=f"Heading {block.level}")
            elif isinstance(block, ParagraphBlock):
                self._paragraph(block.text, block.align)
            elif isinstance(block, ListBlock):
                self._list(block)
            elif isinstance(block, TableBlock):
                self._table(block)
            elif isinstance(block, ImageBlock):
                self._image(block)
            elif isinstance(block, FigurePlaceholderBlock):
                self._placeholder(block)
            elif isinstance(block, CodeBlock):
                self._code(block)
            elif isinstance(block, NoteBlock):
                self._note(block)
            elif isinstance(block, PageBreakBlock):
                self._page_break()
            elif isinstance(block, DividerBlock):
                self._divider()

        if spec.references:
            if chapter_break and seen_chapter:
                self._page_break()
            title = self.p["labels"]["references"]
            if self.p.get("references_numbered"):
                self.doc.add_paragraph(f"{numberer.next(1)} {title}", style="Heading 1")
            else:
                self._front_heading(title)
            num_id = self.numbering.new_list("number")
            for ref in spec.references:
                para = self.doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                self.numbering.apply(para, num_id, 0)
                self._inline(para, ref.text)
                if ref.url:
                    if ref.text:
                        ooxml.set_run_font(para.add_run(" "), self.font, self.body_size)
                    ooxml.add_hyperlink(para, ref.url, ref.url, self.font, self.body_size, self.p["hyperlink_color"])

    def _list(self, block: ListBlock, container=None) -> None:
        container = container or self.doc
        num_id = self.numbering.new_list(block.style)
        for item in block.items:
            para = container.add_paragraph()
            para.alignment = ALIGN[self.p["body"]["align"]]
            para.paragraph_format.space_after = Pt(2)
            para.paragraph_format.first_line_indent = None
            self.numbering.apply(para, num_id, block.level)
            self._inline(para, item)

    def _column_widths(self, block: TableBlock) -> list[float]:
        n = len(block.columns)
        weights = list(block.col_widths[:n]) if len(block.col_widths) >= n else []
        if not weights:
            for c in range(n):
                cells = [block.columns[c]] + [row[c] if c < len(row) else "" for row in block.rows]
                longest = max(max((len(line) for line in cell.split("\n")), default=0) for cell in cells)
                weights.append(min(max(longest, 4), 40))
        total = sum(weights) or 1
        return [self.text_width_cm * w / total for w in weights]

    def _table(self, block: TableBlock) -> None:
        cfg = self.p["table"]
        size = cfg["font_size_pt"]
        self.table_count += 1
        if block.caption:
            self._caption(self.p["caption"]["table_label"], self.table_count, block.caption)
            self.doc.paragraphs[-1].paragraph_format.keep_with_next = True
        n = len(block.columns)
        table = self.doc.add_table(rows=1 + len(block.rows), cols=n)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        ooxml.table_borders(table, cfg["border_color"])
        widths = self._column_widths(block)

        short_col = []
        for c in range(n):
            values = [row[c] if c < len(row) else "" for row in block.rows]
            short_col.append(all(len(v) <= 25 and "\n" not in v for v in values))

        for c, title in enumerate(block.columns):
            cell = table.rows[0].cells[c]
            ooxml.cell_shading(cell, cfg["header_fill"])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.first_line_indent = None
            self._inline(para, title, size_pt=size, bold=True)
        ooxml.repeat_header_row(table.rows[0])

        for r, row in enumerate(block.rows, start=1):
            ooxml.cant_split_row(table.rows[r])
            for c in range(n):
                value = row[c] if c < len(row) else ""
                cell = table.rows[r].cells[c]
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for i, line in enumerate(value.split("\n")):
                    para = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
                    para.paragraph_format.space_after = Pt(0)
                    para.paragraph_format.first_line_indent = None
                    if line.startswith("- "):
                        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        para.paragraph_format.left_indent = Cm(0.4)
                        para.paragraph_format.first_line_indent = Cm(-0.4)
                        self._inline(para, "- " + line[2:], size_pt=size)
                    else:
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER if short_col[c] else WD_ALIGN_PARAGRAPH.LEFT
                        self._inline(para, line, size_pt=size)

        for row in table.rows:
            for c, cell in enumerate(row.cells):
                cell.width = Cm(widths[c])
        spacer = self.doc.add_paragraph()
        spacer.paragraph_format.space_after = Pt(4)

    def _image(self, block: ImageBlock) -> None:
        path = self._resolve(block.path)
        if not path.exists():
            self._placeholder(FigurePlaceholderBlock(type="figure_placeholder", caption=block.caption, description=f"Không tìm thấy ảnh: {block.path}"))
            return
        para = self.doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.keep_with_next = True
        para.paragraph_format.first_line_indent = None
        width = min(block.width_cm, self.text_width_cm)
        para.add_run().add_picture(str(path), width=Cm(width))
        self.figure_count += 1
        self._caption(self.p["caption"]["figure_label"], self.figure_count, block.caption)

    def _placeholder(self, block: FigurePlaceholderBlock) -> None:
        para = self.doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.keep_with_next = True
        para.paragraph_format.first_line_indent = None
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.left_indent = Cm(2)
        para.paragraph_format.right_indent = Cm(2)
        dashed = {"val": "dashed", "sz": "8", "color": "7F7F7F", "space": "18"}
        ooxml.paragraph_border(para, {side: dashed for side in ("top", "left", "bottom", "right")})
        ooxml.paragraph_shading(para, "F7F7F7")
        run = para.add_run("[ CHÈN HÌNH ]")
        run.bold = True
        ooxml.set_run_font(run, self.font, 12)
        if block.description:
            run.add_break()
            desc = para.add_run(block.description)
            desc.italic = True
            ooxml.set_run_font(desc, self.font, 12)
        self.figure_count += 1
        self._caption(self.p["caption"]["figure_label"], self.figure_count, block.caption)

    def _code(self, block: CodeBlock) -> None:
        if block.caption:
            cap = self.doc.add_paragraph(style=CAPTION_STYLE)
            cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
            cap.paragraph_format.keep_with_next = True
            cap.paragraph_format.space_after = Pt(2)
            self._inline(cap, block.caption, size_pt=self.p["caption"]["size_pt"])
        cfg = self.p["code"]
        para = self.doc.add_paragraph(style=CODE_STYLE)
        lines = block.code.rstrip("\n").split("\n")
        for i, line in enumerate(lines):
            run = para.add_run(line)
            ooxml.set_run_font(run, self.p["fonts"]["code"], cfg["size_pt"])
            if i < len(lines) - 1:
                run.add_break()
        ooxml.paragraph_border(para, {"left": {"sz": "18", "color": cfg["border_color"], "space": "6"}})
        ooxml.paragraph_shading(para, cfg["fill"])

    def _note(self, block: NoteBlock) -> None:
        border, fill = NOTE_COLORS[block.kind]
        para = self.doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.first_line_indent = None
        ooxml.paragraph_border(para, {"left": {"sz": "18", "color": border, "space": "6"}})
        ooxml.paragraph_shading(para, fill)
        label = para.add_run(NOTE_LABELS[block.kind])
        label.bold = True
        ooxml.set_run_font(label, self.font, self.body_size)
        self._inline(para, block.text)


def render_spec(spec: ReportSpec, output: str | Path, base_dir: str | Path = ".") -> Path:
    return DocxRenderer(spec.profile, base_dir=base_dir).save(spec, output)
