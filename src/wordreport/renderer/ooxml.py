"""Các thao tác OOXML cấp thấp mà python-docx chưa hỗ trợ trực tiếp.

Gồm: field (TOC, PAGE, SEQ), viền/tô nền đoạn & ô, khung trang, danh sách đánh số
khởi động lại được, hyperlink, cờ cập nhật field khi mở file.
"""

from __future__ import annotations

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph


def _el(tag: str, **attrs: str) -> OxmlElement:
    el = OxmlElement(tag)
    for key, value in attrs.items():
        el.set(qn(key), str(value))
    return el


# ---------------------------------------------------------------------------
# Font
# ---------------------------------------------------------------------------

def set_run_font(run, name: str, size_pt: float | None = None) -> None:
    """Đặt font cho mọi script (ascii, hAnsi, eastAsia, cs) để tiếng Việt không bị đổi font."""
    run.font.name = name
    if size_pt:
        run.font.size = Pt(size_pt)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = _el("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)


def set_style_font(style, name: str, size_pt: float | None = None) -> None:
    style.font.name = name
    if size_pt:
        style.font.size = Pt(size_pt)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = _el("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)
    # Xoá theme font để font tường minh có hiệu lực
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]


# ---------------------------------------------------------------------------
# Fields
# ---------------------------------------------------------------------------

def add_field(paragraph: Paragraph, instr: str, cached_result: str = "", font: str | None = None,
              size_pt: float | None = None, bold: bool = False) -> None:
    """Chèn complex field: begin - instrText - separate - kết quả đệm - end."""

    def run_with(child) -> None:
        run = paragraph.add_run()
        if font:
            set_run_font(run, font, size_pt)
        run.bold = bold
        run._element.append(child)

    run_with(_el("w:fldChar", **{"w:fldCharType": "begin"}))
    instr_el = _el("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = f" {instr} "
    run_with(instr_el)
    run_with(_el("w:fldChar", **{"w:fldCharType": "separate"}))
    result = paragraph.add_run(cached_result)
    result.bold = bold
    if font:
        set_run_font(result, font, size_pt)
    run_with(_el("w:fldChar", **{"w:fldCharType": "end"}))


# Các phần tử của w:settings đứng SAU updateFields theo schema CT_Settings
_SETTINGS_AFTER_UPDATE_FIELDS = [
    "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids", "mathPr",
    "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats",
    "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType",
    "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator",
]


def enable_update_fields_on_open(document) -> None:
    """Word sẽ hỏi cập nhật field (mục lục, số trang) khi mở file."""
    settings = document.settings.element
    zoom = settings.find(qn("w:zoom"))
    if zoom is not None and zoom.get(qn("w:percent")) is None:
        zoom.set(qn("w:percent"), "100")  # template mặc định của python-docx thiếu thuộc tính bắt buộc
    if settings.find(qn("w:updateFields")) is not None:
        return
    element = _el("w:updateFields", **{"w:val": "true"})
    after = {qn(f"w:{name}") for name in _SETTINGS_AFTER_UPDATE_FIELDS}
    for child in settings:
        if child.tag in after:
            child.addprevious(element)
            return
    settings.append(element)


# ---------------------------------------------------------------------------
# Borders & shading
# ---------------------------------------------------------------------------

def paragraph_border(paragraph: Paragraph, sides: dict[str, dict[str, str]]) -> None:
    """sides = {"bottom": {"val": "single", "sz": "6", "color": "000000", "space": "1"}, ...}"""
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = ppr.find(qn("w:pBdr"))
    if pbdr is None:
        pbdr = _el("w:pBdr")
        ppr.append(pbdr)
    for side in ("top", "left", "bottom", "right"):
        if side in sides:
            spec = {"val": "single", "sz": "6", "space": "1", "color": "auto", **sides[side]}
            pbdr.append(_el(f"w:{side}", **{f"w:{k}": v for k, v in spec.items()}))
    _reorder_ppr(ppr)


def paragraph_shading(paragraph: Paragraph, fill: str) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    ppr.append(_el("w:shd", **{"w:val": "clear", "w:color": "auto", "w:fill": fill}))
    _reorder_ppr(ppr)


# Thứ tự phần tử con của w:pPr theo schema (rút gọn các phần tử ta dùng)
_PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
    "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
    "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
    "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
]


def _reorder_ppr(ppr) -> None:
    rank = {qn(f"w:{n}"): i for i, n in enumerate(_PPR_ORDER)}
    children = list(ppr)
    children.sort(key=lambda c: rank.get(c.tag, len(rank)))
    for child in children:
        ppr.remove(child)
        ppr.append(child)


def cell_shading(cell, fill: str) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    tcpr.append(_el("w:shd", **{"w:val": "clear", "w:color": "auto", "w:fill": fill}))


_TBLPR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
    "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
    "tblCaption", "tblDescription", "tblPrChange",
]


def _set_tbl_borders(table, spec: dict[str, str]) -> None:
    tblpr = table._tbl.tblPr
    old = tblpr.find(qn("w:tblBorders"))
    if old is not None:
        tblpr.remove(old)
    borders = _el("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(_el(f"w:{side}", **spec))
    tblpr.append(borders)
    reorder_tblpr(tblpr)


def reorder_tblpr(tblpr) -> None:
    rank = {qn(f"w:{n}"): i for i, n in enumerate(_TBLPR_ORDER)}
    children = sorted(tblpr, key=lambda c: rank.get(c.tag, len(rank)))
    for child in children:
        tblpr.remove(child)
        tblpr.append(child)


def table_borders(table, color: str = "000000", size: str = "4") -> None:
    _set_tbl_borders(table, {"w:val": "single", "w:sz": size, "w:space": "0", "w:color": color})


def table_no_borders(table) -> None:
    _set_tbl_borders(table, {"w:val": "nil"})


def repeat_header_row(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    trpr.append(_el("w:tblHeader", **{"w:val": "true"}))


def cant_split_row(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    trpr.append(_el("w:cantSplit", **{"w:val": "true"}))


def page_border(section, color: str = "000000", size: str = "12") -> None:
    sectpr = section._sectPr
    borders = _el("w:pgBorders", **{"w:offsetFrom": "page"})
    for side in ("top", "left", "bottom", "right"):
        borders.append(_el(f"w:{side}", **{"w:val": "single", "w:sz": size, "w:space": "24", "w:color": color}))
    # pgBorders phải đứng sau pgMar/paperSrc và trước lnNumType/pgNumType/cols...
    anchor = sectpr.find(qn("w:pgMar"))
    if anchor is not None:
        anchor.addnext(borders)
    else:
        sectpr.append(borders)


def page_number_start(section, start: int) -> None:
    sectpr = section._sectPr
    pg = sectpr.find(qn("w:pgNumType"))
    if pg is None:
        pg = _el("w:pgNumType")
        cols = sectpr.find(qn("w:cols"))
        if cols is not None:
            cols.addprevious(pg)
        else:
            sectpr.append(pg)
    pg.set(qn("w:start"), str(start))


# ---------------------------------------------------------------------------
# Hyperlink
# ---------------------------------------------------------------------------

def add_hyperlink(paragraph: Paragraph, text: str, url: str, font: str, size_pt: float, color: str) -> None:
    r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = _el("w:hyperlink", **{"r:id": r_id})
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rpr.append(_el("w:rFonts", **{"w:ascii": font, "w:hAnsi": font, "w:eastAsia": font, "w:cs": font}))
    rpr.append(_el("w:color", **{"w:val": color}))
    rpr.append(_el("w:sz", **{"w:val": str(int(size_pt * 2))}))
    rpr.append(_el("w:u", **{"w:val": "single"}))
    run.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    run.append(t)
    link.append(run)
    paragraph._p.append(link)


# ---------------------------------------------------------------------------
# Numbering (danh sách)
# ---------------------------------------------------------------------------

_LIST_FORMATS = {
    # style: (numFmt, lvlText cho từng cấp)
    "bullet": ("bullet", ["•", "◦", "▪"]),
    "dash": ("bullet", ["-", "+", "•"]),
    "number": ("decimal", ["%1.", "%2)", "%3."]),
    "roman": ("lowerRoman", ["%1.", "%2.", "%3."]),
    "alpha": ("lowerLetter", ["%1.", "%2.", "%3."]),
}


class NumberingManager:
    """Tạo abstractNum cho mỗi kiểu danh sách và một w:num mới cho mỗi danh sách
    (nhờ vậy danh sách đánh số luôn bắt đầu lại từ 1)."""

    def __init__(self, document, font: str) -> None:
        self.numbering = document.part.numbering_part.element
        self.font = font
        self.abstract_ids: dict[str, int] = {}
        existing_abs = [int(e.get(qn("w:abstractNumId"))) for e in self.numbering.findall(qn("w:abstractNum"))]
        existing_num = [int(e.get(qn("w:numId"))) for e in self.numbering.findall(qn("w:num"))]
        self._next_abs = max(existing_abs, default=0) + 100
        self._next_num = max(existing_num, default=0) + 100

    def _abstract(self, style: str) -> int:
        if style in self.abstract_ids:
            return self.abstract_ids[style]
        fmt, texts = _LIST_FORMATS[style]
        abs_id = self._next_abs
        self._next_abs += 1
        abstract = _el("w:abstractNum", **{"w:abstractNumId": str(abs_id)})
        abstract.append(_el("w:multiLevelType", **{"w:val": "hybridMultilevel"}))
        for lvl, text in enumerate(texts):
            lvl_el = _el("w:lvl", **{"w:ilvl": str(lvl)})
            lvl_el.append(_el("w:start", **{"w:val": "1"}))
            lvl_el.append(_el("w:numFmt", **{"w:val": fmt}))
            lvl_el.append(_el("w:lvlText", **{"w:val": text}))
            lvl_el.append(_el("w:lvlJc", **{"w:val": "left"}))
            ppr = _el("w:pPr")
            left = 720 + 360 * lvl
            ppr.append(_el("w:ind", **{"w:left": str(left), "w:hanging": "360"}))
            lvl_el.append(ppr)
            if fmt == "bullet":
                rpr = _el("w:rPr")
                rpr.append(_el("w:rFonts", **{"w:ascii": self.font, "w:hAnsi": self.font, "w:hint": "default"}))
                lvl_el.append(rpr)
            abstract.append(lvl_el)
        # abstractNum phải đứng trước mọi w:num
        first_num = self.numbering.find(qn("w:num"))
        if first_num is not None:
            first_num.addprevious(abstract)
        else:
            self.numbering.append(abstract)
        self.abstract_ids[style] = abs_id
        return abs_id

    def new_list(self, style: str) -> int:
        abs_id = self._abstract(style)
        num_id = self._next_num
        self._next_num += 1
        num = _el("w:num", **{"w:numId": str(num_id)})
        num.append(_el("w:abstractNumId", **{"w:val": str(abs_id)}))
        override = _el("w:lvlOverride", **{"w:ilvl": "0"})
        override.append(_el("w:startOverride", **{"w:val": "1"}))
        num.append(override)
        self.numbering.append(num)
        return num_id

    @staticmethod
    def apply(paragraph: Paragraph, num_id: int, level: int) -> None:
        ppr = paragraph._p.get_or_add_pPr()
        numpr = _el("w:numPr")
        numpr.append(_el("w:ilvl", **{"w:val": str(level)}))
        numpr.append(_el("w:numId", **{"w:val": str(num_id)}))
        ppr.append(numpr)
        _reorder_ppr(ppr)


def hex_color(value: str) -> RGBColor:
    return RGBColor.from_string(value.upper())


def set_outline_level(style, level: int) -> None:
    ppr = style.element.get_or_add_pPr()
    existing = ppr.find(qn("w:outlineLvl"))
    if existing is None:
        existing = _el("w:outlineLvl")
        ppr.append(existing)
    existing.set(qn("w:val"), str(level))
    _reorder_ppr(ppr)
