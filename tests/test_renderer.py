from docx import Document

from wordreport.inspector import lint_docx, outline
from wordreport.numbering import HeadingNumberer, to_roman
from wordreport.renderer import DocxRenderer
from wordreport.spec import ReportSpec
from wordreport.style_profile import list_profiles, load_profile
from wordreport.watermark_remover import find_watermarks


def test_roman_and_heading_numbers():
    assert [to_roman(n) for n in (1, 4, 9, 14, 40)] == ["I", "IV", "IX", "XIV", "XL"]
    n = HeadingNumberer()
    assert [n.next(1), n.next(2), n.next(3), n.next(3), n.next(2), n.next(1), n.next(2)] == [
        "I.", "1.", "1.1.", "1.2.", "2.", "II.", "1."]


def test_profiles_inherit():
    assert {"hcmus-clc", "nd30-a4"} <= set(list_profiles())
    a4 = load_profile("nd30-a4")
    assert a4["page"]["size"] == "a4" and a4["page"]["margin_left_cm"] == 3.0
    assert a4["fonts"]["body"] == "Times New Roman"  # kế thừa từ hcmus-clc
    assert a4["headings"]["level1"]["bold"] is True


def test_render_example(tmp_path, example_spec):
    out = DocxRenderer(example_spec.profile, base_dir=tmp_path).save(example_spec, tmp_path / "r.docx")
    doc = Document(str(out))

    assert len(doc.sections) == 3
    header_text = doc.sections[2].header.paragraphs[0].text
    assert header_text == "Khoa Công nghệ Thông tin\t22CLC03"
    assert doc.sections[1].header.paragraphs[0].text == ""

    heads = [h["text"] for h in outline(out)]
    assert "I. Giới thiệu chung" in heads
    assert "2. Bài 2 – Web, DNS, DHCP Relay và RIP" in heads
    assert "2.2. Bài làm" in heads
    assert heads[-1] == "III. Tài liệu tham khảo"

    xml = doc.element.body.xml
    assert 'TOC \\o "1-3"' in xml
    assert xml.count("SEQ Hình") == 4 and xml.count("SEQ Bảng") == 3
    assert "PAGE" in doc.sections[2].footer._element.xml

    captions = [p.text for p in doc.paragraphs if p.style.name == "Caption" and p.text.startswith(("Hình", "Bảng"))]
    assert "Bảng 1: Danh sách thành viên nhóm" in captions
    assert "Hình 4: Truy cập www.congtyhanhoa.vn từ PC3 và PC5" in captions

    issues = [i for i in lint_docx(out) if i.severity != "info"]
    assert issues == []

    # watermark-remover chạy mặc định: không còn nhãn trình tạo/AI, tác giả là thành viên nhóm
    assert find_watermarks(out) == []
    props = doc.core_properties
    assert props.author == "Nguyễn Minh An, Lê Thu Bình" and props.comments == ""
    assert props.created.year >= 2024


def test_inline_markup_and_missing_image(tmp_path):
    spec = ReportSpec.model_validate({
        "include_toc": False,
        "body": [
            {"type": "heading", "level": 1, "text": "A"},
            {"type": "paragraph", "text": "Chữ **đậm**, *nghiêng*, `code` và [link](https://example.com)."},
            {"type": "image", "path": "khong-ton-tai.png", "caption": "Ảnh thiếu"},
        ],
    })
    out = DocxRenderer("hcmus-clc", base_dir=tmp_path).save(spec, tmp_path / "i.docx")
    doc = Document(str(out))
    para = next(p for p in doc.paragraphs if p.text.startswith("Chữ"))
    runs = {r.text: r for r in para.runs}
    assert runs["đậm"].bold and runs["nghiêng"].italic
    assert runs["code"].font.name == "Consolas"
    assert "w:hyperlink" in para._p.xml
    assert any("CHÈN HÌNH" in p.text and "khong-ton-tai.png" in p.text for p in doc.paragraphs)


def test_lists_restart_numbering(tmp_path):
    spec = ReportSpec.model_validate({"include_toc": False, "body": [
        {"type": "list", "style": "number", "items": ["a", "b"]},
        {"type": "list", "style": "number", "items": ["c"]},
    ]})
    out = DocxRenderer().save(spec, tmp_path / "l.docx")
    doc = Document(str(out))
    num_ids = [p._p.pPr.numPr.numId.val for p in doc.paragraphs if p._p.pPr is not None and p._p.pPr.numPr is not None]
    assert len(num_ids) == 3 and num_ids[0] == num_ids[1] != num_ids[2]
