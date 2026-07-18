"""Tests for models.py"""
import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from docx_editor.models import (
    FormattingData, ParagraphFormatData, ParagraphData,
    ChapterNode, CommentData, RevisionData, MarkdownElement, FormatStore,
    MarkdownElementType
)


class TestFormattingData:
    def test_default_creation(self):
        f = FormattingData()
        assert f.font_name is None
        assert f.size is None
        assert f.bold is None
        assert f.italic is None

    def test_with_values(self):
        f = FormattingData(font_name="Arial", size=21, bold=True, italic=False)
        assert f.font_name == "Arial"
        assert f.size == 21
        assert f.bold is True
        assert f.italic is False

    def test_east_asian_font(self):
        f = FormattingData(font_name_east_asia="SimSun")
        assert f.font_name_east_asia == "SimSun"


class TestParagraphFormatData:
    def test_default_creation(self):
        p = ParagraphFormatData()
        assert p.heading_level is None
        assert p.alignment is None

    def test_heading_detection(self):
        p = ParagraphFormatData(style_name="Heading 1", heading_level=1)
        assert p.style_name == "Heading 1"
        assert p.heading_level == 1

    def test_with_indents(self):
        p = ParagraphFormatData(
            first_line_indent=420,
            left_indent=720,
            line_spacing=1.5
        )
        assert p.first_line_indent == 420
        assert p.line_spacing == 1.5


class TestParagraphData:
    def test_default_creation(self):
        p = ParagraphData(index=0, text="Hello")
        assert p.index == 0
        assert p.text == "Hello"
        assert p.formatting is None
        assert p.runs_data == []

    def test_with_formatting(self):
        fmt = ParagraphFormatData(style_name="Normal")
        p = ParagraphData(index=1, text="Test", formatting=fmt)
        assert p.formatting.style_name == "Normal"


class TestChapterNode:
    def test_create_node(self):
        node = ChapterNode(
            heading_text="3.1 系统架构",
            heading_level=2,
            number_tuple=(3, 1),
            heading_paragraph_index=5
        )
        assert node.heading_text == "3.1 系统架构"
        assert node.number_tuple == (3, 1)

    def test_to_string(self):
        node = ChapterNode(
            heading_text="3.1.1 模块设计",
            heading_level=3,
            number_tuple=(3, 1, 1),
            heading_paragraph_index=10
        )
        assert node.to_string() == "3.1.1"

    def test_to_string_no_number(self):
        node = ChapterNode(
            heading_text="简介",
            heading_level=1,
            number_tuple=None,
            heading_paragraph_index=0
        )
        assert node.to_string() == "简介"

    def test_child_relationship(self):
        parent = ChapterNode(
            heading_text="3 设计",
            heading_level=1, number_tuple=(3,), heading_paragraph_index=3
        )
        child = ChapterNode(
            heading_text="3.1 架构",
            heading_level=2, number_tuple=(3, 1),
            heading_paragraph_index=7, parent=parent
        )
        parent.children.append(child)
        assert child.parent is parent
        assert len(parent.children) == 1
        assert parent.children[0].heading_text == "3.1 架构"

    def test_body_paragraph_indices(self):
        node = ChapterNode(
            heading_text="1 引言",
            heading_level=1, number_tuple=(1,), heading_paragraph_index=0,
            body_paragraph_indices=[1, 2, 3]
        )
        assert node.body_paragraph_indices == [1, 2, 3]


class TestCommentData:
    def test_create(self):
        c = CommentData(id="0", author="测试用户", date="2026-01-01", text="这是一个批注")
        assert c.author == "测试用户"
        assert c.text == "这是一个批注"

    def test_with_paragraph_index(self):
        c = CommentData(id="1", author="User", date="2026-01-01",
                        text="Comment", paragraph_index=5)
        assert c.paragraph_index == 5


class TestRevisionData:
    def test_insertion(self):
        r = RevisionData(
            rev_id="0", author="User", date="2026-01-01",
            type="insertion", text="新文字", paragraph_index=3
        )
        assert r.type == "insertion"
        assert r.text == "新文字"

    def test_deletion(self):
        r = RevisionData(
            rev_id="1", author="User", date="2026-01-01",
            type="deletion", text="旧文字", paragraph_index=5
        )
        assert r.type == "deletion"


class TestMarkdownElement:
    def test_paragraph(self):
        e = MarkdownElement(type=MarkdownElementType.PARAGRAPH, text="Hello")
        assert e.type == MarkdownElementType.PARAGRAPH
        assert e.text == "Hello"

    def test_heading(self):
        e = MarkdownElement(type=MarkdownElementType.HEADING, text="Title", level=1)
        assert e.level == 1

    def test_table(self):
        rows = [["A", "B"], ["1", "2"]]
        merge_map = [["", ">"], ["v", ""]]
        e = MarkdownElement(
            type=MarkdownElementType.TABLE,
            rows=rows, merge_map=merge_map
        )
        assert e.rows[0][0] == "A"
        assert e.merge_map[0][1] == ">"

    def test_mermaid(self):
        e = MarkdownElement(
            type=MarkdownElementType.MERMAID,
            text="graph TD; A-->B;"
        )
        assert e.type == MarkdownElementType.MERMAID

    def test_image(self):
        e = MarkdownElement(
            type=MarkdownElementType.IMAGE,
            image_path="diagram.png", alt_text="架构图"
        )
        assert e.image_path == "diagram.png"
        assert e.alt_text == "架构图"


class TestFormatStore:
    def test_create(self):
        store = FormatStore(docx_path="test.docx")
        assert store.docx_path == "test.docx"
        assert store.document is None
        assert store.paragraphs_data == []
        assert store.formats_json == {}

    def test_with_document(self):
        store = FormatStore(docx_path="test.docx", document="mock_doc")
        assert store.document == "mock_doc"
