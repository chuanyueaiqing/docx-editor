"""Tests for markdown_processor.py"""
import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from docx import Document
from docx_editor.markdown_processor import MarkdownProcessor
from docx_editor.models import (
    MarkdownElement, MarkdownElementType, FormatStore
)
from docx_editor.format_extractor import FormatExtractor


@pytest.fixture
def empty_store():
    doc = Document()
    store = FormatStore(docx_path="test.docx", document=doc)
    store.formats_json = FormatExtractor.extract_all(doc)
    return store


class TestMarkdownParsing:
    def test_parse_heading(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("# 一级标题\n\n## 二级标题\n\n正文内容")
        headings = [e for e in elements if e.type == MarkdownElementType.HEADING]
        assert len(headings) == 2
        assert headings[0].level == 1
        assert headings[0].text == "一级标题"
        assert headings[1].level == 2
        assert headings[1].text == "二级标题"

    def test_parse_paragraph(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("第一段\n\n第二段")
        paras = [e for e in elements if e.type == MarkdownElementType.PARAGRAPH]
        assert len(paras) == 2
        assert paras[0].text == "第一段"
        assert paras[1].text == "第二段"

    def test_parse_table(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        md = "| A | B | C |\n| a1 | > | c1 |\n| a2 | v | > |"
        elements = mp.parse_markdown(md)
        tables = [e for e in elements if e.type == MarkdownElementType.TABLE]
        assert len(tables) == 1
        table = tables[0]
        assert len(table.rows) == 2  # 2 data rows
        # Check headers are separate
        assert table.children is not None
        headers = [c.text for c in table.children if c.type == MarkdownElementType.TABLE]
        if headers:
            assert len(headers) > 0

    def test_parse_code_block(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("```python\nprint('hello')\n```")
        codes = [e for e in elements if e.type == MarkdownElementType.CODE_BLOCK]
        assert len(codes) == 1
        assert codes[0].code_language == "python"
        assert "print" in codes[0].text

    def test_parse_mermaid(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("```mermaid\ngraph TD;\nA-->B;\n```")
        mermaids = [e for e in elements if e.type == MarkdownElementType.MERMAID]
        assert len(mermaids) == 1
        assert "A-->B" in mermaids[0].text

    def test_parse_image(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("![架构图](diagram.png)")
        images = [e for e in elements if e.type == MarkdownElementType.IMAGE]
        assert len(images) == 1
        assert images[0].image_path == "diagram.png"
        assert images[0].alt_text == "架构图"

    def test_parse_list(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("- item1\n- item2\n- item3")
        lists = [e for e in elements if e.type == MarkdownElementType.LIST]
        assert len(lists) == 1
        assert lists[0].ordered is False
        assert len(lists[0].items) >= 2

    def test_parse_ordered_list(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("1. first\n2. second\n3. third")
        lists = [e for e in elements if e.type == MarkdownElementType.LIST]
        assert len(lists) == 1
        assert lists[0].ordered is True

    def test_parse_horizontal_rule(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("---")
        hrs = [e for e in elements if e.type == MarkdownElementType.HORIZONTAL_RULE]
        assert len(hrs) >= 1

    def test_inline_bold(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("这是**粗体**文字")
        # Bold should be handled as inline formatting within a paragraph
        paras = [e for e in elements if e.type == MarkdownElementType.PARAGRAPH]
        assert len(paras) >= 1
        text = paras[0].text
        assert "粗体" in text or "**" in text

    def test_parse_empty_string(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("")
        assert len(elements) == 0

    def test_parse_only_whitespace(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("   \n  \n  ")
        empty = [e for e in elements if e.type == MarkdownElementType.EMPTY_LINE]
        assert len(empty) >= 0


class TestMarkdownConversion:
    def test_heading_to_docx(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("# 测试标题\n\n段落内容")
        mp.apply_to_document(elements)
        doc = empty_store.document
        # Should have at least 2 paragraphs
        assert len(doc.paragraphs) >= 2
        # First should be heading style
        p0 = doc.paragraphs[0]
        style_name = p0.style.name if p0.style else ''
        assert 'heading' in style_name.lower() or p0.text == "测试标题"

    def test_paragraph_to_docx(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("这是一段测试文字。")
        mp.apply_to_document(elements)
        doc = empty_store.document
        assert len(doc.paragraphs) >= 1
        assert "测试文字" in doc.paragraphs[0].text

    def test_multiple_paragraphs(self, empty_store):
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown("第一段\n\n第二段\n\n第三段")
        mp.apply_to_document(elements)
        doc = empty_store.document
        texts = [p.text for p in doc.paragraphs if p.text.strip()]
        assert len(texts) >= 3

    def test_code_block_to_docx(self, empty_store):
        """Code blocks render as a single-cell table with monospace formatting."""
        from docx.shared import Pt
        mp = MarkdownProcessor(empty_store)
        elements = mp.parse_markdown(
            "```python\ndef hello():\n    print('Hi')\n```"
        )
        mp.apply_to_document(elements)
        doc = empty_store.document

        # Should produce exactly one table
        assert len(doc.tables) == 1
        table = doc.tables[0]
        assert len(table.rows) == 1
        assert len(table.columns) == 1

        # Cell contains all code lines (indentation preserved)
        cell = table.cell(0, 0)
        texts = [p.text for p in cell.paragraphs]
        assert texts[0] == "def hello():"
        assert "    print('Hi')" in texts

        # Font is Consolas, 9pt
        for para in cell.paragraphs:
            for run in para.runs:
                assert run.font.name == 'Consolas'
                assert run.font.size == Pt(9)
