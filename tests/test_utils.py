"""Tests for utils.py"""
import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from docx_editor.utils import (
    qn, half_points_to_pt, pt_to_half_points,
    DocxError, ChapterNotFoundError, MermaidRenderError,
    CommentsReadError, Win32ComError
)


class TestQn:
    def test_simple_tag(self):
        result = qn('w:p')
        assert result == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'

    def test_different_prefix(self):
        result = qn('r:id')
        assert '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

    def test_w14_prefix(self):
        result = qn('w14:paraId')
        assert 'paraId' in result

    def test_invalid_prefix(self):
        with pytest.raises(KeyError):
            qn('unknown:tag')


class TestUnitConversion:
    def test_half_points_to_pt(self):
        assert half_points_to_pt(21) == 10.5
        assert half_points_to_pt(40) == 20.0
        assert half_points_to_pt(0) == 0.0

    def test_half_points_to_pt_none(self):
        assert half_points_to_pt(None) is None

    def test_pt_to_half_points(self):
        assert pt_to_half_points(10.5) == 21
        assert pt_to_half_points(20) == 40
        assert pt_to_half_points(0) == 0


class TestExceptions:
    def test_docx_error_hierarchy(self):
        assert issubclass(ChapterNotFoundError, DocxError)
        assert issubclass(MermaidRenderError, DocxError)
        assert issubclass(CommentsReadError, DocxError)
        assert issubclass(Win32ComError, DocxError)

    def test_chapter_not_found(self):
        e = ChapterNotFoundError("Chapter 5.1 not found")
        assert "Chapter 5.1 not found" in str(e)

    def test_mermaid_render_error(self):
        e = MermaidRenderError("mmdc not available")
        assert "mmdc not available" in str(e)
