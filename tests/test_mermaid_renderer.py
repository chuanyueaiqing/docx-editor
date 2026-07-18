"""Tests for mermaid_renderer.py"""
import pytest
import sys
import os
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from docx_editor.mermaid_renderer import MermaidRenderer, MermaidNotAvailableError


class TestMermaidRendererInit:
    def test_create_renderer(self):
        renderer = MermaidRenderer()
        assert renderer is not None
        assert renderer.mmdc_cmd is not None or renderer.mmdc_cmd == ''


class TestMermaidRendererCheck:
    def test_check_mmdc_available(self):
        renderer = MermaidRenderer()
        # Just check it doesn't crash - mmdc may or may not be installed
        available = renderer.is_available()
        assert isinstance(available, bool)

    def test_render_fails_without_mmdc(self):
        renderer = MermaidRenderer()
        if not renderer.is_available():
            with pytest.raises(MermaidNotAvailableError):
                renderer.render("graph TD; A-->B;")


class TestMermaidCodeHandling:
    def test_write_mermaid_file(self):
        renderer = MermaidRenderer()
        code = "graph TD;\n    A-->B;\n    B-->C;"
        # Test the _write_mermaid_file method
        filepath = renderer._write_mermaid_file(code)
        try:
            assert os.path.exists(filepath)
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            assert "graph TD" in content
            assert "A-->B" in content
        finally:
            if os.path.exists(filepath):
                os.unlink(filepath)

    def test_code_hash_consistency(self):
        renderer = MermaidRenderer()
        code1 = "graph TD; A-->B;"
        code2 = "graph TD; A-->B;"
        code3 = "graph TD; A-->C;"
        hash1 = renderer._code_hash(code1)
        hash2 = renderer._code_hash(code2)
        hash3 = renderer._code_hash(code3)
        assert hash1 == hash2  # Same code = same hash
        assert hash1 != hash3  # Different code = different hash
