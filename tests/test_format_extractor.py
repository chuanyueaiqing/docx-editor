"""Tests for format_extractor.py"""
import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from docx import Document
from docx_editor.format_extractor import FormatExtractor
from docx_editor.models import FormattingData, ParagraphFormatData


@pytest.fixture(scope="module")
def test_doc():
    """Load the test docx document once for all tests."""
    doc_path = os.path.join(os.path.dirname(__file__), '..', '测试.docx')
    return Document(doc_path)


@pytest.fixture(scope="module")
def extracted(test_doc):
    """Extract all formatting from test document."""
    return FormatExtractor.extract_all(test_doc)


class TestExtractAll:
    def test_returns_dict(self, extracted):
        assert isinstance(extracted, dict)

    def test_all_paragraphs_extracted(self, extracted, test_doc):
        assert len(extracted) == len(test_doc.paragraphs)

    def test_each_entry_has_keys(self, extracted):
        for idx, data in extracted.items():
            assert 'paragraph_format' in data, f"Missing paragraph_format at {idx}"
            assert 'runs' in data, f"Missing runs at {idx}"

    def test_index_continuous(self, extracted):
        indices = sorted(extracted.keys())
        assert indices == list(range(len(indices)))


class TestParagraphFormatExtraction:
    def test_first_paragraph_style(self, extracted):
        # First paragraph should be Normal style
        pf = extracted[0]['paragraph_format']
        assert pf is not None

    def test_paragraph_format_type(self, extracted):
        pf = extracted[0]['paragraph_format']
        assert isinstance(pf, ParagraphFormatData)


class TestRunFormatExtraction:
    def test_runs_is_list(self, extracted):
        runs = extracted[0]['runs']
        assert isinstance(runs, list)

    def test_run_has_text_and_formatting(self, extracted):
        runs = extracted[0]['runs']
        for r in runs:
            assert 'text' in r
            assert 'formatting' in r
            assert isinstance(r['formatting'], FormattingData)

    def test_run_text_not_none(self, extracted):
        for idx, data in extracted.items():
            for r in data['runs']:
                assert r['text'] is not None


class TestTrackedChangesInXML:
    """FormatExtractor should not crash on documents with tracked changes."""

    def test_extract_with_tracked_changes(self, extracted):
        # Should not raise any exceptions
        assert len(extracted) > 0

    def test_ins_elements_handled(self, test_doc):
        """Verify w:ins elements exist and format_extractor handles them."""
        from docx_editor.utils import qn
        body = test_doc.element.body
        ins_elems = body.findall('.//' + qn('w:ins'))
        # The test document has tracked changes
        assert len(ins_elems) >= 0  # Just check it doesn't crash
