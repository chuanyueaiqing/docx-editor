"""Tests for win32_ops.py

Note: Tests that start/stop Microsoft Word may crash under pytest due to
COM exception handling conflicts. Tests that require actually opening Word
are written as direct calls at module level when run as __main__.
"""
import sys
import os
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from docx_editor.win32_ops import Win32Ops


class TestWin32OpsBasic:
    def test_is_word_available(self):
        """Check if Word is available on this system."""
        available = Win32Ops.is_word_available()
        assert isinstance(available, bool)

    def test_is_word_available_returns_bool(self):
        result = Win32Ops.is_word_available()
        # Just verify it returns a bool without error
        assert result is True or result is False


if __name__ == '__main__':
    """Run comprehensive Win32 tests that require starting Word."""
    doc_path = os.path.join(os.path.dirname(__file__), '..', '测试.docx')

    print(f"Word available: {Win32Ops.is_word_available()}")

    if not Win32Ops.is_word_available():
        print("Microsoft Word not available, skipping tests")
        sys.exit(0)

    # Test context manager lifecycle
    print("\n=== Test: Context manager lifecycle ===")
    with Win32Ops() as ops:
        assert ops.word is not None
    print("PASS")

    # Test open document
    print("\n=== Test: Open document ===")
    with Win32Ops() as ops:
        ops.open_document(doc_path)
        assert ops.wd_doc is not None
    print("PASS")

    # Test track changes toggle
    print("\n=== Test: Track changes toggle ===")
    with Win32Ops() as ops:
        ops.open_document(doc_path)
        ops.enable_track_changes(True)
        ops.enable_track_changes(False)
    print("PASS")

    # Test read comments
    print("\n=== Test: Read comments ===")
    with Win32Ops() as ops:
        ops.open_document(doc_path)
        comments = ops.read_comments()
        assert isinstance(comments, list)
        print(f"  Found {len(comments)} comments")
    print("PASS")

    # Test read revisions
    print("\n=== Test: Read revisions ===")
    with Win32Ops() as ops:
        ops.open_document(doc_path)
        revisions = ops.read_revisions()
        assert isinstance(revisions, list)
        print(f"  Found {len(revisions)} revisions")
        for r in revisions:
            print(f"  - [{r.type}] {r.author}: {r.text[:60]}")
        assert len(revisions) > 0, "测试.docx should have revisions"
    print("PASS")

    print("\n=== All Win32 ops tests PASSED ===")
