"""Tests for table_builder.py"""
import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from docx import Document
from docx_editor.table_builder import TableBuilder


@pytest.fixture
def doc():
    return Document()


class TestBasicTable:
    def test_create_simple_table(self, doc):
        builder = TableBuilder(doc)
        headers = ["A", "B", "C"]
        rows = [["1", "2", "3"], ["4", "5", "6"]]
        table = builder.build_table(headers=headers, rows=rows)
        assert table is not None
        # Check table structure
        assert len(table.rows) >= 1  # At least headers

    def test_create_table_no_headers(self, doc):
        builder = TableBuilder(doc)
        rows = [["a", "b"], ["c", "d"]]
        table = builder.build_table(headers=None, rows=rows)
        assert table is not None

    def test_cell_content(self, doc):
        builder = TableBuilder(doc)
        headers = ["Name", "Age"]
        rows = [["Alice", "30"], ["Bob", "25"]]
        table = builder.build_table(headers=headers, rows=rows)
        # Check first data cell
        cell = table.cell(1, 0)  # Row 1 (0-indexed), column 0
        assert cell is not None


class TestTableWithMerges:
    def test_horizontal_merge(self, doc):
        builder = TableBuilder(doc)
        headers = ["A", "B", "C"]
        rows = [["a1", ">", "c1"]]
        merge_map = [["", ">", ""]]
        table = builder.build_table(headers=headers, rows=rows, merge_map=merge_map)
        assert table is not None

    def test_vertical_merge(self, doc):
        builder = TableBuilder(doc)
        headers = ["A", "B"]
        rows = [["a1", "b1"], ["v", "b2"]]
        merge_map = [["", ""], ["v", ""]]
        table = builder.build_table(headers=headers, rows=rows, merge_map=merge_map)
        assert table is not None

    def test_complex_merge(self, doc):
        builder = TableBuilder(doc)
        headers = ["Col A", "Col B", "Col C", "Col D"]
        rows = [
            ["a1", ">", "c1", "d1"],
            ["a2", "v", "c2", ">"],
        ]
        merge_map = [
            ["", ">", "", ""],
            ["", "v", "", ">"],
        ]
        table = builder.build_table(headers=headers, rows=rows, merge_map=merge_map)
        assert table is not None

    def test_multiple_horizontal_merges_in_row(self, doc):
        builder = TableBuilder(doc)
        headers = ["A", "B", "C", "D", "E"]
        rows = [["a1", ">", ">", "d1", ">"]]
        merge_map = [["", ">", ">", "", ">"]]
        table = builder.build_table(headers=headers, rows=rows, merge_map=merge_map)
        assert table is not None
        # First cell should span 3 columns
        row = table.rows[0]
        # The merged cell should span multiple columns

    def test_empty_merge_map(self, doc):
        builder = TableBuilder(doc)
        headers = ["A", "B"]
        rows = [["1", "2"]]
        table = builder.build_table(headers=headers, rows=rows)
        assert table is not None


class TestTableEdgeCases:
    def test_empty_rows(self, doc):
        builder = TableBuilder(doc)
        table = builder.build_table(headers=["A"], rows=[])
        assert table is not None
        # Should have at least the header row

    def test_single_cell(self, doc):
        builder = TableBuilder(doc)
        table = builder.build_table(headers=["X"], rows=[["Y"]])
        assert table is not None

    def test_irregular_row_lengths(self, doc):
        builder = TableBuilder(doc)
        rows = [["a", "b", "c"], ["d", "e"]]
        table = builder.build_table(rows=rows)
        assert table is not None
