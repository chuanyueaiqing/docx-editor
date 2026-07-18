"""Table builder module.

Builds python-docx tables with merged cells from markdown table data.
Supports custom merge markers:
  - '>' in a cell = horizontal merge (merge with cell to the left)
  - 'v' in a cell = vertical merge (merge with cell above)
"""
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.shared import Pt, Emu, RGBColor
from docx.oxml.ns import qn as docx_qn
from lxml import etree

from .utils import DocxError, TableBuildError, qn


class TableBuilder:
    """Build python-docx tables with merged cells.

    Usage:
        builder = TableBuilder(document)
        table = builder.build_table(
            headers=["A", "B", "C"],
            rows=[["1", ">", "3"], ["4", "v", "6"]],
            merge_map=[["", ">", ""], ["", "v", ""]],
        )
    """

    # OOXML namespace for table cells
    W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    def __init__(self, document: Document):
        self.document = document

    def build_table(
        self,
        headers: Optional[List[str]] = None,
        rows: Optional[List[List[str]]] = None,
        merge_map: Optional[List[List[str]]] = None,
    ) -> Any:
        """Build a docx table with optional merged cells.

        Args:
            headers: Column header texts (optional)
            rows: Data rows as list of cell text lists
            merge_map: 2D array of merge markers matching the rows shape.
                       '>' = merge with cell to the left (horizontal).
                       'v' = merge with cell above (vertical).
                       '' or None = no merge.

        Returns:
            python-docx Table object
        """
        headers = headers or []
        rows = rows or []
        merge_map = merge_map or []

        # Calculate dimensions
        num_cols = max(
            len(headers),
            max((len(r) for r in rows), default=0),
        )
        if num_cols == 0:
            num_cols = 1

        num_data_rows = len(rows)
        total_rows = num_data_rows + (1 if headers else 0)

        if total_rows == 0:
            # Create table with at least 1 row
            total_rows = 1

        # Create the table
        table = self.document.add_table(rows=total_rows, cols=num_cols)
        table.style = 'Table Grid'

        # Fill headers
        if headers:
            for col_idx, header_text in enumerate(headers):
                if col_idx < num_cols:
                    cell = table.cell(0, col_idx)
                    cell.text = header_text
                    self._style_cell(cell, is_header=True)

        # Fill data rows
        for row_idx, row_data in enumerate(rows):
            doc_row_idx = row_idx + (1 if headers else 0)
            for col_idx, cell_text in enumerate(row_data):
                if col_idx < num_cols:
                    cell = table.cell(doc_row_idx, col_idx)
                    cell.text = cell_text

        # Apply merges from merge_map
        if merge_map:
            self._apply_merges(table, merge_map, headers is not None)

        return table

    def _apply_merges(
        self,
        table,
        merge_map: List[List[str]],
        has_headers: bool,
    ):
        """Apply merge markers to the table.

        Process order: horizontal merges first, then vertical merges.
        This matches how Word tracks merged cells (gridSpan for horizontal,
        vMerge for vertical).

        Args:
            table: python-docx Table object
            merge_map: 2D list of merge markers ('>', 'v', or '')
            has_headers: Whether the first row is a header row
        """
        data_start = 1 if has_headers else 0

        # Phase 1: Horizontal merges (left to right, top to bottom)
        for merge_row_idx, merge_row in enumerate(merge_map):
            doc_row_idx = merge_row_idx + data_start

            col = 0
            while col < len(merge_row):
                if merge_row[col] == '>' and col > 0:
                    # Merge with cell to the left
                    try:
                        left_cell = table.cell(doc_row_idx, col - 1)
                        curr_cell = table.cell(doc_row_idx, col)
                        left_cell.merge(curr_cell)
                    except Exception:
                        pass  # Skip problematic merges
                col += 1

        # Phase 2: Vertical merges (top to bottom, left to right)
        for merge_row_idx, merge_row in enumerate(merge_map):
            doc_row_idx = merge_row_idx + data_start

            for col, marker in enumerate(merge_row):
                if marker == 'v' and merge_row_idx > 0:
                    # Merge with cell above
                    try:
                        above_cell = table.cell(doc_row_idx - 1, col)
                        curr_cell = table.cell(doc_row_idx, col)
                        above_cell.merge(curr_cell)
                    except Exception:
                        pass

    def _style_cell(self, cell, is_header: bool = False):
        """Apply basic styling to a table cell.

        Args:
            cell: python-docx Cell object
            is_header: Whether this is a header cell
        """
        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.space_before = Pt(2)
            paragraph.paragraph_format.space_after = Pt(2)
            for run in paragraph.runs:
                if is_header:
                    run.bold = True

    @staticmethod
    def _set_cell_shading(cell, color: str):
        """Set background shading on a cell via XML.

        Args:
            cell: python-docx Cell object
            color: Hex color string (e.g. "D9E2F3")
        """
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shading = tcPr.makeelement(docx_qn('w:shd'), {
            docx_qn('w:fill'): color,
            docx_qn('w:val'): 'clear',
        })
        tcPr.append(shading)
