"""DOCX Editor Module.

A comprehensive module for reading, editing, and manipulating DOCX files.
Supports chapter-based access, format extraction, markdown-to-docx conversion,
table merging with custom merge markers, mermaid diagram rendering,
tracked changes, and comments/revisions reading.

Usage:
    from docx_editor import DocxDocument
    doc = DocxDocument('path/to/file.docx')
    chapter = doc.get_chapter('3.1')
    doc.replace_chapter('3.1', '# New Section\\n\\nContent...')
    doc.save('output.docx')
"""

from .document import DocxDocument

__all__ = ['DocxDocument']
