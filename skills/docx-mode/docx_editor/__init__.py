"""DOCX Editor Module.

A comprehensive module for reading, editing, and manipulating DOCX files.
Supports chapter-based access, format extraction, markdown-to-docx conversion,
table merging with custom merge markers, mermaid diagram rendering,
tracked changes, and comments/revisions reading.

Usage:
    from docx_editor import DocxDocument, DocxCreator, FormatAnalyzer

    # Load and edit an existing document
    doc = DocxDocument('path/to/file.docx')
    chapter = doc.get_chapter('3.1')
    doc.replace_chapter('3.1', '# New Section\\n\\nContent...')
    doc.save('output.docx')

    # Create a new document from markdown
    creator = DocxCreator.create('new.docx', '# Title\\n\\nContent', format_spec={...})

    # Analyze format consistency
    from docx_editor.format_analyzer import FormatAnalyzer
    report = FormatAnalyzer(doc).analyze()
"""

from .document import DocxDocument
from .creator import DocxCreator

__all__ = ['DocxCreator', 'DocxDocument']
