"""Data models for the docx editor module.

All data is carried in dataclass objects for type safety and clarity.
"""
from dataclasses import dataclass, field
from enum import Enum
from typing import (
    Any, Dict, List, Optional, Tuple, Union
)


# ---- Formatting Models ----


@dataclass
class FormattingData:
    """Run-level formatting data, extracted directly from OOXML."""
    font_name: Optional[str] = None
    font_name_east_asia: Optional[str] = None
    font_name_h_ansi: Optional[str] = None
    font_name_cs: Optional[str] = None  # Complex script font
    size: Optional[float] = None        # In half-points (e.g. 21 = 10.5pt)
    size_cs: Optional[float] = None     # Complex script size
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[Any] = None     # bool, str, or None
    color: Optional[str] = None         # Hex color without # (e.g. "FF0000")
    color_theme: Optional[str] = None   # Theme color reference
    color_tint: Optional[str] = None
    color_shade: Optional[str] = None
    strike: Optional[bool] = None
    double_strike: Optional[bool] = None
    superscript: Optional[bool] = None
    subscript: Optional[bool] = None
    small_caps: Optional[bool] = None
    all_caps: Optional[bool] = None
    highlight: Optional[str] = None     # Highlight color
    lang: Optional[str] = None
    east_asian_lang: Optional[str] = None
    bidi: Optional[bool] = None
    vanish: Optional[bool] = None       # Hidden text
    spacing: Optional[int] = None       # Character spacing
    position: Optional[int] = None      # Character position (raised/lowered)
    outline: Optional[bool] = None
    shadow: Optional[bool] = None
    emboss: Optional[bool] = None
    imprint: Optional[bool] = None
    no_proof: Optional[bool] = None     # No spelling/grammar check
    spec_vanish: Optional[bool] = None
    web_hidden: Optional[bool] = None
    rtl: Optional[bool] = None
    complex_script: Optional[bool] = None


@dataclass
class ParagraphFormatData:
    """Paragraph-level formatting data."""
    style_name: Optional[str] = None
    style_id: Optional[str] = None
    heading_level: Optional[int] = None  # 0 = not heading, 1-9 = heading levels
    alignment: Optional[str] = None      # LEFT, CENTER, RIGHT, JUSTIFY, DISTRIBUTE
    first_line_indent: Optional[int] = None   # In twips (1/20 pt)
    left_indent: Optional[int] = None
    right_indent: Optional[int] = None
    hanging_indent: Optional[int] = None
    space_before: Optional[int] = None        # In twips (1/20 pt)
    space_after: Optional[int] = None
    line_spacing: Optional[float] = None
    line_spacing_rule: Optional[str] = None  # SINGLE, DOUBLE, MULTIPLE, AT_LEAST, EXACTLY
    outline_level: Optional[int] = None
    keep_next: Optional[bool] = None
    keep_lines: Optional[bool] = None
    page_break_before: Optional[bool] = None
    widow_control: Optional[bool] = None
    suppress_line_numbers: Optional[bool] = None
    shading: Optional[str] = None       # Background color
    borders: Optional[Dict[str, Any]] = None
    numPr: Optional[Dict[str, Any]] = None  # List numbering properties
    contextual_spacing: Optional[bool] = None
    mirror_indents: Optional[bool] = None
    text_direction: Optional[str] = None


@dataclass
class ParagraphData:
    """Complete paragraph data: content + formatting."""
    paragraph: Any = None               # python-docx Paragraph object reference
    index: int = 0                      # Index in document.paragraphs
    text: str = ''
    formatting: Optional[ParagraphFormatData] = None
    runs_data: List[Tuple[str, FormattingData]] = field(default_factory=list)
    heading_number: Optional[Tuple[int, ...]] = None  # e.g. (3, 1, 1)


# ---- Chapter Tree Model ----


@dataclass
class ChapterNode:
    """Tree node representing a chapter in the document hierarchy."""
    heading_text: str
    heading_level: int                            # 1-9
    number_tuple: Optional[Tuple[int, ...]]       # e.g. (3, 1, 1) or None
    heading_paragraph_index: int                  # Index into document.paragraphs
    body_paragraph_indices: List[int] = field(default_factory=list)
    body_table_indices: List[int] = field(default_factory=list)
    children: List['ChapterNode'] = field(default_factory=list)
    parent: Optional['ChapterNode'] = None
    heading_style_id: Optional[str] = None

    def to_string(self) -> str:
        """Return the chapter number as a string like '3.1.1'.

        If no number tuple is assigned, returns the heading text.
        """
        if self.number_tuple:
            return '.'.join(str(n) for n in self.number_tuple)
        return self.heading_text

    def __repr__(self) -> str:
        return f"Chapter({self.to_string()}, {self.heading_text[:20]!r})"


# ---- Comment & Revision Models ----


@dataclass
class CommentData:
    """Represents a single comment/annotation in the document."""
    id: str
    author: str
    date: str
    text: str
    paragraph_index: Optional[int] = None
    parent_id: Optional[str] = None      # w:parent – IDs of the parent comment if this is a reply


@dataclass
class RevisionData:
    """Represents a single tracked change (insertion or deletion)."""
    rev_id: str
    author: str
    date: str
    type: str                     # "insertion" (w:ins) or "deletion" (w:del)
    text: str
    paragraph_index: int


# ---- Markdown Intermediate Representation ----


class MarkdownElementType(Enum):
    """Types of elements in the markdown intermediate representation."""
    PARAGRAPH = "paragraph"
    HEADING = "heading"
    TABLE = "table"
    IMAGE = "image"
    CODE_BLOCK = "code_block"
    MERMAID = "mermaid"
    HORIZONTAL_RULE = "horizontal_rule"
    LIST = "list"
    LIST_ITEM = "list_item"
    BLOCKQUOTE = "blockquote"
    EMPTY_LINE = "empty_line"
    MATH = "math"                     # $...$ 行内公式
    DISPLAY_MATH = "display_math"     # $$...$$ 行间公式


@dataclass
class MarkdownElement:
    """Intermediate representation between markdown and docx elements."""
    type: MarkdownElementType
    text: Optional[str] = None
    level: Optional[int] = None         # Heading level, list nesting level
    children: Optional[List['MarkdownElement']] = field(default=None)
    rows: Optional[List[List[str]]] = field(default=None)          # For tables
    merge_map: Optional[List[List[str]]] = field(default=None)     # For > and v markers
    image_path: Optional[str] = None
    image_base64: Optional[str] = None
    alt_text: Optional[str] = None
    code_language: Optional[str] = None
    source_info: Optional[dict] = None
    items: Optional[List[str]] = None   # For list items
    ordered: Optional[bool] = None      # Ordered vs unordered list


# ---- Format Store ----


@dataclass
class FormatStore:
    """Bundles python-docx Document with extracted format data."""
    docx_path: str
    document: Any = None                 # python-docx Document
    paragraphs_data: List[ParagraphData] = field(default_factory=list)
    formats_json: Dict[int, dict] = field(default_factory=dict)
