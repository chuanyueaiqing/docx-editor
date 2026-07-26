"""Utility functions and constants for the docx editor module."""

# OOXML namespace constants
NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'rPr': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}


def qn(tag: str) -> str:
    """Convert 'w:p' style tag to Clark notation namespace URI.

    Args:
        tag: Tag in 'prefix:localname' format (e.g. 'w:p')

    Returns:
        Full Clark notation tag with namespace URI

    Raises:
        KeyError: If the prefix is not in NSMAP
    """
    prefix, local = tag.split(':')
    ns = NSMAP[prefix]
    return f'{{{ns}}}{local}'


def get_xml_element_text(element, tag: str) -> str:
    """Get text content from first child XML element matching tag.

    Args:
        element: Parent lxml element
        tag: Tag in 'w:xxx' format

    Returns:
        Text content or empty string
    """
    child = element.find(qn(tag))
    if child is not None and child.text:
        return child.text
    return ''


def half_points_to_pt(half_points) -> float:
    """Convert half-points (as used in OOXML) to points.

    Args:
        half_points: Value in half-points (e.g. 21 = 10.5pt), or None

    Returns:
        Value in points, or None if input was None
    """
    if half_points is None:
        return None
    return float(half_points) / 2.0


def pt_to_half_points(pt: float) -> int:
    """Convert points to half-points (as used in OOXML).

    Args:
        pt: Value in points

    Returns:
        Value in half-points
    """
    return int(pt * 2)


# ---- Custom Exceptions ----


class DocxError(Exception):
    """Base exception for all docx editor errors."""
    pass


class ChapterNotFoundError(DocxError):
    """Raised when a requested chapter number does not exist."""
    pass


class MermaidRenderError(DocxError):
    """Raised when mermaid diagram rendering fails."""
    pass


class CommentsReadError(DocxError):
    """Raised when document comments cannot be read."""
    pass


class Win32ComError(DocxError):
    """Raised when win32com operations fail (Word not installed, etc.)."""
    pass


class MarkdownParseError(DocxError):
    """Raised when markdown parsing fails."""
    pass


class TableBuildError(DocxError):
    """Raised when table building with merges fails."""
    pass


class EquationError(DocxError):
    """Raised when equation/formula operations fail."""
    pass
