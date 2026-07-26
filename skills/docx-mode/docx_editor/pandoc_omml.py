"""Pandoc-based LaTeX to OMML converter.

Converts LaTeX math expressions to OMML (Office Math Markup Language)
via Pandoc, producing native editable Word equations.

How it works (Pandoc 3.x):
  1. Pandoc parses ``$latex$`` / ``$$latex$$`` as inline/display Math AST nodes
  2. The DOCX writer converts Math AST → OMML XML (``<m:oMath>`` / ``<m:oMathPara>``)
  3. We extract the OMML elements from the generated .docx and return them
     as python-docx ``OxmlElement`` objects ready for injection

Note:
  The ``--mathml`` flag is technically a no-op for DOCX output — Pandoc's
  DOCX writer always uses OMML natively, regardless of the ``--mathml``
  flag.  We keep it for explicitness and forward compatibility.

Usage:
    from docx_editor.pandoc_omml import latex_to_omml, is_pandoc_available

    if is_pandoc_available():
        omath_elements = latex_to_omml(r'E=mc^2')
        if omath_elements:
            for elem in omath_elements:
                paragraph._element.append(elem)
"""

import logging
import os
import re
import tempfile
from typing import List, Optional

OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

logger = logging.getLogger(__name__)

# In-memory cache: once we confirm pandoc works, cache the result
_PANDOC_AVAILABLE = None


def is_pandoc_available() -> bool:
    """Check if pypandoc (and pandoc) is installed.

    Results are cached after the first check to avoid repeated
    subprocess calls.
    """
    global _PANDOC_AVAILABLE
    if _PANDOC_AVAILABLE is not None:
        return _PANDOC_AVAILABLE

    try:
        import pypandoc
        pypandoc.get_pandoc_version()
        _PANDOC_AVAILABLE = True
        return True
    except (ImportError, OSError, Exception) as e:
        logger.debug('pandoc not available: %s', e)
        _PANDOC_AVAILABLE = False
        return False


def _convert_and_extract(md_content: str, extract_tag: str) -> Optional[List]:
    """Core conversion: markdown → DOCX → OMML extraction.

    Args:
        md_content: Markdown fragment containing LaTeX math
        extract_tag: XML tag to extract (``m:oMath`` or ``m:oMathPara``)

    Returns:
        List of parsed OxmlElement objects, or None on failure
    """
    import pypandoc
    from docx import Document as DocxLoader
    from lxml import etree
    from docx.oxml import parse_xml

    tmp_path = None
    try:
        # Convert to docx via temp file
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
            tmp_path = tmp.name

        pypandoc.convert_text(
            md_content,
            'docx',
            outputfile=tmp_path,
            extra_args=['--mathml'],
            format='markdown',
        )

        # Load and serialise the document body to XML text
        doc = DocxLoader(tmp_path)
        body_xml = etree.tostring(doc.element.body, encoding='unicode')

        # Extract matching OMML elements via regex
        # (The body_xml is small — one equation — so regex is fine here)
        pattern = rf'<{extract_tag}[ >].*?</{extract_tag}>'
        matches = re.findall(pattern, body_xml, re.DOTALL)

        if not matches:
            logger.warning(
                'No <%s> elements found by pandoc for: %r',
                extract_tag, md_content[:80],
            )
            return None

        # Parse each fragment back to OxmlElement
        result = []
        for xml_fragment in matches:
            # Re-add namespace declaration stripped by re.findall
            # The regex captures e.g. ``<m:oMath ...>`` without the
            # ``xmlns:m`` declaration — inject it on the opening tag.
            base_tag = extract_tag  # e.g. 'm:oMath' or 'm:oMathPara'
            opening = f'<{base_tag}'
            xml_fragment = xml_fragment.replace(
                opening,
                f'{opening} xmlns:m="{OMML_NS}"',
                1,  # only the first (opening) tag
            )
            elem = parse_xml(xml_fragment)
            result.append(elem)

        return result

    except Exception as e:
        logger.warning(
            'Pandoc OMML conversion failed: %s\n'
            '  Input: %r',
            e, md_content[:80],
        )
        return None
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


def set_omml_math_font(omath_elements: List, font_name: str) -> None:
    """Set the math font on all OMML runs in-place.

    Pandoc-generated OMML does not include explicit font names —
    Word renders them with the default math font (Cambria Math).
    This function inserts ``<w:rPr><w:rFonts w:ascii="..." w:hAnsi="..."/>``
    into every ``<m:r>`` element that lacks one.

    Args:
        omath_elements: List of ``<m:oMath>`` or ``<m:oMathPara>`` elements
        font_name: Font name to set (e.g. ``'Cambria Math'``, ``'Times New Roman'``)
    """
    if not font_name or not omath_elements:
        return

    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from lxml import etree

    # Traverse all descendant elements
    for elem in omath_elements:
        for m_r in elem.iter(f'{{{OMML_NS}}}r'):
            # Find or create <m:rPr>
            rPr = m_r.find(f'{{{OMML_NS}}}rPr')
            if rPr is None:
                rPr = OxmlElement(f'm:rPr')
                m_r.insert(0, rPr)

            # Remove existing w:rPr if any (we'll replace it)
            existing_w_rPr = rPr.find(qn('w:rPr'))
            if existing_w_rPr is not None:
                rPr.remove(existing_w_rPr)

            # Insert w:rPr with the desired font
            w_rPr = OxmlElement('w:rPr')
            w_rFonts = OxmlElement('w:rFonts')
            w_rFonts.set(qn('w:ascii'), font_name)
            w_rFonts.set(qn('w:hAnsi'), font_name)
            w_rPr.append(w_rFonts)
            rPr.append(w_rPr)


def latex_to_omml(latex: str, *, math_font: str = '') -> Optional[List]:
    """Convert inline math (``$...$``) to OMML ``<m:oMath>`` elements.

    Args:
        latex: LaTeX math expression (without ``$`` delimiters)
        math_font: Font for math runs. Empty = default (Cambria Math).

    Returns:
        List of ``<m:oMath>`` OxmlElement objects, or ``None`` if
        conversion fails
    """
    md_content = f'${latex}$\n'
    result = _convert_and_extract(md_content, 'm:oMath')
    if result and math_font:
        set_omml_math_font(result, math_font)
    return result


def latex_to_omml_display(latex: str, *, math_font: str = '') -> Optional[List]:
    """Convert display math (``$$...$$``) to OMML ``<m:oMathPara>`` elements.

    Uses ``$$...$$`` so Pandoc generates a display-style equation
    wrapped in ``<m:oMathPara>``.

    Args:
        latex: LaTeX math expression (without ``$$`` delimiters)
        math_font: Font for math runs. Empty = default (Cambria Math).

    Returns:
        List of ``<m:oMathPara>`` OxmlElement objects, or ``None`` if
        conversion fails
    """
    md_content = f'$${latex}$$\n'
    result = _convert_and_extract(md_content, 'm:oMathPara')
    if result:
        if math_font:
            set_omml_math_font(result, math_font)
        return result

    # Pandoc may emit oMath instead of oMathPara for simple expressions
    logger.debug(
        'No oMathPara generated for display math, falling back to oMath'
    )
    return latex_to_omml(latex, math_font=math_font)
