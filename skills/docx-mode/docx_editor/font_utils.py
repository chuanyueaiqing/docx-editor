"""Font utilities for Chinese/English bilingual DOCX formatting.

Provides helpers to set East-Asian (Chinese) and Western fonts on
paragraph runs, and to batch-apply font configs to an entire document.

Usage:
    from docx_editor.font_utils import set_run_font, apply_font_config

    # Configure fonts
    config = {
        'body_cn': '宋体', 'body_en': 'Times New Roman', 'body_size': 12,
        'heading_cn': '黑体', 'heading_en': 'Arial',
        'h1_size': 22, 'h2_size': 16, 'h3_size': 14,
        'line_spacing': 1.5,
    }

    # Apply font to a single run
    set_run_font(run, cn_font='宋体', en_font='Times New Roman', size_pt=12)

    # Batch apply to entire document
    apply_font_config(docx_path, config)
"""

import logging
from typing import Any, Dict, Optional

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

logger = logging.getLogger(__name__)

# Default font configuration
DEFAULT_FONT_CONFIG = {
    'body_cn': '宋体',
    'body_en': 'Times New Roman',
    'body_size': 12,
    'heading_cn': '黑体',
    'heading_en': 'Arial',
    'h1_size': 22,
    'h2_size': 16,
    'h3_size': 14,
    'line_spacing': 1.5,
}


def set_run_font(
    run,
    cn_font: Optional[str] = None,
    en_font: Optional[str] = None,
    size_pt: Optional[float] = None,
):
    """Set East-Asian and Western fonts on a single run.

    Args:
        run: python-docx Run object
        cn_font: Chinese font name (e.g. '宋体', '黑体')
        en_font: Western font name (e.g. 'Times New Roman', 'Arial')
        size_pt: Font size in points
    """
    if en_font:
        run.font.name = en_font
    if size_pt:
        run.font.size = Pt(size_pt)

    # Set East-Asian font via XML
    if cn_font:
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = rpr.makeelement(qn('w:rFonts'), {})
            rpr.append(rfonts)
        rfonts.set(qn('w:eastAsia'), cn_font)


def apply_font_config(docx_path: str, config: Optional[Dict[str, Any]] = None):
    """Batch-apply font configuration to a saved .docx file.

    Scans all paragraphs, detects heading styles, and applies
    the corresponding Chinese + Western font and size.

    Args:
        docx_path: Path to the .docx file
        config: Font configuration dict (uses defaults for missing keys)
    """
    cfg = {**DEFAULT_FONT_CONFIG, **(config or {})}
    doc = Document(docx_path)

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else ''

        if style_name.startswith('Heading 1'):
            cn, en, sz = cfg['heading_cn'], cfg['heading_en'], cfg['h1_size']
        elif style_name.startswith('Heading 2'):
            cn, en, sz = cfg['heading_cn'], cfg['heading_en'], cfg['h2_size']
        elif style_name.startswith('Heading 3'):
            cn, en, sz = cfg['heading_cn'], cfg['heading_en'], cfg['h3_size']
        else:
            cn, en, sz = cfg['body_cn'], cfg['body_en'], cfg['body_size']

        for run in para.runs:
            set_run_font(run, cn_font=cn, en_font=en, size_pt=sz)

        para.paragraph_format.line_spacing = cfg['line_spacing']

    # Set Normal style defaults
    if 'Normal' in doc.styles:
        normal = doc.styles['Normal']
        normal.font.name = cfg['body_en']
        normal.font.size = Pt(cfg['body_size'])
        rpr = normal.element.get_or_add_rPr()
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = rpr.makeelement(qn('w:rFonts'), {})
            rpr.append(rfonts)
        rfonts.set(qn('w:eastAsia'), cfg['body_cn'])

    doc.save(docx_path)
    logger.info('Font config applied to %s', docx_path)
