#!/usr/bin/env python
"""CLI: Create a new DOCX document from Markdown content.

Usage:
    python scripts/create_docx.py --content "# Title\\n\\nBody" --output out.docx
    python scripts/create_docx.py --input doc.md --output out.docx --font-name "SimSun" --font-size 12
"""
import argparse
import os
import sys
import tempfile

# Make docx_editor importable
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_DIR = os.path.abspath(os.path.join(_SCRIPT_DIR, '..', '..', '..'))
# Fallback: if auto-detection doesn't find docx_editor, use known project path
if not os.path.isdir(os.path.join(_PROJECT_DIR, 'docx_editor')):
    _PROJECT_DIR = 'E:/demo/docx'
if _PROJECT_DIR not in sys.path:
    sys.path.insert(0, _PROJECT_DIR)

from docx_editor.creator import DocxCreator


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description='Create a new DOCX document from Markdown.'
    )
    source = parser.add_mutually_exclusive_group()
    source.add_argument('--content', metavar='TEXT',
                        help='Markdown text content')
    source.add_argument('--input', metavar='FILE',
                        help='Read Markdown from a file')

    parser.add_argument('--output', metavar='FILE', required=True,
                        help='Output .docx path')
    parser.add_argument('--template', metavar='FILE',
                        help='Optional .docx template for styles')

    # Format options
    parser.add_argument('--font-name', metavar='FONT', default='Calibri',
                        help='Western font name (default: Calibri)')
    parser.add_argument('--font-ea', metavar='FONT', default='宋体',
                        help='East-Asian font name (default: 宋体)')
    parser.add_argument('--font-size', type=int, default=12,
                        help='Font size in points (default: 12)')
    parser.add_argument('--line-spacing', type=float, default=1.5,
                        help='Line spacing multiplier (default: 1.5)')
    parser.add_argument('--alignment', default='JUSTIFY',
                        choices=('LEFT', 'CENTER', 'RIGHT', 'JUSTIFY'),
                        help='Paragraph alignment (default: JUSTIFY)')
    parser.add_argument('--first-line-indent', type=float, default=24,
                        help='First line indent in points (default: 24)')
    parser.add_argument('--space-after', type=float, default=6,
                        help='Space after paragraphs in points (default: 6)')

    return parser.parse_args(argv)


def build_format_spec(args) -> dict:
    spec = {}
    if args.font_name:
        spec['font_name'] = args.font_name
    if args.font_ea:
        spec['font_name_east_asia'] = args.font_ea
    if args.font_size:
        spec['font_size'] = args.font_size
    if args.line_spacing:
        spec['line_spacing'] = args.line_spacing
    if args.alignment:
        spec['alignment'] = args.alignment
    if args.first_line_indent:
        spec['first_line_indent'] = args.first_line_indent
    if args.space_after:
        spec['space_after'] = args.space_after
    return spec


def read_markdown(args) -> str:
    if args.content:
        return args.content
    if args.input:
        with open(args.input, 'r', encoding='utf-8') as f:
            return f.read()
    # Read from stdin if piped
    if not sys.stdin.isatty():
        return sys.stdin.read()
    return ''


def main():
    args = parse_args()
    md_text = read_markdown(args)

    if not md_text.strip():
        print('Error: No Markdown content provided.', file=sys.stderr)
        sys.exit(1)

    format_spec = build_format_spec(args)
    creator = DocxCreator(template_path=args.template)
    if format_spec:
        creator.set_default_format(format_spec)
    creator.add_markdown(md_text)
    creator.save(args.output)

    para_count = len(creator.document.paragraphs)
    table_count = len(creator.document.tables)
    print(f'Created: {args.output}')
    print(f'  Paragraphs: {para_count}')
    if table_count:
        print(f'  Tables: {table_count}')


if __name__ == '__main__':
    main()
