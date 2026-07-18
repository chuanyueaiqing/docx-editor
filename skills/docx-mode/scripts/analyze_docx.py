#!/usr/bin/env python
"""CLI: Analyse and modify existing DOCX documents.

Usage:
    python scripts/analyze_docx.py --tree path.docx
    python scripts/analyze_docx.py --format-report path.docx
    python scripts/analyze_docx.py --comments-with-context path.docx
    python scripts/analyze_docx.py --delete-comment path.docx <comment_id>
    python scripts/analyze_docx.py --delete-all-comments path.docx
    python scripts/analyze_docx.py --contents path.docx 2.1
    python scripts/analyze_docx.py --replace path.docx 2.1 --content "..." --output out.docx
    python scripts/analyze_docx.py --apply-format path.docx --target body --font-name 黑体
"""
import argparse
import os
import sys

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_SKILL_DIR = os.path.abspath(os.path.join(_SCRIPT_DIR, '..'))
if _SKILL_DIR not in sys.path:
    sys.path.insert(0, _SKILL_DIR)

from docx_editor import DocxDocument
from docx_editor.format_analyzer import FormatAnalyzer


# ======================== Sub-command Implementations ========================


def cmd_tree(path: str):
    """Display the chapter tree."""
    doc = DocxDocument(path)
    tree = doc.chapter_parser.tree_to_string()
    if tree.strip():
        print(tree)
    else:
        print('(No chapters detected)')


def cmd_format_report(path: str):
    """Analyse formatting consistency across chapters."""
    doc = DocxDocument(path)
    report = FormatAnalyzer(doc).analyze()

    print(f'Total chapters:  {report["total_chapters"]}')
    print(f'Body paragraphs: {report["total_body_paragraphs"]}')
    print()

    for ch in report['chapters']:
        num = ch['chapter_number']
        title = ch['chapter_title']
        bf = ch.get('body_format', {})
        hf = ch.get('heading_format', {})
        print(f'── Chapter {num}: {title} ──')
        if bf:
            print(f'  Body ({ch["body_paragraph_count"]} paras, {ch["body_runs_count"]} runs):')
            print(f'    {FormatAnalyzer.format_summary(bf).replace(chr(10), chr(10) + "    ")}')
        else:
            print('  Body: (empty or no format data)')
        if hf:
            print(f'  Heading:')
            hr = hf.get('run_format', {})
            print(f'    Font: {hr.get("font_name", "?")}, Size: {hr.get("size", "?")}')

    inc = report['inconsistencies']
    if inc:
        print()
        print(FormatAnalyzer.inconsistency_summary(inc))

    dom = report['dominant_body_format']
    if dom:
        print()
        print('Document dominant body format:')
        print(f'  {FormatAnalyzer.format_summary(dom).replace(chr(10), chr(10) + "  ")}')


def cmd_contents(path: str, chapter_num: str):
    """Show the text content of a chapter."""
    doc = DocxDocument(path)
    chapter = doc.get_chapter(chapter_num)
    if chapter is None:
        print(f'Error: Chapter "{chapter_num}" not found.', file=sys.stderr)
        sys.exit(1)
    text = doc.get_chapter_text(chapter)
    print(text)

    # Extract and report embedded images
    thumb_dir = os.path.join(
        os.path.dirname(os.path.abspath(path)),
        'thumbnails',
        f'ch{chapter_num.replace(".", "_")}',
    )
    images = doc.get_chapter_images(chapter, thumb_dir)
    if images:
        refs = []
        for embed, img_path in images.items():
            _, ext = os.path.splitext(img_path)
            refs.append(f'  [图片: {embed}{ext}] -> {img_path}')
        print()
        print('📷 本章图片 (' + str(len(images)) + ' 张):')
        print('\n'.join(refs))


def _build_threads(comments):
    """Group comments into conversation threads.

    Uses ``w:parent`` (formal OOXML threading) first.
    If no formal parent exists, falls back to grouping comments
    that share the same ``paragraph_index`` (heuristic threading).

    Returns a list of threads, each thread being a list of
    ``CommentData`` sorted by date (parent first, replies after).
    """
    by_id = {c.id: c for c in comments}
    has_formal_parent = any(c.parent_id and c.parent_id in by_id for c in comments)

    if has_formal_parent:
        # ── Formal threading via w:parent ──
        roots: list = []
        children_of: dict = {}
        for c in comments:
            if c.parent_id and c.parent_id in by_id:
                children_of.setdefault(c.parent_id, []).append(c)
            else:
                roots.append(c)

        threads = []
        for root in roots:
            thread = [root] + children_of.get(root.id, [])
            threads.append(thread)
        linked_ids = {c.id for t in threads for c in t}
        for c in comments:
            if c.id not in linked_ids:
                threads.append([c])
        return _sort_threads(threads, comments)

    else:
        # ── Heuristic: group by paragraph_index ──
        by_para: dict = {}
        unmapped = []
        for c in comments:
            if c.paragraph_index is not None:
                by_para.setdefault(c.paragraph_index, []).append(c)
            else:
                unmapped.append(c)

        threads = []
        for pidx, group in by_para.items():
            group.sort(key=lambda c: c.date or '')
            threads.append(group)
        for c in unmapped:
            threads.append([c])

        return _sort_threads(threads, comments)


def _sort_threads(threads, comments):
    """Sort threads by the first comment's position in the original list."""
    pos = {c.id: i for i, c in enumerate(comments)}
    threads.sort(key=lambda t: pos.get(t[0].id, 999999))
    return threads


def _format_date(d):
    """Strip seconds / timezone for a shorter date display."""
    if not d:
        return ''
    return d.replace('T', ' ').split('Z')[0].split('+')[0][:16]


def _print_comment(c, doc, prefix='', is_last=True, show_context=True,
                   display_num=None):
    """Print a single comment with optional context."""
    connector = '└─ ' if prefix else ''
    indent = prefix + connector
    text_indent = prefix + ('   ' if prefix else '')
    tag = f'[#{display_num}] ' if display_num is not None else ''

    print(f'{indent}{tag}{c.author}  {_format_date(c.date)}')
    for line in c.text.split('\n'):
        print(f'{text_indent}{line}')

    if show_context:
        ctx = doc.get_comment_context(c.id)
        if ctx and ctx.get('context_text'):
            visible = [l for l in ctx['context_text'].split('\n') if l.strip()]
            for cl in visible[:5]:
                print(f'{text_indent}{cl}')
            if len(visible) > 5:
                print(f'{text_indent}... (共 {len(visible)} 段)')
    print()


def cmd_comments_with_context(path: str):
    """List all comments threaded by parent_id or paragraph_index."""
    doc = DocxDocument(path)
    comments = doc.read_comments()

    if not comments:
        print('No comments found.')
        return

    n = len(comments)
    threads = _build_threads(comments)
    thread_count = sum(1 for t in threads if len(t) > 1)

    # Assign display numbers based on document order
    display_of = {c.id: i + 1 for i, c in enumerate(comments)}

    extra = f'，{thread_count} 个对话线程' if thread_count else ''
    print(f'{n} 条批注{extra}\n')

    for tidx, thread in enumerate(threads, 1):
        parent = thread[0]
        ctx = doc.get_comment_context(parent.id)
        ch = ctx['chapter_title'] if ctx and ctx.get('chapter_title') else '(未映射到章节)'

        plural = f'（{len(thread)} 条）' if len(thread) > 1 else ''
        label = f'线程 {tidx}' if len(thread) > 1 else '单独批注'
        print(f'── {label} ── {ch} {plural}──')

        for ri, c in enumerate(thread):
            _print_comment(c, doc,
                          show_context=(ri == 0),
                          is_last=(ri == len(thread) - 1),
                          display_num=display_of.get(c.id))


def cmd_delete_comment(path: str, comment_id: str):
    """Delete a single comment by ID."""
    doc = DocxDocument(path)
    ok = doc.delete_comment(comment_id)
    if ok:
        doc.save()
        print(f'Comment {comment_id} deleted.')
    else:
        print(f'Comment {comment_id} not found.', file=sys.stderr)
        sys.exit(1)


def cmd_delete_all_comments(path: str):
    """Delete all comments from the document."""
    doc = DocxDocument(path)
    n = doc.delete_all_comments()
    if n:
        doc.save()
        print(f'{n} comment(s) deleted.')
    else:
        print('No comments to delete.')


def cmd_replace(path: str, chapter_num: str, md_text: str,
                output: str, track_changes: bool):
    """Replace a chapter's content with Markdown."""
    doc = DocxDocument(path, use_track_changes=track_changes)
    doc.replace_chapter(chapter_num, md_text)
    doc.save(output or path)
    print(f'Saved: {output or path}')


def cmd_apply_format(path: str, target: str, format_spec: dict):
    """Apply formatting to paragraphs matching target."""
    doc = DocxDocument(path)

    # Resolve target -- if it looks like a chapter number, treat as list
    if target == 'all':
        target_arg = 'all'
    elif target == 'body':
        target_arg = 'body'
    elif target == 'heading':
        target_arg = 'heading'
    else:
        # Assume it's a chapter number
        target_arg = [target]

    n = doc.apply_format(format_spec, target=target_arg)
    if n:
        doc.save()
        print(f'Formatting applied to {n} paragraph(s). Saved.')
    else:
        print('No paragraphs matched the target.')


# ======================== Argument Parsing ========================


def make_apply_format_parser(sub):
    """Add --apply-format related arguments to a subparser."""
    sub.add_argument('--target', default='body',
                     help='all | body | heading | <chapter_number>')
    sub.add_argument('--font-name', help='Western font name')
    sub.add_argument('--font-ea', help='East-Asian font name')
    sub.add_argument('--font-size', type=int, help='Font size in points')
    sub.add_argument('--bold', action='store_true', default=None,
                     help='Bold (use --no-bold to unset)')
    sub.add_argument('--no-bold', action='store_false', dest='bold')
    sub.add_argument('--italic', action='store_true', default=None,
                     help='Italic (use --no-italic to unset)')
    sub.add_argument('--no-italic', action='store_false', dest='italic')
    sub.add_argument('--alignment',
                     choices=('LEFT', 'CENTER', 'RIGHT', 'JUSTIFY'))
    sub.add_argument('--line-spacing', type=float)
    sub.add_argument('--first-line-indent', type=float,
                     help='In points')
    sub.add_argument('--space-before', type=float, help='In points')
    sub.add_argument('--space-after', type=float, help='In points')


def build_format_spec(args) -> dict:
    spec = {}
    for key in ('font_name', 'font_name_east_asia', 'font_size',
                'bold', 'italic', 'alignment', 'line_spacing',
                'first_line_indent', 'space_before', 'space_after'):
        val = getattr(args, key.replace('-', '_'), None)
        if val is not None:
            spec_key = key
            # Map font-ea -> font_name_east_asia
            spec[spec_key] = val
    return spec


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description='Analyse and modify DOCX documents.'
    )
    sub = parser.add_subparsers(dest='command', required=True)

    # --tree
    p = sub.add_parser('tree', help='Display chapter tree')
    p.add_argument('path', help='Path to .docx file')

    # --format-report (using subparser name)
    p = sub.add_parser('format-report', help='Analyse formatting consistency')
    p.add_argument('path', help='Path to .docx file')

    # --contents
    p = sub.add_parser('contents', help='Show chapter text content')
    p.add_argument('path', help='Path to .docx file')
    p.add_argument('chapter', help='Chapter number (e.g. "2.1")')

    # --comments-with-context
    p = sub.add_parser('comments-with-context',
                       help='List comments with chapter mapping')
    p.add_argument('path', help='Path to .docx file')

    # --delete-comment
    p = sub.add_parser('delete-comment', help='Delete a comment by ID')
    p.add_argument('path', help='Path to .docx file')
    p.add_argument('comment_id', help='Comment ID')

    # --delete-all-comments
    p = sub.add_parser('delete-all-comments',
                       help='Delete all comments')
    p.add_argument('path', help='Path to .docx file')

    # --replace
    p = sub.add_parser('replace', help='Replace a chapter with Markdown')
    p.add_argument('path', help='Path to .docx file')
    p.add_argument('chapter', help='Chapter number (e.g. "2.1")')
    src = p.add_mutually_exclusive_group(required=True)
    src.add_argument('--content', help='Markdown text')
    src.add_argument('--input', help='Markdown file')
    p.add_argument('--output', help='Output .docx path (default: overwrite)')
    p.add_argument('--track-changes', action='store_true',
                   help='Use Word track-changes mode')

    # --apply-format
    p = sub.add_parser('apply-format', help='Batch-apply formatting')
    p.add_argument('path', help='Path to .docx file')
    make_apply_format_parser(p)

    return parser.parse_args(argv)


def main():
    args = parse_args()

    if args.command == 'tree':
        cmd_tree(args.path)

    elif args.command == 'format-report':
        cmd_format_report(args.path)

    elif args.command == 'contents':
        cmd_contents(args.path, args.chapter)

    elif args.command == 'comments-with-context':
        cmd_comments_with_context(args.path)

    elif args.command == 'delete-comment':
        cmd_delete_comment(args.path, args.comment_id)

    elif args.command == 'delete-all-comments':
        cmd_delete_all_comments(args.path)

    elif args.command == 'replace':
        if args.input:
            with open(args.input, 'r', encoding='utf-8') as f:
                md_text = f.read()
        else:
            md_text = args.content
        cmd_replace(args.path, args.chapter, md_text,
                    args.output, args.track_changes)

    elif args.command == 'apply-format':
        fmt_spec = build_format_spec(args)
        if not fmt_spec:
            print('Error: No format properties specified.', file=sys.stderr)
            sys.exit(1)
        cmd_apply_format(args.path, args.target, fmt_spec)


if __name__ == '__main__':
    main()
