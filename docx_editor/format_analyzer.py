"""Format analysis module.

Provides cross-chapter format comparison and consistency checking
for DOCX documents. Used to detect formatting differences between
chapters and determine the dominant format across the document.
"""
from collections import Counter
from typing import Any, Dict, List, Optional

from .document import DocxDocument
from .models import FormattingData, ParagraphFormatData


class FormatAnalyzer:
    """Analyze and compare formatting across document chapters.

    Usage:
        doc = DocxDocument('path.docx')
        analyzer = FormatAnalyzer(doc)
        report = analyzer.analyze()
        # report['inconsistencies'] -> list of format differences
        # report['dominant_body_format'] -> most common format
    """

    # Run-level fields that matter for body text comparison
    _RUN_FIELDS = [
        'font_name',
        'font_name_east_asia',
        'size',
        'bold',
        'italic',
        'underline',
    ]

    # Paragraph-level fields that matter for body text comparison
    _PARA_FIELDS = [
        'alignment',
        'first_line_indent',
        'left_indent',
        'line_spacing',
        'line_spacing_rule',
        'space_before',
        'space_after',
    ]

    def __init__(self, document: DocxDocument):
        self.doc = document
        self.chapters = document.get_chapter_tree()
        self.flat_chapters = document.chapter_parser.list_chapters()

    def analyze(self) -> Dict[str, Any]:
        """Run full format analysis on the document.

        Returns:
            Dict with:
            - ``chapters``: per-chapter format data
            - ``inconsistencies``: list of format differences across chapters
            - ``dominant_body_format``: most common body format overall
            - ``total_body_paragraphs``: count of body paragraphs analyzed
            - ``total_chapters``: count of chapters
        """
        chapter_data = self._analyze_chapters()
        dominant = self._compute_dominant_format(chapter_data)
        inconsistencies = self._detect_inconsistencies(chapter_data, dominant)

        total_body = sum(
            ch['body_paragraph_count'] for ch in chapter_data
        )

        return {
            'chapters': chapter_data,
            'inconsistencies': inconsistencies,
            'dominant_body_format': dominant,
            'total_body_paragraphs': total_body,
            'total_chapters': len(chapter_data),
        }

    def _analyze_chapters(self) -> List[Dict[str, Any]]:
        """Analyze formatting for every chapter in the document."""
        results = []

        for chapter in self.flat_chapters:
            body_fmts = self._collect_body_formats(chapter)
            heading_fmt = self._collect_heading_format(chapter)

            chapter_data: Dict[str, Any] = {
                'chapter_number': chapter.to_string(),
                'chapter_title': chapter.heading_text,
                'heading_level': chapter.heading_level,
                'body_paragraph_count': len(chapter.body_paragraph_indices),
                'body_runs_count': sum(
                    len(fmts.get('runs', [])) for fmts in body_fmts
                ),
            }

            if body_fmts:
                chapter_data['body_format'] = self._aggregate_formats(body_fmts)
            else:
                chapter_data['body_format'] = {}

            if heading_fmt:
                chapter_data['heading_format'] = heading_fmt
            else:
                chapter_data['heading_format'] = {}

            results.append(chapter_data)

        return results

    def _collect_body_formats(
        self, chapter
    ) -> List[Dict[str, Any]]:
        """Collect formatting data for all body paragraphs in a chapter.

        Returns:
            List of dicts with 'para_format' (ParagraphFormatData) and
            'runs' (list of FormattingData per run)
        """
        formats = []
        for idx in chapter.body_paragraph_indices:
            fmt_entry = self.doc.format_store.formats_json.get(idx)
            if fmt_entry is None:
                continue
            # Only collect non-empty paragraphs to avoid noise
            para = self.doc.document.paragraphs[idx]
            if para.text.strip():
                formats.append(fmt_entry)
        return formats

    def _collect_heading_format(
        self, chapter
    ) -> Optional[Dict[str, Any]]:
        """Collect formatting from the chapter's heading paragraph.

        Returns:
            Dict with 'run_format' and 'para_format', or None if not found
        """
        idx = chapter.heading_paragraph_index
        fmt_entry = self.doc.format_store.formats_json.get(idx)
        if fmt_entry is None:
            return None

        runs = fmt_entry.get('runs', [])
        run_fmt = runs[0]['formatting'] if runs else FormattingData()
        para_fmt = fmt_entry.get('paragraph_format', ParagraphFormatData())

        return {
            'run_format': self._formatting_to_dict(run_fmt),
            'para_format': self._paragraph_format_to_dict(para_fmt),
        }

    def _aggregate_formats(
        self, format_entries: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """Aggregate formatting across multiple paragraphs into one representative format.

        For each field, picks the most common value (mode).

        Args:
            format_entries: List of format dicts with 'paragraph_format' and 'runs'

        Returns:
            Single dict with the most common values for each field
        """
        run_values: Dict[str, list] = {f: [] for f in self._RUN_FIELDS}
        para_values: Dict[str, list] = {f: [] for f in self._PARA_FIELDS}

        for entry in format_entries:
            # Collect run-level data from all runs
            for run_data in entry.get('runs', []):
                fmt = run_data.get('formatting')
                if fmt is None:
                    continue
                fd = self._formatting_to_dict(fmt)
                for field in self._RUN_FIELDS:
                    val = fd.get(field)
                    if val is not None:
                        run_values[field].append(val)

            # Collect paragraph-level data
            para_fmt = entry.get('paragraph_format')
            if para_fmt is not None:
                pd = self._paragraph_format_to_dict(para_fmt)
                for field in self._PARA_FIELDS:
                    val = pd.get(field)
                    if val is not None:
                        para_values[field].append(val)

        # Pick mode for each field
        result: Dict[str, Any] = {}
        for field in self._RUN_FIELDS:
            vals = run_values[field]
            if vals:
                counter = Counter(vals)
                result[field] = counter.most_common(1)[0][0]
        for field in self._PARA_FIELDS:
            vals = para_values[field]
            if vals:
                counter = Counter(vals)
                result[field] = counter.most_common(1)[0][0]

        return result

    def _compute_dominant_format(
        self, chapter_data: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """Compute the single most common body format across all chapters.

        Merges all per-chapter body formats into one dominant format.
        """
        all_bodies: Dict[str, Counter] = {}
        for field in self._RUN_FIELDS + self._PARA_FIELDS:
            all_bodies[field] = Counter()

        for ch in chapter_data:
            bf = ch.get('body_format', {})
            for field in self._RUN_FIELDS + self._PARA_FIELDS:
                val = bf.get(field)
                if val is not None:
                    all_bodies[field][val] += 1

        dominant = {}
        for field, counter in all_bodies.items():
            if counter:
                dominant[field] = counter.most_common(1)[0][0]

        return dominant

    def _detect_inconsistencies(
        self,
        chapter_data: List[Dict[str, Any]],
        dominant: Dict[str, Any],
    ) -> List[Dict[str, Any]]:
        """Detect chapters whose body/heading format differs from the dominant format.

        Args:
            chapter_data: Per-chapter format data from ``_analyze_chapters``
            dominant: Dominant body format from ``_compute_dominant_format``

        Returns:
            List of inconsistency dicts, each with:
            - type: 'body' or 'heading'
            - field: the format field that differs
            - chapter_number: which chapter
            - value: the chapter's value
            - dominant_value: the dominant value
        """
        inconsistencies = []

        for ch in chapter_data:
            ch_num = ch['chapter_number']

            # Check body format
            bf = ch.get('body_format', {})
            for field in self._RUN_FIELDS + self._PARA_FIELDS:
                ch_val = bf.get(field)
                dom_val = dominant.get(field)
                if ch_val is not None and dom_val is not None:
                    if ch_val != dom_val and ch_val != dom_val:
                        inconsistencies.append({
                            'type': 'body',
                            'field': field,
                            'chapter_number': ch_num,
                            'value': ch_val,
                            'dominant_value': dom_val,
                        })

            # Check heading format (compare run_format font/size/bold across chapters)
            hf = ch.get('heading_format', {})
            hf_run = hf.get('run_format', {})
            for field in ('font_name', 'font_name_east_asia', 'size', 'bold'):
                ch_val = hf_run.get(field)
                # Find another heading's same field as reference
                for other_ch in chapter_data:
                    if other_ch['chapter_number'] == ch_num:
                        continue
                    other_hf = other_ch.get('heading_format', {})
                    other_run = other_hf.get('run_format', {})
                    other_val = other_run.get(field)
                    if ch_val is not None and other_val is not None:
                        if ch_val != other_val:
                            # Check if we already reported this
                            already = any(
                                i['type'] == 'heading'
                                and i['field'] == field
                                and i['chapter_number'] == ch_num
                                for i in inconsistencies
                            )
                            if not already:
                                inconsistencies.append({
                                    'type': 'heading',
                                    'field': field,
                                    'chapter_number': ch_num,
                                    'value': ch_val,
                                    'dominant_value': other_val,
                                })
                            break

        return inconsistencies

    # ======================== Format Serialisation ========================

    @staticmethod
    def _formatting_to_dict(fmt: FormattingData) -> Dict[str, Any]:
        """Convert a FormattingData dataclass to a plain dict (non-None fields only)."""
        result: Dict[str, Any] = {}
        fields = [
            'font_name', 'font_name_east_asia', 'font_name_h_ansi',
            'font_name_cs', 'size', 'size_cs', 'bold', 'italic', 'underline',
            'color', 'strike', 'double_strike', 'superscript', 'subscript',
            'small_caps', 'all_caps', 'highlight', 'lang',
        ]
        for f in fields:
            val = getattr(fmt, f, None)
            if val is not None:
                result[f] = val
        return result

    @staticmethod
    def _paragraph_format_to_dict(fmt: ParagraphFormatData) -> Dict[str, Any]:
        """Convert ParagraphFormatData to a plain dict (non-None fields only)."""
        result: Dict[str, Any] = {}
        fields = [
            'style_name', 'style_id', 'alignment',
            'first_line_indent', 'left_indent', 'right_indent',
            'hanging_indent', 'space_before', 'space_after',
            'line_spacing', 'line_spacing_rule', 'outline_level',
        ]
        for f in fields:
            val = getattr(fmt, f, None)
            if val is not None:
                result[f] = val
        return result

    @staticmethod
    def format_summary(format_dict: Dict[str, Any]) -> str:
        """Create a human-readable summary of a format dict.

        Args:
            format_dict: A format dict (e.g. ``dominant_body_format``)

        Returns:
            Multi-line string describing the format
        """
        lines = []
        if format_dict.get('font_name'):
            lines.append(f"字体: {format_dict['font_name']}")
        if format_dict.get('font_name_east_asia'):
            lines.append(f"中文字体: {format_dict['font_name_east_asia']}")
        if format_dict.get('size'):
            pt = format_dict['size'] / 2.0  # half-pts → pts
            lines.append(f"字号: {pt}pt")
        if format_dict.get('bold'):
            lines.append("加粗: 是")
        if format_dict.get('italic'):
            lines.append("斜体: 是")
        if format_dict.get('alignment'):
            lines.append(f"对齐: {format_dict['alignment']}")
        if format_dict.get('first_line_indent'):
            indent = format_dict['first_line_indent'] / 20.0  # twips → pts
            lines.append(f"首行缩进: {indent}pt")
        if format_dict.get('line_spacing'):
            lines.append(f"行距: {format_dict['line_spacing']}")
        if format_dict.get('line_spacing_rule'):
            lines.append(f"行距规则: {format_dict['line_spacing_rule']}")
        if format_dict.get('space_before'):
            lines.append(f"段前间距: {format_dict['space_before'] / 20.0}pt")
        if format_dict.get('space_after'):
            lines.append(f"段后间距: {format_dict['space_after'] / 20.0}pt")
        if format_dict.get('left_indent'):
            lines.append(f"左缩进: {format_dict['left_indent'] / 20.0}pt")

        return '\n'.join(lines) if lines else '(默认格式)'

    @staticmethod
    def inconsistency_summary(
        inconsistencies: List[Dict[str, Any]]
    ) -> str:
        """Create a human-readable summary of all inconsistencies found.

        Args:
            inconsistencies: List from ``analyze()['inconsistencies']``

        Returns:
            Multi-line string describing each inconsistency
        """
        if not inconsistencies:
            return "未发现格式不一致。"

        lines = ["发现以下格式不一致："]
        for inc in inconsistencies:
            ch = inc['chapter_number']
            field_name = {
                'font_name': '西文字体',
                'font_name_east_asia': '中文字体',
                'size': '字号',
                'bold': '加粗',
                'italic': '斜体',
                'underline': '下划线',
                'alignment': '对齐方式',
                'first_line_indent': '首行缩进',
                'line_spacing': '行距',
                'line_spacing_rule': '行距规则',
                'space_before': '段前间距',
                'space_after': '段后间距',
            }.get(inc['field'], inc['field'])

            lines.append(
                f"  • 章节 {ch} 的{inc['type']}格式 — {field_name}: "
                f"当前值「{inc['value']}」，"
                f"主导格式为「{inc['dominant_value']}」"
            )

        return '\n'.join(lines)
