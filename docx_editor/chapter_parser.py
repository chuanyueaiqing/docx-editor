"""Chapter parser module.

Builds a hierarchical tree of chapters from document headings,
supports lookup by chapter number string/tuple, and content access.
"""
import re
from typing import Dict, List, Optional, Tuple, Union

from .models import ChapterNode, FormatStore, ParagraphData
from .utils import ChapterNotFoundError, qn


class ChapterParser:
    """Build and query a chapter tree from document headings.

    Usage:
        parser = ChapterParser(format_store)
        chapter = parser.get_chapter_by_number('3.1')
        contents = parser.get_chapter_contents(chapter)
    """

    # Heading style patterns for detection
    HEADING_PATTERNS = [
        re.compile(r'^第[一二三四五六七八九十百零〇]+[章节篇部]'),  # 第一章, 第一节
        re.compile(r'^第\d+[章节篇部]'),                           # 第1章
        re.compile(r'^\d+(\.\d+)*[\s\.、]'),                      # 1, 1.1, 1.1.1
        re.compile(r'^[A-Z]\.\s'),                                 # A. Introduction
        re.compile(r'^[0-9]+\.\s+'),                               # 1. Introduction
    ]

    CHINESE_NUM_MAP = {
        '零': 0, '一': 1, '二': 2, '三': 3, '四': 4,
        '五': 5, '六': 6, '七': 7, '八': 8, '九': 9,
        '十': 10, '百': 100, '〇': 0,
    }

    # Map Chinese heading unit characters to heading levels
    CHINESE_HEADING_UNIT_LEVEL = {
        '篇': 1,
        '章': 1,
        '部': 1,
        '节': 2,
        '条': 3,
    }

    def __init__(self, format_store: FormatStore):
        self.store = format_store
        self.document = format_store.document
        self.root_chapters: List[ChapterNode] = []
        self._chapter_map: Dict[str, ChapterNode] = {}   # "3.1.1" -> node
        self._paragraph_to_chapter: Dict[int, ChapterNode] = {}
        self._all_nodes_cache: Optional[List[ChapterNode]] = None
        self._build_tree()

    # ---- Tree Building ----

    def _populate_paragraphs_data(self):
        """Build ParagraphData list from document paragraphs and format store."""
        self.store.paragraphs_data = []
        for i, para in enumerate(self.document.paragraphs):
            fmt_data = self.store.formats_json.get(i, {})
            pd = ParagraphData(
                paragraph=para,
                index=i,
                text=para.text,
                formatting=fmt_data.get('paragraph_format'),
            )
            self.store.paragraphs_data.append(pd)

    def _build_tree(self):
        """Walk all paragraphs, identify headings, and build chapter tree."""
        self.root_chapters = []
        self._chapter_map = {}
        self._paragraph_to_chapter = {}
        self._all_nodes_cache = None

        # Ensure paragraphs_data is populated
        if not self.store.paragraphs_data:
            self._populate_paragraphs_data()

        # Collect heading paragraphs and their levels
        headings = self._collect_headings()
        if not headings:
            return

        # Assign number tuples
        self._assign_numbers(headings)

        # Build tree structure
        stack: List[ChapterNode] = []
        for idx, level, text, num_tuple in headings:
            node = ChapterNode(
                heading_text=text,
                heading_level=level,
                number_tuple=num_tuple,
                heading_paragraph_index=idx,
            )

            # Pop stack until we find the parent
            while stack and stack[-1].heading_level >= level:
                stack.pop()

            if stack:
                parent = stack[-1]
                parent.children.append(node)
                node.parent = parent
            else:
                self.root_chapters.append(node)

            stack.append(node)

            # Register in chapter map
            if num_tuple:
                key = '.'.join(str(n) for n in num_tuple)
                self._chapter_map[key] = node

            # Map heading paragraph
            self._paragraph_to_chapter[idx] = node

        # Assign body paragraphs to chapters
        self._assign_body_paragraphs(headings)

    def _collect_headings(self) -> List[Tuple[int, int, str, Optional[Tuple[int, ...]]]]:
        """Collect all heading paragraphs from the document.

        Returns:
            List of (paragraph_index, heading_level, text, number_tuple)
            where number_tuple is provisional (will be reassigned)
        """
        headings = []

        for i, para in enumerate(self.document.paragraphs):
            text = para.text.strip()
            if not text:
                continue

            level = self._detect_heading_level(para, i)
            if level is not None and level > 0 and level <= 9:
                headings.append((i, level, text, None))

        return headings

    def _detect_heading_level(self, para, index: int) -> Optional[int]:
        """Detect heading level from a paragraph.

        Uses multiple signals in order of reliability:
        1. Style is a Heading style
        2. Outline level in paragraph properties
        3. Text pattern matching (numbered headings)

        Args:
            para: python-docx Paragraph object
            index: Paragraph index in document

        Returns:
            Heading level (1-9), or None if not a heading
        """
        # Method 1: Check style
        style = para.style
        if style is not None:
            style_name = style.name or ''
            style_id = style.style_id or ''
            for name in [style_name, style_id]:
                name_lower = name.lower()
                if name_lower.startswith('heading'):
                    try:
                        num = ''.join(c for c in name if c.isdigit())
                        if num:
                            return int(num)
                    except (ValueError, IndexError):
                        pass
                    return 1  # Default heading level

        # Method 2: Check format store for outline level
        fmt_data = self.store.formats_json.get(index, {})
        pf = fmt_data.get('paragraph_format')
        if pf and pf.heading_level is not None:
            return pf.heading_level
        if pf and pf.outline_level is not None:
            return pf.outline_level + 1

        # Method 3: Check paragraph element properties
        p_element = para._element
        ppr = p_element.find(qn('w:pPr'))
        if ppr is not None:
            outline_lvl = ppr.find(qn('w:outlineLvl'))
            if outline_lvl is not None:
                val = outline_lvl.get(qn('w:val'))
                if val:
                    return int(val) + 1

        # Method 4: Check if text matches heading patterns
        text = para.text.strip()
        for pattern in self.HEADING_PATTERNS:
            if pattern.match(text):
                # Map Chinese heading units (章/节/篇/部) to levels
                first_word = text.split()[0] if text.split() else ''
                unit_match = re.search(r'[章节篇部]$', first_word)
                if unit_match:
                    level = self.CHINESE_HEADING_UNIT_LEVEL.get(unit_match.group(), 1)
                    return min(level, 9)
                # Determine level from dotted structure
                dots = first_word.count('.')
                return min(dots + 1, 9)

        return None

    def _extract_number_from_text(self, text: str, level: int) -> Optional[Tuple[int, ...]]:
        """Extract chapter number from heading text.

        Args:
            text: Heading text (e.g. "3.1.1 系统架构", "第一章 引言")
            level: Detected heading level

        Returns:
            Number tuple or None
        """
        # Pattern 1: "3.1.1" or "3.1.1 " or "3.1.1 xxx"
        m = re.match(r'^(\d+(?:\.\d+)*)', text)
        if m:
            parts = m.group(1).split('.')
            return tuple(int(p) for p in parts)

        # Pattern 2: "第一章" -> (1,), "第二章" -> (2,)
        m = re.match(r'^第([一二三四五六七八九十百零〇]+)[章节篇部]', text)
        if m:
            num = self._chinese_to_int(m.group(1))
            return (num,)

        # Pattern 3: "第1章" -> (1,)
        m = re.match(r'^第(\d+)[章节篇部]', text)
        if m:
            return (int(m.group(1)),)

        # Pattern 4: "A." -> try to infer from outline
        m = re.match(r'^([A-Z])\.\s', text)
        if m:
            return (ord(m.group(1)) - ord('A') + 1,)

        return None

    def _chinese_to_int(self, chinese: str) -> int:
        """Convert Chinese numeral string to integer.

        Args:
            chinese: Chinese numeral (e.g. "一", "十二", "一百零五")

        Returns:
            Integer value
        """
        if not chinese:
            return 0

        # Handle single character
        if chinese in self.CHINESE_NUM_MAP and len(chinese) == 1:
            return self.CHINESE_NUM_MAP[chinese]

        total = 0
        current = 0
        for char in chinese:
            if char in self.CHINESE_NUM_MAP:
                val = self.CHINESE_NUM_MAP[char]
                if val >= 10:
                    if current == 0:
                        current = 1
                    total += current * val
                    current = 0
                else:
                    current = val
        total += current
        return total if total > 0 else current

    def _assign_numbers(self, headings: List[Tuple[int, int, str, Optional[Tuple[int, ...]]]]):
        """Assign consistent number tuples to headings.

        Algorithm:
        - Maintain counters for levels 1-9
        - When a heading at level N is found:
          - Increment counter[N]
          - Reset counters[N+1..9] to 0
          - Assign (counter[1], ..., counter[N])

        Args:
            headings: List of (index, level, text, extracted_num)
        """
        counters = [0] * 10  # Index 1-9 used

        for i, (idx, level, text, extracted_num) in enumerate(headings):
            # Increment at this level
            counters[level] += 1
            # Reset higher levels
            for j in range(level + 1, 10):
                counters[j] = 0

            # Build the number tuple
            num_tuple = tuple(counters[1:level + 1])

            # Replace in list (list is mutable, but tuples are not)
            # We need to update the headings list element
            headings[i] = (idx, level, text, num_tuple)

    def _assign_body_paragraphs(self, headings: List[Tuple[int, int, str, Optional[Tuple[int, ...]]]]):
        """Assign body paragraphs (non-heading) to their containing chapter.

        Each heading's **immediate body** consists of paragraphs that are direct
        children of that heading — i.e. paragraphs after it but NOT inside any
        sub-heading.  Sub-headings' body paragraphs belong to those sub-headings.

        Args:
            headings: List of (index, level, text, number_tuple)
        """
        if not headings:
            return

        heading_indices: set = {h[0] for h in headings}

        for i, (h_idx, h_level, h_text, h_num) in enumerate(headings):
            chapter = self._paragraph_to_chapter.get(h_idx)
            if chapter is None:
                continue

            start = h_idx + 1

            # 遇到任何一个下一个标题（不论层级）就停止，保证只取"直属正文"
            if i + 1 < len(headings):
                end = headings[i + 1][0]
            else:
                end = len(self.document.paragraphs)

            # Assign body paragraphs (skip headings)
            for body_idx in range(start, end):
                if body_idx not in heading_indices:
                    chapter.body_paragraph_indices.append(body_idx)
                    self._paragraph_to_chapter[body_idx] = chapter

    def _iter_all_nodes(self) -> List[ChapterNode]:
        """Iterate all nodes in the tree (depth-first)."""
        result = []

        def dfs(nodes):
            for node in nodes:
                result.append(node)
                if node.children:
                    dfs(node.children)

        dfs(self.root_chapters)
        return result

    # ---- Lookup Methods ----

    def get_chapter_by_number(self, number: Union[str, Tuple[int, ...]]) -> Optional[ChapterNode]:
        """Look up a chapter by its number.

        Args:
            number: Either "3.1.1" string or (3, 1, 1) tuple

        Returns:
            ChapterNode or None if not found
        """
        if isinstance(number, str):
            # Normalize: strip leading/trailing whitespace and dots
            number = number.strip().strip('.')
            key = number
        elif isinstance(number, tuple):
            key = '.'.join(str(n) for n in number)
        else:
            return None

        return self._chapter_map.get(key)

    def get_chapter_contents(self, chapter: ChapterNode) -> List[ParagraphData]:
        """Get all paragraphs (heading + body) belonging to a chapter.

        Args:
            chapter: The chapter node

        Returns:
            List of ParagraphData objects in document order
        """
        contents = []

        # Add heading paragraph
        heading_idx = chapter.heading_paragraph_index
        if heading_idx < len(self.store.paragraphs_data):
            contents.append(self.store.paragraphs_data[heading_idx])

        # Add body paragraphs
        for body_idx in chapter.body_paragraph_indices:
            if body_idx < len(self.store.paragraphs_data):
                contents.append(self.store.paragraphs_data[body_idx])

        # Recursively add children's contents (preserving document order)
        for child in chapter.children:
            contents.extend(self.get_chapter_contents(child))

        return contents

    def get_chapter_text(self, chapter: ChapterNode) -> str:
        """Get concatenated text of a chapter (heading + body).

        Args:
            chapter: The chapter node

        Returns:
            Combined text of all paragraphs in the chapter
        """
        paragraphs = self.get_chapter_contents(chapter)
        return '\n'.join(p.text for p in paragraphs if p.text)

    # ---- Deletion ----

    def delete_chapter(self, chapter: ChapterNode):
        """Delete a chapter and all its contents from the document.

        This removes:
        - The heading paragraph
        - All body paragraphs
        - All sub-chapters (children) and their contents

        After deletion, the tree is rebuilt.
        """
        # Collect all indices to delete (this chapter + all descendants)
        indices_to_delete = set()
        self._collect_indices(chapter, indices_to_delete)

        # Delete from highest index to lowest (to preserve indices)
        for idx in sorted(indices_to_delete, reverse=True):
            if idx < len(self.document.paragraphs):
                para = self.document.paragraphs[idx]
                elem = para._element
                elem.getparent().remove(elem)

        # Rebuild tree
        self._rebuild()

    def _collect_indices(self, chapter: ChapterNode, indices: set):
        """Recursively collect paragraph indices for a chapter and all descendants.

        Args:
            chapter: Chapter to collect from
            indices: Set to add indices into
        """
        indices.add(chapter.heading_paragraph_index)
        for body_idx in chapter.body_paragraph_indices:
            indices.add(body_idx)
        for child in chapter.children:
            self._collect_indices(child, indices)

    def _rebuild(self):
        """Rebuild the chapter tree and paragraph data after structural changes."""
        from .format_extractor import FormatExtractor
        self.store.formats_json = FormatExtractor.extract_all(self.document)
        self._populate_paragraphs_data()
        self._build_tree()

    # ---- Utility ----

    def tree_to_string(self, indent: int = 2) -> str:
        """Return a visual representation of the chapter tree.

        Args:
            indent: Number of spaces per indent level

        Returns:
            Formatted tree string
        """
        lines = []

        def dfs(nodes, depth: int):
            for node in nodes:
                prefix = ' ' * (depth * indent)
                num_str = node.to_string()
                lines.append(f"{prefix}{num_str} {node.heading_text}")
                if node.children:
                    dfs(node.children, depth + 1)

        dfs(self.root_chapters, 0)
        return '\n'.join(lines)

    def list_chapters(self) -> List[ChapterNode]:
        """Get flat list of all chapters in document order.

        Returns:
            List of ChapterNode in document order (depth-first)
        """
        if self._all_nodes_cache is None:
            self._all_nodes_cache = self._iter_all_nodes()
        return self._all_nodes_cache
