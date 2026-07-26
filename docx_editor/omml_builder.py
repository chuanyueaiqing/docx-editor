"""LaTeX to OMML (Office Math Markup Language) XML builder.

Generates proper OMML XML elements from LaTeX math expressions,
which can be injected directly into python-docx documents via the
oxml layer.  This produces native Word equations with proper
structure (fractions, radicals, superscripts, etc.).

Usage:
    from docx_editor.omml_builder import latex_to_omml

    # Returns lxml elements ready for insertion
    omath_elements = latex_to_omml(r'E=mc^2')
    omath_elements = latex_to_omml(r'\frac{a}{b}')
"""

import logging
import re
from typing import List, Optional

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .latex_to_unicodemath import GREEK_MAP, SYMBOL_MAP

logger = logging.getLogger(__name__)

# OMML namespace URI
OMML_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

# Default math font for equations
DEFAULT_MATH_FONT = 'Cambria Math'


def _m(tag: str) -> str:
    """Create a Clark-notation tag for the OMML namespace."""
    return f'{{{OMML_URI}}}{tag}'


def _make(tag: str) -> 'OxmlElement':
    """Create an OxmlElement with the given OMML tag (e.g. 'oMath')."""
    return OxmlElement(f'm:{tag}')


def _r(text: str, math_font: str = '') -> 'OxmlElement':
    """Create an OMML run (<m:r>) with text content.

    Includes both ``<m:rPr>`` (math style) and ``<w:rPr>`` (font/size)
    properties, matching what Word natively generates.  This ensures
    proper rendering in both Word and WPS.

    Args:
        text: Run text content
        math_font: Font name for the math run. Empty string means
            use DEFAULT_MATH_FONT.
    """
    if not math_font:
        math_font = DEFAULT_MATH_FONT

    r = _make('r')

    # 1. OMML run properties: plain style
    oMathPr = _make('rPr')
    sty = _make('sty')
    sty.set(_m('val'), 'p')
    oMathPr.append(sty)
    r.append(oMathPr)

    # 2. Word run properties: configurable font + 12pt size
    w_rPr = OxmlElement('w:rPr')
    w_rFonts = OxmlElement('w:rFonts')
    w_rFonts.set(qn('w:ascii'), math_font)
    w_rFonts.set(qn('w:hAnsi'), math_font)
    w_rPr.append(w_rFonts)
    w_sz = OxmlElement('w:sz')
    w_sz.set(qn('w:val'), '24')  # 12pt in half-points
    w_rPr.append(w_sz)
    w_szCs = OxmlElement('w:szCs')
    w_szCs.set(qn('w:val'), '22')
    w_rPr.append(w_szCs)
    r.append(w_rPr)

    # 3. Text content
    t = _make('t')
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    r.append(t)
    return r


def _r_with_greek(text: str) -> List['OxmlElement']:
    """Create multiple runs splitting out Greek/Unicode characters.

    Some characters (like ∑, ∫, α, β, etc.) render best as plain Unicode
    in OMML runs.  This method ensures each run has valid text content.
    """
    return [_r(text)]


# ── Parsing: LaTeX tokenizer ──

# Matches: \command, { } ^ _ &, words, operators, or any single char
TOKEN_PATTERN = re.compile(
    r'\\[a-zA-Z]+'           # LaTeX command
    r'|[{}^_&]'              # special chars
    r'|[a-zA-Z0-9]+'         # word (variable/number)
    r'|[+\-=\/()\[\],;:!@#\'`~|<>*%.]'  # operators / punctuation
    r'|\s+'                  # whitespace
    r'|.'                    # any other single char
)


class OMMLBuilder:
    """Parse LaTeX and build OMML element tree.

    Args:
        latex: LaTeX math expression (without $ delimiters)
        math_font: Font name for the math run. If empty or None,
            uses DEFAULT_MATH_FONT ('Cambria Math').
    """

    def __init__(self, latex: str, math_font: str = ''):
        self.latex = latex
        self.math_font = math_font or DEFAULT_MATH_FONT
        self.pos = 0
        self.tokens: List[str] = []
        self._tokenize()

    def _mr(self, text: str) -> 'OxmlElement':
        """Create a math run using the configured math font."""
        return _r(text, self.math_font)

    def _tokenize(self):
        """Tokenize the LaTeX string."""
        text = self.latex
        i = 0
        while i < len(text):
            m = TOKEN_PATTERN.match(text, i)
            if m:
                self.tokens.append(m.group(0))
                i = m.end()
            else:
                # Single character
                self.tokens.append(text[i])
                i += 1

    def peek(self) -> Optional[str]:
        if self.pos < len(self.tokens):
            return self.tokens[self.pos]
        return None

    def consume(self) -> Optional[str]:
        t = self.peek()
        if t is not None:
            self.pos += 1
        return t

    def expect(self, expected: str) -> str:
        t = self.consume()
        if t != expected:
            raise ValueError(f'Expected {expected!r}, got {t!r}')
        return t

    def skip_whitespace(self):
        """Skip whitespace tokens."""
        while self.peek() in (' ', '\t'):
            self.consume()

    def parse_group(self) -> Optional['OxmlElement']:
        """Parse a group: { expression } or a single token."""
        self.skip_whitespace()
        if self.peek() == '{':
            self.consume()  # {
            args = self.parse_expression(stop_at='}')
            self.skip_whitespace()
            if self.peek() == '}':
                self.consume()  # }
            if len(args) == 1:
                return args[0]
            # Multiple items in group: wrap in a base element
            base = _make('e')
            for a in args:
                base.append(a)
            return base
        else:
            return self.parse_atom()

    def parse_expression(self, stop_at: Optional[str] = None) -> List:
        """Parse a sequence of tokens until stop token or end.

        After each base atom, checks for ^ (superscript) and _ (subscript)
        to attach them to the base properly.
        """
        args = []
        while self.pos < len(self.tokens):
            t = self.peek()
            if t == stop_at:
                break

            # Parse the base element
            elem = self.parse_base()
            if elem is not None:
                args.append(elem)
        return args

    def parse_base(self):
        """Parse a base element with optional ^/_ attachments.

        Handles: atom [^ group] [_ group]  — attaching sup/sub to base.
        """
        base = self.parse_atom()
        if base is None:
            return None

        self.skip_whitespace()

        # Check for ^ and _ after the base
        while self.peek() in ('^', '_'):
            op = self.consume()
            arg = self.parse_group()
            self.skip_whitespace()

            if op == '^':
                sup = _make('sup')
                base_e = _make('e')
                base_e.append(base)
                sup_e = _make('e')
                if arg is not None:
                    sup_e.append(arg)
                sup.append(base_e)
                sup.append(sup_e)
                base = sup
            elif op == '_':
                sub = _make('sub')
                base_e = _make('e')
                base_e.append(base)
                sub_e = _make('e')
                if arg is not None:
                    sub_e.append(arg)
                sub.append(base_e)
                sub.append(sub_e)
                base = sub

        return base

    def parse_atom(self):
        """Parse a single atom (command, variable, operator, or group)."""
        t = self.consume()
        if t is None:
            return None

        # ── LaTeX commands ──
        if t.startswith('\\'):
            cmd = t[1:]  # strip backslash

            # Fractions
            if cmd == 'frac':
                num = self.parse_group()
                den = self.parse_group()
                f = _make('f')
                num_e = _make('num')
                if num is not None:
                    num_e.append(num)
                den_e = _make('den')
                if den is not None:
                    den_e.append(den)
                f.append(num_e)
                f.append(den_e)
                return f

            # Square root
            elif cmd == 'sqrt':
                rad = _make('rad')
                self.skip_whitespace()
                deg = _make('deg')
                if self.peek() == '[':
                    self.consume()  # [
                    deg_expr = self.parse_expression(stop_at=']')
                    if self.peek() == ']':
                        self.consume()
                    for d in deg_expr:
                        deg.append(d)
                rad.append(deg)
                radicand = self.parse_group()
                e = _make('e')
                if radicand is not None:
                    e.append(radicand)
                rad.append(e)
                return rad

            # Sum, Integral, Product (nary operators)
            elif cmd in ('sum', 'int', 'iint', 'iiint', 'oint', 'prod'):
                return self._build_nary(cmd)

            # Greek letters → Unicode (GREEK_MAP keys include backslash)
            elif '\\' + cmd in GREEK_MAP:
                return self._mr(GREEK_MAP['\\' + cmd])

            # Symbols → Unicode (SYMBOL_MAP keys include backslash)
            elif '\\' + cmd in SYMBOL_MAP:
                return self._mr(SYMBOL_MAP['\\' + cmd])

            # Function names (sin, cos, log, etc.) → plain text
            elif cmd in ('sin', 'cos', 'tan', 'cot', 'sec', 'csc',
                         'arcsin', 'arccos', 'arctan',
                         'sinh', 'cosh', 'tanh',
                         'log', 'ln', 'lg',
                         'lim', 'max', 'min', 'sup', 'inf',
                         'det', 'dim', 'Pr', 'exp', 'mod'):
                return self._mr(cmd)

            # \mathbf, \mathrm, \mathcal, etc. — just the content
            elif cmd in ('mathbf', 'mathrm', 'mathcal', 'mathit',
                         'mathsf', 'mathtt', 'text'):
                group = self.parse_group()
                return group

            # \left, \right, \bigl, etc. — ignore
            elif cmd in ('left', 'right', 'bigl', 'bigr', 'Bigl', 'Bigr',
                         'big', 'Big', 'bigg', 'Bigg'):
                return self.parse_atom()

            # \to, \gets, \mapsto
            elif cmd == 'to':
                return self._mr('→')
            elif cmd == 'gets':
                return self._mr('←')
            elif cmd == 'mapsto':
                return self._mr('↦')

            # Spacing commands
            elif cmd in ('quad', 'qquad', ','):
                return self._mr(' ')
            elif cmd == '!':
                return None  # negative space

            # Unknown command — output the raw text
            else:
                logger.debug('Unknown LaTeX command: %s', t)
                return self._mr(t)

        # ── Operators and punctuation ──
        elif t in ('+', '-', '=', '/', '(', ')', '[', ']', '|',
                   '<', '>', ',', '.', '!', '?', ':', ';', '@', '#',
                   '%', '&', '*', '~', '`', '"', "'"):
            return self._mr(t)

        # ── Numbers and letters ──
        else:
            return self._mr(t)

    def _build_nary(self, cmd: str) -> 'OxmlElement':
        """Build an n-ary operator (sum, integral, product)."""
        nary = _make('nary')
        nary_pr = _make('naryPr')

        # Operator symbol
        symb = _make('chr')
        op_map = {
            'sum': '∑', 'int': '∫', 'iint': '∬', 'iiint': '∭',
            'oint': '∮', 'prod': '∏',
        }
        symb.text = op_map.get(cmd, '∑')
        nary_pr.append(symb)

        # Limits location (subscript/superscript for sum, underscript/overscript for int)
        lim_loc = _make('limLoc')
        if cmd in ('sum', 'prod'):
            # Use subscript/superscript for sum and prod (default)
            lim_loc.set(_m('val'), 'subSup')
        else:
            # Use underOver for integrals
            lim_loc.set(_m('val'), 'undOvr')
        nary_pr.append(lim_loc)

        # Grow operator (display style)
        grow = _make('grow')
        grow.set(_m('val'), '1')
        nary_pr.append(grow)

        nary.append(nary_pr)

        # Subscript (lower limit)
        # Check for _ and ^ after the command
        self.skip_whitespace()
        if self.peek() == '_':
            self.consume()
            sub_arg = self.parse_group()
            sub_e = _make('sub')
            if sub_arg is not None:
                sub_e.append(sub_arg)
            nary.append(sub_e)
        else:
            nary.append(_make('sub'))  # empty

        # Superscript (upper limit)
        self.skip_whitespace()
        if self.peek() == '^':
            self.consume()
            sup_arg = self.parse_group()
            sup_e = _make('sup')
            if sup_arg is not None:
                sup_e.append(sup_arg)
            nary.append(sup_e)
        else:
            nary.append(_make('sup'))  # empty

        # Expression (operand)
        e = _make('e')
        nary.append(e)

        return nary

    def build(self) -> List:
        """Build OMML elements from LaTeX.

        Returns:
            List of OMML OxmlElement objects for injection into a paragraph
        """
        elements = self.parse_expression()
        return elements


def latex_to_omml_elements(latex: str, *, math_font: str = '') -> List:
    """Convert LaTeX to a list of OMML OxmlElement objects.

    Args:
        latex: LaTeX math expression (without $ delimiters)
        math_font: Font for math runs. Empty = default (Cambria Math).

    Returns:
        List of OMML XML elements for injection into a paragraph
    """
    builder = OMMLBuilder(latex, math_font=math_font)
    return builder.build()


def create_omath_paragraph(latex: str, *, math_font: str = '') -> 'OxmlElement':
    """Create a full ``<m:oMath>`` element containing the formula.

    Args:
        latex: LaTeX math expression
        math_font: Font for math runs. Empty = default (Cambria Math).

    Returns:
        ``<m:oMath>`` OxmlElement
    """
    omath = _make('oMath')
    elements = latex_to_omml_elements(latex, math_font=math_font)
    for elem in elements:
        if elem is not None:
            omath.append(elem)
    return omath


def create_omath_para_element(latex: str, *, math_font: str = '') -> 'OxmlElement':
    """Create a ``<m:oMathPara>`` element (display math).

    For display math ($$...$$) — wraps oMath in oMathPara.

    Args:
        latex: LaTeX math expression
        math_font: Font for math runs. Empty = default (Cambria Math).

    Returns:
        ``<m:oMathPara>`` OxmlElement
    """
    omath_para = _make('oMathPara')
    omath = create_omath_paragraph(latex, math_font=math_font)
    omath_para.append(omath)
    return omath_para
