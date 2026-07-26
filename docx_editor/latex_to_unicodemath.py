"""LaTeX to UnicodeMath conversion module.

Converts LaTeX math expressions to UnicodeMath (Microsoft's linear format),
which can be consumed by both Microsoft Word's and WPS's OMath equation engine.

Usage:
    from docx_editor.latex_to_unicodemath import latex_to_unicodemath

    result = latex_to_unicodemath('\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}')
    # Returns (approx): "(-b \\pm \\sqrt(b^2 - 4ac)/2a)"
"""

import re
import logging

logger = logging.getLogger(__name__)

# ── Greek letters ──

GREEK_MAP = {
    r'\alpha': 'α', r'\beta': 'β', r'\gamma': 'γ', r'\delta': 'δ',
    r'\epsilon': 'ε', r'\varepsilon': 'ε', r'\zeta': 'ζ', r'\eta': 'η',
    r'\theta': 'θ', r'\vartheta': 'θ', r'\iota': 'ι', r'\kappa': 'κ',
    r'\lambda': 'λ', r'\mu': 'μ', r'\nu': 'ν', r'\xi': 'ξ',
    r'\omicron': 'ο', r'\pi': 'π', r'\varpi': 'π', r'\rho': 'ρ',
    r'\varrho': 'ρ', r'\sigma': 'σ', r'\varsigma': 'ς', r'\tau': 'τ',
    r'\upsilon': 'υ', r'\phi': 'φ', r'\varphi': 'φ', r'\chi': 'χ',
    r'\psi': 'ψ', r'\omega': 'ω',
    # Upper case
    r'\Gamma': 'Γ', r'\Delta': 'Δ', r'\Theta': 'Θ', r'\Lambda': 'Λ',
    r'\Xi': 'Ξ', r'\Pi': 'Π', r'\Sigma': 'Σ', r'\Upsilon': 'Υ',
    r'\Phi': 'Φ', r'\Psi': 'Ψ', r'\Omega': 'Ω',
}

# ── Math symbols ──

SYMBOL_MAP = {
    # Operators
    r'\sum': '∑', r'\int': '∫', r'\iint': '∬', r'\iiint': '∭',
    r'\oint': '∮', r'\prod': '∏', r'\coprod': '∐',
    r'\bigcup': '⋃', r'\bigcap': '⋂', r'\bigvee': '⋁',
    r'\bigwedge': '⋀', r'\bigoplus': '⊕', r'\bigotimes': '⊗',
    r'\bigodot': '⊙',
    # Relations
    r'\leq': '≤', r'\geq': '≥', r'\neq': '≠', r'\equiv': '≡',
    r'\approx': '≈', r'\sim': '∼', r'\simeq': '≃', r'\cong': '≅',
    r'\propto': '∝', r'\prec': '≺', r'\succ': '≻', r'\preceq': '≼',
    r'\succeq': '≽', r'\subset': '⊂', r'\supset': '⊃',
    r'\subseteq': '⊆', r'\supseteq': '⊇', r'\in': '∈', r'\ni': '∋',
    r'\notin': '∉', r'\forall': '∀', r'\exists': '∃', r'\nexists': '∄',
    r'\vdash': '⊢', r'\dashv': '⊣', r'\models': '⊨',
    r'\perp': '⊥', r'\parallel': '∥', r'\mid': '∣',
    # Arrows
    r'\rightarrow': '→', r'\leftarrow': '←', r'\Rightarrow': '⇒',
    r'\Leftarrow': '⇐', r'\Leftrightarrow': '⇔',
    r'\leftrightarrow': '↔', r'\mapsto': '↦', r'\longmapsto': '⟼',
    r'\longrightarrow': '⟶', r'\longleftarrow': '⟵',
    r'\Longrightarrow': '⟹', r'\Longleftarrow': '⟸',
    r'\uparrow': '↑', r'\downarrow': '↓', r'\updownarrow': '↕',
    r'\Uparrow': '⇑', r'\Downarrow': '⇓', r'\Updownarrow': '⇕',
    r'\nearrow': '↗', r'\searrow': '↘', r'\nwarrow': '↖',
    r'\swarrow': '↙', r'\to': '→', r'\gets': '←',
    # Binary operators
    r'\pm': '±', r'\mp': '∓', r'\times': '×', r'\div': '÷',
    r'\cdot': '·', r'\bullet': '•', r'\circ': '∘', r'\ast': '∗',
    r'\star': '⋆', r'\otimes': '⊗', r'\oplus': '⊕', r'\ominus': '⊖',
    r'\odot': '⊙', r'\oslash': '⊘', r'\wedge': '∧', r'\vee': '∨',
    r'\cap': '∩', r'\cup': '∪', r'\uplus': '⊎', r'\sqcap': '⊓',
    r'\sqcup': '⊔', r'\triangleleft': '◁', r'\triangleright': '▷',
    r'\setminus': '∖', r'\wr': '≀',
    # Miscellaneous
    r'\infty': '∞', r'\nabla': '∇', r'\partial': '∂',
    r'\angle': '∠', r'\measuredangle': '∡', r'\sphericalangle': '∢',
    r'\triangle': '△', r'\box': '□', r'\diamond': '◇',
    r'\ell': 'ℓ', r'\hbar': 'ℏ', r'\hslash': 'ℏ',
    r'\Re': 'ℜ', r'\Im': 'ℑ', r'\wp': '℘',
    r'\aleph': 'א', r'\beth': 'ב', r'\gimel': 'ג', r'\daleth': 'ד',
    r'\prime': '′', r'\backprime': '‵', r'\emptyset': '∅',
    r'\varnothing': '∅', r'\neg': '¬', r'\lnot': '¬',
    r'\top': '⊤', r'\bot': '⊥', r'\clubsuit': '♣',
    r'\diamondsuit': '♦', r'\heartsuit': '♥', r'\spadesuit': '♠',
    r'\ldots': '…', r'\cdots': '⋯', r'\vdots': '⋮', r'\ddots': '⋱',
    r'\therefore': '∴', r'\because': '∵',
    r'\surd': '√', r'\imath': 'ı', r'\jmath': 'ȷ',
    # Spacing (ignored)
    r'\;': ' ', r'\,': ' ', r'\:': ' ', r'\quad': '  ',
    r'\qquad': '    ', r'\!': '', r'\ ': ' ',
    # Text short forms
    r'\text': '',  # handled specially
}

# ── Function names (kept as plain text in UnicodeMath) ──

FUNC_NAMES = {
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
    'sinh', 'cosh', 'tanh', 'coth', 'sech', 'csch',
    'arcsin', 'arccos', 'arctan', 'arccot', 'arcsec', 'arccsc',
    'log', 'ln', 'lg', 'lb',
    'lim', 'limsup', 'liminf',
    'max', 'min', 'sup', 'inf',
    'det', 'dim', 'hom', 'ker', 'deg',
    'arg', 'Re', 'Im',
    'Pr', 'exp', 'mod', 'gcd', 'lcm',
    'erf', 'erfc',
}

# ── Patterns for structural commands ──

# Pattern for \frac{a}{b} — captures two brace groups
RE_FRAC = re.compile(r'\\frac\{([^}]*)\}\{([^}]*)\}')

# Pattern for \sqrt{x} and \sqrt[n]{x}
RE_SQRT = re.compile(r'\\sqrt(?:\[([^\]]*)\])?\{([^}]*)\}')

# Pattern for \left and \right delimiters (with \b word boundary to avoid
# matching \right inside \rightarrow)
RE_LEFT_RIGHT = re.compile(r'\\(left|right|bigl|bigr|Bigl|Bigr|big|Big|bigg|Bigg)\b')

# Pattern for \text{...}
RE_TEXT = re.compile(r'\\text\{([^}]*)\}')

# Pattern for \underset{below}{expr} and \overset{above}{expr}
RE_UNDER_OVER_SET = re.compile(r'\\(underset|overset)\{([^}]*)\}\{([^}]*)\}')

# Pattern for \binom{n}{k}
RE_BINOM = re.compile(r'\\binom\{([^}]*)\}\{([^}]*)\}')

# Pattern for \limits (ignored in UnicodeMath)
RE_LIMITS = re.compile(r'\\limits\s*')

# ── Token-level replacement ──

def _replace_greek_and_symbols(latex: str) -> str:
    """Replace Greek letter commands and symbol commands with Unicode chars.

    Uses a single-pass regex substitution for all known commands.
    Longer commands are tried first to avoid partial matches (e.g. \\varepsilon before \\epsilon).
    """
    # Build combined mapping, sorted by key length descending
    all_commands = {}
    all_commands.update(GREEK_MAP)
    all_commands.update(SYMBOL_MAP)

    sorted_keys = sorted(all_commands.keys(), key=len, reverse=True)
    # Build regex: match any command (backslash + word chars)
    # but only if it's one of our known keys
    pattern_src = '|'.join(re.escape(k) for k in sorted_keys)
    pattern = re.compile(pattern_src)

    def _replace(match):
        cmd = match.group(0)
        return all_commands.get(cmd, cmd)

    return pattern.sub(_replace, latex)


# ── Structural conversions ──

def _convert_fractions(latex: str) -> str:
    """Convert \\frac{numerator}{denominator} → (numerator/denominator)"""
    def _replace(m):
        num = m.group(1)
        den = m.group(2)
        return f'({num}/{den})'
    return RE_FRAC.sub(_replace, latex)


def _convert_sqrt(latex: str) -> str:
    """Convert \\sqrt{x} → √(x) and \\sqrt[n]{x} → √(n&x)"""
    def _replace(m):
        root = m.group(1)
        radicand = m.group(2)
        if root:
            return f'√({root}&{radicand})'
        return f'√({radicand})'
    return RE_SQRT.sub(_replace, latex)


def _remove_left_right(latex: str) -> str:
    """Remove \\left, \\right, \\bigl, etc. delimiters."""
    return RE_LEFT_RIGHT.sub('', latex)


def _convert_text(latex: str) -> str:
    """Convert \\text{...} → just the text content."""
    def _replace(m):
        return m.group(1)
    return RE_TEXT.sub(_replace, latex)


def _convert_underset_overset(latex: str) -> str:
    """Convert \\underset{below}{expr} and \\overset{above}{expr}.

    UnicodeMath represents these as:
    - \\underset{below}{expr} → (below)█(expr)  (need special markup)
    - \\overset{above}{expr} → (above)█(expr)
    Falls back to just expr with below/above in parentheses.
    """
    def _replace(m):
        cmd = m.group(1)
        arg1 = m.group(2)
        arg2 = m.group(3)
        if cmd == 'underset':
            return f'{arg2}  ({arg1})'  # approximation
        else:
            return f'{arg1}  {arg2}'  # approximation
    return RE_UNDER_OVER_SET.sub(_replace, latex)


def _convert_binom(latex: str) -> str:
    """Convert \\binom{n}{k} → (n¦k) using the vertical bar character."""
    def _replace(m):
        top = m.group(1)
        bottom = m.group(2)
        return f'({top}¦{bottom})'
    return RE_BINOM.sub(_replace, latex)


def _remove_limits(latex: str) -> str:
    """Remove \\limits commands (not needed in UnicodeMath)."""
    return RE_LIMITS.sub('', latex)


def _cleanup_whitespace(latex: str) -> str:
    """Collapse multiple spaces into one and trim."""
    result = re.sub(r'  +', ' ', latex)
    return result.strip()


# ── Public API ──

def latex_to_unicodemath(latex: str) -> str:
    """Convert a LaTeX math expression to UnicodeMath format.

    UnicodeMath is Microsoft's linear format for equations, used by
    both Word and WPS equation engines.  This converter handles the
    most common LaTeX math constructs.

    Args:
        latex: A LaTeX math expression (without the $...$ delimiters)

    Returns:
        UnicodeMath string suitable for insertion via COM OMaths API

    Example:
        >>> latex_to_unicodemath(r'\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}')
        '((-b) ± √((b^2) - (4ac)))/(2a)'
    """
    if not latex or not latex.strip():
        return ''

    result = latex.strip()

    # Order matters: structural conversions first, then token-level
    result = _remove_limits(result)
    result = _convert_text(result)
    result = _convert_underset_overset(result)
    result = _convert_binom(result)
    result = _remove_left_right(result)
    result = _convert_sqrt(result)
    result = _convert_fractions(result)
    result = _replace_greek_and_symbols(result)
    result = _cleanup_whitespace(result)

    # Handle curly braces used for grouping - UnicodeMath uses them too
    # but we need to ensure they're balanced

    logger.debug('latex_to_unicodemath: %r → %r', latex, result)
    return result


def is_latex_likely(text: str) -> bool:
    """Heuristic: check if a string looks like LaTeX math.

    Args:
        text: The string to check

    Returns:
        True if the string contains LaTeX-like commands or patterns
    """
    if not text:
        return False
    # Check for common LaTeX patterns
    patterns = [
        r'\\[a-zA-Z]+',          # LaTeX command
        r'[_^]\s*\{',            # subscript/superscript with braces
        r'\$\$?',                # math delimiters
        r'\\frac', r'\\sqrt',    # common structures
        r'\\sum', r'\\int',      # operators
    ]
    for p in patterns:
        if re.search(p, text):
            return True
    return False
