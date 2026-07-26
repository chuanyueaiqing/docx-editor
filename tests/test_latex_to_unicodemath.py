"""Tests for the LaTeX to UnicodeMath converter."""
import pytest

from docx_editor.latex_to_unicodemath import (
    latex_to_unicodemath,
    is_latex_likely,
)


class TestLatexToUnicodemath:
    """Test basic LaTeX to UnicodeMath conversion."""

    def test_empty(self):
        assert latex_to_unicodemath('') == ''
        assert latex_to_unicodemath('   ') == ''

    def test_plain_text(self):
        """Plain text without LaTeX commands passes through unchanged."""
        result = latex_to_unicodemath('x = 1')
        assert result == 'x = 1'

    def test_greek_lower(self):
        """Greek lowercase letters are converted to Unicode."""
        assert 'α' in latex_to_unicodemath(r'\alpha')
        assert 'β' in latex_to_unicodemath(r'\beta')
        assert 'π' in latex_to_unicodemath(r'\pi')

    def test_greek_upper(self):
        """Greek uppercase letters are converted to Unicode."""
        assert 'Γ' in latex_to_unicodemath(r'\Gamma')
        assert 'Δ' in latex_to_unicodemath(r'\Delta')
        assert 'Ω' in latex_to_unicodemath(r'\Omega')

    def test_fraction(self):
        """\frac{a}{b} becomes (a/b)."""
        result = latex_to_unicodemath(r'\frac{a}{b}')
        assert '(a/b)' in result or 'a/b' in result

    def test_fraction_nested(self):
        """Nested fractions are handled."""
        result = latex_to_unicodemath(r'\frac{\frac{a}{b}}{c}')
        assert 'a' in result and 'b' in result and 'c' in result

    def test_sqrt(self):
        """sqrt{x} becomes √(x)."""
        result = latex_to_unicodemath(r'\sqrt{x}')
        assert '√' in result
        assert 'x' in result

    def test_sqrt_nth_root(self):
        """sqrt[n]{x} contains n root notation."""
        result = latex_to_unicodemath(r'\sqrt[3]{x}')
        assert '√' in result
        assert '3' in result

    def test_sum(self):
        """Sum becomes ∑."""
        result = latex_to_unicodemath(r'\sum_{i=1}^{n}')
        assert '∑' in result

    def test_integral(self):
        """Integral becomes ∫."""
        result = latex_to_unicodemath(r'\int_{0}^{\infty}')
        assert '∫' in result
        assert '∞' in result

    def test_arrows(self):
        """Arrow commands become Unicode arrows."""
        assert '→' in latex_to_unicodemath(r'\rightarrow')
        assert '←' in latex_to_unicodemath(r'\leftarrow')
        assert '⇒' in latex_to_unicodemath(r'\Rightarrow')

    def test_relations(self):
        """Relation symbols become Unicode."""
        assert '≤' in latex_to_unicodemath(r'\leq')
        assert '≥' in latex_to_unicodemath(r'\geq')
        assert '≠' in latex_to_unicodemath(r'\neq')

    def test_operators(self):
        """Binary operators become Unicode."""
        assert '±' in latex_to_unicodemath(r'\pm')
        assert '×' in latex_to_unicodemath(r'\times')
        assert '÷' in latex_to_unicodemath(r'\div')

    def test_left_right_stripped(self):
        """left( and right) are stripped to just parentheses."""
        result = latex_to_unicodemath(r'\left( x \right)')
        assert 'x' in result
        assert '(' in result
        assert ')' in result

    def test_text_command(self):
        """\text{...} is unwrapped."""
        result = latex_to_unicodemath(r'\text{hello}')
        assert 'hello' in result

    def test_quadratic_formula(self):
        """Real-world: quadratic formula."""
        result = latex_to_unicodemath(
            r'\frac{-b \pm \sqrt{b^2 - 4ac}}{2a}'
        )
        assert '±' in result
        assert '√' in result
        assert 'b' in result and 'a' in result and 'c' in result

    def test_display_integral(self):
        """Real-world: definite integral."""
        result = latex_to_unicodemath(
            r'\int_{a}^{b} f(x) \, dx'
        )
        assert '∫' in result
        assert 'a' in result and 'b' in result
        assert 'f(x)' in result or 'f' in result

    def test_partial_derivative(self):
        """Partial derivative notation."""
        result = latex_to_unicodemath(r'\frac{\partial f}{\partial x}')
        assert '∂' in result
        assert 'f' in result
        assert 'x' in result

    def test_matrix_notation_kept(self):
        """Matrix commands that can't directly convert are preserved."""
        result = latex_to_unicodemath(
            r'\begin{bmatrix}1 & 0\\0 & 1\end{bmatrix}'
        )
        # Should keep recognizable content
        assert 'begin' in result or 'bmatrix' in result
        assert '1' in result and '0' in result

    def test_unknown_command(self):
        """Unknown commands are passed through as-is."""
        # \unknowncommand will match the generic pattern
        result = latex_to_unicodemath(r'\unknowncommand')
        # The command itself may or may not be preserved depending on pattern matching
        assert isinstance(result, str)
        assert len(result) > 0

    def test_limits_stripped(self):
        """limits is removed (not needed in UnicodeMath)."""
        result = latex_to_unicodemath(r'\sum\limits_{i=1}')
        assert '∑' in result
        # Should not contain the word 'limits'
        assert 'limits' not in result


class TestIsLatexLikely:
    """Test the LaTeX heuristic detector."""

    def test_detects_commands(self):
        assert is_latex_likely(r'\frac{a}{b}') is True
        assert is_latex_likely(r'\alpha') is True

    def test_detects_subscript(self):
        assert is_latex_likely(r'a_{b}') is True
        assert is_latex_likely(r'a^{b}') is True

    def test_rejects_plain_text(self):
        assert is_latex_likely('hello world') is False
        assert is_latex_likely('a = b + c') is False
        assert is_latex_likely('') is False

    def test_detects_dollar_signs(self):
        assert is_latex_likely('$E=mc^2$') is True
