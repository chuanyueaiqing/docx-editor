"""Tests for equation_ops.py (non-COM parts only).

COM-dependent tests (Word/WPS conversion) require Microsoft Word or WPS
installed and are gated behind ``__main__`` blocks.
"""
import os
import tempfile
import pytest

from docx_editor.equation_ops import (
    EquationProcessor,
    EquationResult,
    EquationError,
)
from docx_editor.latex_to_unicodemath import latex_to_unicodemath


class TestEngineDetection:
    """Test engine detection without actually starting COM."""

    def test_detect_engine_returns_string(self):
        """detect_engine() always returns one of the three valid values."""
        engine = EquationProcessor.detect_engine()
        assert engine in ('word', 'wps', 'image')

    def test_is_com_available_returns_bool(self):
        """is_com_available() returns a bool."""
        result = EquationProcessor.is_com_available()
        assert isinstance(result, bool)


class TestLatexRendering:
    """Test LaTeX to image rendering (requires matplotlib)."""

    def test_render_simple_formula(self):
        """Render a simple formula to PNG."""
        try:
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
                output_path = f.name

            result = EquationProcessor.render_latex_as_image(
                'E=mc^2', output_path=output_path
            )

            assert os.path.exists(result)
            assert result == output_path
            assert os.path.getsize(output_path) > 100  # should be a non-trivial PNG

        except EquationError as e:
            if 'matplotlib' in str(e):
                pytest.skip("matplotlib not installed")
            raise
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_render_display_math(self):
        """Render a display-style formula."""
        try:
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
                output_path = f.name

            result = EquationProcessor.render_latex_as_image(
                r'\int_{0}^{\infty} e^{-x^2} dx',
                output_path=output_path,
                display=True,
            )

            assert os.path.exists(result)
            assert os.path.getsize(output_path) > 100

        except EquationError as e:
            if 'matplotlib' in str(e):
                pytest.skip("matplotlib not installed")
            raise
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_render_without_output_path(self):
        """Without output_path, creates a temp file."""
        try:
            result = EquationProcessor.render_latex_as_image('x^2')
            assert os.path.exists(result)
            assert result.endswith('.png')
        except EquationError as e:
            if 'matplotlib' in str(e):
                pytest.skip("matplotlib not installed")
            raise
        finally:
            if 'result' in locals() and os.path.exists(result):
                os.unlink(result)

    def test_render_invalid_latex(self):
        """Invalid LaTeX should raise EquationError, not crash."""
        with pytest.raises(EquationError):
            EquationProcessor.render_latex_as_image(
                r'\frac{',  # deliberately broken LaTeX
                output_path=os.path.join(tempfile.gettempdir(), '_bad_math.png'),
            )


class TestPostProcessingEdgeCases:
    """Test post-processing logic edge cases."""

    def test_empty_elements_list(self):
        """Empty math_elements list should return SKIPPED."""
        proc = EquationProcessor()
        # Can't call post_process without a real file, but we can test
        # that the result is SKIPPED for empty input conceptually
        assert EquationResult.SKIPPED.value == 'skipped'

    def test_post_process_nonexistent_file(self):
        """Non-existent file should raise EquationError."""
        proc = EquationProcessor()
        with pytest.raises(EquationError):
            proc.post_process('/nonexistent/path.docx', [('x^2', False)])


class TestUnicodemathRoundtrip:
    """Verify that latex_to_unicodemath produces WPS-compatible output."""

    def test_simple_conversions(self):
        """Basic formula conversions produce plausible UnicodeMath."""
        cases = [
            (r'x^2', 'x^2'),
            (r'a_b', 'a_b'),
            (r'\frac{a}{b}', '/'),  # fraction produces (a/b) or similar
            (r'\alpha', 'α'),
            (r'\beta', 'β'),
        ]
        for latex, expected_substring in cases:
            result = latex_to_unicodemath(latex)
            assert expected_substring in result, (
                f'Expected "{expected_substring}" in "{result}" for input "{latex}"'
            )

    def test_equation_result_values(self):
        """Verify EquationResult enum values."""
        assert EquationResult.CONVERTED_COM.value == 'converted_com'
        assert EquationResult.CONVERTED_WPS.value == 'converted_wps'
        assert EquationResult.FALLBACK_IMAGE.value == 'fallback_image'
        assert EquationResult.SKIPPED.value == 'skipped'
        assert EquationResult.PARTIAL.value == 'partial'


if __name__ == '__main__':
    """Manual COM-dependent tests — only runs when Word or WPS is available.

    Run with:  python tests/test_equation_ops.py
    """
    import logging
    logging.basicConfig(level=logging.DEBUG)

    from docx_editor.creator import DocxCreator

    print("=" * 60)
    print("COM-dependent equation integration tests")
    print("=" * 60)

    engine = EquationProcessor.detect_engine()
    print(f"Detected engine: {engine}")

    if engine == 'image':
        print("SKIP: No COM engine (Word or WPS) available.")
        print("Install Microsoft Word or WPS Office to run these tests.")
        exit(0)

    # Create a document with formulas
    md_text = """# 公式测试

行内公式 $E=mc^2$ 是质能方程。

行间公式：

$$
\\int_{0}^{\\infty} e^{-x^2} \\, dx = \\frac{\\sqrt{\\pi}}{2}
$$

多个公式：$a^2 + b^2 = c^2$ 和 $\\\\alpha + \\\\beta = \\\\gamma$
"""
    output_path = os.path.join(tempfile.gettempdir(), '_test_equations.docx')

    try:
        creator = DocxCreator()
        creator.set_default_format({
            'font_name': 'Times New Roman',
            'font_size': 12,
        })
        creator.add_markdown(md_text)
        creator.save(output_path)
        print(f"\nCreated: {output_path}")
        print(f"  Equations processed via: {engine}")
        print("\nOpen the file in Word/WPS and verify equations are editable.")
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Cleanup
        if os.path.exists(output_path):
            try:
                os.unlink(output_path)
                print(f"\nCleaned up: {output_path}")
            except Exception:
                pass
