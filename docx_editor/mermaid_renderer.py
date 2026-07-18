"""Mermaid diagram rendering module.

Renders Mermaid diagram definitions to PNG images for embedding in DOCX documents.
Uses the @mermaid-js/mermaid-cli (mmdc) package via npx.
"""
import hashlib
import os
import subprocess
import sys
import tempfile
from typing import Optional

from .utils import MermaidRenderError


class MermaidNotAvailableError(MermaidRenderError):
    """Raised when the mermaid CLI tool is not available."""
    pass


class MermaidRenderer:
    """Render Mermaid diagram definitions to PNG images.

    Usage:
        renderer = MermaidRenderer()
        if renderer.is_available():
            png_path = renderer.render("graph TD; A-->B;")
    """

    # Default mmdc command
    DEFAULT_MMDC_CMD = 'npx.cmd' if sys.platform == 'win32' else 'npx'

    def __init__(self, output_dir: Optional[str] = None):
        self.output_dir = output_dir or tempfile.mkdtemp(prefix='mermaid_')
        self.mmdc_cmd = self._find_mmdc()
        self._cache: dict = {}  # code_hash -> output_path

    def _find_mmdc(self) -> Optional[str]:
        """Find the mmdc command. Returns None if not available.

        Checks:
        1. npx mmdc (auto-download from npm)
        2. Direct mmdc in PATH
        """
        # Check if npx is available
        try:
            result = subprocess.run(
                [self.DEFAULT_MMDC_CMD, '--version'],
                capture_output=True, text=True, timeout=10,
            )
            if result.returncode == 0:
                return self.DEFAULT_MMDC_CMD
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

        # Try direct mmdc
        try:
            result = subprocess.run(
                ['mmdc', '--version'],
                capture_output=True, text=True, timeout=10,
            )
            if result.returncode == 0:
                return 'mmdc'
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

        return None

    def is_available(self) -> bool:
        """Check if mmdc is available for rendering.

        Returns:
            True if mmdc can be invoked, False otherwise
        """
        return self.mmdc_cmd is not None

    def render(
        self,
        mermaid_code: str,
        output_path: Optional[str] = None,
        width: int = 800,
        timeout: int = 60,
    ) -> str:
        """Render a mermaid diagram to a PNG file.

        Args:
            mermaid_code: The mermaid diagram definition text
            output_path: Optional path for the output PNG.
                         If not provided, uses temp file.
            width: Output image width in pixels (default: 800)
            timeout: Maximum render time in seconds (default: 60)

        Returns:
            Path to the generated PNG file

        Raises:
            MermaidNotAvailableError: If mmdc is not available
            MermaidRenderError: If rendering fails
        """
        if not self.is_available():
            raise MermaidNotAvailableError(
                "mmdc (mermaid CLI) is not available. "
                "Install it with: npm install -g @mermaid-js/mermaid-cli"
            )

        # Check cache
        code_hash = self._code_hash(mermaid_code)
        if code_hash in self._cache:
            cached = self._cache[code_hash]
            if os.path.exists(cached):
                return cached

        # Write mermaid code to temp .mmd file
        mmd_path = self._write_mermaid_file(mermaid_code)

        # Determine output path
        if output_path is None:
            output_path = os.path.join(
                self.output_dir, f'{code_hash}.png'
            )

        try:
            # Build the mmdc command
            cmd = [
                self.mmdc_cmd,
                '-i', mmd_path,
                '-o', output_path,
                '-w', str(width),
                '-b', 'transparent',
            ]

            result = subprocess.run(
                cmd,
                capture_output=True, text=True, timeout=timeout,
            )

            if result.returncode != 0:
                raise MermaidRenderError(
                    f"mmdc failed (exit code {result.returncode}):\n"
                    f"stdout: {result.stdout}\n"
                    f"stderr: {result.stderr}"
                )

            if not os.path.exists(output_path):
                raise MermaidRenderError(
                    f"mmdc completed but output file not found: {output_path}"
                )

            # Cache result
            self._cache[code_hash] = output_path

            return output_path

        except subprocess.TimeoutExpired:
            raise MermaidRenderError(
                f"mmdc timed out after {timeout}s for mermaid code:\n{mermaid_code[:200]}"
            )
        except FileNotFoundError as e:
            raise MermaidNotAvailableError(
                f"mmdc command not found: {e}"
            )
        finally:
            # Clean up temp .mmd file
            try:
                if os.path.exists(mmd_path):
                    os.unlink(mmd_path)
            except Exception:
                pass

    def render_and_get_image_data(self, mermaid_code: str) -> bytes:
        """Render mermaid and return PNG bytes.

        Args:
            mermaid_code: The mermaid diagram definition text

        Returns:
            PNG image bytes
        """
        path = self.render(mermaid_code)
        with open(path, 'rb') as f:
            data = f.read()
        return data

    def render_to_temp(self, mermaid_code: str) -> str:
        """Render to a temp file and return the path.

        Caller is responsible for cleanup.

        Args:
            mermaid_code: The mermaid diagram definition text

        Returns:
            Path to the generated PNG file
        """
        return self.render(mermaid_code)

    def _write_mermaid_file(self, mermaid_code: str) -> str:
        """Write mermaid code to a temporary .mmd file.

        Args:
            mermaid_code: The mermaid diagram definition text

        Returns:
            Path to the written .mmd file
        """
        code_hash = self._code_hash(mermaid_code)
        mmd_path = os.path.join(self.output_dir, f'{code_hash}.mmd')

        with open(mmd_path, 'w', encoding='utf-8') as f:
            f.write(mermaid_code)

        return mmd_path

    @staticmethod
    def _code_hash(code: str) -> str:
        """Generate a hash for mermaid code for caching purposes.

        Args:
            code: The mermaid diagram definition text

        Returns:
            MD5 hash string (first 16 chars)
        """
        return hashlib.md5(code.encode('utf-8')).hexdigest()[:16]

    def clear_cache(self):
        """Clear the rendering cache."""
        self._cache.clear()
