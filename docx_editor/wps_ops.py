"""WPS Office COM operations module.

Provides WPS COM automation for tracked changes (revision mode),
as an independent alternative to ``Win32Ops`` (which requires
Microsoft Word).

Usage:
    with WpsOps() as ops:
        ops.open_document('path.docx')
        ops.word.ActiveDocument.Compare(
            Name='modified.docx', CompareTarget=2,
        )
        ops.save_document('output.docx')
"""
import logging
import os
from typing import Optional

from .utils import DocxError

logger = logging.getLogger(__name__)


class WpsComError(DocxError):
    """WPS COM operation failed."""


class WpsOps:
    """Manage WPS Application lifecycle via COM.

    Dispatches ``KWPS.Application`` (WPS's COM prog ID) rather than
    ``Word.Application``.  The WPS COM interface is largely compatible
    with Microsoft Word's, so ``Document.Compare()`` and related APIs
    work identically.

    Usage:
        with WpsOps() as ops:
            ops.open_document('path.docx')
            # Use ops.word / ops.wd_doc for COM calls
            ops.save_document('output.docx')
    """

    # WPS.Application CLSID (known value for WPS Office)
    WPS_CLSID = "{000209FF-0000-0000-C000-000000000047}"

    def __init__(self):
        self.word = None
        self.wd_doc = None
        self._initialized = False
        self._pythoncom = None

    def __enter__(self):
        """Initialize COM and start WPS Application."""
        try:
            import pythoncom
            import win32com.client
            self._pythoncom = pythoncom
            self._win32com_client = win32com.client
        except ImportError as e:
            raise WpsComError(
                "pywin32 is not installed. Install with: pip install pywin32"
            ) from e

        try:
            pythoncom.CoInitialize()
            self.word = win32com.client.Dispatch("KWPS.Application")
            self.word.Visible = False
            self._initialized = True
        except Exception as e:
            self._cleanup_com()
            raise WpsComError(
                f"Failed to start WPS Application via KWPS.Application: {e}"
            ) from e

        return self

    def _cleanup_com(self):
        """Clean up COM initialization if __enter__ fails partway through."""
        if self._initialized and self._pythoncom:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
            self._initialized = False

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Clean up COM resources."""
        if self.wd_doc is not None:
            try:
                self.wd_doc.Close()
            except Exception:
                pass
            self.wd_doc = None

        if self.word is not None:
            try:
                self.word.Quit()
            except Exception:
                pass
            self.word = None

        if self._initialized and self._pythoncom:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
            self._initialized = False

    def open_document(self, path: str):
        """Open a document in WPS.

        Args:
            path: Absolute or relative path to the .docx file

        Raises:
            FileNotFoundError: If the path does not exist
            WpsComError: If WPS fails to open the document
        """
        abs_path = os.path.abspath(path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Document not found: {abs_path}")

        try:
            self.wd_doc = self.word.Documents.Open(abs_path)
        except Exception as e:
            raise WpsComError(
                f"WPS failed to open document: {e}"
            ) from e

    def save_document(self, path: Optional[str] = None):
        """Save the current document.

        Args:
            path: Output path. If None, saves to current location.

        Raises:
            WpsComError: If save fails
        """
        if self.wd_doc is None:
            raise RuntimeError("No document open")

        try:
            if path:
                abs_path = os.path.abspath(path)
                self.wd_doc.SaveAs(abs_path)
            else:
                self.wd_doc.Save()
        except Exception as e:
            raise WpsComError(
                f"WPS failed to save document: {e}"
            ) from e

    def enable_track_changes(self, enabled: bool = True):
        """Enable or disable track changes mode.

        Args:
            enabled: True to track changes, False to stop tracking
        """
        if self.wd_doc is None:
            raise RuntimeError("No document open")
        self.wd_doc.TrackRevisions = enabled

    # ---- Detection ---- #

    @staticmethod
    def is_wps_available() -> bool:
        """Check if WPS Office is installed and its COM interface is usable.

        Checks the Windows registry for ``KWPS.Application`` without
        starting WPS itself.

        Returns:
            True if ``KWPS.Application`` is registered in COM
        """
        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT, r"KWPS.Application"
            ) as key:
                return bool(winreg.QueryValue(key, ""))
        except (OSError, ImportError):
            pass

        # Fallback: try COM without creating the application
        try:
            import pythoncom
            pythoncom.CoInitialize()
            try:
                clsid = pythoncom.CLSIDFromProgID("KWPS.Application")
                return clsid is not None
            except Exception:
                return False
            finally:
                pythoncom.CoUninitialize()
        except Exception:
            return False
