"""Win32COM operations module.

Provides Microsoft Word COM automation for:
  - Tracked changes (revision mode) writing
  - Reading comments/annotations
  - Reading tracked changes/revisions

Uses the context manager pattern for safe COM resource management.
Pattern follows the existing tracked_change.py implementation.
"""
import logging
import os
import sys
from typing import List, Optional

from .models import CommentData, RevisionData
from .utils import Win32ComError

logger = logging.getLogger(__name__)


class Win32Ops:
    """Manage win32com Word application lifecycle and operations.

    Usage:
        with Win32Ops() as ops:
            ops.open_document('path.docx')
            ops.enable_track_changes(True)
            comments = ops.read_comments()
            revisions = ops.read_revisions()
    """

    def __init__(self, use_track_changes: bool = True):
        self.use_track_changes = use_track_changes
        self.word = None
        self.wd_doc = None
        self._initialized = False
        self._pythoncom = None

    def __enter__(self):
        """Initialize COM and start Word application."""
        try:
            import pythoncom
            import win32com.client
            self._pythoncom = pythoncom
            self._win32com_client = win32com.client
        except ImportError as e:
            raise Win32ComError(
                "pywin32 is not installed. Install with: pip install pywin32"
            ) from e

        try:
            pythoncom.CoInitialize()
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            self._initialized = True

            # Verify it's actually Microsoft Word, not WPS or other shim
            try:
                app_name = self.word.Name
                if "Microsoft Word" not in app_name:
                    raise Win32ComError(
                        f"Detected '{app_name}' instead of Microsoft Word.\n"
                        "WPS also registers as Word.Application in COM, but its "
                        "COM interface is not fully compatible.\n"
                        "Please install Microsoft Word for track-changes mode, or "
                        "use track_changes=False to use python-docx directly."
                    )
            except Win32ComError:
                raise
            except Exception:
                pass  # If Name property is unavailable, proceed with caution

            # Warn if WPS is also installed (known conflict with Word's SaveAs)
            if self._is_wps_installed():
                logger.warning(
                    "WPS is installed alongside Word. Track-changes SaveAs "
                    "may be affected by WPS file handler conflicts."
                )

        except Win32ComError:
            self._cleanup_com()
            raise
        except Exception as e:
            self._cleanup_com()
            raise Win32ComError(
                f"Failed to start Microsoft Word: {e}\n"
                "Make sure Microsoft Word is installed (not just WPS)."
            ) from e

        return self

    @staticmethod
    def _is_wps_installed() -> bool:
        """Check if WPS is installed alongside Word (known conflict source)."""
        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT, r"KWPS.Application"
            ) as key:
                return bool(winreg.QueryValue(key, ""))
        except OSError:
            return False

    def _cleanup_com(self):
        """Clean up COM initialization if __enter__ fails partway through."""
        if self._initialized and self._pythoncom:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
            self._initialized = False

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Clean up COM resources.

        Follows the exact cleanup pattern from tracked_change.py:
        1. Close document
        2. Quit Word
        3. CoUninitialize
        All with individual error suppression.
        """
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
        """Open a document in Word.

        Args:
            path: Absolute or relative path to the .docx file
        """
        abs_path = os.path.abspath(path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Document not found: {abs_path}")

        self.wd_doc = self.word.Documents.Open(abs_path)

    def save_document(self, path: Optional[str] = None):
        """Save the current document.

        Args:
            path: Output path. If None, saves to current location.

        Raises:
            Win32ComError: If save fails, with specific handling for
                WPS-related conflicts.
        """
        if self.wd_doc is None:
            raise RuntimeError("No document open")

        if path:
            abs_path = os.path.abspath(path)
            try:
                self.wd_doc.SaveAs(abs_path)
            except Exception as e:
                error_str = str(e)
                if 'Kingsoft' in error_str or 'WPS' in error_str:
                    raise Win32ComError(
                        "Document save failed due to WPS conflict.\n"
                        "WPS file handlers are interfering with Microsoft Word's SaveAs.\n"
                        "Solutions:\n"
                        "  1. Uninstall WPS or disable its COM add-ins in Word\n"
                        "  2. Use track_changes=False to use python-docx directly\n"
                        "  3. Open the output file in Word and save manually"
                    ) from e
                raise Win32ComError(
                    f"Failed to save document: {e}"
                ) from e
        else:
            try:
                self.wd_doc.Save()
            except Exception as e:
                raise Win32ComError(
                    f"Failed to save document: {e}"
                ) from e

    def enable_track_changes(self, enabled: bool = True):
        """Enable or disable track changes mode.

        Args:
            enabled: True to track changes, False to stop tracking
        """
        if self.wd_doc is None:
            raise RuntimeError("No document open")
        self.wd_doc.TrackRevisions = enabled

    def insert_text(self, text: str, position: str = 'end'):
        """Insert text at specified position in the document.

        Args:
            text: Text to insert
            position: 'start', 'end', or 'selection' for current selection
        """
        if self.wd_doc is None:
            raise RuntimeError("No document open")

        if position == 'start':
            self.word.Selection.HomeKey(Unit=6)  # wdStory = 6
        elif position == 'end':
            self.word.Selection.EndKey(Unit=6)  # wdStory = 6
        # 'selection' keeps current position

        self.word.Selection.TypeText(text)

    def read_comments(self) -> List[CommentData]:
        """Read all comments/annotations from the document.

        Returns:
            List of CommentData objects
        """
        if self.wd_doc is None:
            raise RuntimeError("No document open")

        comments = []
        try:
            word_comments = self.wd_doc.Comments
            for i in range(1, word_comments.Count + 1):
                wc = word_comments(i)
                comment = CommentData(
                    id=str(wc.Id),
                    author=wc.Author,
                    date=str(wc.Date) if wc.Date else '',
                    text=wc.Range.Text.strip() if wc.Range else '',
                )
                comments.append(comment)
        except Exception as e:
            raise Win32ComError(f"Failed to read comments: {e}")

        return comments

    def read_revisions(self) -> List[RevisionData]:
        """Read all tracked changes/revisions from the document.

        Returns:
            List of RevisionData objects
        """
        if self.wd_doc is None:
            raise RuntimeError("No document open")

        revisions = []
        try:
            word_revisions = self.wd_doc.Revisions
            for i in range(1, word_revisions.Count + 1):
                wr = word_revisions(i)

                # Determine revision type
                rev_type = str(wr.Type)
                if 'Insert' in rev_type or 'insert' in rev_type:
                    rtype = 'insertion'
                elif 'Delete' in rev_type or 'delete' in rev_type:
                    rtype = 'deletion'
                else:
                    rtype = rev_type

                revision = RevisionData(
                    rev_id=str(wr.Index),
                    author=wr.Author,
                    date=str(wr.Date) if wr.Date else '',
                    type=rtype,
                    text=wr.Range.Text.strip() if wr.Range else '',
                    paragraph_index=0,  # Can't easily map without more work
                )
                revisions.append(revision)
        except Exception as e:
            raise Win32ComError(f"Failed to read revisions: {e}")

        return revisions

    # Microsoft Word's well-known CLSID
    MS_WORD_CLSID = "{000209FF-0000-0000-C000-000000000046}"

    @staticmethod
    def is_word_available() -> bool:
        """Check if Microsoft Word (not WPS) is installed and win32com can be used.

        Uses Windows Registry to detect Word installation without starting it.
        Verifies the CLSID matches Microsoft Word's known CLSID to distinguish
        from WPS (which also registers as ``Word.Application`` for compatibility).

        Returns:
            True if Microsoft Word is available, False otherwise
        """
        try:
            import pythoncom
            import win32com.client
        except ImportError:
            return False

        # Check via registry - look for Word.Application in CLSID
        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT,
                r"Word.Application\CLSID"
            ) as key:
                clsid = winreg.QueryValue(key, "")
                # WPS also registers Word.Application but with its own CLSID.
                # Only accept Microsoft Word's known CLSID.
                return bool(clsid) and clsid.upper() == Win32Ops.MS_WORD_CLSID
        except (OSError, FileNotFoundError):
            pass

        # Fallback: try COM without creating the application
        try:
            pythoncom.CoInitialize()
            try:
                clsid = pythoncom.CLSIDFromProgID("Word.Application")
                return str(clsid).upper() == Win32Ops.MS_WORD_CLSID
            except Exception:
                return False
            finally:
                pythoncom.CoUninitialize()
        except Exception:
            return False
