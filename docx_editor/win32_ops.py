"""Win32COM operations module.

Provides Microsoft Word COM automation for:
  - Tracked changes (revision mode) writing
  - Reading comments/annotations
  - Reading tracked changes/revisions

Uses the context manager pattern for safe COM resource management.
Pattern follows the existing tracked_change.py implementation.
"""
import os
import sys
from typing import List, Optional

from .models import CommentData, RevisionData
from .utils import Win32ComError


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
        except Exception as e:
            if self._pythoncom:
                try:
                    self._pythoncom.CoUninitialize()
                except Exception:
                    pass
            raise Win32ComError(
                f"Failed to start Microsoft Word: {e}\n"
                "Make sure Microsoft Word is installed."
            ) from e

        return self

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
        """
        if self.wd_doc is None:
            raise RuntimeError("No document open")

        if path:
            abs_path = os.path.abspath(path)
            self.wd_doc.SaveAs(abs_path)
        else:
            self.wd_doc.Save()

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

    @staticmethod
    def is_word_available() -> bool:
        """Check if Microsoft Word is installed and win32com can be used.

        Uses Windows Registry to detect Word installation without starting it.
        Falls back to checking if the CLSID can be resolved.

        Returns:
            True if Word is available, False otherwise
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
                return bool(clsid)
        except (OSError, FileNotFoundError):
            pass

        # Fallback: try COM without creating the application
        try:
            pythoncom.CoInitialize()
            try:
                clsid = pythoncom.CLSIDFromProgID("Word.Application")
                return clsid is not None
            except Exception:
                return False
            finally:
                pythoncom.CoUninitialize()
        except Exception:
            return False
