"""
AppContext for managing the Word application instance and the active document state.
"""

import threading
from typing import Optional
from dataclasses import dataclass
import win32com.client

@dataclass
class AppContext:
    """
    Application context that holds the Word application instance and the active document.
    This class acts as a state container. The lifecycle of the Word application
    itself is managed by the server's lifespan manager.
    """

    def __init__(self, word_app: win32com.client.CDispatch):
        """
        Initialize the AppContext with a running Word application instance.

        Args:
            word_app: An active Word application dispatch object.
        """
        self.word_app: win32com.client.CDispatch = word_app
        self._active_document: Optional[win32com.client.CDispatch] = None
        self._lock = threading.Lock()

    def get_active_document(self) -> Optional[win32com.client.CDispatch]:
        """Get the current active document."""
        with self._lock:
            return self._active_document

    def set_active_document(self, doc: Optional[win32com.client.CDispatch]) -> None:
        """Set the current active document."""
        with self._lock:
            self._active_document = doc

    def clear_active_document(self) -> None:
        """Clear the current active document."""
        with self._lock:
            self._active_document = None

    def open_document(self, file_path: str) -> None:
        """Open a document in the Word application.

        Args:
            file_path: The absolute path to the document to open.
        """

        import pythoncom
        with self._lock:
            if file_path:
                try:
                    # Convert to absolute path for COM
                    import os
                    abs_path = os.path.abspath(file_path)
                    self._active_document = self.word_app.Documents.Open(abs_path)
                except pythoncom.com_error as e:
                    # This can happen if the file is corrupt, password-protected, or doesn't exist.
                    self.close_document()
                    raise WordDocumentError(
                        f"Word COM error while opening document: {file_path}. Details: {e}"
                    )
                except Exception as e:
                    self.close_document()
                    raise IOError(f"Failed to open document: {file_path}. Error: {e}")
            else:
                self._active_document = self.word_app.Documents.Add()

    def close_document(self) -> None:
        """Close the currently active document."""
        with self._lock:
            if self._active_document:
                self._active_document.Close()
                self._active_document = None


