"""
AppContext for managing the Word application instance and the active document state.
"""

import threading
from typing import Optional

import win32com.client


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


# Create a global instance of AppContext
# This will be initialized when the Word application is started
app_context = None
