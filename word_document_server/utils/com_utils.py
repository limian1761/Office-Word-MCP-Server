"""
COM utilities for Word Document Server.

This module provides utility functions for working with Word documents via COM interface.
"""

from typing import Optional

import win32com.client

from word_document_server.utils.app_context import app_context


def get_active_document() -> Optional[win32com.client.CDispatch]:
    """Get the current active document from the global app context.

    Returns:
        The active document COM object, or None if no active document.
    """
    if app_context is None:
        return None
    return app_context.get_active_document()


def handle_com_error(e: Exception) -> str:
    """Handle COM-related errors consistently across all functions."""
    if "-2147417848" in str(e) or "disconnected" in str(e).lower():
        return "COM Error: The Word application has been disconnected. Please restart the server."
    return str(e)
