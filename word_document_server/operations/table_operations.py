"""
Table operations for Word Document MCP Server.

This module contains functions for table-related operations.
"""

from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client

from word_document_server.errors import ErrorCode, WordDocumentError
from word_document_server.word_backend import WordBackend


def add_table(
    backend: WordBackend, com_range_obj: win32com.client.CDispatch, rows: int, cols: int
):
    """
    Adds a table after a given range.

    Args:
        backend: The WordBackend instance.
        com_range_obj: The range to insert the table after.
        rows: Number of rows for the table.
        cols: Number of columns for the table.

    Returns:
        The newly created table COM object.
    """
    if not backend.document:
        raise RuntimeError("No document open.")

    try:
        # Validate row and column parameters
        if not isinstance(rows, int) or rows <= 0:
            raise ValueError("Row count must be a positive integer")
        if not isinstance(cols, int) or cols <= 0:
            raise ValueError("Column count must be a positive integer")

        # Validate range object
        if not com_range_obj or not hasattr(com_range_obj, "Duplicate"):
            raise ValueError("Invalid range object")

        # Collapse the range to its end point to insert after
        new_range = com_range_obj.Duplicate
        new_range.Collapse(0)  # WdCollapseDirection.wdCollapseEnd
        new_range.InsertParagraphAfter()  # Add a paragraph break to ensure table is on a new line
        return backend.document.Tables.Add(new_range, rows, cols)
    except Exception as e:
        # Check if it's a COM related error
        if "COM" in str(type(e)) or "Dispatch" in str(type(e)):
            from word_document_server.errors import WordDocumentError

            raise WordDocumentError(f"Failed to create table in Word: {str(e)}")
        raise
