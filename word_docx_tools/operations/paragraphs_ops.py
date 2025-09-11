"""
Paragraph operations for Word Document MCP Server.
This module contains functions for paragraph-related operations.
"""

import json
import logging
from typing import Any, Dict, List, Optional

import win32com.client

from ..com_backend.com_utils import handle_com_error, iter_com_collection
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError, log_error,
                                      log_info)
from ..selector.selector import SelectorEngine

logger = logging.getLogger(__name__)


@handle_com_error(ErrorCode.PARAGRAPH_SELECTION_FAILED, "get paragraphs in range")
def get_paragraphs_in_range(
    document: win32com.client.CDispatch, locator: Dict[str, Any]
) -> List[Dict[str, Any]]:
    """
    Retrieves paragraphs within a specific range defined by a locator.

    Args:
        document: The Word document COM object.
        locator: A locator dictionary defining the range to retrieve paragraphs from.

    Returns:
        A list of dictionaries with paragraph details.
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # Use selector engine to find the range based on locator
    selector = SelectorEngine()
    selection = selector.select(document, locator)

    # Get the range object from the first object in selection
    if not selection._com_ranges:
        raise WordDocumentError(
            ErrorCode.OBJECT_NOT_FOUND, "No objects found for the given locator"
        )

    # Get the range from the selection (直接使用_com_ranges中的Range对象)
    range_obj = selection._com_ranges[0]

    # Get paragraphs in the range
    paragraphs: List[Dict[str, Any]] = []
    for i, paragraph in enumerate(iter_com_collection(range_obj.Paragraphs)):
        try:
            paragraph_info = {
                "index": i,
                "text": paragraph.Range.Text.strip(),
                "style_name": paragraph.Style.NameLocal,
                "range_start": paragraph.Range.Start,
                "range_end": paragraph.Range.End,
            }
            paragraphs.append(paragraph_info)
        except Exception as e:
            log_error(
                f"Failed to retrieve paragraph in range at index {i}: {e}",
                exc_info=True,
            )
            continue

    return paragraphs


@handle_com_error(ErrorCode.PARAGRAPH_SELECTION_FAILED, "get paragraphs info")
def get_paragraphs_info(document: win32com.client.CDispatch) -> Dict[str, Any]:
    """
    Retrieves information about the document's paragraphs.

    Args:
        document: The Word document COM object.

    Returns:
        A dictionary with paragraph statistics.
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # Get statistics
    stats = {"total_paragraphs": document.Paragraphs.Count, "styles_used": {}}

    # Count style usage
    for i, paragraph in enumerate(iter_com_collection(document.Paragraphs)):
        try:
            style_name = paragraph.Style.NameLocal
            if style_name in stats["styles_used"]:
                stats["styles_used"][style_name] += 1
            else:
                stats["styles_used"][style_name] = 1
        except Exception as e:
            log_error(
                f"Failed to retrieve paragraph style at index {i}: {e}", exc_info=True
            )
            continue

    # Sort styles by usage
    stats["styles_used"] = dict(
        sorted(stats["styles_used"].items(), key=lambda item: item[1], reverse=True)
    )

    return stats


@handle_com_error(ErrorCode.PARAGRAPH_SELECTION_FAILED, "get all paragraphs")
def get_all_paragraphs(document: win32com.client.CDispatch) -> List[Dict[str, Any]]:
    """
    Retrieves all paragraphs from the document.

    Args:
        document: The Word document COM object.

    Returns:
        A list of dictionaries with paragraph details.
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    paragraphs: List[Dict[str, Any]] = []
    paragraphs_count = document.Paragraphs.Count
    for i in range(1, paragraphs_count + 1):
        try:
            paragraph = document.Paragraphs(i)
            paragraph_info = {
                "index": i - 1,  # 0-based index
                "text": paragraph.Range.Text.strip(),
                "style_name": paragraph.Style.NameLocal,
                "range_start": paragraph.Range.Start,
                "range_end": paragraph.Range.End,
            }
            paragraphs.append(paragraph_info)
        except Exception as e:
            log_error(f"Failed to retrieve paragraph at index {i}: {e}", exc_info=True)
            continue

    return paragraphs
