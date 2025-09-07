"""
Document objects operations for Word Document MCP Server.
This module contains functions for document objects operations including bookmarks, citations, and hyperlinks.
"""

import logging
from typing import TYPE_CHECKING, Any, Dict, List, Optional

import win32com.client

from ..com_backend.com_utils import handle_com_error, safe_com_call
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError, log_error,
                                      log_info)
from ..selector.selector import SelectorEngine

if TYPE_CHECKING:
    from win32com.client import CDispatch
else:
    CDispatch = Any

logger = logging.getLogger(__name__)


def _get_range_from_locator(document: Any, locator: Optional[Dict[str, Any]]) -> Any:
    """Helper function to get a Range object from a locator."""
    # 检查locator是否为None或字典类型
    if locator is not None and not isinstance(locator, dict):
        raise WordDocumentError(ErrorCode.INVALID_INPUT, "Locator must be a dictionary")

    if not locator:
        range_obj = document.Range()
        range_obj.Collapse(False)  # wdCollapseEnd
        return range_obj

    selector = SelectorEngine()
    try:
        selection = selector.select(document, locator)
        # 确保selection是有效的对象
        if not hasattr(selection, "_com_ranges") or not selection._com_ranges:
            raise WordDocumentError(
                ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
            )

        range_obj = selection._com_ranges[0]
        # 验证获取的对象是否为有效的Range对象
        if not hasattr(range_obj, "Start") or not hasattr(range_obj, "End"):
            # 如果不是有效的Range对象，创建一个新的Range对象
            range_obj = document.Range()
            range_obj.Collapse(False)  # wdCollapseEnd

        return range_obj
    except Exception as e:
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR, f"Failed to locate position: {str(e)}"
        )


# === Bookmark Operations ===
@handle_com_error(ErrorCode.OBJECT_TYPE_ERROR, "create bookmark")
def create_bookmark(
    document: win32com.client.CDispatch,
    bookmark_name: str,
    locator: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """创建书签

    Args:
        document: Word文档COM对象
        bookmark_name: 书签名称
        locator: 定位器对象，用于指定书签位置

    Returns:
        包含书签信息的字典

    Raises:
        WordDocumentError: 当创建书签失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not hasattr(document, "Bookmarks") or document.Bookmarks is None:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, "Document does not support bookmarks"
        )

    if not bookmark_name:
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT, "Bookmark name cannot be empty"
        )

    if any(c in bookmark_name for c in [" ", "\t", "\n", "\r"]):
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT,
            "Bookmark name cannot contain whitespace characters",
        )

    if bookmark_name in [bm.Name for bm in document.Bookmarks]:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, f"Bookmark '{bookmark_name}' already exists"
        )

    range_obj = _get_range_from_locator(document, locator)

    try:
        bookmark = document.Bookmarks.Add(bookmark_name, range_obj)
        log_info(f"Successfully created bookmark '{bookmark_name}'")

        return {"bookmark_name": bookmark.Name, "bookmark_index": bookmark.Index}

    except Exception as e:
        log_error(
            f"Failed to create bookmark '{bookmark_name}': {str(e)}", exc_info=True
        )
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR, f"Failed to create bookmark: {str(e)}"
        )


@handle_com_error(ErrorCode.OBJECT_TYPE_ERROR, "get bookmark")
def get_bookmark(
    document: win32com.client.CDispatch, bookmark_name: str
) -> Dict[str, Any]:
    """获取书签信息

    Args:
        document: Word文档COM对象
        bookmark_name: 书签名称

    Returns:
        包含书签信息的字典

    Raises:
        WordDocumentError: 当获取书签失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not hasattr(document, "Bookmarks") or document.Bookmarks is None:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, "Document does not support bookmarks"
        )

    if not bookmark_name:
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT, "Bookmark name cannot be empty"
        )

    try:
        if bookmark_name not in [bm.Name for bm in document.Bookmarks]:
            raise WordDocumentError(
                ErrorCode.OBJECT_NOT_FOUND, f"Bookmark '{bookmark_name}' not found"
            )

        bookmark = document.Bookmarks(bookmark_name)

        range_info = {
            "start": bookmark.Range.Start,
            "end": bookmark.Range.End,
            "text": (
                bookmark.Range.Text[:100] + "..."
                if len(bookmark.Range.Text) > 100
                else bookmark.Range.Text
            ),
        }

        log_info(f"Successfully retrieved bookmark '{bookmark_name}'")

        return {
            "bookmark_name": bookmark.Name,
            "bookmark_index": bookmark.Index,
            "range": range_info,
        }

    except Exception as e:
        if isinstance(e, WordDocumentError):
            raise
        log_error(f"Failed to get bookmark '{bookmark_name}': {str(e)}", exc_info=True)
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR, f"Failed to get bookmark: {str(e)}"
        )


@handle_com_error(ErrorCode.OBJECT_TYPE_ERROR, "delete bookmark")
def delete_bookmark(document: win32com.client.CDispatch, bookmark_name: str) -> None:
    """删除书签

    Args:
        document: Word文档COM对象
        bookmark_name: 要删除的书签名称

    Raises:
        WordDocumentError: 当删除书签失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not hasattr(document, "Bookmarks") or document.Bookmarks is None:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, "Document does not support bookmarks"
        )

    if not bookmark_name:
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT, "Bookmark name cannot be empty"
        )

    try:
        if bookmark_name not in [bm.Name for bm in document.Bookmarks]:
            raise WordDocumentError(
                ErrorCode.OBJECT_NOT_FOUND, f"Bookmark '{bookmark_name}' not found"
            )

        bookmark = document.Bookmarks(bookmark_name)
        bookmark_name_log = bookmark.Name
        bookmark.Delete()

        log_info(f"Successfully deleted bookmark '{bookmark_name_log}'")

    except Exception as e:
        if isinstance(e, WordDocumentError):
            raise
        log_error(
            f"Failed to delete bookmark '{bookmark_name}': {str(e)}", exc_info=True
        )
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR, f"Failed to delete bookmark: {str(e)}"
        )


# === Citation Operations ===


@handle_com_error(ErrorCode.OBJECT_TYPE_ERROR, "create citation")
def create_citation(
    document: win32com.client.CDispatch,
    source_data: Dict[str, Any],
    locator: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """创建引用

    Args:
        document: Word文档COM对象
        source_data: 引用源数据
        locator: 定位器对象，用于指定引用位置

    Returns:
        包含引用信息的字典

    Raises:
        WordDocumentError: 当创建引用失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not hasattr(document, "Bibliography") or document.Bibliography is None:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, "Document does not support bibliography"
        )

    if not source_data:
        raise WordDocumentError(ErrorCode.INVALID_INPUT, "Source data cannot be empty")

    range_obj = _get_range_from_locator(document, locator)

    try:
        source = document.Bibliography.Sources.Add(source_data)
        citation = document.Bibliography.Citations.Add(source, range_obj)

        log_info("Successfully created citation")

        return {"citation_id": citation.ID, "source_tag": source.Tag}

    except Exception as e:
        log_error(f"Failed to create citation: {str(e)}", exc_info=True)
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR, f"Failed to create citation: {str(e)}"
        )


# === Hyperlink Operations ===


@handle_com_error(ErrorCode.OBJECT_TYPE_ERROR, "create hyperlink")
def create_hyperlink(
    document: win32com.client.CDispatch,
    address: str,
    locator: Optional[Dict[str, Any]] = None,
    sub_address: Optional[str] = None,
    screen_tip: Optional[str] = None,
    text_to_display: Optional[str] = None,
) -> Dict[str, Any]:
    """创建超链接

    Args:
        document: Word文档COM对象
        address: 超链接地址
        locator: 定位器对象，用于指定超链接位置
        sub_address: 子地址（如书签名称）
        screen_tip: 屏幕提示文本
        text_to_display: 要显示的文本

    Returns:
        包含超链接信息的字典

    Raises:
        WordDocumentError: 当创建超链接失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not hasattr(document, "Hyperlinks") or document.Hyperlinks is None:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, "Document does not support hyperlinks"
        )

    if not address:
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT, "Hyperlink address cannot be empty"
        )

    range_obj = _get_range_from_locator(document, locator)

    try:
        hyperlink = document.Hyperlinks.Add(
            Anchor=range_obj,
            Address=address,
            SubAddress=sub_address,
            ScreenTip=screen_tip,
            TextToDisplay=text_to_display,
        )

        log_info(f"Successfully created hyperlink to '{address}'")

        return {
            "hyperlink_address": hyperlink.Address,
            "hyperlink_text": hyperlink.TextToDisplay,
        }

    except Exception as e:
        log_error(f"Failed to create hyperlink to '{address}': {str(e)}", exc_info=True)
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR, f"Failed to create hyperlink: {str(e)}"
        )
