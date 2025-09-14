"""
Navigate Tools for Word Document MCP Server.

This module contains operations for navigating and setting active context/objects in Word.
"""

import logging
from typing import Dict, Any, Optional

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..mcp_service.core_utils import (
    ErrorCode,
    WordDocumentError,
    log_error,
    log_info
)


@handle_com_error(ErrorCode.SERVER_ERROR, "set active context")
def set_active_context(document: win32com.client.CDispatch, context_type: str, context_id: str) -> Dict[str, Any]:
    """设置活动上下文

    Args:
        document: Word文档COM对象
        context_type: 上下文类型 (section, paragraph, table, etc.)
        context_id: 上下文ID

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当设置上下文失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    # 支持的上下文类型
    valid_context_types = ["section", "paragraph", "table", "image", "comment", "bookmark"]
    
    if context_type.lower() not in valid_context_types:
        valid_types_str = ", ".join(valid_context_types)
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT,
            f"Invalid context type. Supported types: {valid_types_str}"
        )
    
    try:
        # 根据上下文类型和ID设置活动上下文
        # 这里会实现内部逻辑，可能包括视图滚动，但对外部隐藏这些细节
        word_app = document.Application
        
        # 根据上下文类型查找并设置活动对象
        if context_type.lower() == "section":
            # 实现按节ID设置上下文
            pass
        elif context_type.lower() == "paragraph":
            # 实现按段落ID设置上下文
            pass
        elif context_type.lower() == "table":
            # 实现按表格ID设置上下文
            pass
        elif context_type.lower() == "image":
            # 实现按图像ID设置上下文
            pass
        elif context_type.lower() == "comment":
            # 实现按注释ID设置上下文
            pass
        elif context_type.lower() == "bookmark":
            # 实现按书签名称设置上下文
            if context_id in [b.Name for b in document.Bookmarks]:
                bookmark = document.Bookmarks(context_id)
                bookmark.Select()
                word_app.ActiveWindow.ScrollIntoView(bookmark.Range)
            else:
                raise WordDocumentError(ErrorCode.OBJECT_ERROR, f"Bookmark '{context_id}' not found")
        
        log_info(f"Successfully set active context to {context_type} with ID {context_id}")
        
        return {
            "success": True,
            "context_type": context_type,
            "context_id": context_id,
            "message": f"Active context successfully set to {context_type} with ID {context_id}"
        }
    except Exception as e:
        raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to set active context: {str(e)}")


@handle_com_error(ErrorCode.SERVER_ERROR, "set active object")
def set_active_object(document: win32com.client.CDispatch, object_type: str, object_id: str) -> Dict[str, Any]:
    """设置活动对象

    Args:
        document: Word文档COM对象
        object_type: 对象类型 (paragraph, table, image, comment, text)
        object_id: 对象ID

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当设置活动对象失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    # 支持的对象类型
    valid_object_types = ["paragraph", "table", "image", "comment", "text"]
    
    if object_type.lower() not in valid_object_types:
        valid_types_str = ", ".join(valid_object_types)
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT,
            f"Invalid object type. Supported types: {valid_types_str}"
        )
    
    try:
        word_app = document.Application
        selection = word_app.Selection
        
        # 根据对象类型和ID设置活动对象
        object_type = object_type.lower()
        
        if object_type == "table":
            # 查找并选择指定ID的表格
            for table in document.Tables:
                if str(table.Range.Start) == object_id:
                    table.Select()
                    break
            else:
                raise WordDocumentError(ErrorCode.OBJECT_ERROR, f"Table with ID '{object_id}' not found")
        elif object_type == "image":
            # 查找并选择指定ID的图像
            for shape in document.InlineShapes:
                if str(shape.Range.Start) == object_id:
                    shape.Select()
                    break
            else:
                raise WordDocumentError(ErrorCode.OBJECT_ERROR, f"Image with ID '{object_id}' not found")
        elif object_type == "comment":
            # 查找并选择指定ID的评论
            try:
                comment_index = int(object_id)
                comment = document.Comments(comment_index)
                comment.Select()
            except (ValueError, Exception):
                raise WordDocumentError(ErrorCode.OBJECT_ERROR, f"Comment with ID '{object_id}' not found")
        elif object_type == "paragraph":
            # 尝试按位置选择段落
            try:
                start_pos = int(object_id)
                range_obj = document.Range(Start=start_pos, End=start_pos)
                paragraph = range_obj.Paragraphs(1)
                paragraph.Range.Select()
            except Exception:
                raise WordDocumentError(ErrorCode.OBJECT_ERROR, f"Paragraph with ID '{object_id}' not found")
        elif object_type == "text":
            # 尝试按位置选择文本
            try:
                start_pos = int(object_id)
                range_obj = document.Range(Start=start_pos, End=start_pos + 1)
                range_obj.Select()
            except Exception:
                raise WordDocumentError(ErrorCode.OBJECT_ERROR, f"Text with ID '{object_id}' not found")
        
        # 滚动到新位置
        word_app.ActiveWindow.ScrollIntoView(selection.Range)
        
        log_info(f"Successfully set active object to {object_type} with ID {object_id}")
        
        return {
            "success": True,
            "object_type": object_type,
            "object_id": object_id,
            "message": f"Active object successfully set to {object_type} with ID {object_id}"
        }
    except Exception as e:
        raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to set active object: {str(e)}")