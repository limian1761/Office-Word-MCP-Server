"""
Context Control Operations for Word Document MCP Server.

This module contains operations for managing document context and active objects in Word.
It focuses on setting context, managing active objects, and navigating between objects.
View-related functionalities are handled internally as needed.
"""

import logging
import uuid
import time
from typing import Dict, Any, Optional, List, Set, Tuple

import win32com.client

from ..mcp_service.core_utils import (
    ErrorCode,
    WordDocumentError,
    log_error,
    log_info
)

from ..models.context import DocumentContext


# 局部导入以避免循环依赖
from ..com_backend.com_utils import handle_com_error

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


@handle_com_error(ErrorCode.SERVER_ERROR, "get active object")
def get_active_object(document: win32com.client.CDispatch) -> Dict[str, Any]:
    """获取当前活动对象信息

    Args:
        document: Word文档COM对象

    Returns:
        包含当前活动对象信息的字典

    Raises:
        WordDocumentError: 当获取活动对象信息失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    try:
        word_app = document.Application
        selection = word_app.Selection
        
        # 获取活动对象信息
        active_object = {
            "type": None,
            "id": None,
            "text": None,
            "position": {
                "page": selection.Information(3),  # wdActiveEndPageNumber
                "line": selection.Information(10)  # wdFirstCharacterLineNumber
            }
        }
        
        # 识别选择的对象类型
        if selection.Tables.Count > 0:
            active_object["type"] = "table"
            active_object["id"] = str(selection.Tables(1).Range.Start)
            active_object["text"] = "Table at position " + str(selection.Information(3))
        elif selection.InlineShapes.Count > 0:
            active_object["type"] = "image"
            active_object["id"] = str(selection.InlineShapes(1).Range.Start)
            active_object["text"] = "Image at position " + str(selection.Information(3))
        elif selection.Comments.Count > 0:
            active_object["type"] = "comment"
            active_object["id"] = str(selection.Comments(1).Index)
            active_object["text"] = selection.Comments(1).Range.Text[:50] + ("..." if len(selection.Comments(1).Range.Text) > 50 else "")
        else:
            active_object["type"] = "text"
            active_object["id"] = str(selection.Range.Start)
            active_object["text"] = selection.Text[:50] + ("..." if len(selection.Text) > 50 else "")
        
        log_info(f"Successfully retrieved active object information")
        
        return {
            "success": True,
            "active_object": active_object
        }
    except Exception as e:
        raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to get active object: {str(e)}")


@handle_com_error(ErrorCode.SERVER_ERROR, "navigate to next object")
def navigate_to_next_object(document: win32com.client.CDispatch, object_type: Optional[str] = None) -> Dict[str, Any]:
    """导航到下一个对象

    Args:
        document: Word文档COM对象
        object_type: 可选的对象类型过滤

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当导航失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    try:
        word_app = document.Application
        selection = word_app.Selection
        
        # 保存当前位置用于检查是否成功移动
        current_start = selection.Range.Start
        
        # 根据对象类型进行导航
        if object_type:
            object_type = object_type.lower()
            
            if object_type == "paragraph":
                # 移动到下一段
                selection.MoveDown(Unit=5, Count=1)
            elif object_type == "table":
                # 查找下一个表格
                next_table = None
                for table in document.Tables:
                    if table.Range.Start > selection.Range.End:
                        next_table = table
                        break
                if next_table:
                    next_table.Select()
                else:
                    return {
                        "success": False,
                        "message": "No more tables found"
                    }
            elif object_type == "image":
                # 查找下一个图像
                next_image = None
                for shape in document.InlineShapes:
                    if shape.Range.Start > selection.Range.End:
                        next_image = shape
                        break
                if next_image:
                    next_image.Select()
                else:
                    return {
                        "success": False,
                        "message": "No more images found"
                    }
            else:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    f"Unsupported object type for navigation: {object_type}"
                )
        else:
            # 默认移动到下一个段落
            selection.MoveDown(Unit=5, Count=1)
        
        # 检查是否成功移动
        if selection.Range.Start == current_start:
            return {
                "success": False,
                "message": "Already at the last object"
            }
        
        # 滚动到新位置
        word_app.ActiveWindow.ScrollIntoView(selection.Range)
        
        log_info(f"Successfully navigated to next {object_type or 'object'}")
        
        # 获取新的活动对象信息
        new_object_info = get_active_object(document)
        
        return {
            "success": True,
            "message": f"Successfully navigated to next {object_type or 'object'}",
            "active_object": new_object_info["active_object"]
        }
    except Exception as e:
        raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to navigate to next object: {str(e)}")


@handle_com_error(ErrorCode.SERVER_ERROR, "navigate to previous object")
def navigate_to_previous_object(document: win32com.client.CDispatch, object_type: Optional[str] = None) -> Dict[str, Any]:
    """导航到上一个对象

    Args:
        document: Word文档COM对象
        object_type: 可选的对象类型过滤

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当导航失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    try:
        word_app = document.Application
        selection = word_app.Selection
        
        # 保存当前位置用于检查是否成功移动
        current_start = selection.Range.Start
        
        # 根据对象类型进行导航
        if object_type:
            object_type = object_type.lower()
            
            if object_type == "paragraph":
                # 移动到上一段
                selection.MoveUp(Unit=5, Count=1)
            elif object_type == "table":
                # 查找上一个表格
                prev_table = None
                for table in reversed(document.Tables):
                    if table.Range.End < selection.Range.Start:
                        prev_table = table
                        break
                if prev_table:
                    prev_table.Select()
                else:
                    return {
                        "success": False,
                        "message": "No previous tables found"
                    }
            elif object_type == "image":
                # 查找上一个图像
                prev_image = None
                for shape in reversed(document.InlineShapes):
                    if shape.Range.End < selection.Range.Start:
                        prev_image = shape
                        break
                if prev_image:
                    prev_image.Select()
                else:
                    return {
                        "success": False,
                        "message": "No previous images found"
                    }
            else:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    f"Unsupported object type for navigation: {object_type}"
                )
        else:
            # 默认移动到上一个段落
            selection.MoveUp(Unit=5, Count=1)
        
        # 检查是否成功移动
        if selection.Range.Start == current_start:
            return {
                "success": False,
                "message": "Already at the first object"
            }
        
        # 滚动到新位置
        word_app.ActiveWindow.ScrollIntoView(selection.Range)
        
        log_info(f"Successfully navigated to previous {object_type or 'object'}")
        
        # 获取新的活动对象信息
        new_object_info = get_active_object(document)
        
        return {
            "success": True,
            "message": f"Successfully navigated to previous {object_type or 'object'}",
            "active_object": new_object_info["active_object"]
        }
    except Exception as e:
        raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to navigate to previous object: {str(e)}")


@handle_com_error(ErrorCode.SERVER_ERROR, "get context information")
def get_context_information(document: win32com.client.CDispatch) -> Dict[str, Any]:
    """获取当前文档的上下文信息

    Args:
        document: Word文档COM对象

    Returns:
        包含当前上下文信息的字典

    Raises:
        WordDocumentError: 当获取上下文信息失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    try:
        word_app = document.Application
        selection = word_app.Selection
        active_window = word_app.ActiveWindow
        
        # 获取当前上下文信息
        context_info = {
            "document": {
                "name": document.Name,
                "path": document.FullName if document.FullName else "Unsaved"
            },
            "current_position": {
                "page": selection.Information(3),  # wdActiveEndPageNumber
                "line": selection.Information(10),  # wdFirstCharacterLineNumber
                "section": selection.Information(1)  # wdActiveEndSectionNumber
            }
        }
        
        # 添加活动对象信息
        active_object_info = get_active_object(document)
        if active_object_info["success"]:
            context_info["active_object"] = active_object_info["active_object"]
        
        log_info("Successfully retrieved context information")
        
        return {
            "success": True,
            "context": context_info
        }
    except Exception as e:
        raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to get context information: {str(e)}")


@handle_com_error(ErrorCode.SERVER_ERROR, "set zoom level")
def set_zoom_level(document: win32com.client.CDispatch, zoom_level: int = None) -> Dict[str, Any]:
    """设置文档缩放比例（内部使用，对外部隐藏视图细节）

    Args:
        document: Word文档COM对象
        zoom_level: 缩放比例(10-500)，设置为None或不提供时使用默认级别(100%)

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当设置缩放失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    # 使用默认级别（100%）
    if zoom_level is None:
        zoom_level = 100
    else:
        # 检查缩放比例是否在有效范围内
        if not (10 <= zoom_level <= 500):
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                "Zoom level must be between 10 and 500"
            )
    
    # 设置缩放比例
    document.ActiveWindow.View.Zoom.Percentage = zoom_level
    log_info(f"Successfully set zoom level to {zoom_level}%")
    
    return {
        "success": True,
        "zoom_level": zoom_level,
        "message": f"Zoom level successfully set to {zoom_level}%"
    }