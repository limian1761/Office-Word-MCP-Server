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
                                      log_info, AppContext)
from ..models.context import DocumentContext 
from ..operations.text_operations import insert_text_after_range
from ..operations.text_format_ops import set_paragraph_style

logger = logging.getLogger(__name__)


@handle_com_error(ErrorCode.PARAGRAPH_SELECTION_FAILED, "get paragraphs")
def get_paragraphs(
    document: win32com.client.CDispatch,
    locator: Optional[Dict[str, Any]] = None
) -> List[Dict[str, Any]]:
    """
    Retrieves paragraphs from the document.
    
    If a locator is provided, retrieves paragraphs within the specified range.
    If no locator is provided, retrieves all paragraphs from the document.

    Args:
        document: The Word document COM object.
        locator: Optional. A locator dictionary defining the range to retrieve paragraphs from.

    Returns:
        A list of dictionaries with paragraph summary details.
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    paragraphs: List[Dict[str, Any]] = []
    
    if locator:
        # 使用AppContext获取当前选择范围
        context = AppContext.get_instance()
        
        # 处理不同类型的locator
        if 'type' in locator:
            if locator['type'] == 'paragraph':
                # 获取指定段落
                if 'index' in locator:
                    index = locator['index']
                    if index < 0:
                        # 负索引表示从末尾开始计数
                        paragraph_count = document.Paragraphs.Count
                        if paragraph_count + index >= 0:
                            index = paragraph_count + index + 1
                        else:
                            raise WordDocumentError(
                                ErrorCode.OBJECT_NOT_FOUND,
                                f"Paragraph index out of range: {index}"
                            )
                    if 1 <= index <= document.Paragraphs.Count:
                        paragraph = document.Paragraphs(index)
                        _add_paragraph_info(paragraphs, paragraph, 0)
                        return paragraphs
                    else:
                        raise WordDocumentError(
                            ErrorCode.OBJECT_NOT_FOUND,
                            f"Paragraph index out of range: {index}"
                        )
            else:
                # 处理其他类型的locator
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR,
                    f"Unsupported locator type: {locator['type']}"
                )
        else:
            raise WordDocumentError(
                ErrorCode.OBJECT_TYPE_ERROR,
                "Locator must specify an object type"
            )
    else:
        # Process all paragraphs in the document
        paragraphs_count = document.Paragraphs.Count
        for i in range(1, paragraphs_count + 1):
            try:
                paragraph = document.Paragraphs(i)
                _add_paragraph_info(paragraphs, paragraph, i - 1)  # 0-based index
            except Exception as e:
                log_error(f"Failed to retrieve paragraph at index {i}: {e}", exc_info=True)
                continue

    return paragraphs


def _update_document_context_for_paragraph(paragraph: Any, operation: str = "modify") -> None:
    """
    更新段落对应的DocumentContext
    
    Args:
        paragraph: 段落COM对象
        operation: 操作类型（"modify", "create", "delete"等）
    """
    try:
        app_context = AppContext.get_instance()
        document = paragraph.Document
        
        # 查找段落对应的DocumentContext
        # 基于段落的Range.Start和Range.End查找对应的上下文
        context = app_context.find_context_by_range(
            document=document,
            start=paragraph.Range.Start,
            end=paragraph.Range.End,
            object_type="paragraph"
        )
        
        if context:
            # 更新上下文信息
            if operation == "delete":
                app_context.remove_context_from_tree(context)
            else:
                app_context.update_paragraph_context(context, paragraph)
                # 通知上下文更新处理器
                app_context.notify_context_update(context, operation)
    except Exception as e:
        log_error(f"Failed to update DocumentContext for paragraph operation {operation}: {str(e)}")


def _add_paragraph_info(
    paragraphs: List[Dict[str, Any]], 
    paragraph: Any, 
    index: int
) -> None:
    """
    Helper function to add paragraph information to the list.
    
    Args:
        paragraphs: List to add the paragraph information to.
        paragraph: The paragraph COM object.
        index: The index to assign to the paragraph.
    """
    # 获取段落文本并去除首尾空白
    paragraph_text = paragraph.Range.Text.strip()
    
    # 构建段落概略信息
    paragraph_info = {
        "index": index,
        "style_name": paragraph.Style.NameLocal,
        "range_start": paragraph.Range.Start,
        "range_end": paragraph.Range.End,
        "has_text": len(paragraph_text) > 0
    }
    
    # 如果段落有文字，添加开头和结尾摘要
    if len(paragraph_text) > 0:
        # 获取开头部分（前30个字符）
        paragraph_info["start_text"] = paragraph_text[:30] if len(paragraph_text) > 30 else paragraph_text
        # 获取结尾部分（后20个字符），如果段落较长
        paragraph_info["end_text"] = paragraph_text[-20:] if len(paragraph_text) > 50 else ""
        # 标记是否包含完整文本
        paragraph_info["is_truncated"] = len(paragraph_text) > 30
    else:
        # 对于没有文字的段落，添加特殊标记
        paragraph_info["empty_type"] = "paragraph_break"
        paragraph_info["description"] = "Empty paragraph containing only paragraph break"
        
    paragraphs.append(paragraph_info)


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
        A list of dictionaries with paragraph summary details.
    """
    # 调用统一的get_paragraphs函数
    return get_paragraphs(document, locator)


@handle_com_error(ErrorCode.PARAGRAPH_SELECTION_FAILED, "get all paragraphs")
def get_all_paragraphs(document: win32com.client.CDispatch) -> List[Dict[str, Any]]:
    """
    Retrieves all paragraphs from the document.

    Args:
        document: The Word document COM object.

    Returns:
        A list of dictionaries with paragraph summary details.
    """
    # 调用统一的get_paragraphs函数，不提供locator
    return get_paragraphs(document)


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


@handle_com_error(ErrorCode.PARAGRAPH_SELECTION_FAILED, "get paragraphs details")
def get_paragraphs_details(
    document: win32com.client.CDispatch,
    locator: Optional[Dict[str, Any]] = None,
    include_stats: bool = False
) -> Dict[str, Any]:
    """
    合并版段落信息获取函数，可同时获取段落列表和统计信息。

    Args:
        document: The Word document COM object.
        locator: Optional. A locator dictionary defining the range to retrieve paragraphs from.
        include_stats: Whether to include paragraph statistics in the result.

    Returns:
        A dictionary containing paragraphs list and optionally statistics.
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    result = {}
    
    # 获取段落列表
    paragraphs = get_paragraphs(document, locator)
    result["paragraphs"] = paragraphs
    
    # 如果需要统计信息
    if include_stats:
        # 计算统计信息
        stats = {"total_paragraphs": len(paragraphs), "styles_used": {}}
        
        # 统计样式使用情况
        for paragraph in paragraphs:
            if "style_name" in paragraph:
                style_name = paragraph["style_name"]
                if style_name in stats["styles_used"]:
                    stats["styles_used"][style_name] += 1
                else:
                    stats["styles_used"][style_name] = 1
        
        # 按使用次数排序样式
        stats["styles_used"] = dict(
            sorted(stats["styles_used"].items(), key=lambda item: item[1], reverse=True)
        )
        
        result["stats"] = stats
    
    return result


@handle_com_error(ErrorCode.PARAGRAPH_SELECTION_FAILED, "insert paragraph")
def insert_paragraph_impl(
    document: win32com.client.CDispatch,
    text: str,
    locator: Dict[str, Any],
    style: Optional[str] = None,
    is_independent_paragraph: bool = True
) -> str:
    """
    Inserts a new paragraph at a specific location in the document.

    Args:
        document: The Word document COM object.
        text: The text content of the new paragraph.
        locator: A locator dictionary defining where to insert the paragraph.
        style: Optional paragraph style name.
        is_independent_paragraph: Whether to insert as an independent paragraph.

    Returns:
        A JSON string indicating success and additional information.
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not text:
        raise ValueError("text parameter must be provided for insert_paragraph operation")

    if not locator:
        raise ValueError("locator parameter must be provided for insert_paragraph operation")

    log_info(f"Inserting paragraph: {text}")

    # 使用AppContext获取当前选择范围
    context = AppContext.get_instance()
    range_obj = None
    
    # 处理不同类型的locator
    if 'type' in locator:
        if locator['type'] == 'paragraph':
            # 获取指定段落
            if 'index' in locator:
                index = locator['index']
                if index < 0:
                    # 负索引表示从末尾开始计数
                    paragraph_count = document.Paragraphs.Count
                    if paragraph_count + index >= 0:
                        index = paragraph_count + index + 1
                    else:
                        raise WordDocumentError(
                            ErrorCode.OBJECT_NOT_FOUND,
                            f"Paragraph index out of range: {index}"
                        )
                if 1 <= index <= document.Paragraphs.Count:
                    range_obj = document.Paragraphs(index).Range
                else:
                    raise WordDocumentError(
                        ErrorCode.OBJECT_NOT_FOUND,
                        f"Paragraph index out of range: {index}"
                    )
        elif locator['type'] == 'document_start':
            # 获取文档开头
            range_obj = document.Content
            range_obj.Collapse(True)  # wdCollapseStart
        elif locator['type'] == 'document_end':
            # 获取文档结尾
            range_obj = document.Content
            range_obj.Collapse(False)  # wdCollapseEnd
        else:
            raise WordDocumentError(
                ErrorCode.OBJECT_TYPE_ERROR,
                f"Unsupported locator type: {locator['type']}"
            )
    else:
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR,
            "Locator must specify an object type"
        )
    
    if not range_obj:
        raise WordDocumentError(
            ErrorCode.OBJECT_NOT_FOUND,
            "Failed to locate range for paragraph insertion"
        )

    # If needed as independent paragraph
    if is_independent_paragraph:
        try:
            # Check if current range is already at the end of a paragraph
            if (
                hasattr(range_obj, "Paragraphs")
                and range_obj.Paragraphs.Count > 0
            ):
                current_paragraph = range_obj.Paragraphs(1)
                # Create new paragraph if range is not at the end of a paragraph
                if range_obj.Start != current_paragraph.Range.End - 1:
                    # Insert paragraph mark before current range to create new paragraph
                    range_obj.InsertBefore("\n")
                    # Update range to the new paragraph
                    range_obj.Start = range_obj.Start
                    range_obj.End = range_obj.Start
        except Exception as e:
            log_error(f"Failed to prepare independent paragraph: {str(e)}")

    # Insert paragraph
    result = insert_text_after_range(com_range=range_obj, text=f"\n{text}")

    if style:
        # Apply paragraph style to the newly inserted paragraph
        try:
            # 获取最后一个段落（新插入的段落）
            last_paragraph_index = document.Paragraphs.Count
            if last_paragraph_index > 0:
                new_paragraph = document.Paragraphs(last_paragraph_index)
                set_paragraph_style(new_paragraph.Range, style)
        except Exception as e:
            log_error(f"Failed to apply paragraph style: {str(e)}")
            # 继续执行，因为这不是致命错误

    # 获取新插入的段落并更新DocumentContext
    try:
        last_paragraph_index = document.Paragraphs.Count
        if last_paragraph_index > 0:
            new_paragraph = document.Paragraphs(last_paragraph_index)
            _update_document_context_for_paragraph(new_paragraph, "create")
    except Exception as e:
        log_error(f"Failed to update context after inserting paragraph: {str(e)}")

    return result


@handle_com_error(ErrorCode.PARAGRAPH_SELECTION_FAILED, "delete paragraph")
def delete_paragraph_impl(
    document: win32com.client.CDispatch,
    locator: Dict[str, Any]
) -> str:
    """
    Deletes a paragraph from the document based on the locator.

    Args:
        document: The Word document COM object.
        locator: A locator dictionary defining which paragraph to delete.

    Returns:
        A JSON string indicating success and additional information.
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not locator:
        raise WordDocumentError(
            ErrorCode.PARAMETER_ERROR,
            "locator parameter must be provided for delete_paragraph operation"
        )

    log_info(f"Deleting paragraph with locator: {locator}")

    # 使用AppContext获取当前选择范围
    context = AppContext.get_instance()
    range_obj = None
    
    # 处理不同类型的locator
    if 'type' in locator:
        if locator['type'] == 'paragraph':
            # 获取指定段落
            if 'index' in locator:
                index = locator['index']
                if index < 0:
                    # 负索引表示从末尾开始计数
                    paragraph_count = document.Paragraphs.Count
                    if paragraph_count + index >= 0:
                        index = paragraph_count + index + 1
                    else:
                        raise WordDocumentError(
                            ErrorCode.OBJECT_NOT_FOUND,
                            f"Paragraph index out of range: {index}"
                        )
                if 1 <= index <= document.Paragraphs.Count:
                    range_obj = document.Paragraphs(index).Range
                else:
                    raise WordDocumentError(
                        ErrorCode.OBJECT_NOT_FOUND,
                        f"Paragraph index out of range: {index}"
                    )
        else:
            raise WordDocumentError(
                ErrorCode.OBJECT_TYPE_ERROR,
                f"Unsupported locator type: {locator['type']}"
            )
    else:
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR,
            "Locator must specify an object type"
        )
    
    if not range_obj:
        raise WordDocumentError(
            ErrorCode.OBJECT_NOT_FOUND,
            "Failed to locate range for paragraph deletion"
        )

    # 在删除前获取段落对象并通知删除操作
    try:
        paragraph = range_obj.Paragraphs(1) if range_obj.Paragraphs.Count > 0 else None
    except Exception as e:
        paragraph = None
        log_error(f"Failed to get paragraph before deletion: {str(e)}")

    # Delete the paragraph
    try:
        range_obj.Delete()
        log_info("Paragraph deleted successfully")
        
        # 通知段落删除，更新DocumentContext
        if paragraph:
            _update_document_context_for_paragraph(paragraph, "delete")
        
        return json.dumps(
            {"success": True, "message": "Paragraph deleted successfully"},
            ensure_ascii=False
        )
    except Exception as e:
        log_error(f"Failed to delete paragraph: {str(e)}")
        raise WordDocumentError(
            ErrorCode.OPERATION_ERROR,
            f"Failed to delete paragraph: {str(e)}"
        )


@handle_com_error(ErrorCode.PARAGRAPH_SELECTION_FAILED, "format paragraph")
def format_paragraph_impl(
    document: win32com.client.CDispatch,
    locator: Dict[str, Any],
    formatting: Dict[str, Any]
) -> str:
    """
    Formats a paragraph in the document based on the locator and formatting options.

    Args:
        document: The Word document COM object.
        locator: A locator dictionary defining which paragraph to format.
        formatting: A dictionary of formatting options.

    Returns:
        A JSON string indicating success and additional information.
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not locator:
        raise WordDocumentError(
            ErrorCode.PARAMETER_ERROR,
            "locator parameter must be provided for format_paragraph operation"
        )

    if not formatting:
        raise WordDocumentError(
            ErrorCode.PARAMETER_ERROR,
            "formatting parameter must be provided for format_paragraph operation"
        )

    if "paragraph_style" not in formatting:
        raise ValueError("formatting parameter must contain 'paragraph_style' key")

    log_info(f"Formatting paragraph with locator: {locator}, formatting: {formatting}")

    # 使用AppContext获取当前选择范围
    context = AppContext.get_instance()
    range_obj = None
    
    # 处理不同类型的locator
    if 'type' in locator:
        if locator['type'] == 'paragraph':
            # 获取指定段落
            if 'index' in locator:
                index = locator['index']
                if index < 0:
                    # 负索引表示从末尾开始计数
                    paragraph_count = document.Paragraphs.Count
                    if paragraph_count + index >= 0:
                        index = paragraph_count + index + 1
                    else:
                        raise WordDocumentError(
                            ErrorCode.OBJECT_NOT_FOUND,
                            f"Paragraph index out of range: {index}"
                        )
                if 1 <= index <= document.Paragraphs.Count:
                    range_obj = document.Paragraphs(index).Range
                else:
                    raise WordDocumentError(
                        ErrorCode.OBJECT_NOT_FOUND,
                        f"Paragraph index out of range: {index}"
                    )
        else:
            raise WordDocumentError(
                ErrorCode.OBJECT_TYPE_ERROR,
                f"Unsupported locator type: {locator['type']}"
            )
    else:
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR,
            "Locator must specify an object type"
        )
    
    if not range_obj:
        raise WordDocumentError(
            ErrorCode.OBJECT_NOT_FOUND,
            "Failed to locate range for paragraph formatting"
        )

    # Apply paragraph style
    try:
        set_paragraph_style(range_obj, formatting["paragraph_style"])
        log_info(f"Successfully applied paragraph style '{formatting['paragraph_style']}'")
        
        # 获取格式化的段落并更新DocumentContext
        try:
            paragraph = range_obj.Paragraphs(1) if range_obj.Paragraphs.Count > 0 else None
            if paragraph:
                _update_document_context_for_paragraph(paragraph, "modify")
        except Exception as e:
            log_error(f"Failed to update context after formatting paragraph: {str(e)}")
        
        return json.dumps(
            {
                "success": True,
                "message": f"Successfully applied paragraph style '{formatting['paragraph_style']}'",
                "formatted_count": 1,
                "paragraph_style_applied": formatting["paragraph_style"]
            },
            ensure_ascii=False,
        )
    except Exception as e:
        log_error(f"Failed to apply paragraph style: {str(e)}")
        raise WordDocumentError(
            ErrorCode.FORMATTING_ERROR,
            f"Failed to apply paragraph style: {str(e)}"
        )
