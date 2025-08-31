"""
Text operations for Word Document MCP Server.

This module contains functions for text-related operations.
"""

import json
import logging
from typing import Any, Dict, List, Optional, Union

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..selector.selector import SelectorEngine
from ..utils.core_utils import (ElementNotFoundError, ErrorCode,
                                WordDocumentError, log_error, log_info)
from .text_format_ops import (set_alignment_for_range, set_bold_for_range,
                              set_font_color_for_range,
                              set_font_name_for_range, set_font_size_for_range,
                              set_italic_for_range, set_paragraph_style)

logger = logging.getLogger(__name__)


@handle_com_error(ErrorCode.SERVER_ERROR, "get character count")
def get_character_count(
    document: win32com.client.CDispatch,
    locator: Optional[Dict[str, Any]] = None
) -> int:
    """获取文档或指定元素的字符数

    Args:
        document: Word文档COM对象
        locator: 定位器对象，用于指定要统计字符数的元素

    Returns:
        字符数

    Raises:
        WordDocumentError: 当获取字符数失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    selector = SelectorEngine()

    if locator:
        # 使用定位器找到要统计字符数的元素
        try:
            selection = selector.select(document, locator)
            if hasattr(selection, '_elements') and selection._elements:
                # 获取第一个元素的字符数
                char_count = selection._elements[0].Range.Characters.Count
            else:
                raise WordDocumentError(
                    ErrorCode.ELEMENT_NOT_FOUND, "No element found matching the locator"
                )
        except ElementNotFoundError:
            raise WordDocumentError(
                ErrorCode.ELEMENT_NOT_FOUND, "No element found matching the locator"
            )
    else:
        # 如果没有提供定位器，返回整个文档的字符数
        char_count = document.Range().Characters.Count

    return int(char_count)  # 确保返回整数类型


def get_element_text(element: Any) -> str:
    """获取元素的文本内容

    Args:
        element: 元素对象

    Returns:
        元素的文本内容
    """
    try:
        if hasattr(element, "Range") and hasattr(element.Range, "Text"):
            return str(element.Range.Text)
        elif hasattr(element, "Text"):
            return str(element.Text)
        else:
            return str(element)
    except Exception:
        return ""


@handle_com_error(ErrorCode.FORMATTING_ERROR, "get text from element")
def get_text_from_element(
    document: win32com.client.CDispatch, locator: Dict[str, Any]
) -> str:
    """从指定位置获取文本

    Args:
        document: Word文档COM对象
        locator: 定位器对象

    Returns:
        指定位置的文本内容

    Raises:
        WordDocumentError: 当获取文本失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    selector = SelectorEngine()

    try:
        selection = selector.select(document, locator)
        if hasattr(selection, '_elements') and selection._elements:
            # 获取第一个元素的文本
            element_text = get_element_text(selection._elements[0])
            return str(element_text)
        else:
            raise WordDocumentError(
                ErrorCode.ELEMENT_NOT_FOUND, "No element found matching the locator"
            )
    except ElementNotFoundError:
        raise WordDocumentError(
            ErrorCode.ELEMENT_NOT_FOUND, "No element found matching the locator"
        )


@handle_com_error(ErrorCode.ELEMENT_NOT_FOUND, "insert text")
def insert_text(
    document: win32com.client.CDispatch,
    locator: Dict[str, Any],
    text: str,
    position: str = "after",
) -> bool:
    """在元素位置插入文本

    Args:
        document: Word文档COM对象
        locator: 定位器对象
        text: 要插入的文本
        position: 插入位置 ("before", "after", or "replace")

    Returns:
        操作是否成功
    """
    if not document:
        raise RuntimeError("No document open.")

    if not text:
        raise ValueError("Text cannot be empty.")

    selector = SelectorEngine()
    selection = selector.select(document, locator, expect_single=True)

    element = selection._elements[0]

    if position == "replace":
        element.Range.Text = text
    elif position == "before":
        element.Range.InsertBefore(text)
    else:  # position == "after"
        element.Range.InsertAfter(text)

    return True


@handle_com_error(ErrorCode.ELEMENT_NOT_FOUND, "replace text")
def replace_text(
    document: win32com.client.CDispatch, locator: Dict[str, Any], new_text: str
) -> bool:
    """替换元素文本

    Args:
        document: Word文档COM对象
        locator: 定位器对象
        new_text: 新文本

    Returns:
        操作是否成功
    """
    if not document:
        raise RuntimeError("No document open.")

    if new_text is None:
        raise ValueError("New text cannot be None.")

    selector = SelectorEngine()
    selection = selector.select(document, locator, expect_single=True)

    for element in selection._elements:
        if hasattr(element, "Range"):
            element.Range.Text = new_text

    return True


def insert_text_before_range(com_range: Any, text: str) -> str:
    """在Range对象前插入文本

    Args:
        com_range: Range对象
        text: 要插入的文本

    Returns:
        操作结果的JSON字符串
    """
    try:
        com_range.InsertBefore(text)
        return json.dumps({"success": True, "message": "Text inserted successfully"})
    except Exception as e:
        return json.dumps({"success": False, "message": f"Failed to insert text: {str(e)}"})


def insert_text_after_range(com_range: Any, text: str) -> str:
    """在Range对象后插入文本

    Args:
        com_range: Range对象
        text: 要插入的文本

    Returns:
        操作结果的JSON字符串
    """
    try:
        com_range.InsertAfter(text)
        return json.dumps({"success": True, "message": "Text inserted successfully"})
    except Exception as e:
        return json.dumps({"success": False, "message": f"Failed to insert text: {str(e)}"})


def apply_formatting_to_element(element: Any, formatting: Dict[str, Any]) -> str:
    """对元素应用格式化

    Args:
        element: 元素对象
        formatting: 格式化参数字典

    Returns:
        操作结果的JSON字符串
    """
    try:
        if not hasattr(element, "Range"):
            return json.dumps({"success": False, "message": "Element has no Range attribute"})

        range_obj = element.Range

        # 应用格式化选项
        if "bold" in formatting:
            range_obj.Font.Bold = formatting["bold"]

        if "italic" in formatting:
            range_obj.Font.Italic = formatting["italic"]

        if "font_size" in formatting:
            range_obj.Font.Size = formatting["font_size"]

        if "font_name" in formatting:
            range_obj.Font.Name = formatting["font_name"]

        if "font_color" in formatting:
            # font_color 应该是一个RGB值的元组 (R, G, B)
            if isinstance(formatting["font_color"], (list, tuple)) and len(formatting["font_color"]) == 3:
                range_obj.Font.Color = (
                    formatting["font_color"][0] +
                    (formatting["font_color"][1] << 8) +
                    (formatting["font_color"][2] << 16)
                )

        return json.dumps({"success": True, "message": "Formatting applied successfully"})

    except Exception as e:
        return json.dumps({"success": False, "message": f"Failed to apply formatting: {str(e)}"})


def replace_element_text(element: Any, new_text: str) -> str:
    """替换元素的文本内容

    Args:
        element: 元素对象
        new_text: 新的文本内容

    Returns:
        操作结果的JSON字符串
    """
    try:
        if hasattr(element, "Range"):
            element.Range.Text = new_text
            return json.dumps({"success": True, "message": "Text replaced successfully"})
        return json.dumps({"success": False, "message": "Element has no Range attribute"})
    except Exception as e:
        return json.dumps({"success": False, "message": f"Failed to replace text: {str(e)}"})


def delete_element(element: Any) -> str:
    """删除元素

    Args:
        element: 要删除的元素

    Returns:
        操作结果的JSON字符串
    """
    try:
        if hasattr(element, "Delete"):
            element.Delete()
            return json.dumps({"success": True, "message": "Element deleted successfully"})
        elif hasattr(element, "Range") and hasattr(element.Range, "Delete"):
            element.Range.Delete()
            return json.dumps({"success": True, "message": "Element deleted successfully"})
        return json.dumps({"success": False, "message": "Element cannot be deleted"})
    except Exception as e:
        return json.dumps({"success": False, "message": f"Failed to delete element: {str(e)}"})

