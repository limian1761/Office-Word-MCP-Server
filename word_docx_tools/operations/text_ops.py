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
from ..mcp_service.core_utils import (ErrorCode, ObjectNotFoundError,
                                WordDocumentError, log_error, log_info)
from .text_format_ops import (set_alignment_for_range, set_bold_for_range,
                              set_font_color_for_range,
                              set_font_name_for_range, set_font_size_for_range,
                              set_italic_for_range, set_paragraph_style)

logger = logging.getLogger(__name__)


@handle_com_error(ErrorCode.SERVER_ERROR, "get character count")
def get_character_count(
    document: win32com.client.CDispatch, locator: Optional[Dict[str, Any]] = None
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
            if hasattr(selection, "_com_ranges") and selection._com_ranges:
                # 获取第一个元素（现在保证是Range对象）
                range_obj = selection._com_ranges[0]
                char_count = range_obj.Characters.Count
            else:
                raise WordDocumentError(
                    ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
                )
        except ObjectNotFoundError:
            raise WordDocumentError(
                ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
            )
    else:
        # 如果没有提供定位器，返回整个文档的字符数
        char_count = document.Range().Characters.Count

    return int(char_count)  # 确保返回整数类型


def get_object_text(object: Any) -> str:
    """获取元素的文本内容

    Args:
        object: Range对象（现在保证是Range对象）

    Returns:
        元素的文本内容
    """
    try:
        # 由于我们已经确保object是Range对象，直接访问Text属性
        return str(object.Text)
    except Exception:
        return ""


@handle_com_error(ErrorCode.FORMATTING_ERROR, "get text from object")
def get_text_from_object(
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
        if hasattr(selection, "_com_ranges") and selection._com_ranges:
            # 获取第一个Range对象的文本
            range_obj = selection._com_ranges[0]
            # 由于Range对象保证有Text属性，直接访问
            object_text = str(range_obj.Text)
            return object_text
        else:
            raise WordDocumentError(
                ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
            )
    except ObjectNotFoundError:
        raise WordDocumentError(
            ErrorCode.OBJECT_NOT_FOUND, "No object found matching the locator"
        )


@handle_com_error(ErrorCode.OBJECT_NOT_FOUND, "insert text")
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

    # 获取第一个元素（现在保证是Range对象）
    range_obj = selection._com_ranges[0]

    if position == "replace":
        range_obj.Text = text
    elif position == "before":
        range_obj.InsertBefore(text)
    else:  # position == "after"
        range_obj.InsertAfter(text)

    return True


@handle_com_error(ErrorCode.OBJECT_NOT_FOUND, "replace text")
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

    # Selection._com_ranges中只包含Range对象
    for range_obj in selection._com_ranges:
        try:
            # 由于我们已经确保所有对象都是Range对象，直接访问Text属性
            range_obj.Text = new_text
        except Exception as e:
            logger.warning(f"Failed to replace text for object: {e}")

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
        return json.dumps(
            {"success": False, "message": f"Failed to insert text: {str(e)}"}
        )


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
        return json.dumps(
            {"success": False, "message": f"Failed to insert text: {str(e)}"}
        )


def apply_formatting_to_object(range_obj: Any, formatting: Dict[str, Any]) -> str:
    """对Range对象应用格式化

    Args:
        range_obj: Range对象（现在保证是Range对象）
        formatting: 格式化参数字典

    Returns:
        操作结果的JSON字符串
    """
    try:
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
            if (
                isinstance(formatting["font_color"], (list, tuple))
                and len(formatting["font_color"]) == 3
            ):
                range_obj.Font.Color = (
                    formatting["font_color"][0]
                    + (formatting["font_color"][1] << 8)
                    + (formatting["font_color"][2] << 16)
                )

        return json.dumps(
            {"success": True, "message": "Formatting applied successfully"}
        )

    except Exception as e:
        return json.dumps(
            {"success": False, "message": f"Failed to apply formatting: {str(e)}"}
        )


def replace_object_text(range_obj: Any, new_text: str) -> str:
    """替换Range对象的文本内容

    Args:
        range_obj: Range对象（现在保证是Range对象）
        new_text: 新的文本内容

    Returns:
        操作结果的JSON字符串
    """
    try:
        # 由于我们已经确保所有对象都是Range对象，直接访问Text属性
        range_obj.Text = new_text
        return json.dumps({"success": True, "message": "Text replaced successfully"})
    except Exception as e:
        return json.dumps(
            {"success": False, "message": f"Failed to replace text: {str(e)}"}
        )


def delete_object(range_obj: Any) -> str:
    """删除Range对象

    Args:
        range_obj: Range对象（现在保证是Range对象）

    Returns:
        操作结果的JSON字符串
    """
    try:
        # 由于我们已经确保所有对象都是Range对象，直接调用Delete方法
        range_obj.Delete()
        return json.dumps({"success": True, "message": "Object deleted successfully"})
    except Exception as e:
        return json.dumps(
            {"success": False, "message": f"Failed to delete object: {str(e)}"}
        )
