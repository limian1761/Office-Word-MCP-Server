"""
Text operations for Word Document MCP Server.

This module contains functions for text-related operations.
"""

import json
import logging
from typing import Any, Dict, List, Optional, Union

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..mcp_service.core_utils import (ErrorCode, ObjectNotFoundError,
                                      WordDocumentError, log_error, log_info)
from ..selector.selector import SelectorEngine
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
    result = {
        "success": True,
        "message": "Formatting applied successfully",
        "applied_formats": [],
        "failed_formats": [],
    }

    try:
        # 检查Range对象是否有Font属性
        if not hasattr(range_obj, "Font"):
            return json.dumps(
                {
                    "success": False,
                    "message": "Range object does not have Font property",
                    "applied_formats": [],
                    "failed_formats": list(formatting.keys()),
                }
            )

        font_obj = range_obj.Font

        # 应用粗体
        if "bold" in formatting:
            try:
                font_obj.Bold = formatting["bold"]
                result["applied_formats"].append("bold")
            except Exception as e:
                result["failed_formats"].append(f"bold: {str(e)}")

        # 应用斜体
        if "italic" in formatting:
            try:
                font_obj.Italic = formatting["italic"]
                result["applied_formats"].append("italic")
            except Exception as e:
                result["failed_formats"].append(f"italic: {str(e)}")

        # 应用字体大小
        if "font_size" in formatting:
            try:
                font_obj.Size = float(formatting["font_size"])
                result["applied_formats"].append("font_size")
            except Exception as e:
                result["failed_formats"].append(f"font_size: {str(e)}")

        # 应用字体名称
        if "font_name" in formatting:
            try:
                font_obj.Name = str(formatting["font_name"])
                result["applied_formats"].append("font_name")
            except Exception as e:
                result["failed_formats"].append(f"font_name: {str(e)}")

        # 应用字体颜色
        if "font_color" in formatting:
            try:
                color_value = formatting["font_color"]
                # 支持多种颜色格式
                if isinstance(color_value, (list, tuple)) and len(color_value) == 3:
                    # RGB元组格式 (R, G, B)
                    font_obj.Color = (
                        color_value[0] + (color_value[1] << 8) + (color_value[2] << 16)
                    )
                elif isinstance(color_value, int):
                    # 整数格式的颜色值
                    font_obj.Color = color_value
                elif isinstance(color_value, str):
                    # 字符串格式的颜色值（十六进制或命名颜色）
                    try:
                        # 尝试解析十六进制颜色
                        if color_value.startswith("#"):
                            color_value = color_value[1:]
                        r = int(color_value[0:2], 16)
                        g = int(color_value[2:4], 16)
                        b = int(color_value[4:6], 16)
                        font_obj.Color = r + (g << 8) + (b << 16)
                    except:
                        # 如果解析失败，使用系统默认的颜色映射（如果有）
                        try:
                            # Word VBA中的颜色常量映射
                            color_map = {
                                "black": 0,
                                "white": 16777215,
                                "red": 255,
                                "green": 65280,
                                "blue": 16711680,
                                "yellow": 65535,
                                "cyan": 16776960,
                                "magenta": 16711935,
                                "gray": 12632256,
                            }
                            if color_value.lower() in color_map:
                                font_obj.Color = color_map[color_value.lower()]
                            else:
                                raise ValueError(
                                    f"Unsupported color string: {color_value}"
                                )
                        except:
                            raise ValueError(f"Invalid color format: {color_value}")
                else:
                    raise ValueError(f"Unsupported color type: {type(color_value)}")
                result["applied_formats"].append("font_color")
            except Exception as e:
                result["failed_formats"].append(f"font_color: {str(e)}")

        # 如果有失败的格式，更新成功状态和消息
        if result["failed_formats"]:
            result["success"] = False
            result["message"] = (
                f"Some formatting operations failed: {', '.join(result['failed_formats'][:3])}{'...' if len(result['failed_formats']) > 3 else ''}"
            )

        return json.dumps(result)
    except Exception as e:
        result["success"] = False
        result["message"] = f"Failed to apply formatting: {str(e)}"
        result["failed_formats"] = list(formatting.keys())
        return json.dumps(result)


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
