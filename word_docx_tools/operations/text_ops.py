"""
Text operations for Word Document MCP Server.

This module contains low-level text manipulation functions that are used by higher-level operations.
"""

import json
import logging
from typing import Any, Dict, List, Optional, Union

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..mcp_service.core_utils import (
    ErrorCode, ObjectNotFoundError,
    WordDocumentError, log_error, log_info
)
from ..selector.selector import SelectorEngine
from . import text_format_ops

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
        # 获取文档对象（用于需要文档的格式化操作）
        document = None
        if hasattr(range_obj, "Document"):
            document = range_obj.Document
        
        # 应用字体相关格式
        if hasattr(range_obj, "Font"):
            # 应用粗体
            if "bold" in formatting:
                success = text_format_ops.set_bold_for_range(range_obj, formatting["bold"])
                if success:
                    result["applied_formats"].append("bold")
                else:
                    result["failed_formats"].append("bold")

            # 应用斜体
            if "italic" in formatting:
                success = text_format_ops.set_italic_for_range(range_obj, formatting["italic"])
                if success:
                    result["applied_formats"].append("italic")
                else:
                    result["failed_formats"].append("italic")

            # 应用字体大小
            if "font_size" in formatting:
                try:
                    success = text_format_ops.set_font_size_for_range(range_obj, float(formatting["font_size"]))
                    if success:
                        result["applied_formats"].append("font_size")
                    else:
                        result["failed_formats"].append("font_size")
                except Exception as e:
                    result["failed_formats"].append(f"font_size: {str(e)}")

            # 应用字体名称
            if "font_name" in formatting:
                success = text_format_ops.set_font_name_for_range(range_obj, str(formatting["font_name"]))
                if success:
                    result["applied_formats"].append("font_name")
                else:
                    result["failed_formats"].append("font_name")

            # 应用字体颜色
            if "font_color" in formatting and document:
                success = text_format_ops.set_font_color_for_range(document, range_obj, formatting["font_color"])
                if success:
                    result["applied_formats"].append("font_color")
                else:
                    result["failed_formats"].append("font_color")

        # 应用段落相关格式
        # 设置对齐方式
        if "alignment" in formatting and document:
            success = text_format_ops.set_alignment_for_range(document, range_obj, formatting["alignment"])
            if success:
                result["applied_formats"].append("alignment")
            else:
                result["failed_formats"].append("alignment")

        # 设置段落样式
        if "paragraph_style" in formatting:
            success = text_format_ops.set_paragraph_style(range_obj, formatting["paragraph_style"])
            if success:
                result["applied_formats"].append("paragraph_style")
            else:
                result["failed_formats"].append("paragraph_style")

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
