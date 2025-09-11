"""
Text Operations Module

This module contains the implementation of text operations for Word Document MCP Server,
including both low-level and high-level text manipulation functions.
"""

import json
import logging
from typing import Any, Dict, List, Optional, Union

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..mcp_service.core_utils import (
    ErrorCode,
    ObjectNotFoundError,
    WordDocumentError,
    log_error,
    log_info
)
from ..selector.selector import SelectorEngine
from . import text_format_ops

logger = logging.getLogger(__name__)


# -------------------------- Low-Level Text Operations --------------------------

def insert_text(document: Any, locator: Dict[str, Any], text: str) -> str:
    """在文档中插入文本

    Args:
        document: Word文档对象
        locator: 定位器对象，指定插入位置
        text: 要插入的文本

    Returns:
        操作结果的JSON字符串
    """
    try:
        selector_engine = SelectorEngine()
        selection = selector_engine.select(document, locator)

        if not selection or not selection.get_object_types():
            return json.dumps(
                {"success": False, "message": "Failed to locate object for text insertion"}
            )

        range_obj = selection._com_ranges[0]
        range_obj.InsertAfter(text)
        return json.dumps({"success": True, "message": "Text inserted successfully"})
    except Exception as e:
        return json.dumps(
            {"success": False, "message": f"Failed to insert text: {str(e)}"}
        )

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


# -------------------------- High-Level Text Operations --------------------------

def get_text_from_document(
    active_doc: Any, 
    locator: Optional[Dict[str, Any]] = None,
    start_index: Optional[int] = 0,
    max_length: Optional[int] = 10000
) -> str:
    """从文档中获取文本

    Args:
        active_doc: 活动文档对象
        locator: 定位器对象，用于选择特定元素
        start_index: 开始索引
        max_length: 最大长度

    Returns:
        包含获取文本结果和总长度的JSON字符串
    """
    log_info("Getting text from document")

    if locator:
        # 如果提供了定位器，获取特定元素的文本
        selector_engine = SelectorEngine()
        try:
            selection = selector_engine.select(active_doc, locator)

            if not selection or not selection.get_object_types():
                # 如果找不到元素，返回空文本
                return json.dumps(
                    {"success": True, "text": "", "total_length": 0}, ensure_ascii=False
                )

            # 获取选择区域的文本
            range_obj = selection._com_ranges[0]
            result = range_obj.Text
        except Exception as e:
            # 如果选择过程出错，返回空文本
            log_error(f"Error selecting object: {e}")
            return json.dumps({"success": True, "text": "", "total_length": 0}, ensure_ascii=False)
    else:
        # 如果没有提供定位器，获取整个文档的文本
        result = active_doc.Content.Text

    # 保存原始文本的总长度
    total_length = len(result)

    # 处理start_index和max_length参数
    if start_index is not None and start_index >= 0:
        if start_index < len(result):
            result = result[start_index:]
        else:
            result = ""
    
    if max_length is not None and max_length > 0:
        result = result[:max_length]

    return json.dumps({"success": True, "text": result, "total_length": total_length}, ensure_ascii=False)

def insert_text_into_document(
    active_doc: Any,
    text: str,
    locator: Dict[str, Any],
    position: str = "after"
) -> str:
    """在文档中插入文本

    Args:
        active_doc: 活动文档对象
        text: 要插入的文本
        locator: 定位器对象，指定插入位置
        position: 插入位置，"before"或"after"

    Returns:
        包含插入结果的JSON字符串
    """
    log_info(f"Inserting text: {text}")

    range_obj = _get_selection_range(active_doc, locator, "text insertion")

    # 插入文本
    if position.lower() == "before":
        result = insert_text_before_range(com_range=range_obj, text=text)
    else:
        result = insert_text_after_range(com_range=range_obj, text=text)

    # 检查返回结果是否为字符串（JSON格式），如果是则直接返回
    if isinstance(result, str):
        try:
            result_dict = json.loads(result)
            if not result_dict.get("success", False):
                return json.dumps(result_dict, ensure_ascii=False)
        except json.JSONDecodeError:
            pass
        return result

    # 如果函数返回的是布尔值，则构造结果消息
    return json.dumps(
        {"success": True, "message": "Text inserted successfully"},
        ensure_ascii=False,
    )

def replace_text_in_document(
    active_doc: Any,
    text: str,
    locator: Dict[str, Any]
) -> str:
    """在文档中替换文本

    Args:
        active_doc: 活动文档对象
        text: 替换后的新文本
        locator: 定位器对象，指定要替换的文本位置

    Returns:
        包含替换结果的JSON字符串
    """
    log_info(f"Replacing text with: {text}")

    range_obj = _get_selection_range(active_doc, locator, "text replacement")

    # 替换文本
    result = replace_object_text(range_obj=range_obj, new_text=text)

    return json.dumps(
        {"success": True, "message": "Text replaced successfully"},
        ensure_ascii=False,
    )

def get_character_count_from_document(
    active_doc: Any,
    locator: Optional[Dict[str, Any]] = None
) -> str:
    """获取文档中的字符数

    Args:
        active_doc: 活动文档对象
        locator: 定位器对象，用于选择特定元素

    Returns:
        包含字符数结果的JSON字符串
    """
    if locator:
        range_obj = _get_selection_range(active_doc, locator, "character count")
        text_content = range_obj.Text
    else:
        text_content = active_doc.Content.Text
    
    char_count = len(text_content)
    return json.dumps(
        {"success": True, "character_count": char_count}, ensure_ascii=False
    )

def apply_formatting_to_document_text(
    active_doc: Any,
    formatting: Dict[str, Any],
    locator: Dict[str, Any]
) -> str:
    """对文档中的文本应用格式化

    Args:
        active_doc: 活动文档对象
        formatting: 格式化参数字典
        locator: 定位器对象，指定要格式化的文本位置

    Returns:
        包含格式化结果的JSON字符串
    """
    log_info("Applying formatting")

    range_obj = _get_selection_range(active_doc, locator, "formatting")

    # 应用格式
    result = apply_formatting_to_object(range_obj=range_obj, formatting=formatting)

    return json.dumps(
        {"success": True, "message": "Formatting applied successfully"},
        ensure_ascii=False,
    )


# -------------------------- Helper Functions --------------------------

def _get_selection_range(active_doc: Any, locator: Dict[str, Any], operation_name: str) -> Any:
    """获取选择范围，处理错误

    Args:
        active_doc: 活动文档对象
        locator: 定位器对象
        operation_name: 操作名称，用于错误消息

    Returns:
        获取的Range对象

    Raises:
        ValueError: 当locator为None时
        WordDocumentError: 当无法定位对象或获取对象时
    """
    if locator is None:
        raise ValueError(
            f"locator parameter must be provided for {operation_name} operation"
        )

    selector_engine = SelectorEngine()
    selection = selector_engine.select(active_doc, locator)

    if not selection or not selection.get_object_types():
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR,
            f"Failed to locate object for {operation_name}"
        )

    if not hasattr(selection, "_com_ranges") or not selection._com_ranges:
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR,
            f"Failed to get objects from selection for {operation_name}"
        )

    return selection._com_ranges[0]

def validate_required_params(params: Dict[str, Any], operation_name: str) -> None:
    """验证必需参数

    Args:
        params: 参数字典，键为参数名，值为参数值
        operation_name: 操作名称，用于错误消息

    Raises:
        ValueError: 当必需参数缺失时
    """
    for param_name, param_value in params.items():
        if param_value is None:
            raise ValueError(
                f"{param_name} parameter must be provided for {operation_name} operation"
            )