"""
Text Operations Module

This module contains the core implementation of text operations that were previously
in the tools layer. These functions handle the actual text manipulation logic.
"""

import json
from typing import Any, Dict, Optional

from ..selector.selector import SelectorEngine
from ..mcp_service.core_utils import (
    ErrorCode,
    WordDocumentError,
    log_info,
    log_error
)

# Import the lower-level operations
from .text_ops import (
    insert_text_before_range,
    insert_text_after_range,
    replace_object_text,
    apply_formatting_to_object
)

def get_text_from_document(active_doc: Any, locator: Optional[Dict[str, Any]] = None) -> str:
    """从文档中获取文本

    Args:
        active_doc: 活动文档对象
        locator: 定位器对象，用于选择特定元素

    Returns:
        包含获取文本结果的JSON字符串
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
                    {"success": True, "text": ""}, ensure_ascii=False
                )

            # 获取选择区域的文本
            range_obj = selection._com_ranges[0]
            result = range_obj.Text
        except Exception as e:
            # 如果选择过程出错，返回空文本
            log_error(f"Error selecting object: {e}")
            return json.dumps({"success": True, "text": ""}, ensure_ascii=False)
    else:
        # 如果没有提供定位器，获取整个文档的文本
        result = active_doc.Content.Text

    return json.dumps({"success": True, "text": result}, ensure_ascii=False)

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

def format_document_text(
    active_doc: Any,
    format_type: str,
    format_value: Any,
    locator: Dict[str, Any]
) -> str:
    """对文档中的文本应用单一格式化选项

    Args:
        active_doc: 活动文档对象
        format_type: 格式化类型
        format_value: 格式化值
        locator: 定位器对象，指定要格式化的文本位置

    Returns:
        包含格式化结果的JSON字符串
    """
    log_info(f"Applying text format: {format_type}")

    # 构建只包含一种格式的字典，然后调用apply_formatting_to_object
    formatting = {format_type.lower(): format_value}
    
    range_obj = _get_selection_range(active_doc, locator, "text formatting")
    
    # 应用格式
    result = apply_formatting_to_object(range_obj=range_obj, formatting=formatting)
    
    # 解析结果并返回适当的响应
    try:
        result_dict = json.loads(result)
        if result_dict.get("success", False):
            return json.dumps(
                {"success": True, "message": "Text formatted successfully"},
                ensure_ascii=False,
            )
        else:
            return result
    except json.JSONDecodeError:
        return json.dumps(
            {"success": True, "message": "Text formatted successfully"},
            ensure_ascii=False,
        )

# 辅助函数
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