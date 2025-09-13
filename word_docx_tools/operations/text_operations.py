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

# Import text_format_ops for formatting functions
from . import text_format_ops

# -------------------------- Low-Level Text Operations --------------------------

def get_character_count(document: Any, locator: Optional[Dict[str, Any]] = None) -> str:
    """获取文档或特定对象的字符数

    Args:
        document: Word文档对象
        locator: 可选的定位器对象，指定要计算字符数的对象

    Returns:
        包含字符数的JSON字符串
    """
    try:
        if locator:
            selector_engine = SelectorEngine()
            selection = selector_engine.select(document, locator)

            if not selection or not selection.get_object_types():
                return json.dumps(
                    {"success": False, "message": "Failed to locate object for character count"}
                )

            range_obj = selection._com_ranges[0]
            char_count = len(range_obj.Text)
            return json.dumps({"success": True, "character_count": char_count})
        else:
            # 获取整个文档的字符数
            doc_content = document.Content
            total_char_count = len(doc_content.Text)
            return json.dumps({"success": True, "character_count": total_char_count})
    except Exception as e:
        return json.dumps(
            {"success": False, "message": f"Failed to get character count: {str(e)}"}
        )

def get_object_text(document: Any, locator: Dict[str, Any]) -> str:
    """获取特定对象的文本内容

    Args:
        document: Word文档对象
        locator: 定位器对象，指定要获取文本的对象

    Returns:
        包含文本内容的JSON字符串
    """
    try:
        selector_engine = SelectorEngine()
        selection = selector_engine.select(document, locator)

        if not selection or not selection.get_object_types():
            return json.dumps(
                {"success": False, "message": "Failed to locate object for text extraction"}
            )

        range_obj = selection._com_ranges[0]
        text = range_obj.Text
        return json.dumps({"success": True, "text": text})
    except Exception as e:
        return json.dumps(
            {"success": False, "message": f"Failed to get object text: {str(e)}"}
        )

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
                try:
                    range_obj.Font.Bold = formatting["bold"]
                    result["applied_formats"].append("bold")
                except:
                    result["failed_formats"].append("bold")

            # 应用斜体
            if "italic" in formatting:
                try:
                    range_obj.Font.Italic = formatting["italic"]
                    result["applied_formats"].append("italic")
                except:
                    result["failed_formats"].append("italic")

            # 应用字体大小
            if "font_size" in formatting:
                try:
                    range_obj.Font.Size = float(formatting["font_size"])
                    result["applied_formats"].append("font_size")
                except Exception as e:
                    result["failed_formats"].append(f"font_size: {str(e)}")

            # 应用字体名称
            if "font_name" in formatting:
                try:
                    range_obj.Font.Name = str(formatting["font_name"])
                    result["applied_formats"].append("font_name")
                except:
                    result["failed_formats"].append("font_name")

        # 应用段落相关格式
        # 设置对齐方式
        if "alignment" in formatting:
            try:
                # 简单的对齐方式映射
                alignment_map = {
                    "left": 0,  # wdAlignParagraphLeft
                    "center": 1,  # wdAlignParagraphCenter
                    "right": 2,  # wdAlignParagraphRight
                    "justify": 3  # wdAlignParagraphJustify
                }
                if formatting["alignment"] in alignment_map:
                    range_obj.ParagraphFormat.Alignment = alignment_map[formatting["alignment"]]
                    result["applied_formats"].append("alignment")
                else:
                    result["failed_formats"].append(f"alignment: Invalid alignment value")
            except:
                result["failed_formats"].append("alignment")

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

def get_text_from_document(document: Any, locator: Optional[Dict[str, Any]] = None) -> str:
    """从文档或特定对象中获取文本

    Args:
        document: Word文档对象
        locator: 可选的定位器对象，指定要获取文本的对象

    Returns:
        包含文本内容的JSON字符串
    """
    try:
        if locator:
            # 使用低级函数获取特定对象的文本
            result = get_object_text(document, locator)
            return result
        else:
            # 获取整个文档的文本
            doc_content = document.Content
            text = doc_content.Text
            return json.dumps({"success": True, "text": text})
    except Exception as e:
        return json.dumps(
            {"success": False, "message": f"Failed to get text from document: {str(e)}"}
        )

def insert_text_into_document(document: Any, locator: Dict[str, Any], text: str, position: str = "after") -> str:
    """在文档中插入文本

    Args:
        document: Word文档对象
        locator: 定位器对象，指定插入位置
        text: 要插入的文本
        position: 插入位置，可选值为"before"、"after"或"replace"

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

        if position.lower() == "before":
            result = insert_text_before_range(range_obj, text)
        elif position.lower() == "replace":
            range_obj.Text = text
            result = json.dumps({"success": True, "message": "Text replaced successfully"})
        else:  # 默认"after"
            result = insert_text_after_range(range_obj, text)

        return result
    except Exception as e:
        return json.dumps(
            {"success": False, "message": f"Failed to insert text into document: {str(e)}"}
        )

def replace_text_in_document(document: Any, locator: Dict[str, Any], new_text: str) -> str:
    """替换文档中特定对象的文本

    Args:
        document: Word文档对象
        locator: 定位器对象，指定要替换文本的对象
        new_text: 新的文本内容

    Returns:
        操作结果的JSON字符串
    """
    try:
        selector_engine = SelectorEngine()
        selection = selector_engine.select(document, locator)

        if not selection or not selection.get_object_types():
            return json.dumps(
                {"success": False, "message": "Failed to locate object for text replacement"}
            )

        range_obj = selection._com_ranges[0]
        result = replace_object_text(range_obj, new_text)
        return result
    except Exception as e:
        return json.dumps(
            {"success": False, "message": f"Failed to replace text in document: {str(e)}"}
        )

def format_document_text(document: Any, locator: Dict[str, Any], formatting: Dict[str, Any]) -> str:
    """格式化文档中的文本

    Args:
        document: Word文档对象
        locator: 定位器对象，指定要格式化的文本
        formatting: 格式化参数字典

    Returns:
        操作结果的JSON字符串
    """
    try:
        selector_engine = SelectorEngine()
        selection = selector_engine.select(document, locator)

        if not selection or not selection.get_object_types():
            return json.dumps(
                {"success": False, "message": "Failed to locate object for text formatting"}
            )

        range_obj = selection._com_ranges[0]
        result = apply_formatting_to_object(range_obj, formatting)
        return result
    except Exception as e:
        return json.dumps(
            {"success": False, "message": f"Failed to format document text: {str(e)}"}
        )

# -------------------------- Helper Functions --------------------------

def _get_selection_range(document: Any, locator: Dict[str, Any]) -> Optional[Any]:
    """获取选择范围

    Args:
        document: Word文档对象
        locator: 定位器对象

    Returns:
        Range对象或None
    """
    try:
        selector_engine = SelectorEngine()
        selection = selector_engine.select(document, locator)

        if selection and selection.get_object_types():
            return selection._com_ranges[0]
        return None
    except:
        return None

def validate_required_params(params: Dict[str, Any], required_fields: list) -> Dict[str, Any]:
    """验证必需的参数

    Args:
        params: 参数字典
        required_fields: 必需字段列表

    Returns:
        验证结果字典，包含success和message
    """
    for field in required_fields:
        if field not in params or params[field] is None:
            return {"success": False, "message": f"Missing required parameter: {field}"}
    return {"success": True, "message": "All required parameters are present"}

def get_text_from_document(
    active_doc: Any, 
    locator: Optional[Dict[str, Any]] = None
) -> str:
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
                return json.dumps({"success": True, "text": ""}, ensure_ascii=False)

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
    
    # 检查文本长度，如果超长则添加警告信息
    TEXT_LENGTH_WARNING_THRESHOLD = 10000
    if len(result) > TEXT_LENGTH_WARNING_THRESHOLD:
        warning_message = f"注意：获取的文本长度超过{TEXT_LENGTH_WARNING_THRESHOLD}字符。为了提高性能和避免内存问题，建议使用Locator结合range_start和range_end参数进行多次读取。"
        return json.dumps({"success": True, "text": result, "warning": warning_message}, ensure_ascii=False)

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