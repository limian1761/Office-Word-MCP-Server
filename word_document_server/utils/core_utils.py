"""
Utility functions shared across multiple Word document server modules.
"""

from typing import Any, Dict, List, Optional
import functools

# 添加session上下文管理功能
from mcp.server.fastmcp.server import Context

from word_document_server.utils.errors import (WordDocumentError,
                                         format_error_response,
                                         ErrorCode, ElementNotFoundError)

def validate_locator(locator: Dict[str, Any]) -> Optional[str]:
    """
    Validates the structure of a locator dictionary.

    Args:
        locator: The locator to validate

    Returns:
        An error message if validation fails, or None if validation passes
    """
    if not isinstance(locator, dict):
        return "Error: Locator must be a dictionary"

    if "target" not in locator:
        return "Error: Locator must contain a 'target' field"

    target = locator["target"]
    if not isinstance(target, dict):
        return "Error: Locator 'target' must be a dictionary"

    if "type" not in target:
        return "Error: Locator target must contain a 'type' field"

    return None

def validate_active_document(ctx: Context) -> Optional[str]:
    """验证活动文档是否存在。"""
    if not ctx.session.document_state.get("active_document_path"):
        return "No active document found. Please open a document first."
    return None

def validate_element_type(element_type: str) -> Optional[str]:
    """验证元素类型是否有效。"""
    valid_types = {"paragraphs", "tables", "images", "headings", "styles", "comments"}
    if element_type not in valid_types:
        return f"Invalid element type: {element_type}. Valid types: {', '.join(valid_types)}"
    return None

def parse_color_hex(color: str) -> Optional[int]:
    """解析十六进制颜色代码为RGB整数。"""
    color = color.lstrip("#")
    if len(color) == 6:
        try:
            return int(color, 16)
        except ValueError:
            return None
    return None

def validate_operations(operations: List[Dict[str, Any]]) -> Optional[str]:
    """验证批量操作参数是否有效。"""
    if not isinstance(operations, list):
        return "'operations' must be a list of dictionaries"
    for i, op in enumerate(operations):
        if not isinstance(op, dict):
            return f"Operation at index {i} is not a dictionary"
        if 'locator' not in op or 'formatting' not in op:
            return f"Operation at index {i} missing 'locator' or 'formatting' key"
    return None

def validate_formatting(formatting: Dict[str, Any]) -> Optional[str]:
    """验证格式化参数是否有效。"""
    valid_keys = {"bold", "italic", "font_size", "font_color", "font_name", "alignment", "paragraph_style"}
    for key in formatting:
        if key not in valid_keys:
            return f"Invalid formatting key: {key}. Valid keys: {', '.join(valid_keys)}"
    if "font_size" in formatting and (not isinstance(formatting["font_size"], int) or formatting["font_size"] <= 0):
        return "'font_size' must be a positive integer"
    if "alignment" in formatting and formatting["alignment"].lower() not in ["left", "center", "right"]:
        return "'alignment' must be 'left', 'center', or 'right'"
    return None

def validate_position(position: str) -> Optional[str]:
    """验证位置参数是否有效（适用于caption）。"""
    if position not in ["above", "below"]:
        return "Invalid position. Must be 'above' or 'below'."
    return None

def validate_insert_position(position: str) -> Optional[str]:
    """验证插入位置参数是否有效（适用于object insertion）。"""
    if position not in ["before", "after", "replace"]:
        return "Invalid position. Must be 'before', 'after', or 'replace'."
    return None

# 通用形状类型映射
def get_shape_types():
    """获取Word文档中支持的形状类型映射。"""
    return {
        1: "Picture",  # wdInlineShapePicture
        2: "LinkedPicture",  # wdInlineShapeLinkedPicture
        3: "Chart",  # wdInlineShapeChart
        4: "Diagram",  # wdInlineShapeDiagram
        5: "OLEControlObject",  # wdInlineShapeOLEControlObject
        6: "OLEObject",  # wdInlineShapeOLEObject
        7: "ActiveXControl",  # wdInlineShapeActiveXControl
        8: "SmartArt",  # wdInlineShapeSmartArt
        9: "3DModel",  # wdInlineShape3DModel
    }

def get_color_type(color_code: int) -> str:
    """
    Convert color type code to human-readable string.

    Args:
        color_code: Color type code from Word COM interface

    Returns:
        Human-readable color type
    """
    # Word picture color type constants
    color_types = {
        0: "Color",  # msoPictureColorTypeColor
        1: "Grayscale",  # msoPictureColorTypeGrayscale
        2: "BlackAndWhite",  # msoPictureColorTypeBlackAndWhite
        3: "Watermark",  # msoPictureColorTypeWatermark
    }
    return color_types.get(color_code, "Unknown")

def get_shape_info(shape, index):
    """获取形状的基本信息。"""
    return {
        "index": index,  # 0-based index
        "type": (
            get_shape_types().get(shape.Type, "Unknown")
            if hasattr(shape, "Type")
            else "Unknown"
        ),
        "width": shape.Width if hasattr(shape, "Width") else 0,
        "height": shape.Height if hasattr(shape, "Height") else 0,
    }

def require_active_document_validation(func):
    """验证活动文档的装饰器。"""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        # 获取上下文参数（ctx通常是第一个参数）
        ctx = args[0] if args else kwargs.get('ctx')
        if not ctx:
            return "Context object not found in function parameters."

        # 直接使用同一文件中定义的validate_active_document函数
        error = validate_active_document(ctx)
        if error:
            return error

        return func(*args, **kwargs)
    return wrapper

def standardize_tool_errors(func):
    """标准化工具函数的错误处理装饰器。"""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except ElementNotFoundError as e:
            return f"No elements found matching the locator: {e}. Please try simplifying your locator or use get_document_outline to check the actual document structure."
        except ValueError as e:
            return f"Invalid parameter: {e}"
        except Exception as e:
            return format_error_response(e)
    return wrapper


