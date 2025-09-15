"""Validation utilities for Word Document MCP Server.

This module contains functions for validating input parameters,
locators, object types, formatting, and other aspects of the API.
"""

from typing import Any, Dict, List, Optional

from .errors import ErrorCode


def validate_input_params(
    params: Dict[str, Any], required_params: list
) -> tuple[bool, str]:
    """
    Validates input parameters for MCP tools.

    Args:
        params: The parameters to validate
        required_params: List of required parameter names

    Returns:
        A tuple of (is_valid, error_message)
    """
    # Check for required parameters
    missing_params = [
        param
        for param in required_params
        if param not in params or params[param] is None
    ]
    if missing_params:
        return False, f"Missing required parameter(s): {', '.join(missing_params)}"

    return True, ""


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


def validate_object_type(object_type: str) -> Optional[str]:
    """验证元素类型是否有效."""
    valid_types = {"paragraphs", "tables", "images", "headings", "styles", "comments"}
    if object_type not in valid_types:
        return (
            f"Invalid object type: {object_type}. Valid types: {', '.join(valid_types)}"
        )
    return None


def validate_operations(operations: List[Dict[str, Any]]) -> Optional[str]:
    """验证批量操作参数是否有效."""
    if not isinstance(operations, list):
        return "'operations' must be a list of dictionaries"
    for i, op in enumerate(operations):
        if not isinstance(op, dict):
            return f"Operation at index {i} is not a dictionary"
        if "locator" not in op or "formatting" not in op:
            return f"Operation at index {i} missing 'locator' or 'formatting' key"
    return None


def validate_formatting(formatting: Dict[str, Any]) -> Optional[str]:
    """验证格式化参数是否有效."""
    valid_keys = {
        "bold",
        "italic",
        "font_size",
        "font_color",
        "font_name",
        "alignment",
        "paragraph_style",
    }
    for key in formatting:
        if key not in valid_keys:
            return f"Invalid formatting key: {key}. Valid keys: {', '.join(valid_keys)}"
    if "font_size" in formatting and (
        not isinstance(formatting["font_size"], int) or formatting["font_size"] <= 0
    ):
        return "'font_size' must be a positive integer"
    if "alignment" in formatting and formatting["alignment"].lower() not in [
        "left",
        "center",
        "right",
    ]:
        return "'alignment' must be 'left', 'center', or 'right'"
    return None


def validate_position(position: str) -> Optional[str]:
    """验证位置参数是否有效 (适用于caption)."""
    if position not in ["above", "below"]:
        return "Invalid position. Must be 'above' or 'below'."
    return None


def validate_insert_position(position: str) -> Optional[str]:
    """验证插入位置参数是否有效 (适用于object insertion)."""
    if position not in ["before", "after", "replace"]:
        return "Invalid position. Must be 'before', 'after', or 'replace'."
    return None


def validate_file_path(file_path: Optional[str]) -> str:
    """验证文件路径

    Args:
        file_path: 文件路径

    Returns:
        验证后的文件路径

    Raises:
        ValueError: 当文件路径无效时抛出
    """
    if not file_path:
        raise ValueError("File path must be provided")

    if not os.path.exists(file_path):
        raise ValueError(f"File not found: {file_path}")

    if not os.path.isfile(file_path):
        raise ValueError(f"Path is not a file: {file_path}")

    # 检查文件扩展名是否为Word文档
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in [".docx", ".doc", ".docm"]:
        raise ValueError(f"Unsupported file type: {ext}")

    return file_path


def validate_active_document(ctx):
    """验证活动文档是否存在."""
    if not ctx or not hasattr(ctx, "request_context") or not ctx.request_context:
        return "Request context not found."
    if (
        not hasattr(ctx.request_context, "lifespan_context")
        or not ctx.request_context.lifespan_context
    ):
        return "Lifespan context not found."
    active_doc = ctx.request_context.lifespan_context.get_active_document()
    if not active_doc:
        return "No active document found. Please open a document first."
    return None


def require_active_document_validation(func):
    """验证活动文档的装饰器."""

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        # 获取上下文参数（ctx通常是第一个参数）
        ctx = args[0] if args else kwargs.get("ctx")
        if not ctx:
            return "Context object not found in function parameters."

        # 验证活动文档
        error = validate_active_document(ctx)
        if error:
            return error

        return func(*args, **kwargs)

    return wrapper