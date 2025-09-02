"""
Utility functions shared across multiple Word document server modules.
"""

import functools
import logging
import os
import shutil
from typing import Any, Dict, List, Optional, Tuple

from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from .app_context import AppContext

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("word_docx_tools.log"), logging.StreamHandler()],
)
logger = logging.getLogger("WordDocumentServer")

from enum import Enum

from .errors import ErrorCode, WordDocumentError, DocumentNotFoundError, ObjectNotFoundError, StyleNotFoundError, ImageError, ImageFormatError, CommentError

class CommentIndexError(CommentError):
    """Raised when a comment index is out of range"""

    def __init__(self, comment_index: int, message: Optional[str] = None):
        details = {"comment_index": comment_index}
        super().__init__(ErrorCode.COMMENT_INDEX_ERROR, message, details)


class CommentEmptyError(CommentError):
    """Raised when comment text is empty"""

    def __init__(self, message: Optional[str] = None):
        super().__init__(ErrorCode.COMMENT_EMPTY_ERROR, message)


class ReplyEmptyError(CommentError):
    """Raised when reply text is empty"""

    def __init__(self, message: Optional[str] = None):
        super().__init__(ErrorCode.REPLY_EMPTY_ERROR, message)


def get_active_document(ctx: Context[ServerSession, "AppContext"]) -> Any:
    """获取当前活动的文档对象

    Args:
        ctx: MCP上下文对象

    Returns:
        活动文档对象

    Raises:
        WordDocumentError: 当没有活动文档时抛出
    """
    # 延迟导入以避免循环依赖
    from .app_context import AppContext

    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if not active_doc:
            raise WordDocumentError(
                ErrorCode.DOCUMENT_ERROR,
                "No active document found. Please open a document first.",
            )
        return active_doc
    except Exception as e:
        if isinstance(e, WordDocumentError):
            raise
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, f"Failed to get active document: {str(e)}"
        )


def handle_error(e: Exception) -> tuple[int, str, Dict[str, Any]]:
    """
    Handles errors and returns standardized error information.

    Args:
        e: The exception to handle

    Returns:
        A tuple of (error_code, error_message, error_details)
    """
    # Log the error
    logger.error("Error occurred: %s", str(e), exc_info=True)

    # Handle specific exceptions
    if isinstance(e, WordDocumentError):
        return e.error_code.value[0], e.message, e.details

    # Handle common exceptions
    if isinstance(e, FileNotFoundError):
        return (
            ErrorCode.DOCUMENT_OPEN_ERROR.value[0],
            f"File not found: {str(e)}",
            {"file_path": str(e.filename) if hasattr(e, "filename") else None},
        )

    if isinstance(e, ValueError):
        return (
            ErrorCode.INVALID_INPUT.value[0],
            f"Invalid value: {str(e)}",
            {"error_type": "ValueError"},
        )

    if isinstance(e, PermissionError):
        return (
            ErrorCode.PERMISSION_DENIED.value[0],
            f"Permission denied: {str(e)}",
            {"error_type": "PermissionError"},
        )

    # Handle general exceptions
    return (
        ErrorCode.SERVER_ERROR.value[0],
        f"An unexpected error occurred: {str(e)}",
        {"error_type": type(e).__name__},
    )


def format_error(error_code: ErrorCode, message: str) -> Dict[str, Any]:
    """格式化错误信息

    Args:
        error_code: 错误代码枚举
        message: 错误描述

    Returns:
        格式化后的错误信息字典
    """
    return {
        "error_code": error_code.value,
        "error_type": error_code.name,
        "message": message,
    }


def format_error_response(e: Exception) -> str:
    """
    Formats an error response for the MCP tool.

    Args:
        e: The exception to format

    Returns:
        A formatted error message string
    """
    error_code, error_message, details = handle_error(e)

    # Format the error message with code and message
    response = f"Error [{error_code}]: {error_message}"

    # Add details for common error types
    if error_code == ErrorCode.STYLE_NOT_FOUND.value[0] and "similar_styles" in details:
        similar_styles = details["similar_styles"]
        if similar_styles:
            response += f" Did you mean one of these: {', '.join(similar_styles)}?"

    elif error_code == ErrorCode.NO_ACTIVE_DOCUMENT.value[0]:
        response += " Please use 'open_document' first."

    elif error_code == ErrorCode.OBJECT_NOT_FOUND.value[0] and "locator" in details:
        response += " Please check your locator parameters."

    # Add a generic suggestion for resolving the issue
    response += " For more information, check the server logs."

    # 确保响应是UTF-8编码的安全字符串
    try:
        response.encode("utf-8")
    except UnicodeEncodeError:
        # 如果包含无法编码的字符，则使用安全的ASCII表示
        response = response.encode("utf-8", errors="replace").decode("utf-8")

    return response


def log_info(message: str) -> None:
    """记录信息日志

    Args:
        message: 日志信息
    """
    logger.info(message)


def log_error(message: str, exc_info: bool = False) -> None:
    """记录错误日志

    Args:
        message: 错误信息
        exc_info: 是否包含异常堆栈信息
    """
    if exc_info:
        logger.error(message, exc_info=True)
    else:
        logger.error(message)


def log_warning(message: str) -> None:
    """记录警告日志

    Args:
        message: 警告信息
    """
    logger.warning(message)


def log_debug(message: str) -> None:
    """记录调试日志

    Args:
        message: 调试信息
    """
    logger.debug(message)
    logger.error(message, exc_info=True)


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


def parse_color_hex(color: str) -> Optional[int]:
    """解析十六进制颜色代码为RGB整数."""
    color = color.lstrip("#")
    if len(color) == 6:
        try:
            return int(color, 16)
        except ValueError:
            return None
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


# 通用形状类型映射
def get_shape_types():
    """获取Word文档中支持的形状类型映射."""
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
    """获取形状的基本信息."""
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


def standardize_tool_errors(func):
    """标准化工具函数的错误处理装饰器."""

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except ObjectNotFoundError as e:
            return f"No objects found matching the locator: {e}. Please try simplifying your locator or use get_document_outline to check the actual document structure."
        except ValueError as e:
            return f"Invalid parameter: {e}"
        except Exception as e:
            return format_error_response(e)

    return wrapper


def handle_tool_errors(func):
    """Decorator to handle errors in MCP tools uniformly."""

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except (IOError, ConnectionError, TimeoutError) as e:
            # Log the error with context
            logger.error("Error in tool %s: %s", func.__name__, str(e), exc_info=True)
            # Return formatted error response
            return format_error_response(e)

    return wrapper


def is_file_writeable(filepath: str) -> Tuple[bool, str]:
    """
    Check if a file can be written to.

    Args:
        filepath: Path to the file

    Returns:
        Tuple of (is_writeable, error_message)
    """
    is_writeable = False
    error_message = ""

    # If file doesn't exist, check if directory is writeable
    if not os.path.exists(filepath):
        directory = os.path.dirname(filepath) or "."
        if not os.path.exists(directory):
            error_message = f"Directory {directory} does not exist"
        elif not os.access(directory, os.W_OK):
            error_message = f"Directory {directory} is not writeable"
        else:
            is_writeable = True
    else:
        # If file exists, check if it's writeable
        if not os.access(filepath, os.W_OK):
            error_message = f"File {filepath} is not writeable (permission denied)"
        else:
            # Try to open the file for writing to see if it's locked
            try:
                with open(filepath, "a", encoding="utf-8"):
                    pass
                is_writeable = True
            except (IOError, OSError) as e:
                error_message = f"File {filepath} is not writeable: {str(e)}"

    return is_writeable, error_message


def create_document_copy(
    source_path: str, dest_path: Optional[str] = None
) -> Tuple[bool, str, Optional[str]]:
    """
    Create a copy of a document.

    Args:
        source_path: Path to the source document
        dest_path: Optional path for the new document. If not provided, will use
            source_path + '_copy.docx'

    Returns:
        Tuple of (success, message, new_filepath)
    """
    if not os.path.exists(source_path):
        return False, f"Source document {source_path} does not exist", None

    if not dest_path:
        # Generate a new filename if not provided
        base, ext = os.path.splitext(source_path)
        dest_path = f"{base}_copy{ext}"

    try:
        # Simple file copy
        shutil.copy2(source_path, dest_path)
        return True, f"Document copied to {dest_path}", dest_path
    except (IOError, shutil.Error, OSError) as e:
        return False, f"Failed to copy document: {str(e)}", None


def ensure_docx_extension(filename: str) -> str:
    """
    Ensure filename has .docx extension.

    Args:
        filename: The filename to check

    Returns:
        Filename with .docx extension
    """
    if not filename.endswith(".docx"):
        return filename + ".docx"
    return filename


def get_absolute_path(relative_path: str) -> str:
    """
    Convert a relative path to absolute path.

    Args:
        relative_path: The relative path to convert

    Returns:
        Absolute path
    """
    # Get absolute path based on current working directory
    return os.path.abspath(relative_path)


def get_project_root() -> str:
    """
    Get the project root directory.

    Returns:
        Absolute path to project root
    """
    # Get the directory of the current file
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # Go up two levels to reach project root
    return os.path.abspath(os.path.join(current_dir, "..", ".."))


def get_doc_path(doc_filename: str) -> str:
    """
    Get absolute path to a document in the docs directory.

    Args:
        doc_filename: Filename of the document

    Returns:
        Absolute path to the document
    """
    project_root = get_project_root()
    return os.path.join(project_root, "docs", doc_filename)
