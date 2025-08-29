"""
Utility functions shared across multiple Word document server modules.
"""

from typing import Any, Dict, List, Optional
import functools
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("word_document_server.log"), logging.StreamHandler()],
)
logger = logging.getLogger("WordDocumentServer")

from enum import Enum


class ErrorCode(Enum):
    """Standardized error codes for Word Document Server"""

    # General errors
    SUCCESS = (0, "Operation completed successfully")
    INVALID_INPUT = (1001, "Invalid input parameter")
    NOT_FOUND = (1002, "Requested resource not found")
    PERMISSION_DENIED = (1003, "Permission denied")
    SERVER_ERROR = (1004, "Internal server error")

    # Document errors
    NO_ACTIVE_DOCUMENT = (2001, "No active document")
    DOCUMENT_OPEN_ERROR = (2002, "Failed to open document")
    DOCUMENT_SAVE_ERROR = (2003, "Failed to save document")
    DOCUMENT_FORMAT_ERROR = (2004, "Invalid document format")

    # Element errors
    ELEMENT_NOT_FOUND = (3001, "Element not found")
    ELEMENT_LOCKED = (3002, "Element is locked")
    ELEMENT_TYPE_ERROR = (3003, "Invalid element type")
    PARAGRAPH_SELECTION_FAILED = (3004, "Failed to select paragraph elements")

    # Style errors
    STYLE_NOT_FOUND = (4001, "Style not found")
    STYLE_APPLY_ERROR = (4002, "Failed to apply style")

    # Formatting errors
    FORMATTING_ERROR = (5001, "Formatting error")

    # Image errors
    IMAGE_NOT_FOUND = (6001, "Image not found")
    IMAGE_FORMAT_ERROR = (6002, "Invalid image format")
    IMAGE_LOAD_ERROR = (6003, "Failed to load image")

    # Table errors
    TABLE_ERROR = (7001, "Table operation error")

    # Comment errors
    COMMENT_ERROR = (8001, "Comment operation error")
    COMMENT_INDEX_ERROR = (8002, "Comment index out of range")
    COMMENT_EMPTY_ERROR = (8003, "Comment text cannot be empty")
    REPLY_EMPTY_ERROR = (8004, "Reply text cannot be empty")


class WordDocumentError(Exception):
    """Base exception class for Word Document Server errors"""

    def __init__(
        self,
        error_code: ErrorCode,
        message: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None,
    ):
        self.error_code = error_code
        self.message = message or error_code.value[1]
        self.details = details or {}
        super().__init__(f"[{error_code.value[0]}] {self.message}")


class DocumentNotFoundError(WordDocumentError):
    """Raised when a document is not found"""

    def __init__(self, document_path: str, message: Optional[str] = None):
        details = {"document_path": document_path}
        super().__init__(ErrorCode.DOCUMENT_OPEN_ERROR, message, details)


class ElementNotFoundError(WordDocumentError):
    """Raised when an element is not found"""

    def __init__(self, locator: Dict[str, Any], message: Optional[str] = None):
        details = {"locator": locator}
        super().__init__(ErrorCode.ELEMENT_NOT_FOUND, message, details)


class StyleNotFoundError(WordDocumentError):
    """Raised when a style is not found"""

    def __init__(
        self,
        style_name: str,
        message: Optional[str] = None,
        similar_styles: Optional[list] = None,
    ):
        details = {"style_name": style_name}
        if similar_styles:
            details["similar_styles"] = similar_styles
        super().__init__(ErrorCode.STYLE_NOT_FOUND, message, details)


class ImageError(WordDocumentError):
    """Base exception for image-related errors"""


class ImageFormatError(ImageError):
    """Raised when an image format is invalid"""

    def __init__(self, image_path: str, message: Optional[str] = None):
        details = {"image_path": image_path}
        super().__init__(ErrorCode.IMAGE_FORMAT_ERROR, message, details)


class CommentError(WordDocumentError):
    """Base exception for comment-related errors"""


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

    elif error_code == ErrorCode.ELEMENT_NOT_FOUND.value[0] and "locator" in details:
        response += " Please check your locator parameters."

    # Add a generic suggestion for resolving the issue
    response += " For more information, check the server logs."

    return response


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


def validate_active_document(ctx):
    """验证活动文档是否存在。"""
    if not ctx or not hasattr(ctx, 'request_context') or not ctx.request_context:
        return "Request context not found."
    if not hasattr(ctx.request_context, 'lifespan_context') or not ctx.request_context.lifespan_context:
        return "Lifespan context not found."
    active_doc = ctx.request_context.lifespan_context.get_active_document()
    if not active_doc:
        return "No active document found. Please open a document first."
    return None

def require_active_document_validation(func):
    """验证活动文档的装饰器。"""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        # 获取上下文参数（ctx通常是第一个参数）
        ctx = args[0] if args else kwargs.get('ctx')
        if not ctx:
            return "Context object not found in function parameters."

        # 验证活动文档
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

def handle_tool_errors(func):
    """ Decorator to handle errors in MCP tools uniformly. """
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

