# Error handling module for Word Document Server

import logging
from enum import Enum
from functools import wraps
from typing import Any, Dict, Optional, Tuple

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("word_document_server.log"), logging.StreamHandler()],
)
logger = logging.getLogger("WordDocumentServer")


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

    pass


class ImageFormatError(ImageError):
    """Raised when an image format is invalid"""

    def __init__(self, image_path: str, message: Optional[str] = None):
        details = {"image_path": image_path}
        super().__init__(ErrorCode.IMAGE_FORMAT_ERROR, message, details)


class CommentError(WordDocumentError):
    """Base exception for comment-related errors"""

    pass


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


def handle_error(e: Exception) -> Tuple[int, str, Dict[str, Any]]:
    """
    Handles errors and returns standardized error information.

    Args:
        e: The exception to handle

    Returns:
        A tuple of (error_code, error_message, error_details)
    """
    # Log the error
    logger.error(f"Error occurred: {str(e)}", exc_info=True)

    # Handle specific exceptions
    if isinstance(e, WordDocumentError):
        return e.error_code.value[0], e.message, e.details

    # Handle common exceptions
    elif isinstance(e, FileNotFoundError):
        return (
            ErrorCode.DOCUMENT_OPEN_ERROR.value[0],
            f"File not found: {str(e)}",
            {"file_path": str(e.filename) if hasattr(e, "filename") else None},
        )

    elif isinstance(e, ValueError):
        return (
            ErrorCode.INVALID_INPUT.value[0],
            f"Invalid value: {str(e)}",
            {"error_type": "ValueError"},
        )

    elif isinstance(e, PermissionError):
        return (
            ErrorCode.PERMISSION_DENIED.value[0],
            f"Permission denied: {str(e)}",
            {"error_type": "PermissionError"},
        )

    # Handle general exceptions
    else:
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
) -> Tuple[bool, str]:
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


def handle_tool_errors(func):
    """
    Decorator to handle errors in MCP tools uniformly.
    """

    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            # Log the error with context
            logger.error(f"Error in tool {func.__name__}: {str(e)}", exc_info=True)
            # Return formatted error response
            return format_error_response(e)

    return wrapper
