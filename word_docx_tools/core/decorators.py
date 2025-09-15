"""Decorators for Word Document MCP Server.

This module contains decorators for handling errors, validating
function calls, and other cross-cutting concerns.
"""

import functools
import logging

from .errors import ErrorCode, WordDocumentError

logger = logging.getLogger("WordDocumentServer")


def handle_tool_errors(func):
    """Decorator to handle errors in MCP tools uniformly."""

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            # Log the error with context
            logger.error("Error in tool %s: %s", func.__name__, str(e), exc_info=True)
            # Return formatted error response as a dictionary
            # This ensures compatibility with Pydantic's return type validation
            error_message = format_error_response(e)
            return {"error": error_message}

    return wrapper


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


def handle_error(e: Exception) -> tuple[int, str, dict]:
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


def format_error(error_code: ErrorCode, message: str) -> dict:
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