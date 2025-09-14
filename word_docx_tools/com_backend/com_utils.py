"""
COM utilities for Word Document MCP Server.

This module contains utility functions for working with COM objects,
including error handling and common operations.
"""

import functools
from typing import Any, Callable, List, TypeVar

import win32com.client

from ..mcp_service.errors import ErrorCode, WordDocumentError

T = TypeVar("T")


def handle_com_error(error_code: ErrorCode, operation_name: str):
    """
    Decorator to handle common COM operation errors.

    Args:
        error_code: The error code to use for WordDocumentError
        operation_name: Name of the operation for error messages
    """

    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @functools.wraps(func)
        def wrapper(*args, **kwargs) -> T:
            try:
                return func(*args, **kwargs)
            except Exception as e:
                # 确保error_code是ErrorCode类型的实例
                if isinstance(error_code, ErrorCode):
                    raise WordDocumentError(
                        error_code, f"Failed to {operation_name}: {str(e)}"
                    )
                else:
                    # 如果不是ErrorCode类型，使用SERVER_ERROR作为默认错误码
                    raise WordDocumentError(
                        ErrorCode.SERVER_ERROR, f"Failed to {operation_name}: {str(e)}"
                    )

        return wrapper

    return decorator


def safe_com_call(error_code: ErrorCode, operation_name: str):
    """
    Context manager for safe COM calls.

    Args:
        error_code: The error code to use for WordDocumentError
        operation_name: Name of the operation for error messages

    Usage:
        with safe_com_call(ErrorCode.FORMATTING_ERROR, "set bold formatting"):
            com_range_obj.Bold = 1
    """

    class SafeComCall:
        def __enter__(self):
            pass

        def __exit__(self, exc_type, exc_val, exc_tb):
            if exc_type is not None:
                raise WordDocumentError(
                    error_code, f"Failed to {operation_name}: {exc_val}"
                )
            return False

    return SafeComCall()


def iter_com_collection(collection: Any) -> List[Any]:
    """
    Iterate through a COM collection and return a list of elements.

    Args:
        collection: The COM collection object to iterate through.

    Returns:
        A list containing all elements from the COM collection.
        
    Example:
        paragraphs = iter_com_collection(document.Paragraphs)
    """
    result = []
    try:
        count = collection.Count
        for i in range(1, count + 1):
            try:
                element = collection(i)
                result.append(element)
            except Exception:
                # Skip elements that can't be accessed
                continue
    except Exception:
        # If collection doesn't support Count property, return empty list
        pass
    return result
