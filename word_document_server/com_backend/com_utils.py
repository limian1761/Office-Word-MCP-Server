"""
COM utilities for Word Document MCP Server.

This module contains utility functions for working with COM objects,
including error handling and common operations.
"""

import functools
from typing import Any, Callable, TypeVar

import win32com.client

from word_document_server.errors import ErrorCode, WordDocumentError

T = TypeVar('T')


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
                raise WordDocumentError(
                    error_code, 
                    f"Failed to {operation_name}: {e}"
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
        with safe_com_call(ErrorCode.TEXT_FORMATTING_ERROR, "set bold formatting"):
            com_range_obj.Bold = 1
    """
    class SafeComCall:
        def __enter__(self):
            pass
            
        def __exit__(self, exc_type, exc_val, exc_tb):
            if exc_type is not None:
                raise WordDocumentError(
                    error_code,
                    f"Failed to {operation_name}: {exc_val}"
                )
            return False
    
    return SafeComCall()