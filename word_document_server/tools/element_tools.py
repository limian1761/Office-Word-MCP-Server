"""
Element Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for element operations.
"""

import json
import os
from typing import Any, Dict, List, Optional

# Standard library imports
from dotenv import load_dotenv
# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from word_document_server.mcp_service.core import mcp_server
from word_document_server.selector.selector import SelectorEngine
from word_document_server.utils.core_utils import (
    ErrorCode, WordDocumentError, log_error, log_info, get_active_document)
from word_document_server.utils.app_context import AppContext

# 延迟导入以避免循环导入
def _import_element_operations():
    """延迟导入element操作函数以避免循环导入"""
    from word_document_server.operations.element_selection_ops import (
        batch_apply_formatting, batch_select_elements, delete_element_by_locator,
        get_element_by_id, select_elements)
    return (batch_apply_formatting, batch_select_elements, delete_element_by_locator,
            get_element_by_id, select_elements)


@mcp_server.tool()
async def element_tools(
    ctx: Context,
    operation_type: str = Field(
        ..., 
        description="Type of element operation: select, get_by_id, batch_select, batch_apply_formatting, delete",
    ),
    locator: Optional[Dict[str, Any]] = Field(
        default=None, description="Element locator for element operations"
    ),
    element_id: Optional[str] = Field(
        default=None, description="Element ID for get_by_id operation"
    ),
    locators: Optional[List[Dict[str, Any]]] = Field(
        default=None, description="List of element locators for batch operations"
    ),
    operations: Optional[List[Dict[str, Any]]] = Field(
        default=None, description="List of operations for batch formatting"
    ),
) -> str:
    """
    Unified element operation tool.

    This tool provides a single interface for all element operations:
    - select: Select elements based on locator
    - get_by_id: Get element by ID
    - batch_select: Select multiple elements based on locators
    - batch_apply_formatting: Apply formatting to multiple elements
    - delete: Delete element by locator

    Args:
        ctx: MCP server context
        operation_type: Type of element operation to perform
        locator: Element locator for element operations
        element_id: Element ID for get_by_id operation
        locators: List of element locators for batch operations
        operations: List of operations for batch formatting

    Returns:
        Result of the operation in JSON format
    """
    # Get the active Word document
    app_context = AppContext.get_instance()
    document = app_context.get_active_document()
    
    # Check if there is an active document
    if document is None:
        raise WordDocumentError(
            ErrorCode.NO_ACTIVE_DOCUMENT, "No active document found"
        )
    
    # 延迟导入element操作函数以避免循环导入
    (batch_apply_formatting, batch_select_elements, delete_element_by_locator,
     get_element_by_id, select_elements) = _import_element_operations()

    try:
        if operation_type == "select":
            if locator is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT, "Locator is required for select operation"
                )
            
            result = select_elements(document, [locator])
            return json.dumps({
                "success": True,
                "elements": result,
                "message": "Elements selected successfully"
            }, ensure_ascii=False)

        elif operation_type == "get_by_id":
            if element_id is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT, "Element ID is required for get_by_id operation"
                )
                
            result = get_element_by_id(document, element_id)
            return json.dumps({
                "success": True,
                "element": result,
                "message": "Element retrieved successfully"
            }, ensure_ascii=False)

        elif operation_type == "batch_select":
            if locators is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT, "Locators list is required for batch_select operation"
                )
                
            result = batch_select_elements(document, locators)
            return json.dumps({
                "success": True,
                "elements": result,
                "message": "Elements selected successfully"
            }, ensure_ascii=False)

        elif operation_type == "batch_apply_formatting":
            if operations is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT, "Operations list is required for batch_apply_formatting operation"
                )
                
            result = batch_apply_formatting(document, operations)
            return json.dumps({
                "success": True,
                "results": result,
                "message": "Formatting applied successfully"
            }, ensure_ascii=False)

        elif operation_type == "delete":
            if locator is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT, "Locator is required for delete operation"
                )
                
            result = delete_element_by_locator(document, locator)
            return json.dumps({
                "success": result,
                "message": "Element deleted successfully" if result else "Failed to delete element"
            }, ensure_ascii=False)

        else:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT, f"Unsupported operation type: {operation_type}"
            )

    except Exception as e:
        logger.error(f"Error in element_tools: {str(e)}", exc_info=True)
        raise
