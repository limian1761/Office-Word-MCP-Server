"""
Object Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for object operations.
"""

import json
import logging
import os
from typing import Any, Dict, List, Optional

# Standard library imports
from dotenv import load_dotenv
# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from ..mcp_service.core import mcp_server
from ..selector.selector import SelectorEngine
from ..utils.app_context import AppContext
from ..mcp_service.core_utils import (ErrorCode,
                                                   WordDocumentError,
                                                   get_active_document,
                                                   log_error, log_info)

# Configure logger
logger = logging.getLogger(__name__)

# 延迟导入以避免循环导入
def _import_range_operations():
    """延迟导入range操作函数以避免循环导入"""
    from ..operations.range_ops import (
        batch_apply_formatting, batch_select_objects, delete_object_by_locator,
        get_object_by_id, select_objects)

    return (
        batch_apply_formatting,
        batch_select_objects,
        delete_object_by_locator,
        get_object_by_id,
        select_objects,
    )


@mcp_server.tool()
async def range_tools(
    ctx: Context[ServerSession, AppContext],
    operation_type: str = Field(
        ...,
        description="Type of range operation: select, get_by_id, batch_select, batch_apply_formatting, delete",
    ),
    locator: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Range locator for range operations. Required for: select, delete",
    ),
    object_id: Optional[str] = Field(
        default=None,
        description="Range ID for get_by_id operation. Required for: get_by_id",
    ),
    locators: Optional[List[Dict[str, Any]]] = Field(
        default=None,
        description="List of range locators for batch operations. Required for: batch_select",
    ),
    operations: Optional[List[Dict[str, Any]]] = Field(
        default=None,
        description="List of operations for batch formatting. Required for: batch_apply_formatting",
    ),
) -> str:
    """
    Unified object operation tool.

    This tool provides a single interface for all object operations:
    - select: Select objects based on locator
      * Required parameters: locator
      * Optional parameters: None
    - get_by_id: Get object by ID
      * Required parameters: object_id
      * Optional parameters: None
    - batch_select: Select multiple objects based on locators
      * Required parameters: locators
      * Optional parameters: None
    - batch_apply_formatting: Apply formatting to multiple objects
      * Required parameters: operations
      * Optional parameters: None
    - delete: Delete object by locator
      * Required parameters: locator
      * Optional parameters: None

    Returns:
        Result of the operation in JSON format
    """
    # Get the active Word document
    active_doc = ctx.request_context.lifespan_context.get_active_document()

    # Check if there is an active document
    if active_doc is None:
        raise WordDocumentError(
            ErrorCode.NO_ACTIVE_DOCUMENT, "No active document found"
        )

    # 延迟导入range操作函数以避免循环导入
    (
        batch_apply_formatting,
        batch_select_objects,
        delete_object_by_locator,
        get_object_by_id,
        select_objects,
    ) = _import_range_operations()

    try:
        if operation_type == "select":
            if locator is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT, "Locator is required for select operation"
                )

            log_info(f"Selecting objects with locator: {locator}")
            result = select_objects(active_doc, [locator])
            log_info(f"Successfully selected {len(result) if result else 0} objects")
            return json.dumps(
                {
                    "success": True,
                    "objects": result,
                    "message": "Objects selected successfully",
                },
                ensure_ascii=False,
            )

        elif operation_type == "get_by_id":
            if object_id is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Object ID is required for get_by_id operation",
                )

            log_info(f"Getting object by ID: {object_id}")
            result = get_object_by_id(active_doc, object_id)
            log_info("Object retrieved successfully" if result else "Object not found")
            return json.dumps(
                {
                    "success": True,
                    "object": result,
                    "message": "Object retrieved successfully",
                },
                ensure_ascii=False,
            )

        elif operation_type == "batch_select":
            if locators is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Locators list is required for batch_select operation",
                )

            log_info(
                f"Batch selecting objects with {len(locators) if locators else 0} locators"
            )
            result = batch_select_objects(active_doc, locators)
            log_info(
                f"Successfully selected {len(result) if result else 0} objects in batch"
            )
            return json.dumps(
                {
                    "success": True,
                    "objects": result,
                    "message": "Objects selected successfully",
                },
                ensure_ascii=False,
            )

        elif operation_type == "batch_apply_formatting":
            if operations is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "Operations list is required for batch_apply_formatting operation",
                )

            log_info(
                f"Applying batch formatting with {len(operations) if operations else 0} operations"
            )
            result = batch_apply_formatting(active_doc, operations)
            successful_ops = (
                sum(1 for r in result if r.get("success", False)) if result else 0
            )
            log_info(
                f"Batch formatting completed. {successful_ops}/{len(result) if result else 0} operations successful"
            )
            return json.dumps(
                {
                    "success": True,
                    "results": result,
                    "message": "Formatting applied successfully",
                },
                ensure_ascii=False,
            )

        elif operation_type == "delete":
            if locator is None:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT, "Locator is required for delete operation"
                )

            log_info(f"Deleting object with locator: {locator}")
            result = delete_object_by_locator(active_doc, locator)
            log_info(
                "Object deleted successfully" if result else "Failed to delete object"
            )
            return json.dumps(
                {
                    "success": result,
                    "message": (
                        "Object deleted successfully"
                        if result
                        else "Failed to delete object"
                    ),
                },
                ensure_ascii=False,
            )

        else:
            error_msg = f"Unsupported operation type: {operation_type}"
            log_error(error_msg)
            raise WordDocumentError(ErrorCode.INVALID_INPUT, error_msg)

    except Exception as e:
        log_error(f"Error in object_tools: {str(e)}", exc_info=True)
        raise
