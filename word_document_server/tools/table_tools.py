"""
Table Tools for Word Document MCP Server.
This module provides a unified tool for table-related operations.
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
from word_document_server.operations.table_ops import (create_table,
                                                       get_table_info,
                                                       insert_column,
                                                       insert_row,
                                                       set_cell_text)
from word_document_server.selector.selector import SelectorEngine
# 工具模块
from word_document_server.utils.app_context import AppContext
from word_document_server.utils.core_utils import (
    ErrorCode, WordDocumentError, format_error_response, get_active_document,
    handle_tool_errors, log_error, log_info, require_active_document_validation)


@mcp_server.tool()
@handle_tool_errors
@require_active_document_validation
def table_tools(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    operation_type: Optional[str] = Field(
        default=None,
        description="Type of table operation: create, get_cell, set_cell, get_info, insert_row, insert_column",
    ),
    table_index: Optional[int] = Field(
        default=None,
        description="Table index (1-based) for operations that require specifying a table",
    ),
    rows: Optional[int] = Field(
        default=None, description="Number of rows when creating a table"
    ),
    cols: Optional[int] = Field(
        default=None, description="Number of columns when creating a table"
    ),
    row: Optional[int] = Field(default=None, description="Cell row number (1-based)"),
    col: Optional[int] = Field(
        default=None, description="Cell column number (1-based)"
    ),
    text: Optional[str] = Field(
        default=None, description="Text content for setting cell text"
    ),
    formatting: Optional[Dict[str, Any]] = Field(
        default=None, description="Optional formatting parameters dictionary"
    ),
    locator: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Element locator for specifying position when creating table",
    ),
    position: Optional[str] = Field(
        default=None,
        description="Insertion position, e.g. 'before', 'after' or row/column insertion position",
    ),
    count: Optional[int] = Field(
        default=None, description="Number of rows/columns to insert"
    ),
) -> str:
    """
    Unified table operation tool.

    This tool provides a single interface for all table operations:
    - create: Create a new table
    - get_cell: Get cell text
    - set_cell: Set cell text
    - get_info: Get table information
    - insert_row: Insert rows
    - insert_column: Insert columns

    Returns:
        Operation result based on the operation type
    """
    try:
        # 获取活动文档
        lifespan_context = getattr(ctx.request_context, "lifespan_context", None)
        active_doc = lifespan_context.get_active_document()


        # 根据操作类型执行相应的操作
        if operation_type and operation_type.lower() == "create":
            if rows is None or cols is None:
                raise ValueError(
                    "rows and cols parameters must be provided for create operation"
                )

            log_info(f"Creating table with {rows} rows and {cols} columns")
            result = create_table(
                document=active_doc,
                rows=rows,
                cols=cols,
                locator=locator,
                position=position,
            )
            return str(result)

        elif operation_type and operation_type.lower() == "get_cell":
            if table_index is None or row is None or col is None:
                raise ValueError(
                    "table_index, row, and col parameters must be provided for get_cell operation"
                )

            log_info(f"Getting text from table {table_index}, cell ({row}, {col})")
            result = get_cell_text(
                document=active_doc, table_index=table_index, row=row, col=col
            )
            return str(result)

        elif operation_type and operation_type.lower() == "set_cell":
            if table_index is None or row is None or col is None or text is None:
                raise ValueError(
                    "table_index, row, col, and text parameters must be provided for set_cell operation"
                )

            log_info(f"Setting text in table {table_index}, cell ({row}, {col})")
            result = set_cell_text(
                document=active_doc,
                table_index=table_index,
                row=row,
                col=col,
                text=text,
                formatting=formatting,
            )
            return str(result)

        elif operation_type and operation_type.lower() == "get_info":
            if table_index is None:
                raise ValueError(
                    "table_index parameter must be provided for get_info operation"
                )

            log_info(f"Getting info for table {table_index}")
            result = get_table_info(document=active_doc, table_index=table_index)
            return str(result)

        elif operation_type and operation_type.lower() == "insert_row":
            if table_index is None:
                raise ValueError(
                    "table_index parameter must be provided for insert_row operation"
                )

            log_info(f"Inserting row in table {table_index}")
            result = insert_row(
                document=active_doc,
                table_index=table_index,
                position=position,
                count=count,
            )
            return str(result)

        elif operation_type and operation_type.lower() == "insert_column":
            if table_index is None:
                raise ValueError(
                    "table_index parameter must be provided for insert_column operation"
                )

            log_info(f"Inserting column in table {table_index}")
            result = insert_column(
                document=active_doc,
                table_index=table_index,
                position=position,
                count=count,
            )
            return str(result)

        else:
            raise ValueError(f"Unsupported operation type: {operation_type}")

    except Exception as e:
        log_error(f"Table operation failed: {str(e)}", exc_info=True)
        return str(format_error_response(e))
