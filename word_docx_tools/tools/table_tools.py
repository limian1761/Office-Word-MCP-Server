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
from ..mcp_service.core import mcp_server
from ..operations.table_ops import (create_table,
                                                       get_cell_text,
                                                       get_table_info,
                                                       insert_column,
                                                       insert_row,
                                                       set_cell_text)
from ..selector.selector import SelectorEngine
# 工具模块
from ..utils.app_context import AppContext
from ..mcp_service.core_utils import (
    ErrorCode, WordDocumentError, format_error_response, get_active_document,
    handle_tool_errors, log_error, log_info,
    require_active_document_validation)


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
        description="Table index (larger than 0) for operations that require specifying a table. Required for: get_cell, set_cell, get_info, insert_row, insert_column",
    ),
    rows: Optional[int] = Field(
        default=None,
        description="Number of rows when creating a table. Required for: create",
    ),
    cols: Optional[int] = Field(
        default=None,
        description="Number of columns when creating a table. Required for: create",
    ),
    row: Optional[int] = Field(
        default=None,
        description="Cell row number (1-based). Required for: get_cell, set_cell",
    ),
    col: Optional[int] = Field(
        default=None,
        description="Cell column number (1-based). Required for: get_cell, set_cell",
    ),
    text: Optional[str] = Field(
        default=None,
        description="Text content for setting cell text. Required for: set_cell",
    ),
    formatting: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Optional formatting parameters dictionary. Optional for: set_cell",
    ),
    locator: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Object locator for specifying position when creating table. Required for: create",
    ),
    position: Optional[str] = Field(
        default=None,
        description="Insertion position, e.g. 'before', 'after' or row/column insertion position. Optional for: create, insert_row, insert_column",
    ),
    count: Optional[int] = Field(
        default=None,
        description="Number of rows/columns to insert. Optional for: insert_row, insert_column",
    ),
) -> str:
    """
    Unified table operation tool.

    This tool provides a single interface for all table operations:
    - create: Create a new table
      * Required parameters: rows, cols, locator
      * Optional parameters: position
    - get_cell: Get cell text
      * Required parameters: table_index, row, col
    - set_cell: Set cell text
      * Required parameters: table_index, row, col, text
      * Optional parameters: formatting
    - get_info: Get table information
      * Required parameters: None (不提供table_index则返回所有表格信息)
      * Optional parameters: table_index
    - insert_row: Insert rows
      * Required parameters: table_index
      * Optional parameters: position, count
    - insert_column: Insert columns
      * Required parameters: table_index
      * Optional parameters: position, count

    Returns:
        Operation result based on the operation type
    """
    try:
        # 获取活动文档
        active_doc = ctx.request_context.lifespan_context.get_active_document()

        # 根据操作类型执行相应的操作
        if operation_type and operation_type.lower() == "create":
            if rows is None or cols is None or locator is None:
                raise ValueError(
                    "rows, cols, and locator parameters must be provided for create operation"
                )

            log_info(f"Creating table with {rows} rows and {cols} columns")
            result = create_table(
                document=active_doc,
                rows=rows,
                cols=cols,
                locator=locator,
                position=position,
            )
            log_info("Table created successfully")
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
            log_info("Cell text retrieved successfully")
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
            log_info("Cell text set successfully")
            return str(result)

        elif operation_type and operation_type.lower() == "get_info":
            log_info(f"Getting info for table {table_index if table_index is not None else 'all tables'}")
            result = get_table_info(document=active_doc, table_index=table_index)
            log_info("Table info retrieved successfully")
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
            log_info("Row inserted successfully")
            return str(result)

        elif operation_type and operation_type.lower() == "insert_column":
            if table_index is None:
                raise ValueError(
                    "table_index parameter must be provided for insert_column operation"
                )
            
            # 处理position参数类型矛盾问题
            # 如果position是字符串类型，需要进行适当处理
            insert_position = position
            if isinstance(position, str):
                if position.lower() == "after":
                    # 如果是"after"，则设置为表格当前列数+1，表示在末尾插入
                    insert_position = None  # 由底层函数处理
                else:
                    # 尝试将字符串转换为整数
                    try:
                        insert_position = int(position)
                    except ValueError:
                        raise ValueError("position must be an integer or 'after' for insert_column operation")
            elif position is None:
                # 默认在末尾插入
                insert_position = None

            log_info(f"Inserting column in table {table_index} at position {insert_position}")
            result = insert_column(
                document=active_doc,
                table_index=table_index,
                position=insert_position if insert_position is not None else 9999,  # 使用大数值表示末尾
                count=count,
            )
            log_info("Column inserted successfully")
            return str(result)

        else:
            error_msg = f"Unsupported operation type: {operation_type}"
            log_error(error_msg)
            raise ValueError(error_msg)

    except Exception as e:
        log_error(f"Table operation failed: {str(e)}", exc_info=True)
        return str(format_error_response(e))
