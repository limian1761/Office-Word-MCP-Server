"""
Objects Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for document objects operations.
"""

import json
import os
from typing import Any, Dict, List, Optional

# Standard library imports
from dotenv import load_dotenv
# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
import win32com.client
from pydantic import Field

# Local imports
from word_document_server.mcp_service.core import mcp_server
from word_document_server.operations.document_objects_ops import (
    create_bookmark as op_add_bookmark,
    delete_bookmark as op_delete_bookmark,
    get_bookmark as op_get_bookmark,
    create_citation as op_add_citation,
    create_hyperlink as op_add_hyperlink
)
from word_document_server.utils.app_context import AppContext
from word_document_server.utils.core_utils import (
    ErrorCode, WordDocumentError, format_error_response, get_active_document,
    handle_tool_errors, log_error, log_info, require_active_document_validation)

# 加载环境变量
try:
    load_dotenv()
except Exception as e:
    log_info("python-dotenv not installed, skipping .env file loading")


@mcp_server.tool()
def objects_tools(
    ctx: Context[ServerSession, AppContext] = Field(
        description="MCP context object containing session and application context information"
    ),
    operation_type: str = Field(
        ..., description="Operation type: bookmark_operations, citation_operations, hyperlink_operations"
    ),
    bookmark_name: Optional[str] = Field(
        default=None, description="Name of the bookmark. Required for bookmark_operations"
    ),
    citation_text: Optional[str] = Field(
        default=None, description="Text for the citation. Required for citation_operations"
    ),
    url: Optional[str] = Field(
        default=None, description="URL for the hyperlink. Required for hyperlink_operations"
    ),
    locator: Optional[Dict[str, Any]] = Field(
        default=None, description="Element locator for specifying position. Required for bookmark_operations, citation_operations, hyperlink_operations"
    ),
    sub_operation: Optional[str] = Field(
        default=None, description="Sub-operation type. Required for all operations"
    ),
    display_text: Optional[str] = Field(
        default=None, description="Display text for the hyperlink. Optional for hyperlink_operations"
    ),
    citation_name: Optional[str] = Field(
        default=None, description="Name of the citation. Optional for citation_operations"
    ),
    hyperlink_name: Optional[str] = Field(
        default=None, description="Name of the hyperlink. Optional for hyperlink_operations"
    ),
) -> Dict[str, Any]:
    """
    Document object operation tool

    Supported operation types:
    - bookmark_operations: Bookmark operations (create, get, delete)
    - citation_operations: Citation operations (create)
    - hyperlink_operations: Hyperlink operations (create)

    Args for bookmark_operations:
        Required parameters: bookmark_name, locator, sub_operation
        Optional parameters: None

    Args for citation_operations:
        Required parameters: citation_text, locator, sub_operation
        Optional parameters: citation_name

    Args for hyperlink_operations:
        Required parameters: url, locator, sub_operation
        Optional parameters: display_text, hyperlink_name

    Returns:
        Dictionary of operation results
    """
    try:
        # 验证是否有活动文档
        require_active_document_validation(ctx)

        # 获取活动文档
        document = get_active_document(ctx)

        # 处理不同类型的操作
        result: Dict[str, Any] = {}
        if operation_type == "bookmark_operations":
            result = handle_bookmark_operations(ctx, document, sub_operation, bookmark_name=bookmark_name, locator=locator)
        elif operation_type == "citation_operations":
            result = handle_citation_operations(ctx, document, sub_operation, citation_text=citation_text, locator=locator, citation_name=citation_name)
        elif operation_type == "hyperlink_operations":
            result = handle_hyperlink_operations(ctx, document, sub_operation, url=url, locator=locator, display_text=display_text, hyperlink_name=hyperlink_name)
        else:
            raise ValueError(f"不支持的操作类型: {operation_type}")

        return result
    except Exception as e:
        error_message = format_error_response(e)
        return {"error": error_message}  # 返回包含错误信息的字典


@handle_tool_errors
def handle_objects_operations(
    ctx: Context[ServerSession, AppContext],
    operation_type: str,
    bookmark_name: Optional[str] = None,
    citation_text: Optional[str] = None,
    url: Optional[str] = None,
    locator: Optional[Dict[str, Any]] = None,
    sub_operation: Optional[str] = None,
    display_text: Optional[str] = None,
    citation_name: Optional[str] = None,
    hyperlink_name: Optional[str] = None,
) -> Dict[str, Any]:
    """
    处理文档对象相关操作的统一入口

    Args:
        ctx: MCP上下文对象
        operation_type: 操作类型
        bookmark_name: 书签名称
        citation_text: 引用文本
        url: 超链接URL
        locator: 元素定位器
        sub_operation: 子操作类型
        display_text: 超链接显示文本
        citation_name: 引用名称
        hyperlink_name: 超链接名称

    Returns:
        操作结果字典
    """
    try:
        # 验证是否有活动文档
        require_active_document_validation(ctx)

        # 获取活动文档
        document = get_active_document(ctx)

        # 处理不同类型的操作
        result: Dict[str, Any] = {}
        if operation_type == "bookmark_operations":
            result = handle_bookmark_operations(ctx, document, sub_operation, bookmark_name=bookmark_name, locator=locator)
        elif operation_type == "citation_operations":
            result = handle_citation_operations(ctx, document, sub_operation, citation_text=citation_text, locator=locator, citation_name=citation_name)
        elif operation_type == "hyperlink_operations":
            result = handle_hyperlink_operations(ctx, document, sub_operation, url=url, locator=locator, display_text=display_text, hyperlink_name=hyperlink_name)
        else:
            raise ValueError(f"不支持的操作类型: {operation_type}")

        return result
    except Exception as e:
        error_message = format_error_response(e)
        return {"error": error_message}  # 返回包含错误信息的字典


@handle_tool_errors
def handle_bookmark_operations(
    ctx: Context[ServerSession, AppContext],
    document: win32com.client.CDispatch,
    sub_operation: str,
    **kwargs,
) -> Dict[str, Any]:
    """
    处理书签相关操作

    Args:
        ctx: MCP上下文对象
        document: Word文档COM对象
        sub_operation: 子操作类型
        **kwargs: 操作参数

    Returns:
        操作结果字典
    """
    result: Dict[str, Any] = {}

    if sub_operation == "create":
        bookmark_name = kwargs.get("bookmark_name")
        locator = kwargs.get("locator")
        if bookmark_name:
            result = op_add_bookmark(document, bookmark_name, locator)
        else:
            raise ValueError("bookmark_name is required for create operation")

    elif sub_operation == "get":
        bookmark_name = kwargs.get("bookmark_name")
        if bookmark_name:
            result = op_get_bookmark(document, bookmark_name)
        else:
            raise ValueError("bookmark_name is required for get operation")

    elif sub_operation == "delete":
        bookmark_name = kwargs.get("bookmark_name")
        if bookmark_name:
            result = op_delete_bookmark(document, bookmark_name)
        else:
            raise ValueError("bookmark_name is required for delete operation")

    else:
        raise ValueError(f"不支持的书签操作: {sub_operation}")

    return result


@handle_tool_errors
def handle_citation_operations(
    ctx: Context[ServerSession, AppContext],
    document: win32com.client.CDispatch,
    sub_operation: str,
    **kwargs,
) -> Dict[str, Any]:
    """
    处理引用相关操作

    Args:
        ctx: MCP上下文对象
        document: Word文档COM对象
        sub_operation: 子操作类型
        **kwargs: 操作参数

    Returns:
        操作结果字典
    """
    result: Dict[str, Any] = {}

    if sub_operation == "create":
        citation_text = kwargs.get("citation_text")
        locator = kwargs.get("locator")
        if citation_text:
            result = op_add_citation(document, citation_text, locator)
        else:
            raise ValueError("citation_text is required for create operation")

    elif sub_operation == "get":
        citation_name = kwargs.get("citation_name")
        if citation_name:
            # 由于没有提供get_citation函数，这里暂时抛出异常
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, 
                "Get citation operation is not implemented"
            )
        else:
            raise ValueError("citation_name is required for get operation")

    elif sub_operation == "delete":
        citation_name = kwargs.get("citation_name")
        if citation_name:
            # 由于没有提供delete_citation函数，这里暂时抛出异常
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, 
                "Delete citation operation is not implemented"
            )
        else:
            raise ValueError("citation_name is required for delete operation")

    else:
        raise ValueError(f"不支持的引用操作: {sub_operation}")

    return result


@handle_tool_errors
def handle_hyperlink_operations(
    ctx: Context[ServerSession, AppContext],
    document: win32com.client.CDispatch,
    sub_operation: str,
    **kwargs,
) -> Dict[str, Any]:
    """
    处理超链接相关操作

    Args:
        ctx: MCP上下文对象
        document: Word文档COM对象
        sub_operation: 子操作类型
        **kwargs: 操作参数

    Returns:
        操作结果字典
    """
    result: Dict[str, Any] = {}

    if sub_operation == "create":
        url = kwargs.get("url")
        locator = kwargs.get("locator")
        display_text = kwargs.get("display_text")
        if url and locator:
            result = op_add_hyperlink(document, url, locator, display_text)
        else:
            raise ValueError("url and locator are required for create operation")

    elif sub_operation == "get":
        hyperlink_name = kwargs.get("hyperlink_name")
        if hyperlink_name:
            # 由于没有提供get_hyperlink函数，这里暂时抛出异常
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, 
                "Get hyperlink operation is not implemented"
            )
        else:
            raise ValueError("hyperlink_name is required for get operation")

    elif sub_operation == "delete":
        hyperlink_name = kwargs.get("hyperlink_name")
        if hyperlink_name:
            # 由于没有提供delete_hyperlink函数，这里暂时抛出异常
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, 
                "Delete hyperlink operation is not implemented"
            )
        else:
            raise ValueError("hyperlink_name is required for delete operation")

    else:
        raise ValueError(f"不支持的超链接操作: {sub_operation}")

    return result
