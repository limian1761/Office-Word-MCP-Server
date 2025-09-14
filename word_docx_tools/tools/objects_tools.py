"""
Objects Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for document objects operations.
"""

import json
import os
from typing import Any, Dict, List, Optional

import win32com.client
# Standard library imports
from dotenv import load_dotenv
# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from ..mcp_service.core import mcp_server
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError,
                                      format_error_response,
                                      get_active_document, handle_tool_errors,
                                      log_error, log_info,
                                      require_active_document_validation)
from ..operations.objects_ops import (create_bookmark, create_citation,
                                      create_hyperlink)
from ..mcp_service.app_context import AppContext
from ..operations.navigate_tools import set_active_context, set_active_object

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
        ...,
        description="Operation type: bookmark_operations, citation_operations, hyperlink_operations",
    ),
    bookmark_name: Optional[str] = Field(
        default=None,
        description="Name of the bookmark. Required for bookmark_operations",
    ),
    citation_text: Optional[str] = Field(
        default=None,
        description="Text for the citation. Required for citation_operations",
    ),
    url: Optional[str] = Field(
        default=None,
        description="URL for the hyperlink. Required for hyperlink_operations",
    ),
    context_type: Optional[str] = Field(
        default=None,
        description="Context type for specifying active context (e.g., 'document', 'section', 'paragraph')",
    ),
    context_id: Optional[str] = Field(
        default=None,
        description="Context ID for specifying active context",
    ),
    object_type: Optional[str] = Field(
        default=None,
        description="Object type for specifying active object within context",
    ),
    object_id: Optional[str] = Field(
        default=None,
        description="Object ID for specifying active object within context",
    ),
    sub_operation: Optional[str] = Field(
        default=None, description="Sub-operation type. Required for all operations" 
    ),
    display_text: Optional[str] = Field(
        default=None,
        description="Display text for the hyperlink. Optional for hyperlink_operations",
    ),
    citation_name: Optional[str] = Field(
        default=None,
        description="Name of the citation. Optional for citation_operations",
    ),
    hyperlink_name: Optional[str] = Field(
        default=None,
        description="Name of the hyperlink. Optional for hyperlink_operations",
    ),
) -> Dict[str, Any]:
    """文档对象操作工具

    支持的操作类型：
    - bookmark_operations: 书签操作（创建、获取、删除）
    - citation_operations: 引用操作（创建）
    - hyperlink_operations: 超链接操作（创建）

    bookmark_operations 参数：
        必需参数：bookmark_name, sub_operation
        可选参数：context_type, context_id, object_type, object_id

    citation_operations 参数：
        必需参数：citation_text, sub_operation
        可选参数：citation_name, context_type, context_id, object_type, object_id

    hyperlink_operations 参数：
        必需参数：url, sub_operation
        可选参数：display_text, hyperlink_name, context_type, context_id, object_type, object_id

    返回：
        操作结果的字典
    """
    
    try:
        # 验证是否有活动文档
        require_active_document_validation(ctx)

        # 获取活动文档
        document = get_active_document(ctx)

        # 设置活动上下文和对象
        set_active_context(ctx, context_type, context_id)
        set_active_object(ctx, object_type, object_id)

        # 处理不同类型的操作
        result: Dict[str, Any] = {}
        if operation_type == "bookmark_operations":
            result = handle_bookmark_operations(
                ctx,
                document,
                sub_operation,
                bookmark_name=bookmark_name,
            )
        elif operation_type == "citation_operations":
            result = handle_citation_operations(
                ctx,
                document,
                sub_operation,
                citation_text=citation_text,
                citation_name=citation_name,
            )
        elif operation_type == "hyperlink_operations":
            result = handle_hyperlink_operations(
                ctx,
                document,
                sub_operation,
                url=url,
                display_text=display_text,
                hyperlink_name=hyperlink_name,
            )
        else:
            raise WordDocumentError(
                ErrorCode.UNSUPPORTED_OPERATION,
                f"不支持的操作类型: {operation_type}"
            )

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
        if bookmark_name:
            try:
                # 确保书签名称不包含非法字符
                clean_bookmark_name = bookmark_name
                for char in ["/", "\\", ":", "*", "?", '"', "<", ">", "|"]:
                    clean_bookmark_name = clean_bookmark_name.replace(char, "_")

                # 创建书签
                result = create_bookmark(document, clean_bookmark_name)
            except Exception as e:
                log_error(f"Failed to create bookmark: {str(e)}")
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR, f"Failed to create bookmark: {str(e)}"
                )
        else:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                "bookmark_name is required for create operation"
            )

    elif sub_operation == "get":
        bookmark_name = kwargs.get("bookmark_name")
        if bookmark_name:
            result = get_bookmark(document, bookmark_name)
        else:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT, "bookmark_name is required for get operation"
            )

    elif sub_operation == "delete":
        bookmark_name = kwargs.get("bookmark_name")
        if bookmark_name:
            result = delete_bookmark(document, bookmark_name)
        else:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT, "bookmark_name is required for delete operation"
            )

    else:
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT, f"不支持的书签操作: {sub_operation}")

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
        citation_name = kwargs.get("citation_name", "Citation")
        if citation_text:
            # 修复引用创建问题，改进source_data格式以解决XML数据处理错误
            try:
                # 确保文档支持引用功能
                if not hasattr(document, "Bibliography"):
                    raise WordDocumentError(
                        ErrorCode.DOCUMENT_ERROR,
                        "Current document does not support bibliography/citation features",
                    )
                
                # 创建符合Word引用XML格式的source_data字典
                source_data = {
                    "Tag": citation_name,
                    "Author": "Author",
                    "Title": citation_text,
                    "Type": 1,  # 1代表普通引用类型
                    "Year": "2023",  # 添加必要的年份字段
                    "JournalName": "Journal",  # 添加必要的期刊名称字段
                    "Volume": "1",  # 添加必要的卷号字段
                    "Pages": "1-10",  # 添加必要的页码字段
                }
                # 从AppContext获取已设置的定位信息，不再需要locator参数
                result = create_citation(document, source_data)
            except Exception as e:
                log_error(f"Failed to create citation: {str(e)}")
                # 如果完整的引用创建失败，尝试使用更简单的方法在文档中插入引用文本
                try:
                    from ..operations.text_operations import insert_text
                    # 从AppContext获取定位信息插入文本
                    insert_text(document, f"[{citation_text}]")
                    result = {
                        "warning": "Failed to create proper citation, inserted plain text instead"
                    }
                except Exception as e2:
                    raise WordDocumentError(
                        ErrorCode.OBJECT_TYPE_ERROR,
                        f"Failed to create citation: {str(e2)}",
                    )
        else:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                "citation_text is required for create operation"
            )

    elif sub_operation == "get":
        citation_name = kwargs.get("citation_name")
        if citation_name:
            # 由于没有提供get_citation函数，这里暂时抛出异常
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, "Get citation operation is not implemented"
            )
        else:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT, "citation_name is required for get operation"
            )

    elif sub_operation == "delete":
        citation_name = kwargs.get("citation_name")
        if citation_name:
            # 由于没有提供delete_citation函数，这里暂时抛出异常
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, "Delete citation operation is not implemented"
            )
        else:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT, "citation_name is required for delete operation"
            )

    else:
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT, f"不支持的引用操作: {sub_operation}")

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
        display_text = kwargs.get("display_text", "")
        hyperlink_name = kwargs.get("hyperlink_name", "")
        if url:
            try:
                # 确保URL有效
                clean_url = url.strip("`").strip()
                if not clean_url.startswith(("http://", "https://", "mailto:", "file://")):
                    clean_url = "https://" + clean_url

                # 从AppContext获取已设置的定位信息
                result = create_hyperlink(
                    document,
                    address=clean_url,
                    text_to_display=display_text
                )
            except Exception as e:
                log_error(f"Failed to create hyperlink: {str(e)}")
                # 如果超链接创建失败，尝试使用更简单的方法在文档中插入链接文本
                try:
                    from ..operations.text_operations import insert_text

                    link_text = display_text if display_text else clean_url
                    insert_text(document, f"{link_text}")
                    result = {
                        "warning": "Failed to create proper hyperlink, inserted plain text instead"
                    }
                except Exception as e2:
                    raise WordDocumentError(
                        ErrorCode.OBJECT_TYPE_ERROR,
                        f"Failed to create hyperlink: {str(e2)}",
                    )
        else:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                "url is required for create operation"
            )

    elif sub_operation == "get":
        hyperlink_name = kwargs.get("hyperlink_name")
        if hyperlink_name:
            # 由于没有提供get_hyperlink函数，这里暂时抛出异常
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, "Get hyperlink operation is not implemented"
            )
        else:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT, "hyperlink_name is required for get operation"
            )

    elif sub_operation == "delete":
        hyperlink_name = kwargs.get("hyperlink_name")
        if hyperlink_name:
            # 由于没有提供delete_hyperlink函数，这里暂时抛出异常
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, "Delete hyperlink operation is not implemented"
            )
        else:
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT, "hyperlink_name is required for delete operation"
            )

    else:
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT, f"不支持的超链接操作: {sub_operation}")

    return result
