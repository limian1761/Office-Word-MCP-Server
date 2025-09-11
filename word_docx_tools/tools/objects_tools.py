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
from ..selector.selector import SelectorEngine
from ..selector.locator_parser import LocatorParser
from ..selector.exceptions import LocatorSyntaxError
from ..mcp_service.app_context import AppContext

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
    locator: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Object locator for specifying position. Required for bookmark_operations, citation_operations, hyperlink_operations",
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
    # 导入通用的locator参数检查函数
    from .utils import check_locator_param
    
    try:
        # 验证是否有活动文档
        require_active_document_validation(ctx)

        # 获取活动文档
        document = get_active_document(ctx)

        # 处理不同类型的操作
        result: Dict[str, Any] = {}
        if operation_type == "bookmark_operations":
            result = handle_bookmark_operations(
                ctx,
                document,
                sub_operation,
                bookmark_name=bookmark_name,
                locator=locator,
            )
        elif operation_type == "citation_operations":
            result = handle_citation_operations(
                ctx,
                document,
                sub_operation,
                citation_text=citation_text,
                locator=locator,
                citation_name=citation_name,
            )
        elif operation_type == "hyperlink_operations":
            result = handle_hyperlink_operations(
                ctx,
                document,
                sub_operation,
                url=url,
                locator=locator,
                display_text=display_text,
                hyperlink_name=hyperlink_name,
            )
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
        if bookmark_name and locator:
            try:
                # 检查locator参数
                check_locator_param(locator)
                # 改进书签创建，正确处理Range对象
                # 确保书签名称不包含非法字符
                clean_bookmark_name = bookmark_name
                for char in ["/", "\\", ":", "*", "?", '"', "<", ">", "|"]:
                    clean_bookmark_name = clean_bookmark_name.replace(char, "_")

                engine = SelectorEngine()
                selection = engine.select(document, locator)
                if not selection:
                    raise WordDocumentError(
                        ErrorCode.SELECTOR_ERROR,
                        "Failed to locate position for bookmark",
                    )

                # 正确处理selection对象
                if hasattr(selection, "_com_ranges") and selection._com_ranges:
                    ranges = selection._com_ranges
                else:
                    # 如果是段落对象，使用其Range
                    try:
                        ranges = [selection.Range]
                    except AttributeError:
                        # 如果没有Range属性，尝试直接使用selection
                        ranges = [selection]

                # 创建书签
                range_obj = ranges[0]
                # 如果Range为空，可以插入一个空字符作为书签位置
                if range_obj.Start == range_obj.End:
                    # 如果Range为空，插入一个空字符
                    range_obj.InsertAfter(" ")
                    range_obj.Collapse(True)  # 折叠到开始位置

                bookmark = document.Bookmarks.Add(
                    Name=clean_bookmark_name, Range=range_obj
                )
                result = {"success": True, "bookmark_name": bookmark.Name}
            except Exception as e:
                log_error(f"Failed to create bookmark: {str(e)}")
                raise WordDocumentError(
                    ErrorCode.OBJECT_TYPE_ERROR, f"Failed to create bookmark: {str(e)}"
                )
        else:
            raise ValueError(
                "bookmark_name and locator are required for create operation"
            )

    elif sub_operation == "get":
        bookmark_name = kwargs.get("bookmark_name")
        if bookmark_name:
            result = get_bookmark(document, bookmark_name)
        else:
            raise ValueError("bookmark_name is required for get operation")

    elif sub_operation == "delete":
        bookmark_name = kwargs.get("bookmark_name")
        if bookmark_name:
            result = delete_bookmark(document, bookmark_name)
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
        citation_name = kwargs.get("citation_name", "Citation")
        if citation_text and locator:
            # 修复引用创建问题，改进source_data格式以解决XML数据处理错误
            try:
                # 检查locator参数
                check_locator_param(locator)
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
                result = create_citation(document, source_data, locator)
            except Exception as e:
                log_error(f"Failed to create citation: {str(e)}")
                # 如果完整的引用创建失败，尝试使用更简单的方法在文档中插入引用文本
                try:
                    from ..operations.text_operations import insert_text

                    # 确保locator是字典类型
                    if locator is not None and not isinstance(locator, dict):
                        locator = {"type": "document", "position": "end"}
                    insert_text(document, locator, f"[{citation_text}]")
                    result = {
                        "warning": "Failed to create proper citation, inserted plain text instead"
                    }
                except Exception as e2:
                    raise WordDocumentError(
                        ErrorCode.OBJECT_TYPE_ERROR,
                        f"Failed to create citation: {str(e2)}",
                    )
        else:
            raise ValueError(
                "citation_text and locator are required for create operation"
            )

    elif sub_operation == "get":
        citation_name = kwargs.get("citation_name")
        if citation_name:
            # 由于没有提供get_citation函数，这里暂时抛出异常
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, "Get citation operation is not implemented"
            )
        else:
            raise ValueError("citation_name is required for get operation")

    elif sub_operation == "delete":
        citation_name = kwargs.get("citation_name")
        if citation_name:
            # 由于没有提供delete_citation函数，这里暂时抛出异常
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, "Delete citation operation is not implemented"
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
            try:
                # 检查locator参数
                check_locator_param(locator)
                # 改进超链接创建，确保使用正确的Range对象
                # 移除URL中可能存在的反引号和其他特殊字符
                clean_url = url.strip("`").strip()
                # 确保URL有协议前缀
                if not clean_url.startswith(
                    ("http://", "https://", "file://", "mailto:")
                ):
                    clean_url = "https://" + clean_url

                from ..operations.objects_ops import create_hyperlink

                result = create_hyperlink(
                    document,
                    address=clean_url,
                    locator=locator,
                    text_to_display=display_text,
                )
            except Exception as e:
                log_error(f"Failed to create hyperlink: {str(e)}")
                # 如果超链接创建失败，尝试使用更简单的方法在文档中插入链接文本
                try:
                    from ..operations.text_operations import insert_text

                    link_text = display_text if display_text else clean_url
                    # 确保locator是字典类型
                    if locator is not None and not isinstance(locator, dict):
                        locator = {"type": "document", "position": "end"}
                    insert_text(document, locator, f"{link_text}")
                    result = {
                        "warning": "Failed to create proper hyperlink, inserted plain text instead"
                    }
                except Exception as e2:
                    raise WordDocumentError(
                        ErrorCode.OBJECT_TYPE_ERROR,
                        f"Failed to create hyperlink: {str(e2)}",
                    )
        else:
            raise ValueError("url and locator are required for create operation")

    elif sub_operation == "get":
        hyperlink_name = kwargs.get("hyperlink_name")
        if hyperlink_name:
            # 由于没有提供get_hyperlink函数，这里暂时抛出异常
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, "Get hyperlink operation is not implemented"
            )
        else:
            raise ValueError("hyperlink_name is required for get operation")

    elif sub_operation == "delete":
        hyperlink_name = kwargs.get("hyperlink_name")
        if hyperlink_name:
            # 由于没有提供delete_hyperlink函数，这里暂时抛出异常
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR, "Delete hyperlink operation is not implemented"
            )
        else:
            raise ValueError("hyperlink_name is required for delete operation")

    else:
        raise ValueError(f"不支持的超链接操作: {sub_operation}")

    return result
