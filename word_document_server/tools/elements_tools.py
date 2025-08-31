"""
Word Document Elements Tools
This module provides functions for working with various document elements like bookmarks, citations, and hyperlinks in Word documents.
"""
from typing import Dict, List, Any, Optional, Union
import win32com.client
from mcp.context import Context
from mcp.server import ServerSession
from ..app_context import AppContext
from ..operations.document_objects_ops import (
    create_bookmark as op_create_bookmark,
    get_bookmark as op_get_bookmark,
    delete_bookmark as op_delete_bookmark,
    create_citation as op_create_citation,
    create_hyperlink as op_create_hyperlink
)
from ..utils.tool_utils import (
    require_active_document_validation,
    format_error_response,
    standardize_tool_errors,
    get_document_from_context
)


def elements_tools(
    ctx: Context[ServerSession, AppContext],
    operation_type: str,
    **kwargs
) -> Dict[str, Any]:
    """
    文档元素操作工具函数

    支持的操作类型:
    - bookmark_operations: 书签相关操作
      - create: 创建书签
      - get: 获取书签信息
      - delete: 删除书签
    - citation_operations: 引用相关操作
      - create: 创建引用
    - hyperlink_operations: 超链接相关操作
      - create: 创建超链接

    Args:
        ctx: MCP上下文对象
        operation_type: 操作类型
        **kwargs: 操作参数

    Returns:
        操作结果字典

    Raises:
        ValueError: 当参数无效时抛出
    """
    try:
        # 验证是否有活动文档
        require_active_document_validation(ctx)
        
        # 获取活动文档
        document = get_document_from_context(ctx)
        
        # 处理不同类型的操作
        if operation_type == "bookmark_operations":
            return handle_bookmark_operations(ctx, document, **kwargs)
        elif operation_type == "citation_operations":
            return handle_citation_operations(ctx, document, **kwargs)
        elif operation_type == "hyperlink_operations":
            return handle_hyperlink_operations(ctx, document, **kwargs)
        else:
            raise ValueError(f"不支持的操作类型: {operation_type}")
            
    except Exception as e:
        return format_error_response(str(e))


@standardize_tool_errors
def handle_bookmark_operations(
    ctx: Context[ServerSession, AppContext],
    document: win32com.client.CDispatch,
    sub_operation: str,
    **kwargs
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
    if sub_operation == "create":
        bookmark_name = kwargs.get("bookmark_name")
        locator = kwargs.get("locator")
        
        # 参数验证
        if not bookmark_name:
            raise ValueError("必须提供bookmark_name参数")
            
        return op_create_bookmark(document, bookmark_name, locator)
        
    elif sub_operation == "get":
        bookmark_name = kwargs.get("bookmark_name")
        
        # 参数验证
        if not bookmark_name:
            raise ValueError("必须提供bookmark_name参数")
            
        return op_get_bookmark(document, bookmark_name)
        
    elif sub_operation == "delete":
        bookmark_name = kwargs.get("bookmark_name")
        
        # 参数验证
        if not bookmark_name:
            raise ValueError("必须提供bookmark_name参数")
            
        return op_delete_bookmark(document, bookmark_name)
        
    else:
        raise ValueError(f"不支持的书签操作: {sub_operation}")


@standardize_tool_errors
def handle_citation_operations(
    ctx: Context[ServerSession, AppContext],
    document: win32com.client.CDispatch,
    sub_operation: str,
    **kwargs
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
    if sub_operation == "create":
        source_data = kwargs.get("source_data")
        locator = kwargs.get("locator")
        
        # 参数验证
        if not source_data:
            raise ValueError("必须提供source_data参数")
            
        return op_create_citation(document, source_data, locator)
        
    else:
        raise ValueError(f"不支持的引用操作: {sub_operation}")


@standardize_tool_errors
def handle_hyperlink_operations(
    ctx: Context[ServerSession, AppContext],
    document: win32com.client.CDispatch,
    sub_operation: str,
    **kwargs
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
    if sub_operation == "create":
        address = kwargs.get("address")
        locator = kwargs.get("locator")
        sub_address = kwargs.get("sub_address")
        screen_tip = kwargs.get("screen_tip")
        text_to_display = kwargs.get("text_to_display")
        
        # 参数验证
        if not address:
            raise ValueError("必须提供address参数")
            
        return op_create_hyperlink(
            document, 
            address, 
            locator, 
            sub_address, 
            screen_tip, 
            text_to_display
        )
        
    else:
        raise ValueError(f"不支持的超链接操作: {sub_operation}")