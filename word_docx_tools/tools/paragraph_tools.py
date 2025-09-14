"""Paragraph Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool for paragraph operations.
"""

import json
from typing import Any, Dict, List, Optional

# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from ..mcp_service.core import mcp_server
from ..mcp_service.app_context import AppContext
from ..mcp_service.core_utils import (
    ErrorCode,
    WordDocumentError,
    format_error_response,
    handle_tool_errors,
    log_error,
    log_info,
    require_active_document_validation
)
from ..operations.paragraphs_ops import (
    get_paragraphs_info,
    insert_paragraph_impl,
    delete_paragraph_impl,
    format_paragraph_impl,
    get_paragraphs_details
)
from ..operations.navigate_tools import set_active_context, set_active_object

@mcp_server.tool()
@require_active_document_validation
@handle_tool_errors
def paragraph_tools(
    context: Context,
    session: ServerSession,
    operation_type: str = Field(
        ...,
        description="段落操作类型: insert_paragraph(插入段落), delete_paragraph(删除段落), format_paragraph(格式化段落), get_paragraphs_details(获取段落详情)"
    ),
    text: Optional[str] = Field(
        None,
        description="插入段落的文本内容\n\n        Required for: insert_paragraph"
    ),
    style: Optional[str] = Field(
        None,
        description="段落样式名称\n\n         Optional for: insert_paragraph"
    ),
    is_independent_paragraph: bool = Field(
        True,
        description="是否将插入的段落作为独立段落\n\n         Optional for: insert_paragraph"
    ),
    formatting: Optional[Dict[str, Any]] = Field(
        None,
        description="包含要应用的段落样式的字典\n\n         Required for: format_paragraph"
    ),
    context_type: Optional[str] = Field(
        None,
        description="上下文类型，例如：document, section, paragraph, table"
    ),
    context_id: Optional[int] = Field(
        None,
        description="上下文ID，对应于特定类型的对象ID"
    ),
    object_type: Optional[str] = Field(
        None,
        description="对象类型，例如：paragraph, run, cell"
    ),
    object_id: Optional[int] = Field(
        None,
        description="对象ID，对应于特定类型的对象ID"
    )
) -> str:
    """段落操作工具，支持获取段落信息、插入段落、删除段落和格式化段落等操作。

    支持的操作类型：
    - insert_paragraph: 在指定位置插入新段落
      * 必需参数：text
      * 可选参数：style, is_independent_paragraph, context_type, context_id, object_type, object_id
    - delete_paragraph: 删除指定的段落
      * 必需参数：context_type, context_id, object_type, object_id（至少提供足够的上下文来定位段落）
      * 可选参数：无
    - format_paragraph: 格式化段落
      * 必需参数：formatting, context_type, context_id, object_type, object_id（至少提供足够的上下文来定位段落）
      * 可选参数：无
    - get_paragraphs_details: 获取段落详情（合并版，可同时获取段落列表和统计信息）
      * 必需参数：无
      * 可选参数：context_type, context_id, object_type, object_id

    返回：
        操作结果的JSON字符串
    """
    try:
        log_info(f"Starting paragraph operation: {operation_type}")
        
        # 获取活动文档
        active_doc = AppContext.get_instance().get_active_document()
        if not active_doc:
            raise WordDocumentError(
                ErrorCode.NO_ACTIVE_DOCUMENT,
                "No active document found"
            )
        
        # 设置上下文
        if context_type or context_id:
            set_active_context(context_type, context_id)
        
        if object_type or object_id:
            set_active_object(object_type, object_id)
        
        # 创建新的定位器或使用None
        locator = None
        
        # 执行相应的操作
        if operation_type == "get_paragraphs_details":
            # 获取段落详情
            result = get_paragraphs_details(active_doc, locator)
        elif operation_type == "insert_paragraph":
            # 插入段落
            if not text:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "参数 'text' 是插入段落操作所必需的"
                )
            result = insert_paragraph_impl(active_doc, text, locator, style, is_independent_paragraph)
        elif operation_type == "delete_paragraph":
            # 删除段落
            result = delete_paragraph_impl(active_doc, locator)
        elif operation_type == "format_paragraph":
            # 格式化段落
            if not formatting:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "参数 'formatting' 是格式化段落操作所必需的"
                )
            if "paragraph_style" not in formatting:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    "参数 'formatting' 必须包含 'paragraph_style' 键"
                )
            result = format_paragraph_impl(active_doc, locator, formatting)
        else:
            raise WordDocumentError(
                ErrorCode.UNSUPPORTED_OPERATION,
                f"不支持的段落操作类型: {operation_type}"
            )
        
        log_info(f"Paragraph operation {operation_type} completed successfully")
        return json.dumps({
            "success": True,
            "result": result
        })
    except Exception as e:
        return handle_tool_errors(e)