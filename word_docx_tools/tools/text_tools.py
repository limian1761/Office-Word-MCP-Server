"""
Text Integration Tool for Word Document MCP Server.

This module provides a unified MCP tool interface for text operations,
which delegates the actual implementation to the operations layer.
"""

import json
import os
from typing import Any, Dict, List, Optional, Union

# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from ..mcp_service.core import mcp_server
from ..mcp_service.core_utils import (
    format_error_response,
    handle_tool_errors,
    log_error,
    log_info,
    require_active_document_validation
)
from ..operations.text_operations import (
    get_text_from_document,
    insert_text_into_document,
    replace_text_in_document,
    get_character_count_from_document,
    apply_formatting_to_document_text,
    validate_required_params
)
from ..operations.navigate_tools import set_active_context, set_active_object
from ..mcp_service.app_context import AppContext


# 定位器指南功能已移除，系统现在使用基于AppContext的上下文管理


@mcp_server.tool()
@require_active_document_validation
@handle_tool_errors
def text_tools(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    operation_type: Optional[str] = Field(
        default=None,
        description="Type of text operation: get_text, insert_text, replace_text, get_char_count, apply_formatting",
    ),
    context_type: Optional[str] = Field(
        default=None,
        description="Context type: section, paragraph, table, text, etc.",
    ),
    context_id: Optional[int] = Field(
        default=None,
        description="Context ID for the specific context type",
    ),
    object_type: Optional[str] = Field(
        default=None,
        description="Object type: paragraph, table, text, etc.",
    ),
    object_id: Optional[int] = Field(
        default=None,
        description="Object ID for the specific object type",
    ),
    text: Optional[str] = Field(
        default=None,
        description="Text content for insert or replace operations\n\n    Required for: insert_text, replace_text\n",
    ),
    position: str = Field(
        default="after",
        description="Position for insert operations: before, after, replace\n\n    Used by: insert_text\n",
    ),
    formatting: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Formatting options: bold, italic, font_size, font_name, font_color, alignment, Used for: apply_formatting",
    ),

) -> Any:
    """文本操作工具，支持获取文本内容、插入文本、替换文本、获取字符计数和应用文本格式等操作。

    支持的操作类型：
    - get_text: 从文档或特定上下文获取文本内容
      * 必需参数：无
      * 可选参数：context_type, context_id, object_type, object_id
    - insert_text: 在特定上下文位置插入文本
      * 必需参数：text
      * 可选参数：context_type, context_id, object_type, object_id, position
    - replace_text: 替换特定上下文中的文本内容
      * 必需参数：text
      * 可选参数：context_type, context_id, object_type, object_id
    - get_char_count: 获取文档或特定上下文的字符计数
      * 必需参数：无
      * 可选参数：context_type, context_id, object_type, object_id
    - apply_formatting: 对特定上下文中的文本应用格式设置
      * 必需参数：formatting
      * 可选参数：context_type, context_id, object_type, object_id

    返回：
        操作结果的JSON字符串
    """
    try:
        log_info(f"Starting text operation: {operation_type}")

        # 获取活动文档
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        
        # 设置活动上下文和对象（如果提供了参数）
        if context_type and context_id is not None:
            set_active_context(active_doc, context_type, context_id)
        if object_type and object_id is not None:
            set_active_object(active_doc, object_type, object_id)
        
        # 对于不需要定位的操作（整个文档操作），locator为None
        locator = None
        
        # 根据操作类型调用相应的处理函数
        if operation_type == "get_text":
            # 对于get_text，如果没有指定上下文，则获取整个文档的文本
            return get_text_from_document(active_doc)

        elif operation_type == "insert_text":
            # 验证必需参数
            validate_required_params({"text": text}, "insert_text")
            return insert_text_into_document(active_doc, text, position)

        elif operation_type == "replace_text":
            # 验证必需参数
            validate_required_params({"text": text}, "replace_text")
            return replace_text_in_document(active_doc, text)

        elif operation_type == "get_char_count":
            return get_character_count_from_document(active_doc)

        elif operation_type == "apply_formatting":
            # 验证必需参数
            validate_required_params({"formatting": formatting}, "apply_formatting")
            
            # 只使用formatting参数
            return apply_formatting_to_document_text(active_doc, formatting)

        else:
            raise ValueError(f"Unsupported operation type: {operation_type}")
    except Exception as e:
        log_error(f"Error in text_tools: {e}", exc_info=True)
        return format_error_response(e)
