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
from ..selector.locator_parser import LocatorParser
from ..selector.exceptions import LocatorSyntaxError
from ..operations.paragraphs_ops import (
    get_paragraphs_info,
    insert_paragraph_impl,
    delete_paragraph_impl,
    format_paragraph_impl,
    get_paragraphs_details
)

@mcp_server.tool()
@require_active_document_validation
@handle_tool_errors
def paragraph_tools(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    operation_type: Optional[str] = Field(
        default=None,
        description="段落操作类型: insert_paragraph(插入段落), delete_paragraph(删除段落), format_paragraph(格式化段落), get_paragraphs_details(获取段落详情)",
    ),
    locator: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Locator object for paragraph selection. Returns all paragraph when empty.\n\n    Required for: insert_paragraph, delete_paragraph, format_paragraph\n    Optional for: get_paragraphs_details"),
    text: Optional[str] = Field(
        default=None,
        description="Text content for insert operation\n\n    Required for: insert_paragraph"),
    style: Optional[str] = Field(
        default=None,
        description="Paragraph style name\n\n Optional for: insert_paragraph"),
    is_independent_paragraph: bool = Field(
        default=True,
        description="Whether to insert the paragraph as an independent paragraph\n\n Optional for: insert_paragraph"),
    formatting: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Dictionary containing the paragraph_style to apply.\n\n Required for: format_paragraph\n    Must contain 'paragraph_style' key"),
    include_stats: bool = Field(
        default=False,
        description="Whether to include paragraph statistics in the result\n\n Optional for: get_paragraphs_details"),
) -> Any:
    """段落操作工具，支持获取段落信息、插入段落、删除段落和格式化段落等操作。

    支持的操作类型：
    - insert_paragraph: 在指定位置插入新段落
      * 必需参数：text, locator
      * 可选参数：style, is_independent_paragraph
    - delete_paragraph: 删除指定的段落
      * 必需参数：locator
      * 可选参数：无
    - format_paragraph: 格式化段落
      * 必需参数：locator, formatting
      * 可选参数：无
    - get_paragraphs_details: 获取段落详情（合并版，可同时获取段落列表和统计信息）
      * 必需参数：无
      * 可选参数：locator, include_stats

    返回：
        操作结果的JSON字符串
    """
    # 检查locator参数类型和规范
    def check_locator_param(locator_value):
        if locator_value is not None:
            # 检查是否为字典类型
            if not isinstance(locator_value, dict):
                raise TypeError("locator parameter must be a dictionary")
            
            # 使用LocatorParser验证locator结构
            parser = LocatorParser()
            try:
                parser.validate_locator(locator_value)
            except LocatorSyntaxError:
                # 提示用户参考定位器指南
                raise ValueError("Invalid locator format. Please refer to the locator guide for proper syntax.")
    
    try:
        log_info(f"Starting paragraph operation: {operation_type}")

        # 获取活动文档
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if not active_doc:
            raise WordDocumentError(
                ErrorCode.NO_ACTIVE_DOCUMENT, "No active document found"
            )

        # 根据操作类型调用相应的处理函数
        if operation_type == "insert_paragraph":
            if not text:
                raise WordDocumentError(
                    ErrorCode.PARAMETER_ERROR,
                    "text parameter must be provided for insert_paragraph operation"
                )
            if not locator:
                raise WordDocumentError(
                    ErrorCode.PARAMETER_ERROR,
                    "locator parameter must be provided for insert_paragraph operation"
                )
            
            # 检查locator参数
            check_locator_param(locator)
            
            result = insert_paragraph_impl(active_doc, text, locator, style, is_independent_paragraph)
            return json.dumps(
                {"success": True, "result": result}, ensure_ascii=False
            )

        elif operation_type == "delete_paragraph":
            if not locator:
                raise WordDocumentError(
                    ErrorCode.PARAMETER_ERROR,
                    "locator parameter must be provided for delete_paragraph operation"
                )
            
            # 检查locator参数
            check_locator_param(locator)
            
            result = delete_paragraph_impl(active_doc, locator)
            return result

        elif operation_type == "format_paragraph":
            if not locator:
                raise WordDocumentError(
                    ErrorCode.PARAMETER_ERROR,
                    "locator parameter must be provided for format_paragraph operation"
                )
            
            # 检查locator参数
            check_locator_param(locator)
            
            if not formatting or not isinstance(formatting, dict):
                raise WordDocumentError(
                    ErrorCode.PARAMETER_ERROR,
                    "formatting parameter must be a non-empty dictionary"
                )
            
            result = format_paragraph_impl(active_doc, locator, formatting)
            return json.dumps(
                {"success": True, "result": result}, ensure_ascii=False
            )
        elif operation_type == "get_paragraphs_details":
            # 获取段落详情（合并版，可同时获取段落列表和统计信息）
            # locator是可选参数
            if locator:
                # 检查locator参数
                check_locator_param(locator)
            
            result = get_paragraphs_details(active_doc, locator, include_stats)
            return json.dumps(
                {"success": True, "result": result}, ensure_ascii=False
            )
        else:
            raise WordDocumentError(
                ErrorCode.UNSUPPORTED_OPERATION,
                f"Unsupported operation type: {operation_type}"
            )

    except Exception as e:
        log_error(f"Error in paragraph_tools: {e}", exc_info=True)
        return format_error_response(e)