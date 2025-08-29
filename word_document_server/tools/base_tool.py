"""
Base Tool for Word Document MCP Server.

This module provides a base class for all tool implementations
to reduce code duplication and ensure consistent behavior.
"""

from typing import Any, Dict, List, Optional
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field
from word_document_server.utils.app_context import AppContext
from word_document_server.mcp_service.core import mcp_server, selector
from word_document_server.utils.core_utils import (
    format_error_response,
    handle_tool_errors,
    require_active_document_validation,
    ErrorCode,
    WordDocumentError
)

class BaseWordTool:
    """基类，封装所有工具函数的通用功能"""
    
    def __init__(self):
        pass
    
    @staticmethod
    def get_active_document(ctx: Context[ServerSession, AppContext]):
        """获取活动文档的通用方法"""
        return ctx.request_context.lifespan_context.get_active_document()
    
    @staticmethod
    def get_selector():
        """获取选择器引擎实例"""
        return selector