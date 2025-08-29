"""
Common imports for Word Document MCP Server tools.

This module consolidates common imports used across tool modules
to reduce code duplication.
"""

from typing import Any, Dict, List, Optional
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field
from word_document_server.utils.app_context import AppContext
from word_document_server.utils.core_utils import (
    ElementNotFoundError,
    format_error_response,
    handle_tool_errors,
    require_active_document_validation,
    ErrorCode,
    WordDocumentError,
    CommentEmptyError,
    CommentIndexError,
    ReplyEmptyError
)