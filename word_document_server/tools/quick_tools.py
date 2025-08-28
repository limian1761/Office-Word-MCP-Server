"""
Quick tools for common Word document operations.

This module provides simplified interfaces for frequently used operations,
following the Occam's Razor principle of simplicity.
"""

import json
from typing import Any, Dict, Optional

from mcp.server.fastmcp.server import Context
from pydantic import Field
from word_document_server.core import ServerSession
from word_document_server.utils.app_context import AppContext

from word_document_server.core_utils import mcp_server
from word_document_server.errors import handle_tool_errors
from word_document_server.operations import (add_heading, add_table,
                                            get_all_paragraphs, get_all_tables,
                                            get_all_text,
                                            insert_paragraph_after)


@mcp_server.tool()
@handle_tool_errors
def add_heading_quick(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    text: str = Field(description="The heading text"),
    level: int = Field(description="The heading level (1-9)", default=1),
) -> str:
    """
    Quickly add a heading to the document.

    Returns:
        Success message
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        raise Exception(error)

    active_doc = ctx.request_context.lifespan_context.get_active_document()

    # Add heading at the end of document
    doc_range = active_doc.Range()
    doc_range.Collapse(Direction=0)  # Collapse to end

    heading = add_heading(active_doc, doc_range, text, level)
    return f"Successfully added heading: {text}"


@mcp_server.tool()
@handle_tool_errors
def add_paragraph_quick(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    text: str = Field(description="The paragraph text"),
) -> str:
    """
    Quickly add a paragraph to the document.

    Returns:
        Success message
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        raise Exception(error)

    active_doc = ctx.request_context.lifespan_context.get_active_document()

    # Add paragraph at the end of document
    doc_range = active_doc.Range()
    doc_range.Collapse(Direction=0)  # Collapse to end

    paragraph = insert_paragraph_after(active_doc, doc_range, text)
    return f"Successfully added paragraph: {text}"


@mcp_server.tool()
@handle_tool_errors
def get_document_outline(ctx: Context[ServerSession, AppContext] = Field(description="Context object")) -> str:
    """
    Get a simplified outline of the document - Quick Tool.
    
    This is a quick tool that provides a fast way to get the document outline.

    Returns:
        JSON string with document outline containing heading text and levels.
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document

    error = validate_active_document(ctx)
    if error:
        raise Exception(error)

    active_doc = ctx.request_context.lifespan_context.get_active_document()

    # Get document structure
    structure = get_document_structure(active_doc)

    # Simplify structure to just headings
    outline = [
        {"text": item["text"], "level": item["level"]}
        for item in structure
    ]

    return json.dumps(outline, ensure_ascii=False)
