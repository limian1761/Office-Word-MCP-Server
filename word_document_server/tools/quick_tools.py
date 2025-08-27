"""
Quick tools for common Word document operations.

This module provides simplified interfaces for frequently used operations,
following the Occam's Razor principle of simplicity.
"""
import json
from typing import Dict, Any, Optional
from mcp.server.fastmcp.server import Context
from word_document_server.core_utils import get_backend_for_tool, mcp_server
from word_document_server.errors import handle_tool_errors
from word_document_server.operations import (
    insert_paragraph_after,
    add_heading,
    add_table,
    get_document_structure
)


@mcp_server.tool()
@handle_tool_errors
def add_heading_quick(ctx: Context, text: str, level: int = 1) -> str:
    """
    Quickly add a heading to the document.
    
    Args:
        text: The heading text
        level: The heading level (1-9)
        
    Returns:
        Success message
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document
    error = validate_active_document(ctx)
    if error:
        raise Exception(error)
    
    backend = get_backend_for_tool(ctx, ctx.session.document_state['active_document_path'])
    
    # Add heading at the end of document
    doc_range = backend.document.Range()
    doc_range.Collapse(Direction=0)  # Collapse to end
    
    heading = add_heading(backend, doc_range, text, level)
    return f"Successfully added heading: {text}"


@mcp_server.tool()
@handle_tool_errors
def add_paragraph_quick(ctx: Context, text: str) -> str:
    """
    Quickly add a paragraph to the document.
    
    Args:
        text: The paragraph text
        
    Returns:
        Success message
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document
    error = validate_active_document(ctx)
    if error:
        raise Exception(error)
    
    backend = get_backend_for_tool(ctx, ctx.session.document_state['active_document_path'])
    
    # Add paragraph at the end of document
    doc_range = backend.document.Range()
    doc_range.Collapse(Direction=0)  # Collapse to end
    
    paragraph = insert_paragraph_after(backend, doc_range, text)
    return f"Successfully added paragraph: {text}"


@mcp_server.tool()
@handle_tool_errors
def get_document_outline(ctx: Context) -> str:
    """
    Get a simplified outline of the document.
    
    Returns:
        JSON string with document outline
    """
    # Validate active document
    from word_document_server.core_utils import validate_active_document
    error = validate_active_document(ctx)
    if error:
        raise Exception(error)
    
    backend = get_backend_for_tool(ctx, ctx.session.document_state['active_document_path'])
    
    # Get document structure
    structure = get_document_structure(backend)
    
    # Simplify structure to just headings
    outline = [{"text": item["text"], "level": item["level"]} 
               for item in structure 
               if item["type"] == "heading"]
    
    return json.dumps(outline, ensure_ascii=False)