"""
Tools module for the Word Document MCP Server.

This module contains all the modularized MCP tools organized by functionality.
"""

# Import all tool modules to register their tools with the MCP server
from . import (comment, document, image, quick_tools, table,
               text)
# Document level tools
from .document import (shutdown_word, close_document, open_document,
                       get_document_styles, get_elements)
# Image manipulation tools
from .image import add_caption, get_image_info, insert_object
# Table manipulation tools
from .table import get_text_from_cell, set_cell_value, create_table
# Text manipulation tools
from .text import (apply_formatting, batch_apply_format, create_bulleted_list,
                   find_text, get_text, insert_paragraph, replace_text, insert_text)
# Comment tools
from .comment import (add_comment, get_comments, delete_comment,
                     edit_comment, reply_to_comment, get_comment_thread)
# Quick tools
from .quick_tools import (add_heading_quick, add_paragraph_quick, get_document_outline)

__all__ = [
    # Modules
    "comment",
    "document",
    "image",
    "table",
    "text",
    "quick_tools",
    
    # Document tools
    "open_document",
    "close_document",
    "shutdown_word",
    "get_document_styles",
    "get_elements",
    
    # Text tools
    "get_text",
    "insert_paragraph",
    "insert_text",
    "apply_formatting",
    "batch_apply_format",
    "find_text",
    "replace_text",
    "create_bulleted_list",
    
    # Table tools
    "get_text_from_cell",
    "set_cell_value",
    "create_table",
    
    # Image tools
    "get_image_info",
    "insert_object",
    "add_caption",
    
    # Comment tools
    "add_comment",
    "get_comments",
    "delete_comment",
    "edit_comment",
    "reply_to_comment",
    "get_comment_thread",
    
    # Quick tools
    "add_heading_quick",
    "add_paragraph_quick",
    "get_document_outline"
]