"""
Tools module for the Word Document MCP Server.

This module contains all the modularized MCP tools organized by functionality.
"""

# Import all tool modules to register their tools with the MCP server
from . import (comment, document, image, quick_tools, table,  # Add quick tools
               text)
# Document level tools
from .document import (get_all_text, shutdown_word)
from .document import close_document, open_document
from .quick_tools import get_document_outline
# Image manipulation tools
from .image import add_caption, get_image_info, insert_object
# Table manipulation tools
from .table import add_table, get_text_from_cell, set_cell_value
# Text manipulation tools
from .text import (apply_format, batch_apply_format, create_bulleted_list,
                   delete_element, find_text, get_text, insert_paragraph,
                   replace_text)

__all__ = [
    "comment",
    "document",
    "image",
    "table",
    "text",
    "quick_tools",  # Export quick tools
    # Text tools
    "get_text",
    "insert_paragraph",
    "delete_element",
    "apply_format",
    "batch_apply_format",
    "find_text",
    "replace_text",
    "create_bulleted_list",
    # Table tools
    "get_text_from_cell",
    "set_cell_value",
    "insert_table",
    # Image tools
    "get_image_info",
    "insert_object",
    "add_caption",
    # Document tools
    "open_document",
    "shutdown_word",
    "get_all_text",
    "get_document_outline",
    "open_document",
]
