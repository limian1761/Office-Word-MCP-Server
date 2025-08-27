"""
Tools module for the Word Document MCP Server.

This module contains all the modularized MCP tools organized by functionality.
"""

# Import all tool modules to register their tools with the MCP server
from . import (
    comment,
    document,
    image,
    table,
    text,
    quick_tools  # Add quick tools
)

# Text manipulation tools
from .text import (
    get_text,
    insert_paragraph,
    delete_element,
    apply_format,
    batch_apply_format,
    find_text,
    replace_text,
    create_bulleted_list
)

# Table manipulation tools
from .table import (
    get_text_from_cell,
    set_cell_value,
    insert_table
)

# Image manipulation tools
    from .image import (
    get_image_info,
    insert_object,
    add_caption
)

# Document level tools
from .document import (
    open_document,
    shutdown_word,
    get_document_styles,
    get_document_structure,
    get_all_text,
    get_elements
)


__all__ = [
    "comment",
    "document",
    "image",
    "table",
    "text",
    "quick_tools",  # Export quick tools
    
    # Text tools
    'get_text',
    'insert_paragraph',
    'delete_element',
    'apply_format',
    'batch_apply_format',
    'find_text',
    'replace_text',
    'create_bulleted_list',
    
    # Table tools
    'get_text_from_cell',
    'set_cell_value',
    'insert_table',
    
    # Image tools
    'get_image_info',
    'insert_object',
    'add_caption',
    
    # Document tools
    'open_document',
    'shutdown_word',
    'get_document_styles',
    'get_document_structure',
    'get_all_text',
    'get_elements'
]