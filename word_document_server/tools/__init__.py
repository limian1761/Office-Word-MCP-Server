"""
Tools module for the Word Document MCP Server.

This module contains all the modularized MCP tools organized by functionality.
"""

# Import and re-export tools from different modules

# Text manipulation tools
from .text import (
    insert_paragraph,
    delete_element,
    get_text,
    replace_text,
    find_text,
    apply_format,
    apply_paragraph_style,
    create_bulleted_list,
    batch_apply_format
)

# Table manipulation tools
from .table import (
    get_text_from_cell,
    set_cell_value,
    create_table
)

# Image manipulation tools
from .image import (
    get_image_info,
    insert_inline_picture,
    set_image_size,
    set_image_color_type,
    delete_image,
    add_picture_caption
)

# Document level tools
from .document import (
    open_document,
    shutdown_word,
    save_document,
    close_document,
    get_document_styles,
    get_document_structure,
    set_header_text,
    set_footer_text,
    enable_track_revisions,
    disable_track_revisions,
    accept_all_changes
)

# Comments tools
from .comment import (
    add_comment,
    get_comments,
    delete_comment,
    delete_all_comments,
    edit_comment,
    reply_to_comment,
    get_comment_thread
)

__all__ = [
    # Text tools
    'insert_paragraph',
    'delete_element',
    'get_text',
    'replace_text',
    'find_text',
    'apply_format',
    'apply_paragraph_style',
    'create_bulleted_list',
    'batch_apply_format',
    
    # Table tools
    'get_text_from_cell',
    'set_cell_value',
    'create_table',
    
    # Image tools
    'get_image_info',
    'insert_inline_picture',
    'set_image_size',
    'set_image_color_type',
    'delete_image',
    'add_picture_caption',
    
    # Document tools
    'open_document',
    'shutdown_word',
    'save_document',
    'close_document',
    'get_document_styles',
    'get_document_structure',
    'set_header_text',
    'set_footer_text',
    'enable_track_revisions',
    'disable_track_revisions',
    'accept_all_changes',
    
    # Comments tools
    'add_comment',
    'get_comments',
    'delete_comment',
    'delete_all_comments',
    'edit_comment',
    'reply_to_comment',
    'get_comment_thread'
]

# Version information
__version__ = '1.0.0'