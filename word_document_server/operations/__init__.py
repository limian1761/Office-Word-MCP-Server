"""
Operations package for Word Document MCP Server.

This package contains all the specialized operation modules.
"""

from .comment_operations import *
from .document_operations import *
from .text_formatting import *
from .element_operations import *

__all__ = [
    # Document-level operations
    "get_all_paragraphs",
    "get_paragraphs_in_range",
    "get_all_tables",
    "get_text_from_range",
    "get_runs_in_range",
    "set_header_text",
    "set_footer_text",
    "get_headings",
    "get_all_styles",
    "get_protection_status",
    "unprotect_document",
    "get_document_styles",
    "get_all_text",
    "find_text",
    "replace_text",
    "get_selection_info",
    "get_all_inline_shapes",
    "get_comments",
    
    # Element-level operations
    "add_heading",
    "add_table",
    "set_bold_for_range",
    "set_italic_for_range",
    "set_font_size_for_range",
    "set_font_color_for_range",
    "set_font_name_for_range",
    "insert_paragraph_after",
    "set_alignment_for_range",
    "add_comment",
    "delete_comment",
    "delete_all_comments",
    "edit_comment",
    "reply_to_comment",
    "get_comment_thread",
    "set_picture_element_color_type",
    "replace_element_text",
    "insert_text_before_element",
    "insert_text_after_element",
    "add_element_caption",
    "get_element_text",
    "set_paragraph_style"
]