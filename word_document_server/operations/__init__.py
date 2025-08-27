"""
Operations package for Word Document MCP Server.

This package contains all the specialized operation modules.
"""

from .document_operations import *
from .table_operations import *
from .text_formatting import *
from .image_operations import *
from .comment_operations import *

__all__ = [
    # document_operations exports
    'get_all_paragraphs',
    'get_paragraphs_in_range',
    'get_all_tables',
    'get_text_from_range',
    'get_runs_in_range',
    'get_tables_in_range',
    'get_cells_in_range',
    'set_header_text',
    'set_footer_text',
    'get_headings',
    'enable_track_revisions',
    'get_all_styles',
    'get_protection_status',
    'unprotect_document',
    'get_document_styles',
    'get_document_structure',
    'get_all_text',
    'find_text',
    'replace_text',
    'get_selection_info',
    # table_operations exports
    'add_table',
    # text_formatting exports
    'set_bold_for_range',
    'set_italic_for_range',
    'set_font_size_for_range',
    'set_font_color_for_range',
    'set_font_name_for_range',
    'insert_paragraph_after',
    'create_bulleted_list_relative_to',
    'set_alignment_for_range',
    # image_operations exports
    'get_all_inline_shapes',
    'insert_inline_picture',
    'add_picture_caption',
    # comment_operations exports
    'add_comment',
    'get_comments',
    'get_comments_by_range',
    'delete_comment',
    'delete_all_comments',
    'edit_comment',
    'reply_to_comment',
    'get_comment_thread'
]