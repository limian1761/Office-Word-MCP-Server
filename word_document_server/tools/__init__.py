"""
MCP tool implementations for the Word Document Server.

This package contains the MCP tool implementations that expose functionality
to clients through the Model Context Protocol.
"""

# Document tools
from word_document_server.tools.document_tools import (
    create_document, get_document_info, 
    get_document_outline, list_opened_documents, 
    copy_document, merge_documents
)

# Content tools
from word_document_server.tools.content_tools import (
    add_heading, add_paragraph, add_table, add_picture,
    add_page_break, delete_paragraph,
    search_and_replace, select_paragraphs, add_picture_caption,
    add_paragraph_numbering
)

# Format tools
from word_document_server.tools.format_tools import (
    format_text, create_custom_style, format_table
)

# Protection tools
from word_document_server.tools.protection_tools import (
    protect_document, unprotect_document, verify_document
)

# Footnote tools
from word_document_server.tools.footnote_tools import (
    add_footnote_to_document, add_endnote_to_document,
    convert_footnotes_to_endnotes, customize_footnote_style
)

# Comment tools
from word_document_server.tools.comment_tools import (
    get_all_comments, get_comments_by_author, get_comments_for_paragraph
)

# COM Document tools (Windows only)
try:
    from word_document_server.tools.com_document_tools import (
        get_document_properties_com_tool,
        get_all_paragraphs_com_tool,
        get_paragraphs_by_range_com_tool,
        get_paragraphs_by_page_com_tool,
        analyze_paragraph_distribution_com_tool,
        check_com_availability_tool
    )
except ImportError:
    # COM tools not available on non-Windows platforms
    pass
