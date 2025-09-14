"""
Initialization module for Word Document MCP Server tools.

This module imports all available tools and makes them available for registration
with the MCP server.
"""

# Import all tools to register them with the MCP server
from .comment_tools import comment_tools
from .document_tools import document_tools
from .image_tools import image_tools
from .navigate_tools import navigate_tools
from .objects_tools import objects_tools
from .paragraph_tools import paragraph_tools
from .range_tools import range_tools
from .styles_tools import styles_tools
from .table_tools import table_tools
from .text_tools import text_tools
from .view_control_tools import view_control_tools

__all__ = [
    "comment_tools",
    "document_tools",
    "image_tools",
    "navigate_tools",
    "objects_tools",
    "paragraph_tools",
    "range_tools",
    "styles_tools",
    "table_tools",
    "text_tools",
    "view_control_tools",
]
