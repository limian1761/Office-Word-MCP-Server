"""
Initialization module for Word Document MCP Server tools.

This module imports all available tools and makes them available for registration
with the MCP server.
"""

# Import all tools to register them with the MCP server
from .comment_tools import comment_tools
from .document_tools import document_tools
from .image_tools import image_tools
from .objects_tools import objects_tools
from .range_tools import range_tools
from .styles_tools import styles_tools
from .table_tools import table_tools
from .text_tools import text_tools
from .watch_and_execute import watch_and_execute
__all__ = [
    "comment_tools",
    "document_tools",
    "image_tools",
    "objects_tools",
    "range_tools",
    "styles_tools",
    "table_tools",
    "text_tools",
    "watch_and_execute"
]