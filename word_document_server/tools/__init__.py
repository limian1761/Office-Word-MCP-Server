"""Office Word MCP Server 工具模块

包含所有Word文档操作相关的函数。
"""

# 导出各个工具模块的函数
from .comment_tools import comment_tools
from .document_tools import document_tools
from .image_tools import image_tools
from .objects_tools import objects_tools
from .range_tools import range_tools
from .styles_tools import styles_tools
from .table_tools import table_tools
from .text_tools import text_tools

__all__ = [
    "comment_tools",
    "document_tools",
    "image_tools",
    "objects_tools",
    "styles_tools",
    "table_tools",
    "text_tools",
]
