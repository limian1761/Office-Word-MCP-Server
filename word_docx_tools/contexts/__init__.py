"""
Contexts 包 - Word文档上下文管理核心功能

此包提供了文档上下文的创建、管理、查询、更新和事件处理等核心功能，
用于支持Word文档操作的上下文感知能力。
"""

from .context_control import DocumentContext
from .context_manager import ContextManager
from .search_utils import (
    search_contexts_by_type,
    get_context_hierarchy,
    find_contexts_by_metadata,
    get_context_by_position,
    get_all_contexts_of_type
)
from .document_change_handler import DocumentChangeHandler
from .metadata_processing import (
    MetadataProcessor,
    metadata_processor,
    create_document_metadata,
    create_section_metadata,
    create_paragraph_metadata,
    create_table_metadata,
    create_image_metadata
)

__all__ = [
    # 核心上下文类
    'DocumentContext',
    'ContextManager',
    
    # 搜索和查询工具
    'search_contexts_by_type',
    'get_context_hierarchy',
    'find_contexts_by_metadata',
    'get_context_by_position',
    'get_all_contexts_of_type',
    
    # 变更处理
    'DocumentChangeHandler',
    
    # 元数据处理
    'MetadataProcessor',
    'metadata_processor',
    'create_document_metadata',
    'create_section_metadata',
    'create_paragraph_metadata',
    'create_table_metadata',
    'create_image_metadata'
]

# 版本信息
__version__ = '1.0.0'