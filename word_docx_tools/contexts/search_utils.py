import time
from typing import Dict, List, Optional, Any, Set, Tuple
from win32com.client import CDispatch
import logging
from ..common.exceptions import DocumentContextError, ErrorCode
from ..com_backend.com_utils import handle_com_error
from .context_control import DocumentContext
from .context_manager import get_context_manager


def search_contexts_by_type(context_type: str, include_children: bool = False) -> List[Dict[str, Any]]:
    """
    按类型搜索上下文对象
    
    Args:
        context_type: 上下文类型（如'paragraph', 'table', 'image', 'section'等）
        include_children: 是否包含子上下文的详细信息
    
    Returns:
        符合条件的上下文对象列表（字典形式）
    """
    start_time = time.time()
    results = []

    try:
        context_manager = get_context_manager()
        all_contexts = context_manager.get_all_contexts()
        
        # 过滤出指定类型的上下文
        for context in all_contexts:
            if context.metadata.get('type') == context_type:
                # 转换为字典格式
                context_dict = _context_to_dict(context, include_children)
                results.append(context_dict)
        
        logger.info(f"Search for context type '{context_type}' found {len(results)} results")
    except Exception as e:
        logger.error(f"Error searching contexts by type '{context_type}': {e}")
        raise DocumentContextError(
            error_code=ErrorCode.SEARCH_FAILED,
            message=f"Failed to search contexts by type: {str(e)}"
        )
    finally:
        # 记录性能指标
        context_manager = get_context_manager()
        context_manager._record_operation_time(
            'search_contexts_by_type', 
            time.time() - start_time, 
            success=True, 
            result_count=len(results)
        )
    
    return results


def get_context_hierarchy(context_id: str, depth: int = -1) -> Dict[str, Any]:
    """
    获取上下文的层次结构
    
    Args:
        context_id: 上下文ID
        depth: 层次深度，-1表示获取所有层次
    
    Returns:
        包含层次结构的字典
    """
    start_time = time.time()
    result = None

    try:
        context_manager = get_context_manager()
        context = context_manager.find_context_by_id(context_id)
        
        if not context:
            logger.warning(f"Context with ID {context_id} not found")
            raise DocumentContextError(
                error_code=ErrorCode.CONTEXT_NOT_FOUND,
                message=f"Context with ID {context_id} not found"
            )
        
        # 构建层次结构
        result = _build_context_hierarchy(context, context_manager, depth)
        
        logger.info(f"Successfully retrieved context hierarchy for ID {context_id}")
    except Exception as e:
        logger.error(f"Error getting context hierarchy for ID {context_id}: {e}")
        if not isinstance(e, DocumentContextError):
            raise DocumentContextError(
                error_code=ErrorCode.HIERARCHY_RETRIEVAL_FAILED,
                message=f"Failed to get context hierarchy: {str(e)}"
            )
        raise
    finally:
        # 记录性能指标
        context_manager = get_context_manager()
        context_manager._record_operation_time(
            'get_context_hierarchy', 
            time.time() - start_time, 
            success=result is not None
        )
    
    return result


def _context_to_dict(context: DocumentContext, include_children: bool = False) -> Dict[str, Any]:
    """
    将上下文对象转换为字典格式
    
    Args:
        context: 上下文对象
        include_children: 是否包含子上下文信息
    
    Returns:
        上下文信息的字典
    """
    try:
        context_dict = {
            'id': context.context_id,
            'title': context.title,
            'type': context.metadata.get('type'),
            'metadata': context.metadata.copy(),
            'has_children': len(context.child_contexts) > 0
        }
        
        # 添加位置信息（如果可用）
        if hasattr(context, 'range_obj') and context.range_obj:
            try:
                context_dict['start'] = context.range_obj.Start
                context_dict['end'] = context.range_obj.End
            except Exception:
                pass
        
        # 如果需要，添加子上下文信息
        if include_children and hasattr(context, 'child_contexts'):
            context_dict['children'] = [
                _context_to_dict(child) for child in context.child_contexts
            ]
        
        return context_dict
    except Exception as e:
        logger.error(f"Error converting context to dict: {e}")
        return {
            'id': context.context_id if hasattr(context, 'context_id') else 'unknown',
            'title': 'Error converting context',
            'type': 'error',
            'error': str(e)
        }


def _build_context_hierarchy(context: DocumentContext, 
                              context_manager: Any, 
                              depth: int = -1) -> Dict[str, Any]:
    """
    构建上下文的层次结构
    
    Args:
        context: 上下文对象
        context_manager: 上下文管理器
        depth: 层次深度，-1表示获取所有层次
    
    Returns:
        包含层次结构的字典
    """
    # 基本信息
    hierarchy = _context_to_dict(context)
    
    # 如果达到最大深度或没有子上下文，返回基本信息
    if depth == 0 or len(context.child_contexts) == 0:
        return hierarchy
    
    # 递归构建子上下文层次结构
    hierarchy['children'] = []
    for child in context.child_contexts:
        child_hierarchy = _build_context_hierarchy(
            child, context_manager, depth - 1 if depth > 0 else -1
        )
        hierarchy['children'].append(child_hierarchy)
    
    return hierarchy


def find_contexts_by_metadata(metadata_filters: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    根据元数据筛选上下文
    
    Args:
        metadata_filters: 元数据筛选条件，键值对形式
    
    Returns:
        符合条件的上下文对象列表（字典形式）
    """
    start_time = time.time()
    results = []

    try:
        context_manager = get_context_manager()
        all_contexts = context_manager.get_all_contexts()
        
        # 根据元数据筛选上下文
        for context in all_contexts:
            # 检查是否满足所有筛选条件
            match = True
            for key, value in metadata_filters.items():
                if key not in context.metadata or context.metadata[key] != value:
                    match = False
                    break
            
            if match:
                results.append(_context_to_dict(context))
        
        logger.info(f"Metadata search found {len(results)} matching contexts")
    except Exception as e:
        logger.error(f"Error searching contexts by metadata: {e}")
        raise DocumentContextError(
            error_code=ErrorCode.SEARCH_FAILED,
            message=f"Failed to search contexts by metadata: {str(e)}"
        )
    finally:
        # 记录性能指标
        context_manager = get_context_manager()
        context_manager._record_operation_time(
            'find_contexts_by_metadata', 
            time.time() - start_time, 
            success=True, 
            result_count=len(results)
        )
    
    return results


def find_contexts_by_range(start_pos: int, end_pos: int) -> List[Dict[str, Any]]:
    """
    查找指定范围内的上下文
    
    Args:
        start_pos: 起始位置
        end_pos: 结束位置
    
    Returns:
        范围内的上下文对象列表（字典形式）
    """
    start_time = time.time()
    results = []

    try:
        context_manager = get_context_manager()
        all_contexts = context_manager.get_all_contexts()
        
        # 查找与指定范围有交集的上下文
        for context in all_contexts:
            # 检查上下文是否有位置信息
            if not hasattr(context, 'range_obj') or not context.range_obj:
                continue
            
            try:
                context_start = context.range_obj.Start
                context_end = context.range_obj.End
                
                # 检查是否有交集
                if not (context_end < start_pos or context_start > end_pos):
                    results.append(_context_to_dict(context))
            except Exception:
                # 如果获取位置信息失败，跳过此上下文
                continue
        
        logger.info(f"Range search found {len(results)} matching contexts")
    except Exception as e:
        logger.error(f"Error searching contexts by range: {e}")
        raise DocumentContextError(
            error_code=ErrorCode.SEARCH_FAILED,
            message=f"Failed to search contexts by range: {str(e)}"
        )
    finally:
        # 记录性能指标
        context_manager = get_context_manager()
        context_manager._record_operation_time(
            'find_contexts_by_range', 
            time.time() - start_time, 
            success=True, 
            result_count=len(results)
        )
    
    return results


def search_contexts(keyword: str, search_fields: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    """
    全文搜索上下文
    
    Args:
        keyword: 搜索关键词
        search_fields: 要搜索的字段列表，默认为['title', 'metadata']
    
    Returns:
        包含关键词的上下文对象列表（字典形式）
    """
    start_time = time.time()
    results = []
    
    # 默认搜索字段
    if not search_fields:
        search_fields = ['title', 'metadata']

    try:
        context_manager = get_context_manager()
        all_contexts = context_manager.get_all_contexts()
        
        # 转换关键词为小写以进行不区分大小写的搜索
        keyword_lower = keyword.lower()
        
        # 搜索上下文
        for context in all_contexts:
            # 检查是否在任何指定字段中包含关键词
            match = False
            
            if 'title' in search_fields and keyword_lower in context.title.lower():
                match = True
            
            if 'metadata' in search_fields:
                # 检查元数据中的字符串值
                for key, value in context.metadata.items():
                    if isinstance(value, str) and keyword_lower in value.lower():
                        match = True
                        break
            
            if match:
                results.append(_context_to_dict(context))
        
        logger.info(f"Keyword search for '{keyword}' found {len(results)} matching contexts")
    except Exception as e:
        logger.error(f"Error performing keyword search: {e}")
        raise DocumentContextError(
            error_code=ErrorCode.SEARCH_FAILED,
            message=f"Failed to perform keyword search: {str(e)}"
        )
    finally:
        # 记录性能指标
        context_manager = get_context_manager()
        context_manager._record_operation_time(
            'search_contexts', 
            time.time() - start_time, 
            success=True, 
            result_count=len(results)
        )
    
    return results