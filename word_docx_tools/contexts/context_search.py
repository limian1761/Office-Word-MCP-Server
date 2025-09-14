from typing import Dict, Any, Optional, List, Union
import time
from ..utils.exceptions import WordDocumentError, ErrorCode
from ..utils.logger import log_info, log_error, log_debug
from ..utils.decorators import handle_com_error, record_operation_time
from .context_control import DocumentContext

@handle_com_error(ErrorCode.SERVER_ERROR, "search contexts by type")
@record_operation_time
def search_contexts_by_type(
    document: object,
    object_type: str,
    filters: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """按类型搜索上下文

    Args:
        document: Word文档COM对象
        object_type: 对象类型（如'paragraph', 'table', 'image'等）
        filters: 可选的过滤条件

    Returns:
        包含搜索结果的字典
    """
    log_info(f"Searching contexts by type: {object_type}")
    
    start_time = time.time()
    
    try:
        # 获取根上下文
        root_context = DocumentContext.create_root_context(document)
        
        # 执行搜索
        results = []
        
        # 递归搜索函数
        def search_recursive(context: DocumentContext):
            # 检查当前上下文类型
            context_type = context.metadata.get("type")
            if context_type and context_type.lower() == object_type.lower():
                # 应用过滤器
                if filters:
                    if apply_filters(context, filters):
                        results.append(context.to_dict())
                else:
                    results.append(context.to_dict())
            
            # 递归搜索子上下文
            for child in context.child_contexts:
                search_recursive(child)
        
        # 开始搜索
        search_recursive(root_context)
        
        # 计算耗时
        elapsed_time = time.time() - start_time
        
        log_info(f"Search completed in {elapsed_time:.2f} seconds, found {len(results)} results")
        
        return {
            "success": True,
            "results": results,
            "count": len(results),
            "elapsed_time": elapsed_time
        }
    except Exception as e:
        log_error(f"Failed to search contexts: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "apply filters")
def apply_filters(context: DocumentContext, filters: Dict[str, Any]) -> bool:
    """应用过滤条件到上下文

    Args:
        context: DocumentContext对象
        filters: 过滤条件字典

    Returns:
        布尔值，表示上下文是否满足过滤条件
    """
    for key, value in filters.items():
        # 检查metadata中是否存在该属性
        if key not in context.metadata:
            return False
        
        # 比较值
        # 这里实现简单的相等比较，可以扩展为更复杂的比较逻辑
        if context.metadata[key] != value:
            return False
    
    return True

@handle_com_error(ErrorCode.SERVER_ERROR, "get context hierarchy")
@record_operation_time
def get_context_hierarchy(
    document: object,
    context_id: Optional[str] = None
) -> Dict[str, Any]:
    """获取上下文层次结构

    Args:
        document: Word文档COM对象
        context_id: 可选的上下文ID，如果提供则获取该上下文及其祖先的层次结构

    Returns:
        包含层次结构的字典
    """
    log_info(f"Getting context hierarchy for ID: {context_id or 'root'}")
    
    try:
        # 获取根上下文
        root_context = DocumentContext.create_root_context(document)
        
        if context_id:
            # 查找指定上下文
            target_context = root_context.find_child_context_by_id(context_id)
            
            if not target_context:
                raise WordDocumentError(ErrorCode.CONTEXT_ERROR, f"Context with ID {context_id} not found")
            
            # 构建从目标上下文到根的路径
            path = []
            current = target_context
            
            while current:
                path.append({
                    "context_id": current.context_id,
                    "title": current.title,
                    "type": current.metadata.get("type")
                })
                current = current.parent_context
            
            # 反转路径以从根开始
            path.reverse()
            
            return {
                "success": True,
                "hierarchy": path,
                "target_context": target_context.to_dict()
            }
        else:
            # 返回完整的层次结构
            hierarchy = {
                "context_id": root_context.context_id,
                "title": root_context.title,
                "type": root_context.metadata.get("type", "document"),
                "children": []
            }
            
            # 递归构建层次结构
            build_hierarchy_recursive(root_context, hierarchy["children"])
            
            return {
                "success": True,
                "hierarchy": hierarchy
            }
    except Exception as e:
        log_error(f"Failed to get context hierarchy: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "build hierarchy recursive")
def build_hierarchy_recursive(
    context: DocumentContext,
    parent_children: List[Dict[str, Any]]
) -> None:
    """递归构建上下文层次结构

    Args:
        context: 当前DocumentContext对象
        parent_children: 父级的children列表，用于添加当前上下文的层次结构
    """
    for child in context.child_contexts:
        # 创建当前子上下文的层次结构表示
        child_hierarchy = {
            "context_id": child.context_id,
            "title": child.title,
            "type": child.metadata.get("type"),
            "children": []
        }
        
        # 递归处理子上下文的子节点
        build_hierarchy_recursive(child, child_hierarchy["children"])
        
        # 添加到父级的children列表
        parent_children.append(child_hierarchy)

@handle_com_error(ErrorCode.SERVER_ERROR, "context to dict")
def context_to_dict(
    context: DocumentContext,
    include_children: bool = True,
    depth: int = -1
) -> Dict[str, Any]:
    """将上下文对象转换为字典

    Args:
        context: DocumentContext对象
        include_children: 是否包含子上下文
        depth: 包含子上下文的深度，-1表示无限制

    Returns:
        上下文的字典表示
    """
    result = {
        "context_id": context.context_id,
        "title": context.title,
        "metadata": context.metadata.copy(),
    }
    
    # 添加子上下文（如果需要）
    if include_children and (depth == -1 or depth > 0):
        child_depth = depth - 1 if depth > 0 else -1
        result["children"] = [
            context_to_dict(child, include_children=True, depth=child_depth)
            for child in context.child_contexts
        ]
    
    return result

@handle_com_error(ErrorCode.SERVER_ERROR, "find contexts by property")
@record_operation_time
def find_contexts_by_property(
    document: object,
    property_name: str,
    property_value: Any,
    object_type: Optional[str] = None
) -> Dict[str, Any]:
    """根据属性查找上下文

    Args:
        document: Word文档COM对象
        property_name: 属性名称
        property_value: 属性值
        object_type: 可选的对象类型过滤

    Returns:
        包含搜索结果的字典
    """
    log_info(f"Finding contexts by property: {property_name}={property_value}")
    
    start_time = time.time()
    
    try:
        # 获取根上下文
        root_context = DocumentContext.create_root_context(document)
        
        # 执行搜索
        results = []
        
        # 递归搜索函数
        def search_recursive(context: DocumentContext):
            # 检查对象类型（如果提供了）
            if object_type:
                context_type = context.metadata.get("type")
                if not context_type or context_type.lower() != object_type.lower():
                    return
            
            # 检查属性值
            if property_name in context.metadata and context.metadata[property_name] == property_value:
                results.append(context.to_dict())
            
            # 递归搜索子上下文
            for child in context.child_contexts:
                search_recursive(child)
        
        # 开始搜索
        search_recursive(root_context)
        
        # 计算耗时
        elapsed_time = time.time() - start_time
        
        log_info(f"Find operation completed in {elapsed_time:.2f} seconds, found {len(results)} results")
        
        return {
            "success": True,
            "results": results,
            "count": len(results),
            "elapsed_time": elapsed_time
        }
    except Exception as e:
        log_error(f"Failed to find contexts by property: {str(e)}")
        raise