from typing import Dict, Any, Optional, List, Union
import time
from ..utils.exceptions import WordDocumentError, ErrorCode
from ..utils.logger import log_info, log_error, log_debug
from ..utils.decorators import handle_com_error, record_operation_time
from .context_control import DocumentContext

class MetadataProcessor:
    """元数据处理器，负责处理文档上下文的元数据"""
    
    def __init__(self):
        # 支持的元数据类型和验证规则
        self.supported_metadata_types = {
            "string": str,
            "number": (int, float),
            "boolean": bool,
            "list": list,
            "dict": dict,
            "date": str  # ISO 8601格式的日期字符串
        }
        
        # 常用的元数据字段和默认值
        self.common_metadata_fields = {
            "created_at": lambda: time.time(),
            "updated_at": lambda: time.time(),
            "created_by": "system",
            "updated_by": "system",
            "type": "unknown"
        }
    
    def validate_metadata(self, metadata: Dict[str, Any]) -> bool:
        """验证元数据的有效性
        
        Args:
            metadata: 元数据字典
            
        Returns:
            布尔值，表示元数据是否有效
        """
        # 这里可以实现更复杂的元数据验证逻辑
        # 简化实现：检查元数据是否为字典
        return isinstance(metadata, dict)
    
    def sanitize_metadata(self, metadata: Dict[str, Any]) -> Dict[str, Any]:
        """清理元数据，移除无效或敏感的字段
        
        Args:
            metadata: 原始元数据字典
            
        Returns:
            清理后的元数据字典
        """
        # 检查元数据是否为字典
        if not isinstance(metadata, dict):
            return {}
        
        # 移除敏感字段或无效值
        sanitized = {}
        for key, value in metadata.items():
            # 跳过None值
            if value is None:
                continue
            
            # 跳过内部使用的字段（以_开头）
            if isinstance(key, str) and key.startswith('_'):
                continue
            
            # 添加到清理后的元数据
            sanitized[key] = value
        
        return sanitized
    
    def merge_metadata(
        self,
        base_metadata: Dict[str, Any],
        new_metadata: Dict[str, Any],
        overwrite: bool = True
    ) -> Dict[str, Any]:
        """合并元数据
        
        Args:
            base_metadata: 基础元数据
            new_metadata: 新元数据
            overwrite: 是否覆盖已有字段
            
        Returns:
            合并后的元数据
        """
        merged = base_metadata.copy()
        
        for key, value in new_metadata.items():
            # 检查是否应该覆盖已有字段
            if overwrite or key not in merged:
                merged[key] = value
        
        return merged
    
    def extract_metadata_from_object(self, obj: object) -> Dict[str, Any]:
        """从对象中提取元数据
        
        Args:
            obj: 要提取元数据的对象
            
        Returns:
            提取的元数据字典
        """
        metadata = {}
        
        # 尝试提取常见属性
        try:
            # 检查对象是否有Name属性
            if hasattr(obj, 'Name'):
                metadata['name'] = obj.Name
            
            # 检查对象是否有Title属性
            if hasattr(obj, 'Title'):
                metadata['title'] = obj.Title
            
            # 检查对象是否有ID属性
            if hasattr(obj, 'ID'):
                metadata['id'] = obj.ID
            
            # 检查对象是否有Range属性（Word对象常见）
            if hasattr(obj, 'Range') and obj.Range:
                range_metadata = {
                    'start': obj.Range.Start,
                    'end': obj.Range.End,
                    'text': obj.Range.Text[:100]  # 只保留前100个字符
                }
                metadata['range'] = range_metadata
            
        except Exception as e:
            log_error(f"Failed to extract metadata from object: {str(e)}")
        
        return metadata

# 创建全局元数据处理器实例
metadata_processor = MetadataProcessor()

@handle_com_error(ErrorCode.SERVER_ERROR, "get metadata")
@record_operation_time
def get_metadata(
    document: object,
    context_id: str,
    keys: Optional[List[str]] = None
) -> Dict[str, Any]:
    """获取上下文的元数据

    Args:
        document: Word文档COM对象
        context_id: 上下文ID
        keys: 可选的元数据键列表，如果提供则只返回指定键的值

    Returns:
        包含元数据的字典
    """
    log_info(f"Getting metadata for context: {context_id}")
    
    try:
        # 获取根上下文
        root_context = DocumentContext.create_root_context(document)
        
        # 查找目标上下文
        target_context = root_context.find_child_context_by_id(context_id)
        
        if not target_context:
            raise WordDocumentError(ErrorCode.CONTEXT_ERROR, f"Context with ID {context_id} not found")
        
        # 获取元数据
        metadata = target_context.metadata.copy()
        
        # 如果指定了键列表，只返回指定的键
        if keys:
            filtered_metadata = {}
            for key in keys:
                if key in metadata:
                    filtered_metadata[key] = metadata[key]
            metadata = filtered_metadata
        
        log_info(f"Successfully retrieved metadata for context: {context_id}")
        
        return {
            "success": True,
            "metadata": metadata,
            "context_id": context_id
        }
    except Exception as e:
        log_error(f"Failed to get metadata: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "update metadata")
@record_operation_time
def update_metadata(
    document: object,
    context_id: str,
    new_metadata: Dict[str, Any],
    overwrite: bool = True
) -> Dict[str, Any]:
    """更新上下文的元数据

    Args:
        document: Word文档COM对象
        context_id: 上下文ID
        new_metadata: 新的元数据
        overwrite: 是否覆盖已有字段

    Returns:
        包含更新结果的字典
    """
    log_info(f"Updating metadata for context: {context_id}")
    
    try:
        # 获取根上下文
        root_context = DocumentContext.create_root_context(document)
        
        # 查找目标上下文
        target_context = root_context.find_child_context_by_id(context_id)
        
        if not target_context:
            raise WordDocumentError(ErrorCode.CONTEXT_ERROR, f"Context with ID {context_id} not found")
        
        # 清理新元数据
        sanitized_metadata = metadata_processor.sanitize_metadata(new_metadata)
        
        # 验证元数据
        if not metadata_processor.validate_metadata(sanitized_metadata):
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                "Invalid metadata format. Metadata must be a dictionary."
            )
        
        # 合并元数据
        updated_metadata = metadata_processor.merge_metadata(
            target_context.metadata,
            sanitized_metadata,
            overwrite
        )
        
        # 更新上下文的元数据
        target_context.metadata = updated_metadata
        
        # 更新更新时间
        target_context.metadata["updated_at"] = time.time()
        
        log_info(f"Successfully updated metadata for context: {context_id}")
        
        return {
            "success": True,
            "message": "Metadata updated successfully",
            "context_id": context_id,
            "updated_fields": list(sanitized_metadata.keys())
        }
    except Exception as e:
        log_error(f"Failed to update metadata: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "remove metadata field")
@record_operation_time
def remove_metadata_field(
    document: object,
    context_id: str,
    field_name: str
) -> Dict[str, Any]:
    """移除上下文元数据中的指定字段

    Args:
        document: Word文档COM对象
        context_id: 上下文ID
        field_name: 要移除的字段名

    Returns:
        包含移除结果的字典
    """
    log_info(f"Removing metadata field '{field_name}' for context: {context_id}")
    
    try:
        # 获取根上下文
        root_context = DocumentContext.create_root_context(document)
        
        # 查找目标上下文
        target_context = root_context.find_child_context_by_id(context_id)
        
        if not target_context:
            raise WordDocumentError(ErrorCode.CONTEXT_ERROR, f"Context with ID {context_id} not found")
        
        # 检查字段是否存在
        if field_name not in target_context.metadata:
            return {
                "success": True,
                "message": f"Field '{field_name}' does not exist in metadata",
                "context_id": context_id
            }
        
        # 移除字段
        del target_context.metadata[field_name]
        
        # 更新更新时间
        target_context.metadata["updated_at"] = time.time()
        
        log_info(f"Successfully removed metadata field '{field_name}' for context: {context_id}")
        
        return {
            "success": True,
            "message": f"Metadata field '{field_name}' removed successfully",
            "context_id": context_id
        }
    except Exception as e:
        log_error(f"Failed to remove metadata field: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "batch update metadata")
@record_operation_time
def batch_update_metadata(
    document: object,
    metadata_updates: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """批量更新多个上下文的元数据

    Args:
        document: Word文档COM对象
        metadata_updates: 元数据更新列表，每个元素包含context_id和metadata

    Returns:
        包含批量更新结果的字典
    """
    log_info(f"Starting batch metadata update with {len(metadata_updates)} operations")
    
    start_time = time.time()
    
    # 操作结果记录
    results = {
        "success": [],
        "failed": []
    }
    
    try:
        # 获取根上下文
        root_context = DocumentContext.create_root_context(document)
        
        for update in metadata_updates:
            context_id = update.get("context_id")
            new_metadata = update.get("metadata", {})
            overwrite = update.get("overwrite", True)
            
            try:
                # 查找目标上下文
                target_context = root_context.find_child_context_by_id(context_id)
                
                if not target_context:
                    results["failed"].append({
                        "context_id": context_id,
                        "error": "Context not found"
                    })
                    continue
                
                # 清理新元数据
                sanitized_metadata = metadata_processor.sanitize_metadata(new_metadata)
                
                # 验证元数据
                if not metadata_processor.validate_metadata(sanitized_metadata):
                    results["failed"].append({
                        "context_id": context_id,
                        "error": "Invalid metadata format"
                    })
                    continue
                
                # 合并元数据
                updated_metadata = metadata_processor.merge_metadata(
                    target_context.metadata,
                    sanitized_metadata,
                    overwrite
                )
                
                # 更新上下文的元数据
                target_context.metadata = updated_metadata
                
                # 更新更新时间
                target_context.metadata["updated_at"] = time.time()
                
                # 记录成功
                results["success"].append({
                    "context_id": context_id,
                    "updated_fields": list(sanitized_metadata.keys())
                })
            except Exception as e:
                # 记录失败
                results["failed"].append({
                    "context_id": context_id,
                    "error": str(e)
                })
        
        # 计算耗时
        elapsed_time = time.time() - start_time
        
        log_info(f"Batch metadata update completed in {elapsed_time:.2f} seconds")
        
        return {
            "success": True,
            "message": f"Batch metadata update completed with {len(results['success'])} successes and {len(results['failed'])} failures",
            "results": results,
            "elapsed_time": elapsed_time
        }
    except Exception as e:
        log_error(f"Failed to complete batch metadata update: {str(e)}")
        raise