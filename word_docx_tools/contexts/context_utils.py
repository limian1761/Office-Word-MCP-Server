from typing import Dict, Any, Optional, List, Union, Callable, Tuple
import time
import hashlib
import threading
from functools import wraps
import json

from ..utils.exceptions import WordDocumentError, ErrorCode
from ..utils.logger import log_info, log_error, log_debug
from .context_control import DocumentContext

# 缓存配置
DEFAULT_CACHE_TTL = 60  # 默认缓存过期时间(秒)
DEFAULT_CACHE_SIZE = 100  # 默认缓存大小

class ContextCache:
    """上下文缓存管理类"""
    
    def __init__(self, ttl: int = DEFAULT_CACHE_TTL, max_size: int = DEFAULT_CACHE_SIZE):
        self.ttl = ttl
        self.max_size = max_size
        self.cache = {}
        self.lock = threading.RLock()
    
    def set(self, key: str, value: Any, ttl: Optional[int] = None) -> None:
        """设置缓存项
        
        Args:
            key: 缓存键
            value: 缓存值
            ttl: 缓存过期时间(秒)，None表示使用默认值
        """
        with self.lock:
            # 如果缓存已满，移除最早的项
            if len(self.cache) >= self.max_size:
                # 获取最早添加的键
                oldest_key = min(self.cache.keys(), key=lambda k: self.cache[k]['added_at'])
                del self.cache[oldest_key]
                
            # 存储缓存项
            self.cache[key] = {
                'value': value,
                'added_at': time.time(),
                'ttl': ttl if ttl is not None else self.ttl
            }
    
    def get(self, key: str) -> Any:
        """获取缓存项
        
        Args:
            key: 缓存键
            
        Returns:
            缓存值，如果不存在或已过期则返回None
        """
        with self.lock:
            if key not in self.cache:
                return None
            
            item = self.cache[key]
            current_time = time.time()
            
            # 检查是否过期
            if current_time - item['added_at'] > item['ttl']:
                # 删除过期项
                del self.cache[key]
                return None
            
            return item['value']
    
    def delete(self, key: str) -> None:
        """删除缓存项
        
        Args:
            key: 缓存键
        """
        with self.lock:
            if key in self.cache:
                del self.cache[key]
    
    def clear(self) -> None:
        """清空缓存"""
        with self.lock:
            self.cache.clear()
    
    def size(self) -> int:
        """获取缓存大小
        
        Returns:
            缓存项数量
        """
        with self.lock:
            # 先清理过期项
            self._clean_expired()
            return len(self.cache)
    
    def _clean_expired(self) -> None:
        """清理过期的缓存项"""
        current_time = time.time()
        expired_keys = [
            key for key, item in self.cache.items()
            if current_time - item['added_at'] > item['ttl']
        ]
        
        for key in expired_keys:
            del self.cache[key]

# 创建全局上下文缓存实例
context_cache = ContextCache()

class PerformanceMonitor:
    """性能监控工具类"""
    
    def __init__(self):
        self.performance_records = {}
        self.lock = threading.RLock()
    
    def start(self, operation_name: str) -> float:
        """开始监控操作性能
        
        Args:
            operation_name: 操作名称
            
        Returns:
            开始时间戳
        """
        start_time = time.time()
        
        with self.lock:
            if operation_name not in self.performance_records:
                self.performance_records[operation_name] = []
        
        return start_time
    
    def end(self, operation_name: str, start_time: float, success: bool = True) -> Dict[str, Any]:
        """结束监控操作性能
        
        Args:
            operation_name: 操作名称
            start_time: 开始时间戳
            success: 操作是否成功
            
        Returns:
            性能记录
        """
        end_time = time.time()
        duration = end_time - start_time
        
        with self.lock:
            record = {
                'start_time': start_time,
                'end_time': end_time,
                'duration': duration,
                'success': success
            }
            
            if operation_name in self.performance_records:
                self.performance_records[operation_name].append(record)
            else:
                self.performance_records[operation_name] = [record]
        
        return record
    
    def get_stats(self, operation_name: Optional[str] = None) -> Dict[str, Any]:
        """获取性能统计信息
        
        Args:
            operation_name: 可选的操作名称，如果不提供则返回所有操作的统计
            
        Returns:
            性能统计信息
        """
        with self.lock:
            if operation_name:
                # 获取特定操作的统计
                if operation_name not in self.performance_records:
                    return {
                        'operation_name': operation_name,
                        'count': 0,
                        'success_count': 0,
                        'failure_count': 0,
                        'avg_duration': 0,
                        'min_duration': float('inf'),
                        'max_duration': 0
                    }
                
                records = self.performance_records[operation_name]
                return self._calculate_stats(operation_name, records)
            else:
                # 获取所有操作的统计
                all_stats = {}
                for op_name, records in self.performance_records.items():
                    all_stats[op_name] = self._calculate_stats(op_name, records)
                
                return all_stats
    
    def _calculate_stats(self, operation_name: str, records: List[Dict[str, Any]]) -> Dict[str, Any]:
        """计算指定操作的性能统计
        
        Args:
            operation_name: 操作名称
            records: 性能记录列表
            
        Returns:
            性能统计信息
        """
        if not records:
            return {
                'operation_name': operation_name,
                'count': 0,
                'success_count': 0,
                'failure_count': 0,
                'avg_duration': 0,
                'min_duration': float('inf'),
                'max_duration': 0
            }
        
        # 计算统计数据
        count = len(records)
        success_count = sum(1 for r in records if r['success'])
        failure_count = count - success_count
        
        durations = [r['duration'] for r in records]
        avg_duration = sum(durations) / count if count > 0 else 0
        min_duration = min(durations) if count > 0 else 0
        max_duration = max(durations) if count > 0 else 0
        
        return {
            'operation_name': operation_name,
            'count': count,
            'success_count': success_count,
            'failure_count': failure_count,
            'avg_duration': avg_duration,
            'min_duration': min_duration,
            'max_duration': max_duration
        }
    
    def clear(self) -> None:
        """清除所有性能记录"""
        with self.lock:
            self.performance_records.clear()

# 创建全局性能监控实例
performance_monitor = PerformanceMonitor()

@handle_com_error(ErrorCode.SERVER_ERROR, "cache context data")
def cache_context_data(
    cache_key: str,
    data: Any,
    ttl: Optional[int] = None
) -> Dict[str, Any]:
    """缓存上下文数据

    Args:
        cache_key: 缓存键
        data: 要缓存的数据
        ttl: 缓存过期时间(秒)

    Returns:
        包含缓存结果的字典
    """
    try:
        # 检查数据是否可序列化
        if data is not None:
            # 尝试序列化数据以确保它可以被缓存
            if isinstance(data, DocumentContext):
                # 对于DocumentContext对象，转换为字典
                serialized_data = data.to_dict()
            else:
                # 对于其他类型，尝试JSON序列化
                try:
                    json.dumps(data)
                    serialized_data = data
                except (TypeError, OverflowError):
                    raise WordDocumentError(
                        ErrorCode.INVALID_INPUT,
                        "Cannot cache data that is not JSON serializable"
                    )
        else:
            serialized_data = data
        
        # 设置缓存
        context_cache.set(cache_key, serialized_data, ttl)
        
        log_debug(f"Successfully cached context data with key: {cache_key}")
        
        return {
            "success": True,
            "message": "Context data cached successfully",
            "cache_key": cache_key
        }
    except Exception as e:
        log_error(f"Failed to cache context data: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "get cached context data")
def get_cached_context_data(cache_key: str) -> Dict[str, Any]:
    """获取缓存的上下文数据

    Args:
        cache_key: 缓存键

    Returns:
        包含缓存数据的字典
    """
    try:
        # 获取缓存
        cached_data = context_cache.get(cache_key)
        
        if cached_data is None:
            return {
                "success": False,
                "message": "No cached context data found",
                "cache_key": cache_key,
                "data": None
            }
        
        log_debug(f"Successfully retrieved cached context data with key: {cache_key}")
        
        return {
            "success": True,
            "message": "Cached context data retrieved successfully",
            "cache_key": cache_key,
            "data": cached_data
        }
    except Exception as e:
        log_error(f"Failed to get cached context data: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "invalidate cached context data")
def invalidate_cached_context_data(cache_key: str) -> Dict[str, Any]:
    """使缓存的上下文数据失效

    Args:
        cache_key: 缓存键

    Returns:
        包含操作结果的字典
    """
    try:
        # 删除缓存
        context_cache.delete(cache_key)
        
        log_debug(f"Successfully invalidated cached context data with key: {cache_key}")
        
        return {
            "success": True,
            "message": "Cached context data invalidated successfully",
            "cache_key": cache_key
        }
    except Exception as e:
        log_error(f"Failed to invalidate cached context data: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "clear context cache")
def clear_context_cache() -> Dict[str, Any]:
    """清空上下文缓存

    Returns:
        包含操作结果的字典
    """
    try:
        # 清空缓存
        context_cache.clear()
        
        log_info("Successfully cleared context cache")
        
        return {
            "success": True,
            "message": "Context cache cleared successfully"
        }
    except Exception as e:
        log_error(f"Failed to clear context cache: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "get context cache stats")
def get_context_cache_stats() -> Dict[str, Any]:
    """获取上下文缓存统计信息

    Returns:
        包含缓存统计的字典
    """
    try:
        # 获取缓存大小
        cache_size = context_cache.size()
        
        log_debug(f"Context cache statistics: size={cache_size}")
        
        return {
            "success": True,
            "message": "Context cache statistics retrieved successfully",
            "cache_size": cache_size
        }
    except Exception as e:
        log_error(f"Failed to get context cache statistics: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "get performance stats")
def get_performance_stats(operation_name: Optional[str] = None) -> Dict[str, Any]:
    """获取操作性能统计信息

    Args:
        operation_name: 可选的操作名称

    Returns:
        包含性能统计的字典
    """
    try:
        # 获取性能统计
        stats = performance_monitor.get_stats(operation_name)
        
        log_info(f"Performance statistics retrieved for operation: {operation_name or 'all'}")
        
        return {
            "success": True,
            "message": "Performance statistics retrieved successfully",
            "operation_name": operation_name,
            "stats": stats
        }
    except Exception as e:
        log_error(f"Failed to get performance statistics: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "clear performance stats")
def clear_performance_stats() -> Dict[str, Any]:
    """清空性能统计记录

    Returns:
        包含操作结果的字典
    """
    try:
        # 清空性能记录
        performance_monitor.clear()
        
        log_info("Successfully cleared performance statistics")
        
        return {
            "success": True,
            "message": "Performance statistics cleared successfully"
        }
    except Exception as e:
        log_error(f"Failed to clear performance statistics: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "generate context ID")
def generate_context_id(prefix: str = "ctx", seed: Optional[str] = None) -> str:
    """生成唯一的上下文ID

    Args:
        prefix: ID前缀
        seed: 可选的种子字符串，用于生成更具确定性的ID

    Returns:
        唯一的上下文ID字符串
    """
    try:
        # 生成基于时间的随机ID
        timestamp = str(int(time.time() * 1000))
        random_part = hashlib.md5(f"{timestamp}{seed or ''}".encode()).hexdigest()[:8]
        
        # 组合前缀和随机部分
        context_id = f"{prefix}_{timestamp}_{random_part}"
        
        log_debug(f"Generated context ID: {context_id}")
        
        return context_id
    except Exception as e:
        log_error(f"Failed to generate context ID: {str(e)}")
        raise

@handle_com_error(ErrorCode.SERVER_ERROR, "is valid context ID")
def is_valid_context_id(context_id: str) -> bool:
    """验证上下文ID是否有效

    Args:
        context_id: 要验证的上下文ID

    Returns:
        布尔值，表示ID是否有效
    """
    try:
        # 简单验证规则：非空、长度合理、格式正确
        if not context_id or not isinstance(context_id, str):
            return False
        
        # 检查长度
        if len(context_id) < 8 or len(context_id) > 100:
            return False
        
        # 检查格式（包含至少一个下划线）
        if '_' not in context_id:
            return False
        
        # 检查前缀部分
        parts = context_id.split('_', 1)
        if not parts or not parts[0].isalnum():
            return False
        
        return True
    except Exception as e:
        log_error(f"Failed to validate context ID: {str(e)}")
        return False

@handle_com_error(ErrorCode.SERVER_ERROR, "convert context to dictionary")
def convert_context_to_dictionary(
    context: DocumentContext,
    include_children: bool = False,
    depth: int = 2
) -> Dict[str, Any]:
    """将DocumentContext对象转换为字典

    Args:
        context: DocumentContext对象
        include_children: 是否包含子上下文
        depth: 递归深度

    Returns:
        转换后的字典
    """
    try:
        # 检查是否为DocumentContext对象
        if not isinstance(context, DocumentContext):
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                "Input must be a DocumentContext object"
            )
        
        # 调用DocumentContext的to_dict方法
        if include_children:
            result = context.to_dict_full(depth)
        else:
            result = context.to_dict()
        
        log_debug(f"Successfully converted context to dictionary: {context.context_id}")
        
        return result
    except Exception as e:
        log_error(f"Failed to convert context to dictionary: {str(e)}")
        raise