"""Document context models for Word Document MCP Server.

This module contains the DocumentContext class which represents
and manages context information within Word documents.
"""

import logging
import uuid
import time
from typing import Dict, Any, Optional, List, Set, Tuple

import win32com.client

from ..mcp_service.errors import ErrorCode, WordDocumentError
from ..mcp_service.core_utils import log_error, log_info


class DocumentContext:
    """增强版文档上下文类，用于表示和管理Word文档中的上下文信息
    
    属性:
        context_id: 上下文唯一标识符
        title: 上下文标题
        range: 上下文对应的文档范围对象
        object_list: 上下文包含的对象列表
        parent_context: 父上下文对象
        child_contexts: 子上下文对象列表
        metadata: 上下文元数据字典
        last_updated: 最后更新时间戳
        _cached_dict: 缓存的字典表示，用于性能优化
        _cache_valid: 缓存有效性标志
    """
    
    def __init__(self, title: str = "", range_obj: Optional[Any] = None):
        """初始化文档上下文对象
        
        参数:
            title: 上下文标题
            range_obj: Word文档Range对象
        """
        self.context_id = str(uuid.uuid4())  # 生成唯一上下文ID
        self.title = title
        self.range = range_obj  # Word文档Range对象
        self.object_list: List[Dict[str, Any]] = []  # 对象列表，存储对象信息字典
        self.parent_context: Optional['DocumentContext'] = None  # 父上下文
        self.child_contexts: List['DocumentContext'] = []  # 子上下文列表
        self.metadata: Dict[str, Any] = {}  # 上下文元数据
        self.last_updated = time.time()  # 最后更新时间戳
        self._cached_dict: Optional[Dict[str, Any]] = None  # 缓存的字典表示
        self._cache_valid = False  # 缓存有效性标志
    
    def _invalidate_cache(self) -> None:
        """使缓存失效，并更新最后更新时间"""
        self._cache_valid = False
        self.last_updated = time.time()
    
    def _update_metadata(self, key: str, value: Any) -> None:
        """更新元数据
        
        参数:
            key: 元数据键
            value: 元数据值
        """
        self.metadata[key] = value
        self._invalidate_cache()
    
    def update_multiple_metadata(self, metadata_dict: Dict[str, Any]) -> None:
        """批量更新元数据
        
        参数:
            metadata_dict: 元数据字典
        """
        self.metadata.update(metadata_dict)
        self._invalidate_cache()
    
    def add_object(self, object_info: Dict[str, Any]) -> None:
        """添加对象到上下文
        
        参数:
            object_info: 对象信息字典
        """
        # 确保对象有唯一标识符
        if 'id' not in object_info:
            object_info['id'] = str(uuid.uuid4())
        if 'type' not in object_info:
            object_info['type'] = 'unknown'
            
        self.object_list.append(object_info)
        self._invalidate_cache()
    
    def add_child_context(self, child_context: 'DocumentContext') -> None:
        """添加子上下文
        
        参数:
            child_context: 子上下文对象
        """
        if child_context not in self.child_contexts:
            self.child_contexts.append(child_context)
            child_context.parent_context = self
            # 更新子上下文的元数据，标记其在树中的位置
            child_context._update_metadata('parent_id', self.context_id)
            self._invalidate_cache()
    
    def remove_child_context(self, child_context: 'DocumentContext') -> None:
        """移除子上下文
        
        参数:
            child_context: 要移除的子上下文对象
        """
        if child_context in self.child_contexts:
            self.child_contexts.remove(child_context)
            child_context.parent_context = None
            # 更新子上下文的元数据
            child_context.metadata.pop('parent_id', None)
            self._invalidate_cache()
    
    def find_child_context_by_id(self, context_id: str) -> Optional['DocumentContext']:
        """通过ID查找子上下文
        
        参数:
            context_id: 要查找的上下文ID
        
        返回:
            找到的子上下文对象，未找到则返回None
        """
        for child in self.child_contexts:
            if child.context_id == context_id:
                return child
        return None
    
    def find_object_by_id(self, object_id: str) -> Optional[Dict[str, Any]]:
        """通过ID查找对象
        
        参数:
            object_id: 要查找的对象ID
        
        返回:
            找到的对象信息字典，未找到则返回None
        """
        for obj in self.object_list:
            if obj.get('id') == object_id:
                return obj
        return None
    
    def update_object(self, object_id: str, updated_info: Dict[str, Any]) -> bool:
        """更新对象信息
        
        参数:
            object_id: 要更新的对象ID
            updated_info: 要更新的对象信息
        
        返回:
            更新是否成功
        """
        for i, obj in enumerate(self.object_list):
            if obj.get('id') == object_id:
                self.object_list[i].update(updated_info)
                self._invalidate_cache()
                return True
        return False
    
    def remove_object(self, object_id: str) -> bool:
        """移除对象
        
        参数:
            object_id: 要移除的对象ID
        
        返回:
            移除是否成功
        """
        for i, obj in enumerate(self.object_list):
            if obj.get('id') == object_id:
                del self.object_list[i]
                self._invalidate_cache()
                return True
        return False
    
    def batch_add_objects(self, objects_info: List[Dict[str, Any]]) -> None:
        """批量添加对象
        
        参数:
            objects_info: 对象信息字典列表
        """
        # 为没有ID的对象生成唯一ID
        for obj in objects_info:
            if 'id' not in obj:
                obj['id'] = str(uuid.uuid4())
            if 'type' not in obj:
                obj['type'] = 'unknown'
                
        self.object_list.extend(objects_info)
        self._invalidate_cache()
    
    def to_dict(self) -> Dict[str, Any]:
        """将上下文对象转换为字典格式，使用缓存优化性能
        
        返回:
            包含上下文信息的字典
        """
        # 如果缓存有效，直接返回缓存的字典
        if self._cache_valid and self._cached_dict:
            return self._cached_dict
        
        result = {
            "context_id": self.context_id,
            "title": self.title,
            "has_range": self.range is not None,
            "object_count": len(self.object_list),
            "child_count": len(self.child_contexts),
            "has_parent": self.parent_context is not None,
            "metadata": self.metadata.copy(),
            "last_updated": self.last_updated
        }
        
        # 如果有Range对象，可以添加一些基本信息
        if self.range:
            try:
                result["range_info"] = {
                    "start": self.range.Start,
                    "end": self.range.End,
                    "text_preview": self.range.Text[:50] + ("..." if len(self.range.Text) > 50 else "")
                }
            except Exception:
                result["range_info"] = {"error": "Failed to get range details"}
        
        # 添加对象列表的简要信息
        result["objects_preview"] = [
            {"type": obj.get("type", "unknown"), "id": obj.get("id", "unknown")}
            for obj in self.object_list[:5]  # 只包含前5个对象的预览
        ]
        
        # 更新缓存
        self._cached_dict = result
        self._cache_valid = True
        
        return result
    
    def to_dict_full(self, include_children: bool = False) -> Dict[str, Any]:
        """将上下文对象转换为包含完整信息的字典格式
        
        参数:
            include_children: 是否包含子上下文的完整信息
        
        返回:
            包含完整上下文信息的字典
        """
        result = self.to_dict()
        
        # 添加完整对象列表
        result["objects"] = self.object_list.copy()
        
        # 如果需要，添加子上下文的完整信息
        if include_children:
            result["children"] = [child.to_dict_full(include_children) for child in self.child_contexts]
        
        return result
    
    def update_document_context_for_style(self, style_updates: List[Tuple[str, Dict[str, Any]]]) -> Dict[str, Any]:
        """批量更新文档样式相关的上下文信息
        
        参数:
            style_updates: 包含(对象ID, 更新信息)的元组列表
        
        返回:
            包含操作结果的字典
        """
        results = {
            "success_count": 0,
            "failure_count": 0,
            "errors": []
        }
        
        # 批量更新样式信息
        for obj_id, update_info in style_updates:
            try:
                if self.update_object(obj_id, update_info):
                    results["success_count"] += 1
                else:
                    results["failure_count"] += 1
                    results["errors"].append({"object_id": obj_id, "error": "Object not found"})
            except Exception as e:
                results["failure_count"] += 1
                results["errors"].append({"object_id": obj_id, "error": str(e)})
        
        # 只在批量操作完成后使缓存失效一次
        if results["success_count"] > 0:
            self._invalidate_cache()
        
        return results
    
    @classmethod
    def from_document_selection(cls, document: win32com.client.CDispatch, title: str = "Selected Context") -> 'DocumentContext':
        """从文档当前选择创建上下文对象
        
        参数:
            document: Word文档COM对象
            title: 上下文标题
        
        返回:
            创建的DocumentContext对象
        
        异常:
            WordDocumentError: 当获取选择失败时抛出
        """
        if not document:
            raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
        
        try:
            selection = document.Application.Selection
            context = cls(title=title, range_obj=selection.Range)
            
            # 设置元数据
            context.update_multiple_metadata({
                "source": "selection",
                "creation_time": time.time(),
                "document_name": document.Name
            })
            
            # 批量获取选择范围内的对象信息
            objects_to_add = []
            
            # 收集表格信息
            if selection.Tables.Count > 0:
                for i, table in enumerate(selection.Tables):
                    objects_to_add.append({
                        "type": "table",
                        "id": str(table.Range.Start),
                        "index": i,
                        "range_start": table.Range.Start,
                        "range_end": table.Range.End
                    })
            
            # 收集图像信息
            if selection.InlineShapes.Count > 0:
                for i, shape in enumerate(selection.InlineShapes):
                    objects_to_add.append({
                        "type": "image",
                        "id": str(shape.Range.Start),
                        "index": i,
                        "range_start": shape.Range.Start,
                        "range_end": shape.Range.End
                    })
            
            # 收集注释信息
            if selection.Comments.Count > 0:
                for i, comment in enumerate(selection.Comments):
                    objects_to_add.append({
                        "type": "comment",
                        "id": str(comment.Index),
                        "index": i,
                        "author": comment.Author,
                        "text": comment.Range.Text[:100] + ("..." if len(comment.Range.Text) > 100 else "")
                    })
            
            # 批量添加对象，提高性能
            context.batch_add_objects(objects_to_add)
            
            return context
        except Exception as e:
            raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to create context from selection: {str(e)}")
    
    @classmethod
    def create_root_context(cls, document: win32com.client.CDispatch) -> 'DocumentContext':
        """创建文档的根上下文
        
        参数:
            document: Word文档COM对象
        
        返回:
            创建的根上下文对象
        
        异常:
            WordDocumentError: 当创建失败时抛出
        """
        if not document:
            raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
        
        try:
            # 创建根上下文，范围为整个文档
            root_context = cls(title="Root Document Context", range_obj=document.Content)
            
            # 设置根上下文元数据
            root_context.update_multiple_metadata({
                "source": "document",
                "document_name": document.Name,
                "document_path": document.FullName if hasattr(document, 'FullName') else "Unsaved",
                "page_count": document.ComputeStatistics(2),  # wdStatisticPages
                "word_count": document.ComputeStatistics(1)   # wdStatisticWords
            })
            
            return root_context
        except Exception as e:
            raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to create root context: {str(e)}")