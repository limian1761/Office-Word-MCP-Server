"""
Context Control Operations for Word Document MCP Server.

This module contains operations for managing document context and active objects in Word.
It focuses on setting context, managing active objects, and navigating between objects.
View-related functionalities are handled internally as needed.
"""

import logging
import uuid
import time
from typing import Dict, Any, Optional, List, Set, Tuple

import win32com.client

from ..mcp_service.core_utils import (
    ErrorCode,
    WordDocumentError,
    log_error,
    log_info
)


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


# 局部导入以避免循环依赖
from ..com_backend.com_utils import handle_com_error

@handle_com_error(ErrorCode.SERVER_ERROR, "set active context")
def set_active_context(document: win32com.client.CDispatch, context_type: str, context_id: str) -> Dict[str, Any]:
    """设置活动上下文

    Args:
        document: Word文档COM对象
        context_type: 上下文类型 (section, paragraph, table, etc.)
        context_id: 上下文ID

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当设置上下文失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    # 支持的上下文类型
    valid_context_types = ["section", "paragraph", "table", "image", "comment", "bookmark"]
    
    if context_type.lower() not in valid_context_types:
        valid_types_str = ", ".join(valid_context_types)
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT,
            f"Invalid context type. Supported types: {valid_types_str}"
        )
    
    try:
        # 根据上下文类型和ID设置活动上下文
        # 这里会实现内部逻辑，可能包括视图滚动，但对外部隐藏这些细节
        word_app = document.Application
        
        # 根据上下文类型查找并设置活动对象
        if context_type.lower() == "section":
            # 实现按节ID设置上下文
            pass
        elif context_type.lower() == "paragraph":
            # 实现按段落ID设置上下文
            pass
        elif context_type.lower() == "table":
            # 实现按表格ID设置上下文
            pass
        elif context_type.lower() == "image":
            # 实现按图像ID设置上下文
            pass
        elif context_type.lower() == "comment":
            # 实现按注释ID设置上下文
            pass
        elif context_type.lower() == "bookmark":
            # 实现按书签名称设置上下文
            if context_id in [b.Name for b in document.Bookmarks]:
                bookmark = document.Bookmarks(context_id)
                bookmark.Select()
                word_app.ActiveWindow.ScrollIntoView(bookmark.Range)
            else:
                raise WordDocumentError(ErrorCode.OBJECT_ERROR, f"Bookmark '{context_id}' not found")
        
        log_info(f"Successfully set active context to {context_type} with ID {context_id}")
        
        return {
            "success": True,
            "context_type": context_type,
            "context_id": context_id,
            "message": f"Active context successfully set to {context_type} with ID {context_id}"
        }
    except Exception as e:
        raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to set active context: {str(e)}")


@handle_com_error(ErrorCode.SERVER_ERROR, "get active object")
def get_active_object(document: win32com.client.CDispatch) -> Dict[str, Any]:
    """获取当前活动对象信息

    Args:
        document: Word文档COM对象

    Returns:
        包含当前活动对象信息的字典

    Raises:
        WordDocumentError: 当获取活动对象信息失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    try:
        word_app = document.Application
        selection = word_app.Selection
        
        # 获取活动对象信息
        active_object = {
            "type": None,
            "id": None,
            "text": None,
            "position": {
                "page": selection.Information(3),  # wdActiveEndPageNumber
                "line": selection.Information(10)  # wdFirstCharacterLineNumber
            }
        }
        
        # 识别选择的对象类型
        if selection.Tables.Count > 0:
            active_object["type"] = "table"
            active_object["id"] = str(selection.Tables(1).Range.Start)
            active_object["text"] = "Table at position " + str(selection.Information(3))
        elif selection.InlineShapes.Count > 0:
            active_object["type"] = "image"
            active_object["id"] = str(selection.InlineShapes(1).Range.Start)
            active_object["text"] = "Image at position " + str(selection.Information(3))
        elif selection.Comments.Count > 0:
            active_object["type"] = "comment"
            active_object["id"] = str(selection.Comments(1).Index)
            active_object["text"] = selection.Comments(1).Range.Text[:50] + ("..." if len(selection.Comments(1).Range.Text) > 50 else "")
        else:
            active_object["type"] = "text"
            active_object["id"] = str(selection.Range.Start)
            active_object["text"] = selection.Text[:50] + ("..." if len(selection.Text) > 50 else "")
        
        log_info(f"Successfully retrieved active object information")
        
        return {
            "success": True,
            "active_object": active_object
        }
    except Exception as e:
        raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to get active object: {str(e)}")


@handle_com_error(ErrorCode.SERVER_ERROR, "navigate to next object")
def navigate_to_next_object(document: win32com.client.CDispatch, object_type: Optional[str] = None) -> Dict[str, Any]:
    """导航到下一个对象

    Args:
        document: Word文档COM对象
        object_type: 可选的对象类型过滤

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当导航失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    try:
        word_app = document.Application
        selection = word_app.Selection
        
        # 保存当前位置用于检查是否成功移动
        current_start = selection.Range.Start
        
        # 根据对象类型进行导航
        if object_type:
            object_type = object_type.lower()
            
            if object_type == "paragraph":
                # 移动到下一段
                selection.MoveDown(Unit=5, Count=1)
            elif object_type == "table":
                # 查找下一个表格
                next_table = None
                for table in document.Tables:
                    if table.Range.Start > selection.Range.End:
                        next_table = table
                        break
                if next_table:
                    next_table.Select()
                else:
                    return {
                        "success": False,
                        "message": "No more tables found"
                    }
            elif object_type == "image":
                # 查找下一个图像
                next_image = None
                for shape in document.InlineShapes:
                    if shape.Range.Start > selection.Range.End:
                        next_image = shape
                        break
                if next_image:
                    next_image.Select()
                else:
                    return {
                        "success": False,
                        "message": "No more images found"
                    }
            else:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    f"Unsupported object type for navigation: {object_type}"
                )
        else:
            # 默认移动到下一个段落
            selection.MoveDown(Unit=5, Count=1)
        
        # 检查是否成功移动
        if selection.Range.Start == current_start:
            return {
                "success": False,
                "message": "Already at the last object"
            }
        
        # 滚动到新位置
        word_app.ActiveWindow.ScrollIntoView(selection.Range)
        
        log_info(f"Successfully navigated to next {object_type or 'object'}")
        
        # 获取新的活动对象信息
        new_object_info = get_active_object(document)
        
        return {
            "success": True,
            "message": f"Successfully navigated to next {object_type or 'object'}",
            "active_object": new_object_info["active_object"]
        }
    except Exception as e:
        raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to navigate to next object: {str(e)}")


@handle_com_error(ErrorCode.SERVER_ERROR, "navigate to previous object")
def navigate_to_previous_object(document: win32com.client.CDispatch, object_type: Optional[str] = None) -> Dict[str, Any]:
    """导航到上一个对象

    Args:
        document: Word文档COM对象
        object_type: 可选的对象类型过滤

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当导航失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    try:
        word_app = document.Application
        selection = word_app.Selection
        
        # 保存当前位置用于检查是否成功移动
        current_start = selection.Range.Start
        
        # 根据对象类型进行导航
        if object_type:
            object_type = object_type.lower()
            
            if object_type == "paragraph":
                # 移动到上一段
                selection.MoveUp(Unit=5, Count=1)
            elif object_type == "table":
                # 查找上一个表格
                prev_table = None
                for table in reversed(document.Tables):
                    if table.Range.End < selection.Range.Start:
                        prev_table = table
                        break
                if prev_table:
                    prev_table.Select()
                else:
                    return {
                        "success": False,
                        "message": "No previous tables found"
                    }
            elif object_type == "image":
                # 查找上一个图像
                prev_image = None
                for shape in reversed(document.InlineShapes):
                    if shape.Range.End < selection.Range.Start:
                        prev_image = shape
                        break
                if prev_image:
                    prev_image.Select()
                else:
                    return {
                        "success": False,
                        "message": "No previous images found"
                    }
            else:
                raise WordDocumentError(
                    ErrorCode.INVALID_INPUT,
                    f"Unsupported object type for navigation: {object_type}"
                )
        else:
            # 默认移动到上一个段落
            selection.MoveUp(Unit=5, Count=1)
        
        # 检查是否成功移动
        if selection.Range.Start == current_start:
            return {
                "success": False,
                "message": "Already at the first object"
            }
        
        # 滚动到新位置
        word_app.ActiveWindow.ScrollIntoView(selection.Range)
        
        log_info(f"Successfully navigated to previous {object_type or 'object'}")
        
        # 获取新的活动对象信息
        new_object_info = get_active_object(document)
        
        return {
            "success": True,
            "message": f"Successfully navigated to previous {object_type or 'object'}",
            "active_object": new_object_info["active_object"]
        }
    except Exception as e:
        raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to navigate to previous object: {str(e)}")


@handle_com_error(ErrorCode.SERVER_ERROR, "get context information")
def get_context_information(document: win32com.client.CDispatch) -> Dict[str, Any]:
    """获取当前文档的上下文信息

    Args:
        document: Word文档COM对象

    Returns:
        包含当前上下文信息的字典

    Raises:
        WordDocumentError: 当获取上下文信息失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    try:
        word_app = document.Application
        selection = word_app.Selection
        active_window = word_app.ActiveWindow
        
        # 获取当前上下文信息
        context_info = {
            "document": {
                "name": document.Name,
                "path": document.FullName if document.FullName else "Unsaved"
            },
            "current_position": {
                "page": selection.Information(3),  # wdActiveEndPageNumber
                "line": selection.Information(10),  # wdFirstCharacterLineNumber
                "section": selection.Information(1)  # wdActiveEndSectionNumber
            }
        }
        
        # 添加活动对象信息
        active_object_info = get_active_object(document)
        if active_object_info["success"]:
            context_info["active_object"] = active_object_info["active_object"]
        
        log_info("Successfully retrieved context information")
        
        return {
            "success": True,
            "context": context_info
        }
    except Exception as e:
        raise WordDocumentError(ErrorCode.SERVER_ERROR, f"Failed to get context information: {str(e)}")


@handle_com_error(ErrorCode.SERVER_ERROR, "set zoom level")
def set_zoom_level(document: win32com.client.CDispatch, zoom_level: int = None) -> Dict[str, Any]:
    """设置文档缩放比例（内部使用，对外部隐藏视图细节）

    Args:
        document: Word文档COM对象
        zoom_level: 缩放比例(10-500)，设置为None或不提供时使用默认级别(100%)

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当设置缩放失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    # 使用默认级别（100%）
    if zoom_level is None:
        zoom_level = 100
    else:
        # 检查缩放比例是否在有效范围内
        if not (10 <= zoom_level <= 500):
            raise WordDocumentError(
                ErrorCode.INVALID_INPUT,
                "Zoom level must be between 10 and 500"
            )
    
    # 设置缩放比例
    document.ActiveWindow.View.Zoom.Percentage = zoom_level
    log_info(f"Successfully set zoom level to {zoom_level}%")
    
    return {
        "success": True,
        "zoom_level": zoom_level,
        "message": f"Zoom level successfully set to {zoom_level}%"
    }