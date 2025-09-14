"""
Selector utilities for Word Document MCP Server.

This module contains utility functions for handling document selection and object
location using locators. It provides a unified interface for selecting ranges
in Word documents across different operation modules.
"""

from typing import Any, Dict, Optional, Union

from ..mcp_service.errors import ErrorCode, WordDocumentError


def get_selection_range(
    document: Any,
    locator: Optional[Dict[str, Any]] = None,
    operation_name: str = "operation",
    position: str = "after"
) -> Any:
    """
    获取文档中基于定位器的选择范围
    
    这是一个统一的选择器函数，用于替换各个操作模块中重复的_get_selection_range实现。
    
    Args:
        document: Word文档COM对象
        locator: 定位器对象，用于指定选择范围
        operation_name: 操作名称，用于错误消息
        position: 位置参数，可选值："before", "after", "replace"
        
    Returns:
        选择范围的COM对象
        
    Raises:
        WordDocumentError: 如果定位失败或文档不可用
        ValueError: 如果提供的参数无效
    """
    if not document:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, 
            f"No active document found for {operation_name}"
        )
    
    try:
        # 验证document对象是否有效
        if not hasattr(document, "Application") or document.Application is None:
            raise WordDocumentError(
                ErrorCode.DOCUMENT_ERROR, 
                "Document or Application object not available"
            )
        
        # 如果没有提供locator，使用当前选择位置或整个文档
        if locator is None:
            try:
                # 尝试使用Application.Selection.Range
                return document.Application.Selection.Range
            except Exception:
                # 如果失败，使用整个文档内容
                return document.Content
        
        # 处理不同类型的定位器
        locator_type = locator.get('type', '').lower()
        
        # 处理文档开始位置
        if locator_type == 'document_start':
            range_obj = document.Range()
            range_obj.Collapse(True)  # wdCollapseStart
            return range_obj
        
        # 处理文档结束位置
        elif locator_type == 'document_end':
            range_obj = document.Range()
            range_obj.Collapse(False)  # wdCollapseEnd
            return range_obj
        
        # 处理文档通用位置
        elif locator_type == 'document':
            # 定位到文档的特定位置
            position_value = locator.get('position', '').lower()
            if position_value == 'start':
                return document.Range(0, 0)
            elif position_value == 'end':
                return document.Range(document.Content.End - 1, document.Content.End - 1)
            else:
                return document.Content
        
        # 处理段落定位
        elif locator_type == 'paragraph' or 'paragraph' in locator:
            # 尝试从locator中获取段落索引
            index = None
            if 'paragraph' in locator and not isinstance(locator['paragraph'], dict):
                # 直接指定段落索引的情况
                index = locator['paragraph']
            else:
                # 尝试从多个可能的参数中获取索引
                for param_name in ['value', 'index', 'id']:
                    if param_name in locator:
                        index = locator[param_name]
                        break
            
            if index is None:
                raise ValueError("Paragraph locator must contain 'value', 'index', 'id' or 'paragraph'")
            
            try:
                # 转换为整数索引
                index = int(index)
                
                # 处理索引范围检查
                paragraph_count = document.Paragraphs.Count
                
                # 支持负索引（从末尾开始计数）
                if index < 0:
                    index = paragraph_count + index + 1
                    if index <= 0:
                        index = 1
                
                # 检查索引是否超出范围
                if index <= 0:
                    raise ValueError("Paragraph index must be positive")
                
                if index > paragraph_count:
                    # 索引超出范围时的处理策略
                    if position in ['after', 'end']:
                        # 如果请求的是文档末尾，返回末尾范围
                        range_obj = document.Range()
                        range_obj.Collapse(False)  # wdCollapseEnd
                        return range_obj
                    else:
                        raise IndexError(f"Paragraph index {index} out of range (1-{paragraph_count})")
                
                # 返回指定段落的范围
                return document.Paragraphs(index).Range
                
            except (ValueError, TypeError) as e:
                raise WordDocumentError(
                    ErrorCode.OBJECT_NOT_FOUND, 
                    f"Invalid paragraph index: {e}"
                )
            except IndexError as e:
                raise WordDocumentError(
                    ErrorCode.OBJECT_NOT_FOUND, 
                    str(e)
                )
        
        # 处理默认情况 - 返回当前选择位置
        try:
            return document.Application.Selection.Range
        except Exception:
            # 如果失败，返回文档末尾
            range_obj = document.Range()
            range_obj.Collapse(False)  # wdCollapseEnd
            return range_obj
        
    except WordDocumentError:
        # 重新抛出WordDocumentError异常
        raise
    except Exception as e:
        # 捕获其他所有异常并转换为WordDocumentError
        raise WordDocumentError(
            ErrorCode.SELECTION_ERROR, 
            f"Failed to get selection range for {operation_name}: {str(e)}"
        )


def validate_locator(locator: Any, expected_types: Optional[List[str]] = None) -> Dict[str, Any]:
    """
    验证定位器对象的有效性
    
    Args:
        locator: 要验证的定位器对象
        expected_types: 期望的定位器类型列表（可选）
        
    Returns:
        验证后的定位器字典
        
    Raises:
        ValueError: 如果定位器无效或类型不匹配
    """
    if locator is None:
        return {}
        
    if not isinstance(locator, dict):
        raise ValueError(f"Locator must be a dictionary, got {type(locator).__name__}")
        
    if expected_types and 'type' in locator:
        locator_type = locator['type'].lower()
        if locator_type not in [t.lower() for t in expected_types]:
            raise ValueError(
                f"Locator type must be one of {expected_types}, got '{locator_type}'"
            )
            
    return locator