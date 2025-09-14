"""
上下文控制工具模块，用于管理Word文档的上下文和活动对象。

此模块提供了设置上下文、管理活动对象和在对象之间导航的功能，
视图相关的功能被隐藏在幕后，以提供更简洁的接口。
"""
import os
import logging
from typing import Dict, Any, Optional, Union

# Standard library imports
from dotenv import load_dotenv
# Third-party imports
from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from pydantic import Field

# Local imports
from ..mcp_service.core import mcp_server
from ..mcp_service.app_context import AppContext
from ..mcp_service.core_utils import ErrorCode, WordDocumentError
# 导入操作相关模块
from ..contexts.context_control import set_active_context, get_active_object, navigate_to_next_object, navigate_to_previous_object, get_context_information, set_zoom_level

# 加载.env文件中的环境变量
load_dotenv()

logger = logging.getLogger(__name__)

# 导入COM相关模块
import win32com.client
from win32com.client import CDispatch
from pythoncom import com_error  # pylint: disable=no-name-in-module

def _get_app_context() -> AppContext:
    """获取应用上下文实例。"""
    return AppContext.get_instance()

def _get_word_app() -> CDispatch:
    """获取Word应用程序实例。"""
    app_context = _get_app_context()
    word_app = app_context.get_word_app(create_if_needed=False)
    if not word_app:
        raise WordDocumentError(ErrorCode.APPLICATION_ERROR, "未找到Word应用程序实例")
    return word_app

def _get_active_document() -> CDispatch:
    """获取当前活动文档。"""
    word_app = _get_word_app()
    try:
        if word_app.Documents.Count == 0:
            raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "没有打开的文档")
        return word_app.ActiveDocument
    except com_error as e:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, f"获取活动文档失败: {str(e)}")

def _get_current_selection_range(document: CDispatch = None):
    """获取当前选择范围。"""
    if document is None:
        document = _get_active_document()
    
    try:
        word_app = document.Application
        return word_app.Selection.Range
    except Exception as e:
        raise WordDocumentError(ErrorCode.OBJECT_ERROR, f"获取选择范围失败: {str(e)}")

def scroll_to_current_object() -> Dict[str, Any]:
    """
    滚动视图到当前工作对象（内部实现）。
    
    Returns:
        Dict[str, Any]: 操作结果，包含是否成功、当前对象信息等
    """
    try:
        document = _get_active_document()
        
        # 获取当前选择范围
        range_obj = _get_current_selection_range(document)
        
        # 滚动视图到当前对象
        word_app = _get_word_app()
        word_app.ActiveWindow.ScrollIntoView(range_obj)
        
        # 选中当前对象
        range_obj.Select()
        
        # 获取当前活动对象信息
        active_object_info = get_active_object(document)
        if not active_object_info['success']:
            raise WordDocumentError(ErrorCode.OBJECT_ERROR, "获取当前活动对象信息失败")
        
        current_object = active_object_info['active_object']
        logger.info(f"已滚动到当前对象: {current_object['type']}")
        
        return {
            'success': True,
            'message': '已滚动到当前工作对象',
            'current_object': current_object,
            'context': get_context_information(document)['context']
        }
    except Exception as e:
        logger.error(f"滚动到当前对象失败: {str(e)}")
        return {
            'success': False,
            'error': str(e),
            'error_code': ErrorCode.VIEW_ERROR
        }

def move_to_next_object() -> Dict[str, Any]:
    """
    移动到下一个文档对象（调用context_control模块中的函数）。
    
    Returns:
        Dict[str, Any]: 操作结果，包含是否成功、当前对象信息等
    """
    try:
        # 调用context_control模块中的函数
        result = navigate_to_next_object()
        
        # 滚动视图到新的当前对象
        if result['success']:
            scroll_result = scroll_to_current_object()
            if scroll_result['success']:
                logger.info(f"已移动到下一个对象: {result.get('current_object', {}).get('type', '未知')}")
            else:
                result['warning'] = '对象导航成功但视图滚动失败'
        
        return result
    except Exception as e:
        logger.error(f"移动到下一个对象失败: {str(e)}")
        return {
            'success': False,
            'error': str(e),
            'error_code': ErrorCode.VIEW_ERROR
        }

def move_to_previous_object() -> Dict[str, Any]:
    """
    移动到上一个文档对象（调用context_control模块中的函数）。
    
    Returns:
        Dict[str, Any]: 操作结果，包含是否成功、当前对象信息等
    """
    try:
        # 调用context_control模块中的函数
        result = navigate_to_previous_object()
        
        # 滚动视图到新的当前对象
        if result['success']:
            scroll_result = scroll_to_current_object()
            if scroll_result['success']:
                logger.info(f"已移动到上一个对象: {result.get('current_object', {}).get('type', '未知')}")
            else:
                result['warning'] = '对象导航成功但视图滚动失败'
        
        return result
    except Exception as e:
        logger.error(f"移动到上一个对象失败: {str(e)}")
        return {
            'success': False,
            'error': str(e),
            'error_code': ErrorCode.VIEW_ERROR
        }

def move_to_next_section() -> Dict[str, Any]:
    """
    移动到下一个大纲节点（章节），并滚动视图聚焦。
    
    Returns:
        Dict[str, Any]: 操作结果，包含是否成功、当前章节信息等
    """
    try:
        document = _get_active_document()
        word_app = document.Application
        selection = word_app.Selection
        
        # 保存当前位置用于检查是否成功移动
        current_start = selection.Range.Start
        
        # 查找下一个标题段落
        found = False
        for paragraph in document.Paragraphs:
            # 检查段落是否为标题样式
            if paragraph.Style.NameLocal.startswith('Heading') or paragraph.Style.NameLocal.startswith('标题'):
                if paragraph.Range.Start > current_start:
                    paragraph.Range.Select()
                    found = True
                    break
        
        if not found:
            return {
                'success': False,
                'message': '已经是最后一个章节'
            }
        
        # 滚动视图到新的当前章节
        result = scroll_to_current_object()
        
        if result['success']:
            logger.info(f"已移动到下一个章节")
            
        return result
    except Exception as e:
        logger.error(f"移动到下一个章节失败: {str(e)}")
        return {
            'success': False,
            'error': str(e),
            'error_code': ErrorCode.VIEW_ERROR
        }

def move_to_previous_section() -> Dict[str, Any]:
    """
    移动到上一个大纲节点（章节），并滚动视图聚焦。
    
    Returns:
        Dict[str, Any]: 操作结果，包含是否成功、当前章节信息等
    """
    try:
        document = _get_active_document()
        word_app = document.Application
        selection = word_app.Selection
        
        # 保存当前位置用于检查是否成功移动
        current_start = selection.Range.Start
        
        # 查找上一个标题段落
        found = False
        prev_paragraph = None
        for paragraph in document.Paragraphs:
            # 检查段落是否为标题样式
            if paragraph.Style.NameLocal.startswith('Heading') or paragraph.Style.NameLocal.startswith('标题'):
                if paragraph.Range.Start < current_start:
                    prev_paragraph = paragraph
                else:
                    break
        
        if not prev_paragraph:
            return {
                'success': False,
                'message': '已经是第一个章节'
            }
        
        # 选中找到的上一个标题段落
        prev_paragraph.Range.Select()
        
        # 滚动视图到新的当前章节
        result = scroll_to_current_object()
        
        if result['success']:
            logger.info(f"已移动到上一个章节")
            
        return result
    except Exception as e:
        logger.error(f"移动到上一个章节失败: {str(e)}")
        return {
            'success': False,
            'error': str(e),
            'error_code': ErrorCode.VIEW_ERROR
        }

def get_current_context() -> Dict[str, Any]:
    """
    获取当前文档的上下文信息（调用context_control模块中的函数）。
    
    Returns:
        Dict[str, Any]: 当前上下文信息，包括当前章节、关注范围、当前工作对象等
    """
    try:
        # 调用context_control模块中的函数
        return get_context_information()
    except Exception as e:
        logger.error(f"获取当前上下文失败: {str(e)}")
        return {
            'success': False,
            'error': str(e),
            'error_code': ErrorCode.VIEW_ERROR
        }

def set_zoom(percentage: int = None) -> Dict[str, Any]:
    """
    设置文档视图的缩放比例（调用context_control模块中的函数）。
    
    Args:
        percentage: 缩放百分比（10-500之间），不提供时使用默认级别(100%)
        
    Returns:
        Dict[str, Any]: 操作结果
    """
    try:
        # 调用context_control模块中的函数
        return set_zoom_level(percentage)
    except Exception as e:
        logger.error(f"设置缩放比例失败: {str(e)}")
        return {
            'success': False,
            'error': str(e),
            'error_code': ErrorCode.VIEW_ERROR
        }

@mcp_server.tool()
async def view_control_tools(
    ctx: Context[ServerSession, AppContext] = Field(description="Context object"),
    operation_type: str = Field(
        ...,
        description="Type of context control operation to perform: scroll_to_current, next_object, previous_object, next_section, previous_section, get_context, set_zoom, get_active_object",
    ),
    percentage: Optional[int] = Field(
        default=100,
        description="Zoom percentage (10-500) for set_zoom operation",
    ),
    params: Optional[Dict[str, Any]] = Field(
        default=None,
        description="Parameters for test compatibility"
    )
) -> Dict[str, Any]:
    """上下文控制工具

    支持的操作类型：
    - scroll_to_current: 滚动到当前工作对象
    - next_object: 移动到下一个文档对象
    - previous_object: 移动到上一个文档对象
    - next_section: 移动到下一个大纲节点（章节）
    - previous_section: 移动到上一个大纲节点（章节）
    - get_context: 获取当前文档的上下文信息
    - set_zoom: 设置文档视图的缩放比例
      * 必需参数：percentage (10-500之间)
    - get_active_object: 获取当前活动对象信息
    
    注意：视图相关的功能已隐藏到幕后，主要提供上下文和活动对象的管理功能。
    """
    try:
        # 处理params参数，兼容测试用例
        if params:
            if 'operation_type' in params:
                operation_type = params['operation_type']
            if 'percentage' in params:
                percentage = params['percentage']
                
        operations = {
            'scroll_to_current': scroll_to_current_object,
            'next_object': move_to_next_object,
            'previous_object': move_to_previous_object,
            'next_section': move_to_next_section,
            'previous_section': move_to_previous_section,
            'get_context': get_current_context,
            'set_zoom': lambda: set_zoom(percentage),
            'get_active_object': lambda: get_active_object()
        }
        
        if operation_type not in operations:
            return {
                'success': False,
                'error': f'不支持的操作类型: {operation_type}',
                'error_code': ErrorCode.INVALID_PARAMETER
            }
        
        return operations[operation_type]()
    except Exception as e:
        logger.error(f"上下文控制操作失败 ({operation_type}): {str(e)}")
        return {
            'success': False,
            'error': str(e),
            'error_code': ErrorCode.VIEW_ERROR
        }