"""
上下文导航工具模块

此模块提供文档对象导航、选择和查看相关的功能，支持在文档中定位和操作不同类型的对象。
"""

import logging
from typing import Optional, Dict, Any, Union
from win32com.client import CDispatch
from ..utils.com_error_handler import handle_com_error
from ..utils.logger import get_logger
from .context_control import DocumentContext

# 获取日志记录器
logger = get_logger(__name__)


@handle_com_error
def set_active_context(context: Optional[DocumentContext], word_app: Optional[CDispatch] = None) -> bool:
    """
    设置活动上下文并在Word中选择相应的对象
    
    Args:
        context: 要设置为活动的上下文对象
        word_app: Word应用程序实例，如果为None则尝试从上下文中获取
        
    Returns:
        设置是否成功
    """
    if not context:
        logger.warning("No context provided to set as active")
        return False
    
    try:
        # 尝试从上下文中获取Word应用实例
        if not word_app and hasattr(context, 'word_app'):
            word_app = context.word_app
        
        if not word_app:
            logger.error("No Word application instance available")
            return False
        
        # 尝试选择上下文范围
        if hasattr(context, 'range') and context.range:
            try:
                # 选择范围
                context.range.Select()
                # 滚动到视图中
                word_app.ActiveWindow.ScrollIntoView(context.range)
                
                # 对于书签类型的上下文，可以添加书签选择逻辑
                if context.metadata.get('type') == 'bookmark':
                    bookmark_name = context.metadata.get('bookmark_name')
                    if bookmark_name and hasattr(word_app.ActiveDocument, 'Bookmarks'):
                        try:
                            word_app.ActiveDocument.Bookmarks(bookmark_name).Select()
                        except Exception:
                            logger.warning(f"Bookmark '{bookmark_name}' not found, using range selection instead")
                
                logger.debug(f"Activated context: {context.title} (ID: {context.context_id})")
                return True
            except Exception as e:
                logger.error(f"Failed to select context range: {e}")
                return False
        
        logger.warning("Context does not have a valid range to select")
        return False
    except Exception as e:
        logger.error(f"Failed to set active context: {e}")
        return False


@handle_com_error
def get_active_object(word_app: CDispatch) -> Optional[Dict[str, Any]]:
    """
    获取当前活动对象的信息
    
    Args:
        word_app: Word应用程序实例
        
    Returns:
        包含活动对象信息的字典，如果没有活动对象则返回None
    """
    try:
        active_doc = word_app.ActiveDocument
        if not active_doc:
            logger.warning("No active document")
            return None
        
        selection = word_app.Selection
        if not selection:
            logger.warning("No selection available")
            return None
        
        # 获取选择范围的基本信息
        try:
            start_pos = selection.Start
            end_pos = selection.End
            is_collapsed = selection.Start == selection.End
        except Exception:
            start_pos = end_pos = 0
            is_collapsed = True
        
        object_info = {
            'start': start_pos,
            'end': end_pos,
            'is_collapsed': is_collapsed
        }
        
        # 尝试识别选择的对象类型
        try:
            # 检查是否为表格
            if selection.Information(12) and selection.Tables.Count > 0:  # wdWithInTable
                table = selection.Tables(1)
                object_info.update({
                    'type': 'table',
                    'details': {
                        'rows': table.Rows.Count,
                        'columns': table.Columns.Count,
                        'cell_count': table.Rows.Count * table.Columns.Count
                    }
                })
            # 检查是否为图片
            elif selection.InlineShapes.Count > 0:
                image = selection.InlineShapes(1)
                object_info.update({
                    'type': 'image',
                    'details': {
                        'width': image.Width,
                        'height': image.Height,
                        'type': str(image.Type) if hasattr(image, 'Type') else 'Unknown'
                    }
                })
            # 默认视为段落
            else:
                # 获取段落文本预览
                try:
                    text_preview = selection.Text[:30] + ("..." if len(selection.Text) > 30 else "")
                except Exception:
                    text_preview = ""
                
                # 尝试获取段落样式
                try:
                    style_name = selection.ParagraphFormat.Style.Name
                    is_heading = style_name.startswith('Heading')
                except Exception:
                    style_name = 'Normal'
                    is_heading = False
                
                object_info.update({
                    'type': 'paragraph',
                    'details': {
                        'text_preview': text_preview,
                        'style_name': style_name,
                        'is_heading': is_heading
                    }
                })
        except Exception as e:
            logger.error(f"Error identifying active object type: {e}")
            object_info['type'] = 'unknown'
        
        logger.debug(f"Retrieved active object information: {object_info['type']}")
        return object_info
    except Exception as e:
        logger.error(f"Failed to get active object: {e}")
        return None


@handle_com_error
def navigate_to_next_object(word_app: CDispatch, object_type: str = 'any') -> bool:
    """
    导航到文档中的下一个指定类型的对象
    
    Args:
        word_app: Word应用程序实例
        object_type: 对象类型，可选值: 'any', 'paragraph', 'table', 'image'
        
    Returns:
        导航是否成功
    """
    if not word_app or not hasattr(word_app, 'Selection'):
        logger.error("Invalid Word application instance")
        return False
    
    try:
        selection = word_app.Selection
        active_doc = word_app.ActiveDocument
        
        if not selection or not active_doc:
            logger.warning("No selection or active document available")
            return False
        
        current_pos = selection.End
        
        # 根据对象类型执行导航
        if object_type == 'table' or object_type == 'any':
            # 查找下一个表格
            for table in active_doc.Tables:
                if table.Range.Start > current_pos:
                    table.Range.Select()
                    logger.debug(f"Navigated to next table at position {table.Range.Start}")
                    return True
        
        if (object_type == 'image' or object_type == 'any') and hasattr(active_doc, 'InlineShapes'):
            # 查找下一个图片
            for image in active_doc.InlineShapes:
                if image.Range.Start > current_pos:
                    image.Range.Select()
                    logger.debug(f"Navigated to next image at position {image.Range.Start}")
                    return True
        
        if object_type == 'paragraph' or object_type == 'any':
            # 查找下一个段落
            # 如果当前不在文档末尾，则移动到下一个段落
            if selection.End < active_doc.Content.End:
                selection.Collapse(Direction=0)  # wdCollapseEnd
                selection.MoveDown(Unit=5, Count=1)  # wdParagraph
                
                # 验证是否成功移动到新段落
                if selection.Start > current_pos:
                    logger.debug(f"Navigated to next paragraph at position {selection.Start}")
                    return True
        
        logger.info("No more objects of specified type found")
        return False
    except Exception as e:
        logger.error(f"Failed to navigate to next object: {e}")
        return False


@handle_com_error
def navigate_to_previous_object(word_app: CDispatch, object_type: str = 'any') -> bool:
    """
    导航到文档中的上一个指定类型的对象
    
    Args:
        word_app: Word应用程序实例
        object_type: 对象类型，可选值: 'any', 'paragraph', 'table', 'image'
        
    Returns:
        导航是否成功
    """
    if not word_app or not hasattr(word_app, 'Selection'):
        logger.error("Invalid Word application instance")
        return False
    
    try:
        selection = word_app.Selection
        active_doc = word_app.ActiveDocument
        
        if not selection or not active_doc:
            logger.warning("No selection or active document available")
            return False
        
        current_pos = selection.Start
        
        # 根据对象类型执行导航
        if object_type == 'table' or object_type == 'any':
            # 查找上一个表格（从后往前遍历）
            last_table = None
            for table in active_doc.Tables:
                if table.Range.End < current_pos:
                    last_table = table
                else:
                    break  # 因为表格是按顺序排列的，一旦超过当前位置就可以停止
            
            if last_table:
                last_table.Range.Select()
                logger.debug(f"Navigated to previous table at position {last_table.Range.Start}")
                return True
        
        if (object_type == 'image' or object_type == 'any') and hasattr(active_doc, 'InlineShapes'):
            # 查找上一个图片（从后往前遍历）
            last_image = None
            for image in active_doc.InlineShapes:
                if image.Range.End < current_pos:
                    last_image = image
                else:
                    break  # 因为图片是按顺序排列的，一旦超过当前位置就可以停止
            
            if last_image:
                last_image.Range.Select()
                logger.debug(f"Navigated to previous image at position {last_image.Range.Start}")
                return True
        
        if object_type == 'paragraph' or object_type == 'any':
            # 查找上一个段落
            # 如果当前不在文档开头，则移动到上一个段落
            if selection.Start > 0:
                selection.Collapse(Direction=1)  # wdCollapseStart
                selection.MoveUp(Unit=5, Count=1)  # wdParagraph
                
                # 验证是否成功移动到新段落
                if selection.End < current_pos:
                    logger.debug(f"Navigated to previous paragraph at position {selection.Start}")
                    return True
        
        logger.info("No previous objects of specified type found")
        return False
    except Exception as e:
        logger.error(f"Failed to navigate to previous object: {e}")
        return False


@handle_com_error
def get_context_information(word_app: CDispatch) -> Dict[str, Any]:
    """
    获取当前文档的上下文信息，包括文档信息、位置信息和活动对象
    
    Args:
        word_app: Word应用程序实例
        
    Returns:
        包含上下文信息的字典
    """
    context_info = {
        'document': {},
        'position': {},
        'active_object': None,
        'success': True
    }
    
    try:
        # 获取文档信息
        active_doc = word_app.ActiveDocument
        if active_doc:
            try:
                context_info['document'] = {
                    'name': getattr(active_doc, 'Name', 'Untitled'),
                    'full_name': getattr(active_doc, 'FullName', ''),
                    'path': getattr(active_doc, 'Path', ''),
                    'saved': getattr(active_doc, 'Saved', False),
                    'characters': getattr(active_doc.Content, 'Characters', 0).Count,
                    'paragraphs': getattr(active_doc, 'Paragraphs', 0).Count,
                    'sections': getattr(active_doc, 'Sections', 0).Count,
                    'tables': getattr(active_doc, 'Tables', 0).Count
                }
                
                # 添加页数信息（如果可用）
                if hasattr(active_doc, 'ComputeStatistics'):
                    try:
                        context_info['document']['pages'] = active_doc.ComputeStatistics(2)  # wdStatisticPages
                    except Exception:
                        context_info['document']['pages'] = 0
            except Exception as e:
                logger.error(f"Error retrieving document information: {e}")
                context_info['document'] = {'error': str(e)}
        else:
            context_info['document'] = {'error': 'No active document'}
        
        # 获取位置信息
        selection = word_app.Selection
        if selection:
            try:
                context_info['position'] = {
                    'start': selection.Start,
                    'end': selection.End,
                    'page': selection.Information(3) if hasattr(selection, 'Information') else 0,  # wdActiveEndPageNumber
                    'line': selection.Information(10) if hasattr(selection, 'Information') else 0  # wdFirstCharacterLineNumber
                }
            except Exception as e:
                logger.error(f"Error retrieving position information: {e}")
                context_info['position'] = {'error': str(e)}
        else:
            context_info['position'] = {'error': 'No selection available'}
        
        # 获取活动对象信息
        context_info['active_object'] = get_active_object(word_app)
        
        logger.debug("Retrieved context information successfully")
    except Exception as e:
        logger.error(f"Failed to get context information: {e}")
        context_info['success'] = False
        context_info['error'] = str(e)
    
    return context_info


@handle_com_error
def set_zoom_level(word_app: CDispatch, zoom_level: Union[int, float] = None) -> bool:
    """
    设置文档视图的缩放比例
    
    Args:
        word_app: Word应用程序实例
        zoom_level: 缩放比例（10-500之间的整数或浮点数），如果为None则使用默认值
        
    Returns:
        设置是否成功
    """
    try:
        # 验证Word应用实例
        if not word_app or not hasattr(word_app, 'ActiveWindow'):
            logger.error("Invalid Word application instance")
            return False
        
        # 获取活动窗口
        active_window = word_app.ActiveWindow
        if not active_window or not hasattr(active_window, 'View'):
            logger.error("No active window or view available")
            return False
        
        # 确定缩放级别
        if zoom_level is None:
            # 使用默认缩放级别
            zoom_level = 100
        else:
            # 验证并限制缩放级别范围
            try:
                zoom_level = float(zoom_level)
                # Word的缩放范围通常是10-500%
                zoom_level = max(10.0, min(500.0, zoom_level))
            except ValueError:
                logger.error(f"Invalid zoom level: {zoom_level}")
                return False
        
        # 设置缩放级别
        if hasattr(active_window.View, 'Zoom') and hasattr(active_window.View.Zoom, 'Percentage'):
            active_window.View.Zoom.Percentage = zoom_level
            logger.debug(f"Set document zoom level to {zoom_level}%")
            return True
        
        logger.error("Zoom settings not available in this Word version")
        return False
    except Exception as e:
        logger.error(f"Failed to set zoom level: {e}")
        return False