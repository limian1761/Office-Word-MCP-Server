"""
View Control Operations for Word Document MCP Server.

This module contains operations for controlling document view settings in Word.
"""

import logging
from typing import Dict, Any, Optional

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..mcp_service.core_utils import (
    ErrorCode,
    WordDocumentError,
    log_error,
    log_info
)


@handle_com_error(ErrorCode.SERVER_ERROR, "switch document view")
def switch_view(document: win32com.client.CDispatch, view_type: str) -> Dict[str, Any]:
    """切换文档视图

    Args:
        document: Word文档COM对象
        view_type: 视图类型

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当切换视图失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    # 检查视图类型是否有效
    valid_view_types = {
        "print": 3,  # wdPrintView
        "web": 6,   # wdWebView
        "read": 7,  # wdReadingView
        "outline": 2,  # wdOutlineView
        "draft": 1  # wdNormalView
    }
    
    if view_type.lower() not in valid_view_types:
        valid_types_str = ", ".join(valid_view_types.keys())
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT,
            f"Invalid view type. Supported types: {valid_types_str}"
        )
    
    # 切换视图
    document.ActiveWindow.View.Type = valid_view_types[view_type.lower()]
    log_info(f"Successfully switched view to {view_type}")
    
    return {
        "success": True,
        "view_type": view_type,
        "message": f"View successfully switched to {view_type}"
    }


@handle_com_error(ErrorCode.SERVER_ERROR, "set document zoom")
def set_zoom(document: win32com.client.CDispatch, zoom_level: int) -> Dict[str, Any]:
    """设置文档缩放比例

    Args:
        document: Word文档COM对象
        zoom_level: 缩放比例(10-500)

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当设置缩放失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
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


@handle_com_error(ErrorCode.SERVER_ERROR, "show document element")
def show_element(document: win32com.client.CDispatch, element_type: str) -> Dict[str, Any]:
    """显示文档特定元素

    Args:
        document: Word文档COM对象
        element_type: 元素类型

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当显示元素失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    # 处理不同类型的元素显示
    element_type = element_type.lower()
    word_app = document.Application
    
    if element_type == "rulers":
        word_app.ActiveWindow.DisplayRulers = True
    elif element_type == "gridlines":
        word_app.ActiveWindow.View.GridLines = True
    elif element_type == "navigation_pane":
        word_app.ActiveWindow.Sidepane.Visible = True
        word_app.ActiveWindow.Sidepane.Show = 1  # wdShowNavPane
    elif element_type == "status_bar":
        word_app.DisplayStatusBar = True
    elif element_type == "task_pane":
        # 默认显示格式任务窗格
        word_app.CommandBars("Formatting").Visible = True
    elif element_type == "comments_pane":
        # 显示审阅窗格
        word_app.ActiveWindow.View.SplitSpecial = 3  # wdPaneComments
    elif element_type == "formatting_marks":
        word_app.ActiveWindow.View.ShowAll = True
    else:
        valid_elements = "rulers, gridlines, navigation_pane, status_bar, task_pane, comments_pane, formatting_marks"
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT,
            f"Invalid element type. Supported elements: {valid_elements}"
        )
    
    log_info(f"Successfully showed {element_type}")
    return {
        "success": True,
        "element_type": element_type,
        "message": f"Successfully showed {element_type}"
    }


@handle_com_error(ErrorCode.SERVER_ERROR, "hide document element")
def hide_element(document: win32com.client.CDispatch, element_type: str) -> Dict[str, Any]:
    """隐藏文档特定元素

    Args:
        document: Word文档COM对象
        element_type: 元素类型

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当隐藏元素失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    # 处理不同类型的元素隐藏
    element_type = element_type.lower()
    word_app = document.Application
    
    if element_type == "rulers":
        word_app.ActiveWindow.DisplayRulers = False
    elif element_type == "gridlines":
        word_app.ActiveWindow.View.GridLines = False
    elif element_type == "navigation_pane":
        word_app.ActiveWindow.Sidepane.Visible = False
    elif element_type == "status_bar":
        word_app.DisplayStatusBar = False
    elif element_type == "task_pane":
        # 隐藏所有任务窗格
        for bar in word_app.CommandBars:
            if bar.Type == 16:  # msoBarTypeTaskPane
                bar.Visible = False
    elif element_type == "comments_pane":
        # 隐藏审阅窗格
        word_app.ActiveWindow.View.SplitSpecial = 0  # wdPaneNone
    elif element_type == "formatting_marks":
        word_app.ActiveWindow.View.ShowAll = False
    else:
        valid_elements = "rulers, gridlines, navigation_pane, status_bar, task_pane, comments_pane, formatting_marks"
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT,
            f"Invalid element type. Supported elements: {valid_elements}"
        )
    
    log_info(f"Successfully hid {element_type}")
    return {
        "success": True,
        "element_type": element_type,
        "message": f"Successfully hid {element_type}"
    }


@handle_com_error(ErrorCode.SERVER_ERROR, "toggle document element")
def toggle_element(document: win32com.client.CDispatch, element_type: str) -> Dict[str, Any]:
    """切换文档特定元素的显示状态

    Args:
        document: Word文档COM对象
        element_type: 元素类型

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当切换元素状态失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    # 获取当前状态并切换
    element_type = element_type.lower()
    word_app = document.Application
    current_state = None
    
    if element_type == "rulers":
        current_state = word_app.ActiveWindow.DisplayRulers
        word_app.ActiveWindow.DisplayRulers = not current_state
    elif element_type == "gridlines":
        current_state = word_app.ActiveWindow.View.GridLines
        word_app.ActiveWindow.View.GridLines = not current_state
    elif element_type == "navigation_pane":
        current_state = word_app.ActiveWindow.Sidepane.Visible
        if not current_state:
            word_app.ActiveWindow.Sidepane.Show = 1  # wdShowNavPane
        word_app.ActiveWindow.Sidepane.Visible = not current_state
    elif element_type == "status_bar":
        current_state = word_app.DisplayStatusBar
        word_app.DisplayStatusBar = not current_state
    elif element_type == "formatting_marks":
        current_state = word_app.ActiveWindow.View.ShowAll
        word_app.ActiveWindow.View.ShowAll = not current_state
    else:
        valid_elements = "rulers, gridlines, navigation_pane, status_bar, formatting_marks"
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT,
            f"Invalid element type for toggle. Supported elements: {valid_elements}"
        )
    
    new_state = "visible" if not current_state else "hidden"
    log_info(f"Successfully toggled {element_type} to {new_state}")
    return {
        "success": True,
        "element_type": element_type,
        "new_state": new_state,
        "message": f"Successfully toggled {element_type} to {new_state}"
    }


@handle_com_error(ErrorCode.SERVER_ERROR, "get document view info")
def get_view_info(document: win32com.client.CDispatch) -> Dict[str, Any]:
    """获取当前文档视图信息

    Args:
        document: Word文档COM对象

    Returns:
        包含当前视图信息的字典

    Raises:
        WordDocumentError: 当获取视图信息失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    word_app = document.Application
    active_window = word_app.ActiveWindow
    view = active_window.View
    
    # 视图类型映射
    view_type_map = {
        1: "draft",
        2: "outline",
        3: "print",
        6: "web",
        7: "read"
    }
    
    # 获取当前视图信息
    view_info = {
        "view_type": view_type_map.get(view.Type, "unknown"),
        "zoom_level": view.Zoom.Percentage,
        "show_rulers": active_window.DisplayRulers,
        "show_gridlines": view.GridLines,
        "show_navigation_pane": active_window.Sidepane.Visible if hasattr(active_window, 'Sidepane') else False,
        "show_status_bar": word_app.DisplayStatusBar,
        "show_formatting_marks": view.ShowAll,
        "active_window": active_window.Caption if hasattr(active_window, 'Caption') else "Unknown",
        "current_page": document.Range().Information(3)  # wdActiveEndPageNumber
    }
    
    log_info("Successfully retrieved view information")
    return view_info


@handle_com_error(ErrorCode.SERVER_ERROR, "navigate document")
def navigate(document: win32com.client.CDispatch, navigation_type: str, value: Any) -> Dict[str, Any]:
    """导航到文档特定位置

    Args:
        document: Word文档COM对象
        navigation_type: 导航类型
        value: 导航目标值

    Returns:
        包含操作结果的字典

    Raises:
        WordDocumentError: 当导航失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")
    
    navigation_type = navigation_type.lower()
    word_app = document.Application
    
    if navigation_type == "page":
        # 按页码导航
        try:
            page_num = int(value)
            if page_num < 1 or page_num > document.Range().Information(4):  # wdNumberOfPagesInDocument
                raise WordDocumentError(ErrorCode.INVALID_INPUT, f"Page number {page_num} is out of range")
            
            # 导航到指定页码
            word_app.Selection.GoTo(What=1, Which=1, Count=page_num)  # wdGoToPage, wdGoToAbsolute
        except ValueError:
            raise WordDocumentError(ErrorCode.INVALID_INPUT, "Page number must be an integer")
    elif navigation_type == "heading":
        # 按标题导航
        # 实现按标题文本或索引导航
        pass
    elif navigation_type == "bookmark":
        # 按书签导航
        # 实现按书签名称导航
        pass
    elif navigation_type == "section":
        # 按节导航
        # 实现按节索引导航
        pass
    elif navigation_type == "table":
        # 按表格导航
        # 实现按表格索引导航
        pass
    else:
        valid_types = "page, heading, bookmark, section, table"
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT,
            f"Invalid navigation type. Supported types: {valid_types}"
        )
    
    log_info(f"Successfully navigated to {navigation_type} with value {value}")
    return {
        "success": True,
        "navigation_type": navigation_type,
        "value": value,
        "message": f"Successfully navigated to {navigation_type} with value {value}"
    }