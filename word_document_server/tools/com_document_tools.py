"""
Word COM接口工具 - 可选的增强功能
提供比python-docx更精确的Word文档处理能力
"""
import asyncio
from typing import Dict, Any, Optional
import sys

# 根据平台导入相应的工具
if sys.platform == "win32":
    from ..utils.com_document_utils import (
        get_document_properties_com,
        get_all_paragraphs_com,
        get_paragraphs_by_range_com,
        get_paragraphs_by_page_com,
        analyze_paragraph_distribution_com,
        is_com_available
    )
    from ..utils.com_utils import get_active_document
else:
    # 非Windows平台的占位符
    def is_com_available():
        return False

    def get_document_properties_com(*args, **kwargs):
        return {"error": "COM接口仅在Windows平台可用"}

    def get_all_paragraphs_com(*args, **kwargs):
        return {"error": "COM接口仅在Windows平台可用"}

    def get_paragraphs_by_range_com(*args, **kwargs):
        return {"error": "COM接口仅在Windows平台可用"}

    def get_paragraphs_by_page_com(*args, **kwargs):
        return {"error": "COM接口仅在Windows平台可用"}

    def analyze_paragraph_distribution_com(*args, **kwargs):
        return {"error": "COM接口仅在Windows平台可用"}


async def get_document_properties_com_tool() -> Dict[str, Any]:
    """
    使用Word COM接口获取文档属性（Windows专用）
    
    提供比python-docx更精确的文档属性，包括Word内置统计信息
    
    Returns:
        包含详细文档属性的字典
    """
    try:
        # 检查COM接口可用性
        if not is_com_available():
            return {
                "error": "Word COM接口不可用",
                "suggestion": "请确保已安装Microsoft Word并运行Windows系统"
            }
        
        # 仅使用活动文档
        active_doc = get_active_document()
        if not active_doc:
            return {"error": "没有活动文档"}
        
        # 在异步环境中运行COM操作
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(
            None, 
            get_document_properties_com, 
            active_doc.FullName
        )
        
        return {"data": result}
        
    except Exception as e:
        return {"error": f"获取文档属性失败: {str(e)}"}


async def get_all_paragraphs_com_tool() -> Dict[str, Any]:
    """
    使用Word COM接口获取所有段落（Windows专用）
    
    提供更精确的段落格式信息，包括Word特有的格式属性
    
    Returns:
        包含所有段落详细信息的字典
    """
    try:
        if not is_com_available():
            return {
                "error": "Word COM接口不可用",
                "suggestion": "请确保已安装Microsoft Word并运行Windows系统"
            }
        
        # 仅使用活动文档
        active_doc = get_active_document()
        if not active_doc:
            return {"error": "没有活动文档"}
        
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(
            None, 
            get_all_paragraphs_com, 
            active_doc.FullName
        )
        
        return {"data": result}
        
    except Exception as e:
        return {"error": f"获取段落失败: {str(e)}"}


async def get_paragraphs_by_range_com_tool(
    start_index: int = 0, 
    end_index: Optional[int] = None
) -> Dict[str, Any]:
    """
    使用Word COM接口获取指定范围段落（Windows专用）
    
    提供更精确的段落边界和格式信息
    
    Args:
        start_index: 起始段落索引（包含）
        end_index: 结束段落索引（不包含）
    
    Returns:
        包含指定范围段落信息的字典
    """
    try:
        if not is_com_available():
            return {
                "error": "Word COM接口不可用",
                "suggestion": "请确保已安装Microsoft Word并运行Windows系统"
            }
        
        # 仅使用活动文档
        active_doc = get_active_document()
        if not active_doc:
            return {"error": "没有活动文档"}
        
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(
            None, 
            get_paragraphs_by_range_com, 
            active_doc.FullName,
            start_index, 
            end_index
        )
        
        return {"data": result}
        
    except Exception as e:
        return {"error": f"获取段落范围失败: {str(e)}"}


async def get_paragraphs_by_page_com_tool(
    page_number: int = 1, 
    page_size: int = 100
) -> Dict[str, Any]:
    """
    使用Word COM接口分页获取段落（Windows专用）
    
    提供更精确的分页处理，支持Word的页面概念
    
    Args:
        page_number: 页码（从1开始）
        page_size: 每页段落数量
    
    Returns:
        包含分页段落信息的字典
    """
    try:
        if not is_com_available():
            return {
                "error": "Word COM接口不可用",
                "suggestion": "请确保已安装Microsoft Word并运行Windows系统"
            }
        
        # 仅使用活动文档
        active_doc = get_active_document()
        if not active_doc:
            return {"error": "没有活动文档"}
        
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(
            None, 
            get_paragraphs_by_page_com, 
            active_doc.FullName,
            page_number, 
            page_size
        )
        
        return {"data": result}
        
    except Exception as e:
        return {"error": f"分页获取段落失败: {str(e)}"}


async def analyze_paragraph_distribution_com_tool() -> Dict[str, Any]:
    """
    使用Word COM接口分析段落分布（Windows专用）
    
    提供更精确的统计分析，包括Word内置统计
    
    Returns:
        包含段落统计分析信息的字典
    """
    try:
        if not is_com_available():
            return {
                "error": "Word COM接口不可用",
                "suggestion": "请确保已安装Microsoft Word并运行Windows系统"
            }
        
        # 仅使用活动文档
        active_doc = get_active_document()
        if not active_doc:
            return {"error": "没有活动文档"}
        
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(
            None, 
            analyze_paragraph_distribution_com, 
            active_doc.FullName
        )
        
        return {"data": result}
        
    except Exception as e:
        return {"error": f"分析段落分布失败: {str(e)}"}


async def check_com_availability_tool() -> Dict[str, Any]:
    """
    检查Word COM接口的可用性
    
    Returns:
        包含COM接口状态信息的字典
    """
    try:
        available = is_com_available()
        return {
            "available": available,
            "platform": sys.platform,
            "message": "Word COM接口可用" if available else "Word COM接口不可用",
            "requirements": [
                "Windows操作系统",
                "Microsoft Word已安装",
                "pywin32库已安装"
            ]
        }
    except Exception as e:
        return {
            "available": False,
            "error": str(e),
            "platform": sys.platform
        }