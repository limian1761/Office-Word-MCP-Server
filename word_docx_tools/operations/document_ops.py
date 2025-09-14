"""
Document operations for Word Document MCP Server.

This module contains functions for document-level operations.
"""

import logging
import os
import traceback
import json  # 添加json导入
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Union

import win32com.client
from win32com.client import CDispatch

from ..com_backend.com_utils import handle_com_error, safe_com_call
from ..mcp_service.core_utils import ErrorCode, WordDocumentError, AppContext

logger = logging.getLogger(__name__)


# === Document Management Operations ===


@handle_com_error(ErrorCode.DOCUMENT_ERROR, "create document")
def create_document(
    word_app: Optional[CDispatch] = None,
    template_path: Optional[str] = None,
    visible: bool = True,
) -> CDispatch:
    """
    Create a new Word document.

    Args:
        word_app: Optional Word application instance. If provided, uses this instance to create the document.
        template_path: Optional path to a template file.
        visible: Whether to make the Word application visible.

    Returns:
        The created document COM object.
    """
    try:
        # Use provided Word application instance or raise error if not provided
        if not word_app:
            raise WordDocumentError(
                ErrorCode.SERVER_ERROR,
                "Word application instance must be provided through get_word_app()",
            )
        # Try to set visibility with error handling
        try:
            word_app.Visible = visible
        except AttributeError:
            # Ignore if Visible property cannot be set
            logger.warning("Could not set Word application visibility")
            pass

        # Create the document
        if template_path:
            logger.info(f"Creating document from template: {template_path}")
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"Template file not found: {template_path}")
            doc = word_app.Documents.Add(Template=template_path)
        else:
            logger.info("Creating blank document")
            doc = word_app.Documents.Add()

        logger.info("Successfully created new document")
        
        # 触发DocumentContext的更新，实现文档与DocumentContext的实时映射
        try:
            app_context = AppContext.get_instance()
            app_context.set_active_document(doc)
            app_context.on_document_opened()
        except Exception as e:
            logger.warning(f"Failed to update DocumentContext after creating document: {str(e)}")
            
        return doc

    except Exception as e:
        logger.error(f"Failed to create document: {str(e)}")
        logger.error(f"Error type: {type(e).__name__}")
        logger.error(f"Traceback: {traceback.format_exc()}")
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, f"Failed to create document: {str(e)}"
        )


@handle_com_error(ErrorCode.DOCUMENT_ERROR, "open document")
def open_document(
    word_app: CDispatch,
    file_path: str,
    visible: bool = True,
    password: Optional[str] = None,
) -> CDispatch:
    """
    Open a Word document.

    Args:
        word_app: Word application instance (must be provided through get_word_app()).
        file_path: Path to the Word document file.
        visible: Whether to make the Word application visible.
        password: Optional password for protected documents.

    Returns:
        The opened document COM object.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Document file not found: {file_path}")

    # Word application instance must be provided
    if not word_app:
        raise WordDocumentError(
            ErrorCode.SERVER_ERROR,
            "Word application instance must be provided through get_word_app()",
        )
    word_app.Visible = visible

    # Open the document
    if password:
        doc = word_app.Documents.Open(file_path, PasswordDocument=password)
    else:
        doc = word_app.Documents.Open(file_path)
    
    # 触发DocumentContext的更新，实现文档与DocumentContext的实时映射
    try:
        app_context = AppContext.get_instance()
        app_context.set_active_document(doc)
        app_context.on_document_opened()
    except Exception as e:
        logger.warning(f"Failed to update DocumentContext after opening document: {str(e)}")

    return doc


@handle_com_error(ErrorCode.DOCUMENT_ERROR, "close document")
def close_document(document: CDispatch, save_changes: bool = True) -> bool:
    """
    Closes a Word document.

    Args:
        document: The document COM object to close.
        save_changes: Whether to save changes before closing.

    Returns:
        True if the document was closed successfully.
    """
    try:
        if document:
            document.Close(SaveChanges=save_changes)
            logger.info("Document closed successfully")
            return True
        else:
            logger.warning("No document to close")
            return False
    except Exception as e:
        logger.error(f"Failed to close document: {str(e)}")
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, f"Failed to close document: {str(e)}"
        )


@handle_com_error(ErrorCode.DOCUMENT_ERROR, "save document")
def save_document(
    document: CDispatch,
    file_path: Optional[str] = None,
) -> str:
    """
    Save a Word document.

    Args:
        document: The document COM object to save.
        file_path: Optional file path to save to. If not provided, saves to the current path.

    Returns:
        The file path where the document was saved.
    """
    try:
        if not document:
            raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No document to save")

        if file_path:
            # 确保file_path是字符串类型
            if not isinstance(file_path, (str, bytes, os.PathLike)):
                file_path = str(file_path)
                logger.warning(f"Converted file_path to string: {file_path}")
            
            # 检查文件路径是否为空
            if not file_path:
                raise WordDocumentError(ErrorCode.INVALID_INPUT, "File path cannot be empty")
            
            # 尝试使用SaveAs2方法，如果失败则回退到SaveAs方法
            try:
                # 确保路径使用正确的分隔符
                file_path = os.path.normpath(file_path)
                
                # 确保目标目录存在
                directory = os.path.dirname(file_path)
                if directory and not os.path.exists(directory):
                    os.makedirs(directory)
                    logger.info(f"Created directory: {directory}")
                
                # 对文件名进行额外的验证和处理，以确保符合Word的要求
                file_name = os.path.basename(file_path)
                
                # 检查文件名是否包含Word不支持的字符
                invalid_chars = '\\/:*?"<>|'
                for char in invalid_chars:
                    if char in file_name:
                        # 替换不支持的字符
                        file_name = file_name.replace(char, '_')
                        logger.warning(f"Replaced invalid characters in filename: {file_name}")
                
                # 如果文件名被修改，更新完整路径
                if file_name != os.path.basename(file_path):
                    file_path = os.path.join(directory, file_name)
                
                # 确保文件扩展名为.docx或.doc
                base_name, ext = os.path.splitext(file_name)
                if not ext or ext.lower() not in ['.docx', '.doc']:
                    file_path = os.path.join(directory, base_name + '.docx')
                    logger.warning(f"Added .docx extension to file path: {file_path}")
                
                try:
                    # 尝试使用SaveAs2（Word 2010及以上版本）
                    document.SaveAs2(FileName=file_path)
                except AttributeError:
                    # 如果SaveAs2不可用，尝试使用SaveAs（旧版Word）
                    logger.warning("SaveAs2 method not available, falling back to SaveAs")
                    document.SaveAs(FileName=file_path)
                
                logger.info(f"Document saved to: {file_path}")
            except Exception as e:
                logger.error(f"Failed to save document with path: {file_path}")
                raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, f"Failed to save document: {str(e)}")
        else:
            document.Save()
            file_path = document.FullName
            if file_path is None:
                raise WordDocumentError(ErrorCode.SERVER_ERROR, "Could not determine file path after saving")
            logger.info(f"Document saved to: {file_path}")

        return file_path
    except Exception as e:
        logger.error(f"Failed to save document: {str(e)}")
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, f"Failed to save document: {str(e)}")


# === Document Structure Operations ===


def count_objects_by_type(document: CDispatch, object_type: str) -> int:
    """统计特定类型的元素数量

    Args:
        document: Word文档COM对象
        object_type: 元素类型 ("paragraphs", "tables", "images", etc.)

    Returns:
        元素数量
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        count = 0

        if object_type == "paragraphs":
            count = document.Paragraphs.Count
        elif object_type == "tables":
            count = document.Tables.Count
        elif object_type == "images" or object_type == "inline_shapes":
            count = (
                document.InlineShapes.Count if hasattr(document, "InlineShapes") else 0
            )
        elif object_type == "sections":
            count = document.Sections.Count
        elif object_type == "pages":
            # 近似页数计算
            count = document.Range().Information(
                4
            )  # 4 corresponds to wdNumberOfPagesInDocument
        else:
            raise ValueError(f"Unsupported object type: {object_type}")

        return count

    except Exception as e:
        logger.error(f"Error in count_objects_by_type: {e}")
        raise WordDocumentError(
            ErrorCode.OBJECT_TYPE_ERROR,
            f"Failed to count objects of type '{object_type}': {str(e)}",
        )


@handle_com_error(ErrorCode.SERVER_ERROR, "get document outline")
def get_document_outline(document: CDispatch) -> str:
    """获取文档大纲信息，通过段落级别来判断是否是大纲

    Args:
        document: Word文档COM对象

    Returns:
        包含文档大纲层级结构的JSON字符串
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        outline_structure = []
        
        # 遍历所有段落，提取大纲信息
        for i in range(1, document.Paragraphs.Count + 1):
            paragraph = document.Paragraphs(i)
            
            # 获取段落样式信息
            style_name = paragraph.Style.NameLocal if hasattr(paragraph.Style, 'NameLocal') else ""
            
            # 获取段落级别（大纲级别）
            outline_level = 0
            if hasattr(paragraph, 'OutlineLevel'):
                outline_level = paragraph.OutlineLevel
            
            # 根据用户需求，1~9为标题样式，10为正文
            if 1 <= outline_level <= 9:
                outline_item = {
                    "index": i,
                    "text": paragraph.Range.Text.strip(),
                    "outline_level": outline_level,
                    "style_name": style_name,
                    "page_number": paragraph.Range.Information(3) if hasattr(paragraph.Range, 'Information') else 0  # wdActiveEndPageNumber
                }
                outline_structure.append(outline_item)

        # 构建层次化的大纲结构
        hierarchical_outline = build_hierarchical_outline_by_level(outline_structure)

        return json.dumps({
            "outline_items": hierarchical_outline,
            "total_headings": len(outline_structure),
            "document_statistics": {
                "paragraphs": document.Paragraphs.Count,
                "tables": document.Tables.Count if hasattr(document, 'Tables') and document.Tables is not None else 0,
                "sections": document.Sections.Count if hasattr(document, 'Sections') and document.Sections is not None else 0,
                "pages": document.Range().Information(4) if hasattr(document.Range(), 'Information') else 0  # wdNumberOfPagesInDocument
            }
        }, ensure_ascii=False, indent=2)

    except Exception as e:
        logger.error(f"Error in get_document_outline: {e}")
        raise WordDocumentError(
            ErrorCode.SERVER_ERROR, f"Failed to get document outline: {str(e)}")


def build_hierarchical_outline_by_level(heading_list: list) -> list:
    """将扁平的标题列表构建为层次化的大纲结构
    
    Args:
        heading_list: 标题信息列表
    
    Returns:
        层次化的大纲结构列表
    """
    if not heading_list:
        return []
        
    # 按照outline_level排序
    sorted_headings = sorted(heading_list, key=lambda x: (x["outline_level"], x["index"]))
    
    hierarchical_outline: list = []
    current_stack = [hierarchical_outline]
    current_level = 0
    
    for heading in sorted_headings:
        level = heading["outline_level"]
        
        # 创建当前标题的条目，复制所有属性并添加children字段
        current_item = heading.copy()
        current_item["children"] = []
        
        # 调整堆栈以匹配当前层级
        while level <= current_level and len(current_stack) > 1:
            current_stack.pop()
            current_level -= 1
        
        # 如果当前层级比堆栈中的最高层级大1，则添加新层级
        if level > current_level:
            current_stack[-1].append(current_item)
            current_stack.append(current_item["children"])
            current_level = level
        else:
            # 否则，添加到当前层级
            current_stack[-1].append(current_item)
    
    return hierarchical_outline


def find_and_replace_text(
    document: CDispatch,
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    match_whole_word: bool = False,
) -> int:
    """在文档中查找并替换文本

    Args:
        document: Word文档COM对象
        find_text: 要查找的文本
        replace_text: 替换文本
        match_case: 是否匹配大小写
        match_whole_word: 是否匹配整个单词

    Returns:
        替换的次数
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        if not find_text:
            raise ValueError("Find text cannot be empty.")

        # 使用Word的查找和替换功能
        find = document.Content.Find
        find.ClearFormatting()
        find.Text = find_text
        find.Replacement.ClearFormatting()
        find.Replacement.Text = replace_text
        find.Forward = True
        find.Wrap = 1  # wdFindContinue
        find.Format = False
        find.MatchCase = match_case
        find.MatchWholeWord = match_whole_word
        find.MatchWildcards = False
        find.MatchSoundsLike = False
        find.MatchAllWordForms = False

        # 执行替换所有
        count = 0
        while find.Execute(Replace=2):  # 2 = wdReplaceOne
            count += 1

        return count

    except Exception as e:
        logger.error(f"Error in find_and_replace_text: {e}")
        raise WordDocumentError(
            ErrorCode.SERVER_ERROR, f"Failed to find and replace text: {str(e)}"
        )
