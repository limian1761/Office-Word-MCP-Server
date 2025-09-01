"""
Document operations for Word Document MCP Server.

This module contains functions for document-level operations.
"""

import os
import logging
import traceback
from typing import Any, Dict, List, Optional, Union, TYPE_CHECKING
from win32com.client import CDispatch
import win32com.client

from ..com_backend.com_utils import handle_com_error, safe_com_call
from ..utils.core_utils import ErrorCode, WordDocumentError

logger = logging.getLogger(__name__)


# === Document Management Operations ===


@handle_com_error(ErrorCode.DOCUMENT_ERROR, "create document")
def create_document(
    word_app: Optional[CDispatch] = None,
    visible: bool = True,
    template_path: Optional[str] = None
) -> CDispatch:
    """
    Creates a new Word document.

    Args:
        word_app: Optional existing Word application object. If not provided, creates a new one.
        visible: Whether to make the Word application visible.
        template_path: Optional path to a template file to use for the new document.

    Returns:
        The created document COM object.
    """
    try:
        # Create or use existing Word application instance
        if not word_app:
            logger.info("Creating new Word application instance for document creation")
            word_app = win32com.client.Dispatch('Word.Application')
            logger.info("Successfully created Word application instance")
        
        assert word_app is not None
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
        return doc
        
    except Exception as e:
        logger.error(f"Failed to create document: {str(e)}")
        logger.error(f"Error type: {type(e).__name__}")
        logger.error(f"Traceback: {traceback.format_exc()}")
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR,
            f"Failed to create document: {str(e)}"
        )


@handle_com_error(ErrorCode.DOCUMENT_ERROR, "open document")
def open_document(
    document: Optional[CDispatch],
    file_path: str,
    visible: bool = True,
    password: Optional[str] = None,
) -> CDispatch:
    """
    Opens a Word document.

    Args:
        document: Optional existing document object. If provided, uses its parent Word application.
        file_path: Path to the Word document file.
        visible: Whether to make the Word application visible.
        password: Optional password for protected documents.

    Returns:
        The opened document COM object.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Document file not found: {file_path}")

    # Try to get Word application from existing document, otherwise create new instance
    word_app = None
    if document:
        try:
            word_app = document.Application
        except Exception:
            pass
            
    if not word_app:
        word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = visible

    # Open the document
    if password:
        doc = word_app.Documents.Open(file_path, PasswordDocument=password)
    else:
        doc = word_app.Documents.Open(file_path)

    return doc


@handle_com_error(ErrorCode.DOCUMENT_ERROR, "close document")
def close_document(
    document: CDispatch, save_changes: bool = True
) -> bool:
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
    document: CDispatch, file_path: Optional[str] = None
) -> str:
    """
    Saves a Word document.

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
            document.SaveAs2(FileName=file_path)
            logger.info(f"Document saved to: {file_path}")
        else:
            document.Save()
            file_path = document.FullName
            logger.info(f"Document saved to: {file_path}")

        return file_path
    except Exception as e:
        logger.error(f"Failed to save document: {str(e)}")
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, f"Failed to save document: {str(e)}"
        )


# === Document Structure Operations ===


def count_elements_by_type(
    document: CDispatch, element_type: str
) -> int:
    """统计特定类型的元素数量

    Args:
        document: Word文档COM对象
        element_type: 元素类型 ("paragraphs", "tables", "images", etc.)

    Returns:
        元素数量
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        count = 0

        if element_type == "paragraphs":
            count = document.Paragraphs.Count
        elif element_type == "tables":
            count = document.Tables.Count
        elif element_type == "images" or element_type == "inline_shapes":
            count = (
                document.InlineShapes.Count if hasattr(document, "InlineShapes") else 0
            )
        elif element_type == "sections":
            count = document.Sections.Count
        elif element_type == "pages":
            # 近似页数计算
            count = document.Range().Information(
                4
            )  # 4 corresponds to wdNumberOfPagesInDocument
        else:
            raise ValueError(f"Unsupported element type: {element_type}")

        return count

    except Exception as e:
        logger.error(f"Error in count_elements_by_type: {e}")
        raise WordDocumentError(
            ErrorCode.ELEMENT_TYPE_ERROR,
            f"Failed to count elements of type '{element_type}': {str(e)}",
        )


def get_document_structure(document: CDispatch) -> str:
    """获取文档结构概览

    Args:
        document: Word文档COM对象

    Returns:
        包含文档结构信息的JSON字符串
    """
    try:
        if not document:
            raise RuntimeError("No document open.")

        structure = {
            "paragraphs": document.Paragraphs.Count,
            "tables": document.Tables.Count,
            "inline_shapes": (
                document.InlineShapes.Count if hasattr(document, "InlineShapes") else 0
            ),
            "sections": document.Sections.Count,
            "comments": document.Comments.Count,
            "words": document.Words.Count,
            "characters": document.Characters.Count,
        }

        import json

        return json.dumps(structure, ensure_ascii=False, indent=2)

    except Exception as e:
        logger.error(f"Error in get_document_structure: {e}")
        raise WordDocumentError(
            ErrorCode.SERVER_ERROR, f"Failed to get document structure: {str(e)}"
        )


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
