"""
Other Operations for Word Document MCP Server.

This module contains miscellaneous operations that don't fit into other categories.
"""

import logging
import os
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Union

import win32com.client

from ..com_backend.com_utils import handle_com_error
from ..mcp_service.core_utils import (ErrorCode, WordDocumentError, log_error,
                                      log_info)

if TYPE_CHECKING:
    from win32com.client import CDispatch


@handle_com_error(ErrorCode.SERVER_ERROR, "compare documents")
def compare_documents(
    original_document: win32com.client.CDispatch,
    compared_document_path: str,
    save_comparison_path: Optional[str] = None,
) -> Dict[str, Any]:
    # 由于装饰器已处理异常，函数内部的try-except可以简化，但保持参数验证
    """比较两个文档

    Args:
        original_document: 原始Word文档COM对象
        compared_document_path: 要比较的文档路径
        save_comparison_path: 保存比较结果的路径（可选）

    Returns:
        包含比较结果的字典

    Raises:
        WordDocumentError: 当比较文档失败时抛出
    """
    if not original_document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not compared_document_path:
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT, "Compared document path cannot be empty"
        )

    if not os.path.exists(compared_document_path):
        raise WordDocumentError(
            ErrorCode.NOT_FOUND,
            f"Compared document not found: {compared_document_path}",
        )

    # 打开要比较的文档
    word_app = original_document.Application
    compared_doc = word_app.Documents.Open(compared_document_path)

    # 执行比较
    comparison_result = word_app.CompareDocuments(
        original_document,
        compared_doc,
        CompareFormatting=True,
        CompareCaseChanges=True,
        CompareWhitespace=True,
        CompareTables=True,
        CompareHeaders=True,
        CompareFootnotes=True,
        CompareTextboxes=True,
        CompareFields=True,
        CompareComments=True,
        CompareMoves=True,
        RevisedAuthor="MCP Server",
        IgnoreAllComparisonWarnings=True,
    )

    # 统计差异
    differences_count = 0
    if hasattr(comparison_result, "Revisions"):
        differences_count = comparison_result.Revisions.Count

    # 保存比较结果（如果指定了路径）
    saved_path = None
    if save_comparison_path:
        comparison_result.SaveAs2(save_comparison_path)
        saved_path = save_comparison_path

    # 记录原始文档名称
    original_document_name = original_document.Name

    # 关闭比较文档
    compared_doc.Close(SaveChanges=False)

    # 重新激活原始文档
    for doc in word_app.Documents:
        if doc.Name == original_document_name:
            doc.Activate()
            break

    log_info(f"Successfully compared documents, found {differences_count} differences")
    return {
        "differences_count": differences_count,
        "compared_document_name": compared_doc.Name,
        "saved_path": saved_path,
    }


@handle_com_error(ErrorCode.SERVER_ERROR, "convert document format")
def convert_document_format(
    document: win32com.client.CDispatch, output_path: str, format_type: str
) -> bool:
    """将文档转换为指定格式

    Args:
        document: Word文档COM对象
        output_path: 输出文件路径
        format_type: 输出格式类型

    Returns:
        成功转换为指定格式的布尔值

    Raises:
        WordDocumentError: 当转换文档格式失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not output_path:
        raise WordDocumentError(ErrorCode.INVALID_INPUT, "Output path cannot be empty")

    format_constants = {
        "txt": 0,  # wdFormatText
        "doc": 0,  # wdFormatDocument
        "docx": 12,  # wdFormatXMLDocument
        "rtf": 6,  # wdFormatRTF
        "odt": 23,  # wdFormatOpenDocumentText
        "pdf": 17,  # wdExportFormatPDF
    }

    if format_type not in format_constants:
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT, f"Invalid format type: {format_type}"
        )

    if not hasattr(document, "SaveAs2"):
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, "Document does not support SaveAs2 method"
        )

    document.SaveAs2(output_path, format_constants[format_type])
    log_info(f"Successfully converted document to {format_type}: {output_path}")
    return True


@handle_com_error(ErrorCode.SERVER_ERROR, "export to pdf")
def export_to_pdf(document: win32com.client.CDispatch, output_path: str) -> bool:
    """将文档导出为PDF

    Args:
        document: Word文档COM对象
        output_path: 输出PDF文件路径

    Returns:
        成功导出为PDF的布尔值

    Raises:
        WordDocumentError: 当导出PDF失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not output_path:
        raise WordDocumentError(ErrorCode.INVALID_INPUT, "Output path cannot be empty")

    # 确保输出路径以.pdf结尾
    if not output_path.lower().endswith(".pdf"):
        output_path += ".pdf"

    # 设置导出参数
    export_format = 17  # wdExportFormatPDF

    # 导出为PDF
    document.ExportAsFixedFormat(
        OutputFileName=output_path,
        ExportFormat=export_format,
        OpenAfterExport=False,
        OptimizeFor=0,  # wdExportOptimizeForPrint
        Range=0,  # wdExportAllDocument
        Item=0,  # wdExportDocumentContent
        IncludeDocProps=True,
        KeepIRM=True,
        CreateBookmarks=1,  # wdExportCreateHeadingBookmarks
        DocStructureTags=True,
        BitmapMissingFonts=True,
        UseISO19005_1=True,
    )

    log_info(f"Successfully exported document to PDF: {output_path}")
    return True


@handle_com_error(ErrorCode.SERVER_ERROR, "print document")
def print_document(
    document: win32com.client.CDispatch,
    printer_name: Optional[str] = None,
    copies: int = 1,
) -> bool:
    """打印文档

    Args:
        document: Word文档COM对象
        printer_name: 打印机名称（可选）
        copies: 打印份数（可选）

    Returns:
        成功打印的布尔值

    Raises:
        WordDocumentError: 当打印文档失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if not hasattr(document, "Application") or document.Application is None:
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, "Document does not have an Application instance"
        )

    word_app = document.Application
    if printer_name:
        document.PrintOut(Printer=printer_name)
    else:
        document.PrintOut()

    log_info(f"Successfully printed document {copies} times")
    return True


@handle_com_error(ErrorCode.SERVER_ERROR, "protect document")
def protect_document(
    document: win32com.client.CDispatch,
    password: str,
    protection_type: str = "readonly",
) -> bool:
    """保护文档

    Args:
        document: Word文档COM对象
        password: 保护密码
        protection_type: 保护类型，可以是"readonly", "comments", "tracked_changes", "forms"

    Returns:
        成功保护的布尔值

    Raises:
        WordDocumentError: 当保护文档失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    if protection_type not in ["readonly", "comments", "tracked_changes", "forms"]:
        raise WordDocumentError(
            ErrorCode.INVALID_INPUT, f"Invalid protection type: {protection_type}"
        )

    protection_constants = {
        "readonly": 1,  # wdAllowOnlyReading
        "comments": 2,  # wdAllowOnlyComments
        "tracked_changes": 3,  # wdAllowOnlyRevisions
        "forms": 4,  # wdAllowOnlyFormFields
    }

    # 保护文档
    document.Protect(
        Type=protection_constants[protection_type], NoReset=True, Password=password
    )

    log_info(f"Successfully protected document with {protection_type} protection")
    return True


@handle_com_error(ErrorCode.SERVER_ERROR, "get document statistics")
def get_document_statistics(document: win32com.client.CDispatch) -> Dict[str, Any]:
    """获取文档的全面统计信息

    Args:
        document: Word文档COM对象

    Returns:
        包含文档统计信息的字典

    Raises:
        WordDocumentError: 当获取文档统计信息失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # 获取文档基本统计信息
    stats = {
        "paragraphs": (
            document.Paragraphs.Count
            if hasattr(document, "Paragraphs") and document.Paragraphs is not None
            else 0
        ),
        "tables": (
            document.Tables.Count
            if hasattr(document, "Tables") and document.Tables is not None
            else 0
        ),
        "inline_shapes": (
            document.InlineShapes.Count
            if hasattr(document, "InlineShapes") and document.InlineShapes is not None
            else 0
        ),
        "sections": (
            document.Sections.Count
            if hasattr(document, "Sections") and document.Sections is not None
            else 0
        ),
        "comments": (
            document.Comments.Count
            if hasattr(document, "Comments") and document.Comments is not None
            else 0
        ),
        "words": (
            document.Words.Count
            if hasattr(document, "Words") and document.Words is not None
            else 0
        ),
        "characters": (
            document.Characters.Count
            if hasattr(document, "Characters") and document.Characters is not None
            else 0
        ),
        "pages": document.Range().Information(4),  # wdNumberOfPagesInDocument
        "bookmarks": document.Bookmarks.Count if hasattr(document, "Bookmarks") else 0,
    }

    log_info(f"Successfully retrieved document statistics")
    return stats


@handle_com_error(ErrorCode.SERVER_ERROR, "unprotect document")
def unprotect_document(
    document: win32com.client.CDispatch, password: Optional[str] = None
) -> bool:
    """解除文档保护

    Args:
        document: Word文档COM对象
        password: 文档密码，如果没有密码则为None

    Returns:
        操作是否成功

    Raises:
        WordDocumentError: 当解除文档保护失败时抛出
    """
    if not document:
        raise WordDocumentError(ErrorCode.DOCUMENT_ERROR, "No active document found")

    # 检查文档是否受保护
    # 注意：ProtectionType == -1 表示没有保护，其他值表示有保护
    if document.ProtectionType != -1:  # wdNoProtection
        # 解除保护
        document.Unprotect(Password=password)

        log_info("Successfully unprotected document")
        return True
    else:
        log_info("Document was not protected")
        return False
