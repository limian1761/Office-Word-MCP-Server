"""
Utility functions for the Word Document Server.

This package contains utility modules for file operations and document handling.
"""

from word_document_server.utils.core_utils import (
    CommentEmptyError, CommentError, CommentIndexError, DocumentNotFoundError,
    ElementNotFoundError, ErrorCode, ImageError, ImageFormatError,
    ReplyEmptyError, StyleNotFoundError, WordDocumentError,
    create_document_copy, ensure_docx_extension, format_error,
    format_error_response, get_absolute_path, get_active_document,
    get_color_type, get_doc_path, get_project_root, get_shape_info,
    get_shape_types, handle_error, handle_tool_errors, is_file_writeable,
    log_error, log_info, parse_color_hex, require_active_document_validation,
    standardize_tool_errors, validate_active_document, validate_element_type,
    validate_file_path, validate_formatting, validate_input_params,
    validate_insert_position, validate_locator, validate_operations,
    validate_position)
