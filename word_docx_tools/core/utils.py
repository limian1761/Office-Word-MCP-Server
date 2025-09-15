"""Utility functions for Word Document MCP Server.

This module contains general utility functions used across the application.
"""

import logging
import os
import shutil
from typing import Any, Dict, List, Optional, Tuple

from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession

from .app_context import AppContext
from .errors import ErrorCode, WordDocumentError

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("word_docx_tools.log"), logging.StreamHandler()],
)
logger = logging.getLogger("WordDocumentServer")


def get_active_document(ctx: Context[ServerSession, "AppContext"]) -> Any:
    """获取当前活动的文档对象

    Args:
        ctx: MCP上下文对象

    Returns:
        活动文档对象

    Raises:
        WordDocumentError: 当没有活动文档时抛出
    """
    try:
        active_doc = ctx.request_context.lifespan_context.get_active_document()
        if not active_doc:
            raise WordDocumentError(
                ErrorCode.DOCUMENT_ERROR,
                "No active document found. Please open a document first.",
            )
        return active_doc
    except Exception as e:
        if isinstance(e, WordDocumentError):
            raise
        raise WordDocumentError(
            ErrorCode.DOCUMENT_ERROR, f"Failed to get active document: {str(e)}"
        )


def log_info(message: str) -> None:
    """记录信息日志

    Args:
        message: 日志信息
    """
    logger.info(message)


def log_error(message: str, exc_info: bool = False) -> None:
    """记录错误日志

    Args:
        message: 错误信息
        exc_info: 是否包含异常堆栈信息
    """
    if exc_info:
        logger.error(message, exc_info=True)
    else:
        logger.error(message)


def log_warning(message: str) -> None:
    """记录警告日志

    Args:
        message: 警告信息
    """
    logger.warning(message)


def is_file_writeable(filepath: str) -> Tuple[bool, str]:
    """
    Check if a file can be written to.

    Args:
        filepath: Path to the file

    Returns:
        Tuple of (is_writeable, error_message)
    """
    is_writeable = False
    error_message = ""

    # If file doesn't exist, check if directory is writeable
    if not os.path.exists(filepath):
        directory = os.path.dirname(filepath) or "."
        if not os.path.exists(directory):
            error_message = f"Directory {directory} does not exist"
        elif not os.access(directory, os.W_OK):
            error_message = f"Directory {directory} is not writeable"
        else:
            is_writeable = True
    else:
        # If file exists, check if it's writeable
        if not os.access(filepath, os.W_OK):
            error_message = f"File {filepath} is not writeable (permission denied)"
        else:
            # Try to open the file for writing to see if it's locked
            try:
                with open(filepath, "a", encoding="utf-8"):
                    pass
                is_writeable = True
            except (IOError, OSError) as e:
                error_message = f"File {filepath} is not writeable: {str(e)}"

    return is_writeable, error_message


def create_document_copy(
    source_path: str, dest_path: Optional[str] = None
) -> Tuple[bool, str, Optional[str]]:
    """
    Create a copy of a document.

    Args:
        source_path: Path to the source document
        dest_path: Optional path for the new document. If not provided, will use
            source_path + '_copy.docx'

    Returns:
        Tuple of (success, message, new_filepath)
    """
    if not os.path.exists(source_path):
        return False, f"Source document {source_path} does not exist", None

    if not dest_path:
        # Generate a new filename if not provided
        base, ext = os.path.splitext(source_path)
        dest_path = f"{base}_copy{ext}"

    try:
        # Simple file copy
        shutil.copy2(source_path, dest_path)
        return True, f"Document copied to {dest_path}", dest_path
    except (IOError, shutil.Error, OSError) as e:
        return False, f"Failed to copy document: {str(e)}", None


def ensure_docx_extension(filename: str) -> str:
    """
    Ensure filename has .docx extension.

    Args:
        filename: The filename to check

    Returns:
        Filename with .docx extension
    """
    if not filename.endswith(".docx"):
        return filename + ".docx"
    return filename


def get_absolute_path(relative_path: str) -> str:
    """
    Convert a relative path to absolute path.

    Args:
        relative_path: The relative path to convert

    Returns:
        Absolute path
    """
    # Get absolute path based on current working directory
    return os.path.abspath(relative_path)


def get_project_root() -> str:
    """
    Get the project root directory.

    Returns:
        Absolute path to project root
    """
    # Get the directory of the current file
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # Go up two levels to reach project root
    return os.path.abspath(os.path.join(current_dir, "..", ".."))


def get_doc_path(doc_filename: str) -> str:
    """
    Get absolute path to a document in the docs directory.

    Args:
        doc_filename: Filename of the document

    Returns:
        Absolute path to the document
    """
    project_root = get_project_root()
    return os.path.join(project_root, "docs", doc_filename)


def get_shape_types():
    """获取Word文档中支持的形状类型映射."""
    return {
        1: "Picture",  # wdInlineShapePicture
        2: "LinkedPicture",  # wdInlineShapeLinkedPicture
        3: "Chart",  # wdInlineShapeChart
        4: "Diagram",  # wdInlineShapeDiagram
        5: "OLEControlObject",  # wdInlineShapeOLEControlObject
        6: "OLEObject",  # wdInlineShapeOLEObject
        7: "ActiveXControl",  # wdInlineShapeActiveXControl
        8: "SmartArt",  # wdInlineShapeSmartArt
        9: "3DModel",  # wdInlineShape3DModel
    }


def get_color_type(color_code: int) -> str:
    """
    Convert color type code to human-readable string.

    Args:
        color_code: Color type code from Word COM interface

    Returns:
        Human-readable color type
    """
    # Word picture color type constants
    color_types = {
        0: "Color",  # msoPictureColorTypeColor
        1: "Grayscale",  # msoPictureColorTypeGrayscale
        2: "BlackAndWhite",  # msoPictureColorTypeBlackAndWhite
        3: "Watermark",  # msoPictureColorTypeWatermark
    }
    return color_types.get(color_code, "Unknown")


def get_shape_info(shape, index):
    """获取形状的基本信息."""
    return {
        "index": index,  # 0-based index
        "type": (
            get_shape_types().get(shape.Type, "Unknown")
            if hasattr(shape, "Type")
            else "Unknown"
        ),
        "width": shape.Width if hasattr(shape, "Width") else 0,
        "height": shape.Height if hasattr(shape, "Height") else 0,
    }


def parse_color_hex(color: str) -> Optional[int]:
    """解析十六进制颜色代码为RGB整数."""
    color = color.lstrip("#")
    if len(color) == 6:
        try:
            return int(color, 16)
        except ValueError:
            return None
    return None