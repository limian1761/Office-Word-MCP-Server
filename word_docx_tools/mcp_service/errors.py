from enum import Enum
from typing import Any, Dict, List, Optional


class ErrorCode(Enum):
    """Standardized error codes for Word Document Server"""

    # General errors
    SUCCESS = (0, "Operation completed successfully")
    INVALID_INPUT = (1001, "Invalid input parameter")
    NOT_FOUND = (1002, "Requested resource not found")
    PERMISSION_DENIED = (1003, "Permission denied")
    SERVER_ERROR = (1004, "Internal server error")
    UNSUPPORTED_OPERATION = (1005, "Unsupported operation")

    # Document errors
    NO_ACTIVE_DOCUMENT = (2001, "No active document")
    DOCUMENT_OPEN_ERROR = (2002, "Failed to open document")
    DOCUMENT_SAVE_ERROR = (2003, "Failed to save document")
    DOCUMENT_FORMAT_ERROR = (2004, "Invalid document format")
    DOCUMENT_ERROR = (2005, "Document operation error")
    DOCUMENT_CLOSE_ERROR = (2006, "Failed to close document")

    # Object errors
    OBJECT_NOT_FOUND = (3001, "Object not found")
    OBJECT_LOCKED = (3002, "Object is locked")
    OBJECT_TYPE_ERROR = (3003, "Invalid object type")
    PARAGRAPH_SELECTION_FAILED = (3004, "Failed to select paragraph objects")
    OBJECT_DELETION_FAILED = (3005, "Failed to delete object")

    # Style errors
    STYLE_NOT_FOUND = (4001, "Style not found")
    STYLE_APPLY_ERROR = (4002, "Failed to apply style")
    STYLE_CREATION_ERROR = (4003, "Failed to create style")

    # Formatting errors
    FORMATTING_ERROR = (5001, "Formatting error")
    FONT_SETTING_ERROR = (5002, "Failed to set font properties")
    ALIGNMENT_ERROR = (5003, "Failed to set alignment")

    # Image errors
    IMAGE_NOT_FOUND = (6001, "Image not found")
    IMAGE_FORMAT_ERROR = (6002, "Invalid image format")
    IMAGE_LOAD_ERROR = (6003, "Failed to load image")
    IMAGE_INSERTION_ERROR = (6004, "Failed to insert image")
    IMAGE_RESIZE_ERROR = (6005, "Failed to resize image")

    # Table errors
    TABLE_ERROR = (7001, "Table operation error")
    TABLE_CREATION_ERROR = (7002, "Failed to create table")
    CELL_ACCESS_ERROR = (7003, "Failed to access table cell")
    ROW_INSERTION_ERROR = (7004, "Failed to insert row")
    COLUMN_INSERTION_ERROR = (7005, "Failed to insert column")

    # Comment errors
    COMMENT_ERROR = (8001, "Comment operation error")
    COMMENT_INDEX_ERROR = (8002, "Comment index out of range")
    COMMENT_EMPTY_ERROR = (8003, "Comment text cannot be empty")
    REPLY_EMPTY_ERROR = (8004, "Reply text cannot be empty")
    COMMENT_DELETION_ERROR = (8005, "Failed to delete comment")

    # Selector errors
    SELECTOR_ERROR = (9001, "Selector operation error")
    LOCATOR_PARSE_ERROR = (9002, "Failed to parse locator")
    AMBIGUOUS_LOCATOR_ERROR = (9003, "Ambiguous locator - multiple objects found")
    RELATION_ERROR = (9004, "Invalid relation in locator")


class WordDocumentError(Exception):
    """Base exception class for Word Document Server errors"""

    def __init__(
        self,
        error_code: ErrorCode,
        message: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None,
    ):
        self.error_code = error_code
        self.message = message or error_code.value[1]
        self.details = details or {}
        super().__init__(f"[{error_code.value[0]}] {self.message}")

    def to_dict(self) -> Dict[str, Any]:
        """Convert error to dictionary representation"""
        return {
            "error_code": self.error_code.value[0],
            "error_name": self.error_code.name,
            "message": self.message,
            "details": self.details,
        }


class DocumentNotFoundError(WordDocumentError):
    """Raised when a document is not found"""

    def __init__(self, document_path: str, message: Optional[str] = None):
        details = {"document_path": document_path}
        super().__init__(ErrorCode.DOCUMENT_OPEN_ERROR, message, details)


class CommentError(WordDocumentError):
    """Raised when a comment operation fails"""

    def __init__(
        self, message: Optional[str] = None, details: Optional[Dict[str, Any]] = None
    ):
        super().__init__(ErrorCode.COMMENT_ERROR, message, details)


class ObjectNotFoundError(WordDocumentError):
    """Raised when an object is not found"""

    def __init__(self, locator: Dict[str, Any], message: Optional[str] = None):
        details = {"locator": locator}
        super().__init__(ErrorCode.OBJECT_NOT_FOUND, message, details)


class StyleNotFoundError(WordDocumentError):
    """Raised when a style is not found"""

    def __init__(
        self,
        style_name: str,
        message: Optional[str] = None,
        similar_styles: Optional[List[str]] = None,
    ):
        details: Dict[str, Any] = {"style_name": style_name}
        if similar_styles:
            details["similar_styles"] = similar_styles
        super().__init__(ErrorCode.STYLE_NOT_FOUND, message, details)


class ImageError(WordDocumentError):
    """Base exception for image-related errors"""


class ImageFormatError(ImageError):
    """Raised when an image format is invalid"""

    def __init__(self, image_path: str, message: Optional[str] = None):
        details = {"image_path": image_path}
        super().__init__(ErrorCode.IMAGE_FORMAT_ERROR, message, details)


class SelectorError(WordDocumentError):
    """Raised when selector operations fail"""

    def __init__(self, message: str, locator: Optional[Dict[str, Any]] = None):
        details = {}
        if locator:
            details["locator"] = locator
        super().__init__(ErrorCode.SELECTOR_ERROR, message, details)


class AmbiguousLocatorError(SelectorError):
    """Raised when a locator matches multiple objects but only one is expected"""

    def __init__(self, locator: Dict[str, Any], count: int):
        message = f"Ambiguous locator - found {count} objects, expected 1"
        details = {"locator": locator, "found_objects": count}
        super().__init__(message, locator)
        self.error_code = ErrorCode.AMBIGUOUS_LOCATOR_ERROR
        self.details = details


class UnsupportedOperationError(WordDocumentError):
    """Raised when an unsupported operation is requested"""

    def __init__(self, operation: str, reason: Optional[str] = None):
        message = f"Unsupported operation: {operation}"
        if reason:
            message += f" - {reason}"
        details = {"operation": operation, "reason": reason}
        super().__init__(ErrorCode.UNSUPPORTED_OPERATION, message, details)
