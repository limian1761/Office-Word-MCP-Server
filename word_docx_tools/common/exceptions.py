from ..mcp_service.errors import ErrorCode, WordDocumentError


class DocumentContextError(WordDocumentError):
    """Raised when there's an error with document context operations"""
    
    def __init__(self,
                 message: str = "Document context error",
                 error_code: ErrorCode = ErrorCode.SERVER_ERROR,
                 details: dict = None):
        """
        初始化文档上下文错误
        
        Args:
            message: 错误消息
            error_code: 错误代码，默认为SERVER_ERROR
            details: 附加错误详情
        """
        super().__init__(error_code, message, details or {})