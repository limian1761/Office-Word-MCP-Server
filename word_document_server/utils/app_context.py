"""
AppContext for managing the Word application instance and the active document state.
"""

import logging
from typing import Optional, cast
import os
import traceback
import pythoncom
import win32com.client
from win32com.client.dynamic import CDispatch
from pythoncom import com_error

from word_document_server.utils.core_utils import WordDocumentError, ErrorCode

# Configure logger
logger = logging.getLogger(__name__)

class AppContext:
    """
    Application context that holds the Word application instance and the active document.
    This class acts as a state container. The lifecycle of the Word application
    itself is managed by the server's lifespan manager.
    """

    # Singleton instance
    _instance = None

    def __new__(cls, *args, **kwargs):
        # Create singleton only when there's no instance and no arguments
        if cls._instance is None:
            cls._instance = super(AppContext, cls).__new__(cls)
        # If instance exists but arguments are provided, return a new instance
        elif args or kwargs:
            return super(AppContext, cls).__new__(cls)
        return cls._instance

    @staticmethod
    def get_instance() -> 'AppContext':
        """Get the singleton instance of AppContext, creating it if it doesn't exist"""
        if AppContext._instance is None:
            # Create a parameterless instance and let __init__ handle initialization
            AppContext._instance = AppContext()
        return AppContext._instance

    def __init__(self, word_app: Optional[CDispatch] = None):
        """
        Initialize the AppContext with a running Word application instance.

        Args:
            word_app: An active Word application dispatch object.
        """
        # Prevent duplicate initialization
        if hasattr(self, "_initialized"):
            return

        # Initialize attributes first
        self._temp_word_app: Optional[CDispatch] = None
        self._active_document: Optional[CDispatch] = None
        self._word_app: Optional[CDispatch] = None
        
        # Store the provided word_app if any
        if word_app is not None:
            self._word_app = word_app
        
        self._initialized = True

    def get_word_app(self, create_if_needed: bool = False) -> Optional[CDispatch]:
        """
        Get the Word application instance, optionally creating it if needed.
        
        Args:
            create_if_needed: Whether to create a new Word app instance if one doesn't exist.
            
        Returns:
            The Word application instance or None if not available and not created.
        """
        # Return existing Word app if available
        if self._word_app is not None:
            return self._word_app
            
        # If we shouldn't create and don't have one, return None
        if not create_if_needed:
            return None
            
        # Don't try to attach to existing instances, always create a new one
        try:
            self._word_app = win32com.client.Dispatch("Word.Application")
            logger.info("Started a new Word application instance.")
            return self._word_app
        except Exception as e:
            logger.error(f"Failed to start Word Application: {e}")
            return None

    def get_active_document(self) -> Optional[CDispatch]:
        """Get the current active document."""
        return self._active_document

    def set_active_document(self, doc: Optional[CDispatch]) -> None:
        """Set the current active document."""
        self._active_document = doc

        
    def clear_active_document(self) -> None:
        """Clear the current active document."""
        self._active_document = None

    def open_document(self, file_path: str) -> None:
        """Open a document in the Word application.

        Args:
            file_path: The absolute path to the document to open.
        """

        if file_path:
            try:
                logger.info(f"尝试打开文档: {file_path}")

                # Convert to absolute path for COM
                abs_path = os.path.abspath(file_path)
                logger.info(f"绝对路径: {abs_path}")

                # Check if file exists
                if not os.path.exists(abs_path):
                    raise FileNotFoundError(f"文件不存在: {abs_path}")

                logger.info(f"文件存在: {os.path.exists(abs_path)}")
                logger.info(f"文件大小: {os.path.getsize(abs_path)} bytes")

                # Get or create Word application instance as needed
                word_app = self.get_word_app(create_if_needed=True)
                if word_app is None:
                    raise RuntimeError("无法获取或创建Word应用程序实例")

                # Make the application visible
                word_app.Visible = True

                # Try to open document using main application's Documents collection
                logger.info("尝试访问Documents集合...")
                documents = word_app.Documents
                logger.info(f"Documents对象类型: {type(documents)}")

                # Try to open document
                self._active_document = documents.Open(abs_path)
                logger.info(
                    f"使用主应用程序Documents集合打开文档成功: {self._active_document.Name}"
                )

                # Enable track changes
                if self._active_document and self._active_document.TrackRevisions:
                    self._active_document.TrackRevisions = False
            except pythoncom.com_error as e:
                error_code = e.args[0]
                error_message = e.args[1]
                logger.error(f"COM错误: {error_code}, {error_message}")
                traceback.print_exc()
                # 不再自动关闭文档，让调用者决定是否关闭
                # self.close_document()
                raise WordDocumentError(
                    ErrorCode.DOCUMENT_OPEN_ERROR,
                    f"Word COM error while opening document: {file_path}. Details: {e}"
                )
            except Exception as e:
                logger.error(f"打开文档时发生异常: {str(e)}")
                traceback.print_exc()
                # 不再自动关闭文档，让调用者决定是否关闭
                # self.close_document()
                raise IOError(f"Failed to open document: {file_path}. Error: {e}")
        else:
            # Get or create Word application instance as needed
            word_app = self.get_word_app(create_if_needed=True)
            if word_app is None:
                raise RuntimeError("无法获取或创建Word应用程序实例")
                
            word_app.Visible = True
            self._active_document = word_app.Documents.Add()

    def close_document(self) -> None:
        """Close the currently active document."""
        if self._active_document:
            try:
                # 只关闭活动文档，不退出Word应用
                self._active_document.Close(SaveChanges=0)  # 0 = wdDoNotSaveChanges
            except Exception as e:
                logger.error(f"关闭文档时出错: {str(e)}")
            finally:
                self._active_document = None

        # 不再关闭临时Word实例，避免Word应用意外关闭
        # 这样可以确保Word应用程序保持运行状态
