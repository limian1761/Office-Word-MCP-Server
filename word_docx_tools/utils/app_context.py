"""
AppContext for managing the Word application instance and the active document state.
"""

import logging
import os
import shutil
import traceback
from typing import Optional, cast

import pythoncom
import win32com.client
from pythoncom import com_error
from win32com.client.dynamic import CDispatch

from ..mcp_service.errors import ErrorCode, WordDocumentError

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
    def get_instance() -> "AppContext":
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

    def _clear_com_cache(self):
        """Clear win32com cache to resolve CLSIDToPackageMap errors"""
        try:
            # 获取gen_py目录路径，使用win32com的内置路径
            import win32com

            gen_path = os.path.join(win32com.__gen_path__, "win32com", "gen_py")
            logger.info(f"Checking win32com cache at: {gen_path}")

            # 检查目录是否存在
            if os.path.exists(gen_path):
                logger.info(f"Clearing win32com cache at: {gen_path}")
                # 尝试删除目录
                shutil.rmtree(gen_path)
                logger.info("Cleared win32com cache successfully")
                return True
            else:
                logger.warning(f"win32com cache directory not found: {gen_path}")
                # 尝试查找其他可能的缓存位置
                temp_gen_path = os.path.join(
                    os.environ.get("TEMP", "C:\\temp"), "gen_py"
                )
                if os.path.exists(temp_gen_path):
                    logger.info(
                        f"Clearing win32com cache at temp location: {temp_gen_path}"
                    )
                    shutil.rmtree(temp_gen_path)
                    logger.info(
                        "Cleared win32com cache from temp location successfully"
                    )
                    return True
                return False
        except PermissionError as e:
            logger.error(f"Permission denied when clearing win32com cache: {e}")
            return False
        except Exception as e:
            logger.error(f"Failed to clear win32com cache: {e}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            return False

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
            logger.info("Attempting to create Word Application instance...")
            # 确保每次都使用新的win32com.client导入
            import win32com.client

            self._word_app = win32com.client.Dispatch("Word.Application")
            logger.info("Started a new Word application instance.")
            return self._word_app
        except AttributeError as e:
            logger.error(f"COM cache error detected: {e}")
            logger.error(f"Error type: {type(e).__name__}")
            logger.error(f"Traceback: {traceback.format_exc()}")

            # Try to clear COM cache and retry
            if self._clear_com_cache():
                try:
                    logger.info(
                        "Retrying Word Application creation after cache clear..."
                    )
                    # 重新导入win32com.client以确保使用清除后的缓存
                    import importlib

                    import win32com.client

                    importlib.reload(win32com.client)

                    self._word_app = win32com.client.Dispatch("Word.Application")
                    logger.info(
                        "Successfully created Word application instance after cache clear."
                    )
                    return self._word_app
                except Exception as retry_e:
                    logger.error(
                        f"Failed to start Word Application after cache clear: {retry_e}"
                    )
                    logger.error(f"Retry error type: {type(retry_e).__name__}")
                    logger.error(f"Retry traceback: {traceback.format_exc()}")
            else:
                logger.error(
                    "Failed to clear COM cache, cannot retry Word Application creation"
                )
            return None
        except Exception as e:
            logger.error(f"Failed to start Word Application: {e}")
            logger.error(f"Error type: {type(e).__name__}")
            logger.error(f"Traceback: {traceback.format_exc()}")
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
                    f"Word COM error while opening document: {file_path}. Details: {e}",
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
