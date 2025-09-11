"""
AppContext for managing the Word application instance and the active document state.
"""

import logging
import os
import shutil
import traceback
from typing import Optional, cast

import pythoncom
from pythoncom import com_error
from win32com.client.dynamic import CDispatch

from .errors import ErrorCode, WordDocumentError

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

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(AppContext, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    @classmethod
    def get_instance(cls) -> "AppContext":
        """Get the singleton instance of AppContext, creating it if it doesn't exist"""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        """
        Initialize the AppContext with a running Word application instance.
        """
        # Prevent duplicate initialization
        if self._initialized:
            return

        # Initialize attributes first
        self._temp_word_app: Optional[CDispatch] = None
        self._active_document: Optional[CDispatch] = None
        self._word_app: Optional[CDispatch] = None

        self._initialized = True

    def set_word_app(self, word_app: Optional[CDispatch] = None) -> None:
        """
        Set the Word application instance for the context.

        Args:
            word_app: An active Word application dispatch object.
        """
        self._word_app = word_app

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

    def close_document(self):
        """关闭当前活动文档"""
        try:
            if self._active_document is not None:
                self._active_document.Close(SaveChanges=0)  # 不保存更改
                self._active_document = None
                return True
            return False
        except Exception as e:
            logger.error(f"Error closing document: {e}")
            return False

    def quit_word_app(self):
        """退出Word应用程序"""
        try:
            if self._word_app is not None:
                self._word_app.Quit()
                self._word_app = None
                self._active_document = None
                return True
            return False
        except Exception as e:
            logger.error(f"Error quitting Word application: {e}")
            return False
