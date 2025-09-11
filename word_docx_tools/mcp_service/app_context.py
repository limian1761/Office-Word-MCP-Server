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
        # Return existing Word app if available and validate it's still functional
        if self._word_app is not None:
            if self._validate_word_app(self._word_app):
                logger.debug("Returning existing valid Word application instance.")
                return self._word_app
            else:
                logger.warning("Existing Word application instance is invalid, will attempt to create a new one.")
                self._word_app = None

        # If we shouldn't create and don't have one, return None
        if not create_if_needed:
            return None

        # Try multiple connection methods with retries
        for attempt in range(3):  # Try up to 3 times
            try:
                logger.info(f"Attempt {attempt+1}/3 to create Word Application instance...")
                
                # 1. Try the standard Dispatch method first
                result = self._create_word_app_with_dispatch()
                if result:
                    return result
                
                # 2. If standard method fails, clear cache and try again
                if attempt == 0:
                    if self._clear_com_cache():
                        logger.info("Retrying after COM cache clear...")
                        result = self._create_word_app_with_dispatch(reload_module=True)
                        if result:
                            return result
                    else:
                        logger.warning("Failed to clear COM cache, moving to next method")
                
                # 3. Try with DispatchEx which creates a separate process
                logger.info("Trying with DispatchEx...")
                result = self._create_word_app_with_dispatchex()
                if result:
                    return result
                
                # 4. Try with early binding
                logger.info("Trying with early binding (gencache)...")
                result = self._create_word_app_with_early_binding()
                if result:
                    return result
                
                # If all methods failed for this attempt, wait before retrying
                if attempt < 2:  # Don't wait after the last attempt
                    import time
                    wait_time = 1.5 ** attempt  # Exponential backoff
                    logger.info(f"All methods failed, waiting {wait_time:.2f}s before retry...")
                    time.sleep(wait_time)
                
            except Exception as e:
                logger.error(f"Unexpected error during attempt {attempt+1}: {e}")
                logger.error(f"Error type: {type(e).__name__}")
                logger.error(f"Traceback: {traceback.format_exc()}")
                
        logger.error("Failed to create Word Application instance after multiple attempts.")
        return None
        
    def _create_word_app_with_dispatch(self, reload_module: bool = False) -> Optional[CDispatch]:
        """Create Word app using standard Dispatch method."""
        try:
            # 确保每次都使用新的win32com.client导入
            import win32com.client
            
            if reload_module:
                import importlib
                importlib.reload(win32com.client)
                logger.info("Reloaded win32com.client module.")
            
            self._word_app = win32com.client.Dispatch("Word.Application")
            logger.info("Successfully created Word application instance with Dispatch.")
            return self._word_app
        except Exception as e:
            logger.warning(f"Dispatch method failed: {e}")
            return None
            
    def _create_word_app_with_dispatchex(self) -> Optional[CDispatch]:
        """Create Word app using DispatchEx method (creates a separate process)."""
        try:
            import win32com.client
            
            self._word_app = win32com.client.DispatchEx("Word.Application")
            logger.info("Successfully created Word application instance with DispatchEx.")
            return self._word_app
        except Exception as e:
            logger.warning(f"DispatchEx method failed: {e}")
            return None
            
    def _create_word_app_with_early_binding(self) -> Optional[CDispatch]:
        """Create Word app using early binding (gencache)."""
        try:
            import win32com.client.gencache
            
            # Ensure we have the gen_py files for Word
            win32com.client.gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 7)
            self._word_app = win32com.client.gencache.Dispatch("Word.Application")
            logger.info("Successfully created Word application instance with early binding.")
            return self._word_app
        except Exception as e:
            logger.warning(f"Early binding method failed: {e}")
            return None
            
    def _validate_word_app(self, word_app: CDispatch) -> bool:
        """Validate that the Word application instance is still functional."""
        if not word_app:
            return False
            
        try:
            # Simple property access to verify the COM object is still alive
            app_version = word_app.Version
            logger.debug(f"Word application validation successful, version: {app_version}")
            return True
        except Exception as e:
            logger.warning(f"Word application validation failed: {e}")
            return False

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
