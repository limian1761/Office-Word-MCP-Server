"""
COM Backend Adapter Layer for Word Document MCP Server.

This module encapsulates all interactions with the Word COM interface,
providing a clean, Pythonic API for higher-level components. It is designed
to be used as a context manager to ensure proper resource management.
"""

import logging
from typing import Optional

import win32com.client
from pythoncom import com_error  # pylint: disable=no-name-in-module


class WordBackend:
    """
    Backend adapter for interacting with Word COM interface.

    This class is designed to be used as a context manager (`with` statement)
    to ensure that the Word application is properly initialized and cleaned up.
    """

    def __init__(self, visible: bool = False):
        """
        Initialize the Word backend adapter.

        Args:
            visible (bool): Whether to make the Word application visible.
        """
        self.visible = visible
        self.word_app: Optional[win32com.client.CDispatch] = None
        self.document: Optional[win32com.client.CDispatch] = None

    @staticmethod
    async def connect(visible: bool = True) -> "WordBackend":
        """
        Static method to connect to Word application and open/create document.

        Args:
            visible (bool): Whether to make the Word application visible.

        Returns:
            WordBackend: Connected WordBackend instance
        """
        backend = WordBackend(visible)
        await backend.start()
        return backend

    async def disconnect(self) -> None:
        """
        Method to disconnect from Word application and cleanup resources.
        """
        # 只关闭文档，不退出Word应用，避免应用意外关闭
        if self.document:
            try:
                # 关闭当前文档
                self.document.Close(SaveChanges=0)  # 0 = wdDoNotSaveChanges
            except Exception as e:
                logging.warning(f"关闭文档时出错: {e}")
            finally:
                self.document = None

        # 不再退出Word应用，让应用保持运行状态
        self.word_app = None

        logging.info(
            "Word backend cleaned up (document closed, Word application kept running)."
        )

    async def start(self):
        """
        Starts a new Word application instance.
        Opens or creates a document.
        """
        try:
            # Pre-flight check to ensure Word COM server is available
            # Use a temporary instance for checking only, don't store it
            temp_app = win32com.client.Dispatch("Word.Application")
            temp_app.Quit()  # Immediately quit the temporary instance
        except com_error as e:
            raise RuntimeError(f"Word COM server is not available: {e}") from e

        try:
            # Always use get_word_app to get Word application instance
            from ..mcp_service.app_context import AppContext
            app_context = AppContext.get_instance()
            self.word_app = app_context.get_word_app(create_if_needed=True)
            if not self.word_app:
                raise RuntimeError("Failed to get Word application instance through get_word_app()")
            logging.info("Got Word application instance through get_word_app().")
        except Exception as e:
            raise RuntimeError(f"Failed to get Word Application instance: {e}") from e

        self.word_app.Visible = self.visible
