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
            # First, try to get an active instance of Word
            self.word_app = win32com.client.GetActiveObject("Word.Application")
            logging.info("Attached to an existing Word application instance.")
        except com_error:
            # If that fails, start a new instance
            try:
                self.word_app = win32com.client.Dispatch("Word.Application")
                logging.info("Started a new Word application instance.")
            except Exception as e:
                raise RuntimeError(f"Failed to start Word Application: {e}") from e

        self.word_app.Visible = self.visible
