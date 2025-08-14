"""
Utilities for interacting with the Word application via COM.
"""
import win32com.client
import pythoncom
import os

# Global variable to hold the Word application instance
_word_app = None
# Global variable to hold the current active document
_active_document = None

def get_word_app():
    """
    Gets or creates a new Word Application instance.
    This function ensures that only one instance of the Word application is running.
    """
    global _word_app
    if _word_app is None:
        try:
            # CoInitialize is necessary for multi-threaded environments
            pythoncom.CoInitialize()
            _word_app = win32com.client.Dispatch("Word.Application")
            _word_app.Visible = False  # Run in the background
        except Exception as e:
            # Handle cases where Word might not be installed
            raise RuntimeError("Microsoft Word application could not be started. Please ensure it is installed.") from e
    return _word_app

def quit_word_app():
    """
    Quits the Word application instance if it is running.
    """
    global _word_app, _active_document
    if _word_app is not None:
        try:
            _word_app.Quit(SaveChanges=0) # WdSaveOptions.wdDoNotSaveChanges = 0
        except Exception:
            # Ignore errors on quit
            pass
        finally:
            _word_app = None
            _active_document = None
            pythoncom.CoUninitialize()

def open_document(filename: str):
    """
    Opens a Word document and returns the document object.
    """
    app = get_word_app()
    abs_path = os.path.abspath(filename)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"The file '{filename}' does not exist at '{abs_path}'.")
    try:
        doc = app.Documents.Open(abs_path)
        return doc
    except Exception as e:
        raise IOError(f"Failed to open document '{filename}'.") from e

def create_document(filename: str):
    """
    Creates a new Word document and saves it.
    """
    app = get_word_app()
    abs_path = os.path.abspath(filename)
    try:
        doc = app.Documents.Add()
        doc.SaveAs(abs_path)
        return doc
    except Exception as e:
        raise IOError(f"Failed to create document '{filename}'.") from e

def get_active_document():
    """
    Gets the currently active document.
    """
    global _active_document
    if _active_document is not None:
        return _active_document
    app = get_word_app()
    if app.Documents.Count > 0:
        return app.ActiveDocument
    return None

def set_active_document(doc):
    """
    Sets the current active document.
    """
    global _active_document
    _active_document = doc
