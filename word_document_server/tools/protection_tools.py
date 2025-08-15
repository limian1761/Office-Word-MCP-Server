"""
Protection tools for Word Document Server.
"""
import os
from typing import Optional
from mcp.server.fastmcp.server import Context
from word_document_server.app import app
from word_document_server.utils.app_context import AppContext
from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension
from word_document_server.utils.com_utils import handle_com_error

# Word protection type constants
wdNoProtection = -1
wdAllowOnlyComments = 1
wdAllowOnlyFormFields = 2
wdAllowOnlyRevisions = 0
wdAllowOnlyReading = 3

@app.tool()
async def protect_document(password: str, protection_type: int = wdAllowOnlyReading, context: Context = None) -> str:
    """Add password protection to a Word document."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        doc.Protect(Type=protection_type, NoReset=True, Password=password)
        doc.Save()
        return f"Document protected successfully."
    except Exception as e:
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def unprotect_document(password: str, context: Context = None) -> str:
    """Remove password protection from a Word document."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        if doc.ProtectionType != wdNoProtection:
            doc.Unprotect(Password=password)
            doc.Save()
            return f"Document unprotected successfully."
        else:
            return f"Document is not protected."
    except Exception as e:
        # Check for incorrect password error (HRESULT: 0x800A141F)
        if hasattr(e, 'excepinfo') and e.excepinfo and e.excepinfo[5] == -2146823137:
             return f"Failed to unprotect document: Incorrect password."
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass

@app.tool()
async def add_digital_signature(signer_name: str) -> str:
    """Add a digital signature to a Word document."""
    # This is a placeholder. Digital signatures with pywin32 are complex
    # and require access to the Windows Certificate Store.
    return "Digital signature functionality is not yet fully implemented."

@app.tool()
async def verify_document(password: Optional[str] = None, context: Context = None) -> str:
    """Verify document protection status."""
    doc = None
    try:
        # 从Context获取活动文档
        app_context = context.request_context.lifespan_context.get(AppContext)
        doc = app_context.get_active_document()
        if not doc:
            return "No active document found."

        protection_type = doc.ProtectionType
        if protection_type == wdNoProtection:
            return "Document is not protected."
        elif protection_type == wdAllowOnlyReading:
            return "Document is protected as read-only."
        elif protection_type == wdAllowOnlyComments:
            return "Document is protected, allowing only comments."
        elif protection_type == wdAllowOnlyFormFields:
            return "Document is protected, allowing only form fields."
        elif protection_type == wdAllowOnlyRevisions:
            return "Document is protected, allowing only revisions."
        else:
            return f"Document has an unknown protection type: {protection_type}"
            
    except Exception as e:
        if hasattr(e, 'excepinfo') and e.excepinfo and e.excepinfo[5] == -2146823137:
             return f"Failed to verify document: Incorrect password provided."
        return handle_com_error(e)
    finally:
        # No need to close the active document
        pass
