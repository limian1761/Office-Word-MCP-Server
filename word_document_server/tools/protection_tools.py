"""
Protection tools for Word Document Server using COM.
"""
import os
from typing import Optional
from word_document_server.utils import com_utils
from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension

# Word protection type constants
wdNoProtection = -1
wdAllowOnlyComments = 1
wdAllowOnlyFormFields = 2
wdAllowOnlyRevisions = 0
wdAllowOnlyReading = 3

async def protect_document(filename: str, password: str, protection_type: int = wdAllowOnlyReading) -> str:
    """Add password protection to a Word document using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot protect document: {error_message}"

    doc = None
    try:
        doc = com_utils.open_document(filename)
        doc.Protect(Type=protection_type, NoReset=True, Password=password)
        doc.Save()
        return f"Document {filename} protected successfully."
    except Exception as e:
        return f"Failed to protect document: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def unprotect_document(filename: str, password: str) -> str:
    """Remove password protection from a Word document using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}"

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if doc.ProtectionType != wdNoProtection:
            doc.Unprotect(Password=password)
            doc.Save()
            return f"Document {filename} unprotected successfully."
        else:
            return f"Document {filename} is not protected."
    except Exception as e:
        # Check for incorrect password error (HRESULT: 0x800A141F)
        if hasattr(e, 'excepinfo') and e.excepinfo and e.excepinfo[5] == -2146823137:
             return f"Failed to unprotect document {filename}: Incorrect password."
        return f"Failed to unprotect document: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def add_digital_signature(filename: str, signer_name: str) -> str:
    """Add a digital signature to a Word document using COM."""
    # This is a placeholder. Digital signatures with pywin32 are complex
    # and require access to the Windows Certificate Store.
    return "Digital signature functionality is not yet fully implemented with COM."

async def verify_document(filename: str, password: Optional[str] = None) -> str:
    """Verify document protection status using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    doc = None
    try:
        # If a password is provided, try to open with it.
        # This implicitly checks a read-password.
        app = com_utils.get_word_app()
        abs_path = os.path.abspath(filename)
        doc = app.Documents.Open(abs_path, PasswordDocument=password or "")
        
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
        return f"Failed to verify document: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)