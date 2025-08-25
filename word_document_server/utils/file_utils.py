"""
File utility functions for Word Document Server.
"""
import os
import shutil
from typing import Optional, Tuple


def check_file_writeable(filepath: str) -> Tuple[bool, str]:
    """
    Check if a file can be written to.
    
    Args:
        filepath: Path to the file
        
    Returns:
        Tuple of (is_writeable, error_message)
    """
    # If file doesn't exist, check if directory is writeable
    if not os.path.exists(filepath):
        directory = os.path.dirname(filepath)
        # If no directory is specified (empty string), use current directory
        if directory == '':
            directory = '.'
        if not os.path.exists(directory):
            return False, f"Directory {directory} does not exist"
        if not os.access(directory, os.W_OK):
            return False, f"Directory {directory} is not writeable"
        return True, ""
    
    # If file exists, check if it's writeable
    if not os.access(filepath, os.W_OK):
        return False, f"File {filepath} is not writeable (permission denied)"
    
    # Try to open the file for writing to see if it's locked
    try:
        with open(filepath, 'a'):
            pass
        return True, ""
    except IOError as e:
        return False, f"File {filepath} is not writeable: {str(e)}"
    except Exception as e:
        return False, f"Unknown error checking file permissions: {str(e)}"


def create_document_copy(source_path: str, dest_path: Optional[str] = None) -> Tuple[bool, str, Optional[str]]:
    """
    Create a copy of a document.
    
    Args:
        source_path: Path to the source document
        dest_path: Optional path for the new document. If not provided, will use source_path + '_copy.docx'
        
    Returns:
        Tuple of (success, message, new_filepath)
    """
    if not os.path.exists(source_path):
        return False, f"Source document {source_path} does not exist", None
    
    if not dest_path:
        # Generate a new filename if not provided
        base, ext = os.path.splitext(source_path)
        dest_path = f"{base}_copy{ext}"
    
    try:
        # Simple file copy
        shutil.copy2(source_path, dest_path)
        return True, f"Document copied to {dest_path}", dest_path
    except Exception as e:
        return False, f"Failed to copy document: {str(e)}", None


def ensure_docx_extension(filename: str) -> str:
    """
    Ensure filename has .docx extension.
    
    Args:
        filename: The filename to check
    
    Returns:
        Filename with .docx extension
    """
    if not filename.endswith('.docx'):
        return filename + '.docx'
    return filename


def get_absolute_path(relative_path: str) -> str:
    """
    Convert a relative path to absolute path.
    
    Args:
        relative_path: The relative path to convert
    
    Returns:
        Absolute path
    """
    # Get absolute path based on current working directory
    return os.path.abspath(relative_path)


def get_project_root() -> str:
    """
    Get the project root directory.
    
    Returns:
        Absolute path to project root
    """
    # Get the directory of the current file
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # Go up two levels to reach project root
    return os.path.abspath(os.path.join(current_dir, '..', '..'))


def get_doc_path(doc_filename: str) -> str:
    """
    Get absolute path to a document in the docs directory.
    
    Args:
        doc_filename: Filename of the document
    
    Returns:
        Absolute path to the document
    """
    project_root = get_project_root()
    return os.path.join(project_root, 'docs', doc_filename)
