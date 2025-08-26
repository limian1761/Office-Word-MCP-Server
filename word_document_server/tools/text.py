import os
import json
from typing import Dict, Any, List, Optional
from mcp.server.fastmcp.server import Context
from word_document_server.core_utils import get_backend_for_tool, mcp_server
from word_document_server.selector import SelectorEngine, AmbiguousLocatorError
from word_document_server.errors import ElementNotFoundError, format_error_response, WordDocumentError
import pywintypes

import logging

# 配置日志记录
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
# 创建控制台处理程序
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
# 创建日志格式
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
console_handler.setFormatter(formatter)
# 添加控制台处理程序到日志记录器
logger.addHandler(console_handler)


@mcp_server.tool()
def insert_paragraph(ctx: Context, locator: Dict[str, Any], text: str, position: str = "after", style: Optional[str] = None) -> str:
    """
    Inserts a new paragraph with the given text relative to the element found by the locator.

    Args:
        locator: The Locator object to find the anchor element.
        text: The text to insert.
        position: "before" or "after" the anchor element.
        style: Optional, the paragraph style name to apply.

    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selector_engine = SelectorEngine()
        selection = selector_engine.select(backend, locator, expect_single=True)
        selection.insert_paragraph(text, position, style)
        backend.document.Save()
        return "Successfully inserted paragraph."
    except ElementNotFoundError as e:
        return f"Error [2002]: No elements found matching the locator: {e}" + ". Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except AmbiguousLocatorError as e:
        return f"Error [3001]: The locator matched multiple elements: {e}" + ". Please refine your locator to match a single element."
    except ValueError as e:
        return f"Error [1001]: Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def delete_element(ctx: Context, locator: Dict[str, Any], password: Optional[str] = None) -> str:
    """
    Deletes the element(s) found by the locator.

    Args:
        locator: The Locator object to find the target element(s) to delete.
        password: Optional password to unlock the document if it's protected.

    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selector_engine = SelectorEngine()
        selection = selector_engine.select(backend, locator)
        
        # Validate that we have elements to delete
        if not selection._elements:
            return "Error [2002]: No elements found matching the locator." + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
        
        element_count = len(selection._elements)
        
        try:
            # 尝试删除元素
            selection.delete()
        except Exception as e:
            # 记录错误日志
            logger.error(f"Error occurred: {str(e)}", exc_info=True)
            # 直接汇报错误原因
            return format_error_response(e)
        
        backend.document.Save()
        return f"Successfully deleted {element_count} element(s)."
    except Exception as e:
        return format_error_response(e)

@mcp_server.tool()
def get_text(ctx: Context, locator: Optional[Dict[str, Any]] = None, start_pos: Optional[int] = None, end_pos: Optional[int] = None) -> str:
    """
    Retrieves the text from all elements found by the locator or from a specific range.

    Args:
        locator: Optional, the Locator object to find the target element(s).
        start_pos: Optional, the start position of the text range.
        end_pos: Optional, the end position of the text range.
        
    Returns:
        A string containing the retrieved text content or an error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        if locator:
            # Use locator to select specific elements
            try:
                selector_engine = SelectorEngine()
                selection = selector_engine.select(backend, locator)
                text = selection.get_text()
            except ElementNotFoundError as e:
                return f"Error [2002]: No elements found matching the locator: {e}" + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
        elif start_pos is not None and end_pos is not None:
            # Get text from a specific range
            if start_pos < 0 or end_pos <= start_pos:
                return "Error [1001]: Invalid range parameters. start_pos must be >= 0 and end_pos must be > start_pos."
            
            max_range = backend.document.Content.End
            if end_pos > max_range:
                return f"Error [1001]: Invalid end_pos. Maximum allowed value is {max_range}."
            
            text = backend.get_text_from_range(start_pos, end_pos)
        else:
            # Get all text from the document
            text = backend.get_all_text()
        
        return text
    except ValueError as e:
        return f"Error [1001]: Invalid parameter: {e}"
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def replace_text(ctx: Context, locator: Dict[str, Any], new_text: str) -> str:
    """
    Replaces the text content of the element(s) found by the locator with new text.

    Args:
        locator: The Locator object to find the target element(s) to replace.
        new_text: The new text to replace the existing content.

    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selector_engine = SelectorEngine()
        selection = selector_engine.select(backend, locator)
        
        if not selection._elements:
            return "Error [2002]: No elements found matching the locator." + " Please try simplifying your locator or use get_document_structure to check the actual document structure."        
        selection.replace_text(new_text)
        backend.document.Save()
        return "Successfully replaced text."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def find_text(ctx: Context, find_text: str, match_case: bool = False, match_whole_word: bool = False, match_wildcards: bool = False, match_synonyms: bool = False, ignore_punct: bool = False, ignore_space: bool = False) -> str:
    """
    Finds occurrences of text in the active document.

    Args:
        find_text: The text to search for.
        match_case: Whether to match case exactly (default: False).
        match_whole_word: Whether to match whole words only (default: False).
        match_wildcards: Whether to allow wildcard characters (default: False).
        match_synonyms: Whether to match synonyms (default: False). Note: This parameter is currently not supported.
        ignore_punct: Whether to ignore punctuation differences (default: False).
        ignore_space: Whether to ignore spacing differences (default: False).
    
    Returns:
        A JSON string containing an array of found matches with their position and context.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    if not find_text:
        return "Error [1001]: Search text cannot be empty."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        # Get the document range
        doc_range = backend.document.Content
        
        # Set up the find object
        find = doc_range.Find
        find.ClearFormatting()
        find.Text = find_text
        find.MatchCase = match_case
        find.MatchWholeWord = match_whole_word
        find.MatchWildcards = match_wildcards
        # Removed MatchSynonyms as it's not supported by Word COM object
        find.IgnorePunct = ignore_punct
        find.IgnoreSpace = ignore_space
        
        # Store all matches
        matches = []
        match_index = 0
        
        # Execute the find until no more matches are found
        while find.Execute():
            match_range = find.Parent
            match_text = match_range.Text
            
            # Get context around the match (10 characters before and after)
            start_pos = max(0, match_range.Start - 10)
            end_pos = min(backend.document.Content.End, match_range.End + 10)
            context_range = backend.document.Range(Start=start_pos, End=end_pos)
            context_text = context_range.Text
            
            # Calculate paragraph index
            paragraph_index = -1
            for i, para in enumerate(backend.document.Paragraphs):
                if para.Range.Start <= match_range.Start and para.Range.End >= match_range.End:
                    paragraph_index = i
                    break
            
            # Add match information to the list
            matches.append({
                "index": match_index,
                "text": match_text,
                "start_pos": match_range.Start,
                "end_pos": match_range.End,
                "paragraph_index": paragraph_index,
                "context_preview": context_text
            })
            
            match_index += 1
            
            # Move past the current match to avoid infinite loops
            if match_range.End < backend.document.Content.End:
                doc_range = backend.document.Range(Start=match_range.End + 1, End=backend.document.Content.End)
                find = doc_range.Find
                find.ClearFormatting()
                find.Text = find_text
                find.MatchCase = match_case
                find.MatchWholeWord = match_whole_word
                find.MatchWildcards = match_wildcards
                # Removed MatchSynonyms as it's not supported by Word COM object
                find.IgnorePunct = ignore_punct
                find.IgnoreSpace = ignore_space
            else:
                break
        
        # Convert to JSON string
        result = {
            "matches_found": len(matches),
            "matches": matches
        }
        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        error_result = {
            "error": f"An unexpected error occurred during text search: {e}",
            "matches_found": 0,
            "matches": []
        }
        return json.dumps(error_result, ensure_ascii=False)


@mcp_server.tool()
def apply_format(ctx: Context, locator: Dict[str, Any], formatting: Dict[str, Any]) -> str:
    """
    Applies specified formatting to the element(s) found by the locator.

    Args:
        locator: The Locator object to find the target element(s).
        formatting: A dictionary of formatting options to apply.
                    Example: {"bold": True, "alignment": "center"}

    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selector_engine = SelectorEngine()
        selection = selector_engine.select(backend, locator)
        
        if not selection._elements:
            return "Error [2002]: No elements found matching the locator." + " Please try simplifying your locator or use get_document_structure to check the actual document structure."        
        selection.apply_format(formatting)
        # Add None check for document
        if backend.document is None:
            raise ValueError("Failed to save document: No active document.")
        backend.document.Save()
        return "Successfully applied formatting."
    except ElementNotFoundError as e:
        return f"Error [2002]: {e}" + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except ValueError as e:
        return f"Error [1001]: Invalid formatting parameter: {e}"
    except Exception as e:
        # Handle common formatting errors with more specific messages
        error_message = str(e)
        if "COM" in str(type(e)) or "Dispatch" in str(type(e)):
            return "Error [8001]: Failed to apply formatting. This may occur if Word is in an unstable state. " + "Try closing and reopening the document, or simplifying your formatting request."
        elif "Invalid request" in error_message:
            return "Error [3001]: Invalid formatting request. Please check that your formatting parameters are valid."
        elif "Unsupported" in error_message:
            return "Error [4002]: Some formatting options are not supported for the selected elements."
        elif "Permission denied" in error_message:
            return "Error [4003]: Permission denied when applying formatting. The document may be read-only or protected."
        return format_error_response(e)


@mcp_server.tool()
def apply_paragraph_style(ctx: Context, locator: Dict[str, Any], style_name: str) -> str:
    """
    Applies a paragraph style to the elements found by the locator.

    Args:
        locator: The Locator object to find the target element(s).
        style_name: The name of the paragraph style to apply.

    Returns:
        A success or error message with validation information.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    if not style_name:
        return "Error [1001]: Style name cannot be empty."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Get all available styles to validate
        available_styles = backend.get_all_styles()
        style_exists = any(style['name'] == style_name for style in available_styles)
        
        if not style_exists:
            # Suggest similar styles if the requested one doesn't exist
            similar_styles = []
            for style in available_styles:
                if style_name.lower() in style['name'].lower():
                    similar_styles.append(style['name'])
            
            if similar_styles:
                return f"Error [1001]: Style '{style_name}' does not exist. Did you mean one of these: {', '.join(similar_styles)}?"
            else:
                return f"Error [1001]: Style '{style_name}' does not exist. Available styles include: {', '.join([style['name'] for style in available_styles[:5]])}..."
        
        # Find the elements to apply the style to
        selector_engine = SelectorEngine()
        selection = selector_engine.select(backend, locator)
        if not selection._elements:
            return "Error [2002]: No elements found matching the locator." + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
        
        # Apply the style and collect results for validation
        applied_count = 0
        for element in selection._elements:
            try:
                original_style = element.Style.NameLocal
                element.Style = style_name
                applied_count += 1
            except Exception as e:
                # Log the error but continue processing
                import logging
                logging.error(f"Failed to apply style '{style_name}' to an element: {str(e)}")
        
        # Save the document - add None check
        if backend.document is None:
            raise ValueError("Failed to save document: No active document.")
        backend.document.Save()
        
        # Return success with validation information
        return f"Successfully applied style '{style_name}' to {applied_count} out of {len(selection._elements)} element(s)."
    except ElementNotFoundError as e:
        return f"Error [2002]: No elements found matching the locator: {e}" + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except Exception as e:
        return f"An unexpected error occurred while applying paragraph style: {e}"


@mcp_server.tool()
def create_bulleted_list(ctx: Context, locator: Dict[str, Any], items: List[str], position: str = "after") -> str:
    """
    Creates a new bulleted list relative to the element found by the locator.

    Args:
        locator: The Locator object to find the anchor element.
        items: A list of strings to become the list items.
        position: "before" or "after" the anchor element.

    Returns:
        A success or error message.
    """
    # Get active document path from session state
    active_doc_path = None
    if hasattr(ctx.session, 'document_state'):
        active_doc_path = ctx.session.document_state.get('active_document_path')
    if not active_doc_path:
        return "Error [2001]: No active document. Please use 'open_document' first."

    # Validate items parameter
    if not isinstance(items, list) or not items:
        return "Error [1001]: Invalid 'items' parameter. Expected a non-empty list of strings."

    # Validate position parameter
    if position not in ["before", "after"]:
        return "Error [1001]: Invalid 'position' parameter. Must be 'before' or 'after'."

    try:
        backend = get_backend_for_tool(ctx, active_doc_path)
        selector_engine = SelectorEngine()
        selection = selector_engine.select(backend, locator, expect_single=True)
        selection.create_bulleted_list(items, position)
        # Add None check for document
        if backend.document is None:
            raise ValueError("Failed to save document: No active document.")
        backend.document.Save()
        return f"Successfully created bulleted list with {len(items)} items."
    except ElementNotFoundError as e:
        return f"Error [2002]: No elements found matching the locator: {e}" + " Please try simplifying your locator or use get_document_structure to check the actual document structure."
    except Exception as e:
        return format_error_response(e)


@mcp_server.tool()
def batch_apply_format(ctx: Context, operations: List[Dict[str, Any]], save_document: bool = True) -> str:
    """
    Applies formatting to multiple elements in a single batch operation.
    This is more efficient than calling apply_format multiple times.

    Args:
        operations: A list of operations, each containing 'locator' and 'formatting' keys.
        save_document: Whether to save the document after applying all formats (default: True).

    Returns:
        A summary of the batch operation results.
    """
    try:
        # Get active document path from session state
        active_doc_path = None
        if hasattr(ctx.session, 'document_state'):
            active_doc_path = ctx.session.document_state.get('active_document_path')
        if not active_doc_path:
            return "Error [2001]: No active document. Please use 'open_document' first."

        # Validate operations parameter
        if not isinstance(operations, list):
            return "Error [1001]: Invalid 'operations' parameter. Expected a list of operations."
        
        if not operations:
            return "Error [1001]: No operations provided. Please provide at least one formatting operation."
        
        backend = get_backend_for_tool(ctx, active_doc_path)
        
        # Track operation results
        results: Dict[str, Any] = {
            'total_operations': len(operations),
            'successful_operations': 0,
            'failed_operations': 0,
            'failures': []
        }
        
        # Process each operation in batch
        for i, operation in enumerate(operations):
            try:
                # Validate operation structure
                if not isinstance(operation, dict) or 'locator' not in operation or 'formatting' not in operation:
                    raise ValueError("Each operation must contain 'locator' and 'formatting' keys.")
                
                # Find the elements to apply formatting to
                selector_engine = SelectorEngine()
                selection = selector_engine.select(backend, operation['locator'])
                if not selection._elements:
                    raise ElementNotFoundError(operation['locator'], f"No elements found for operation {i}.")
                
                # Apply formatting
                selection.apply_format(operation['formatting'])
                results['successful_operations'] += 1
                
            except Exception as e:
                results['failed_operations'] += 1
                results['failures'].append({
                    'operation_index': i,
                    'error': str(e)
                })
                # Continue with next operation
                continue
        
        # Save the document only once after all operations
        if save_document:
            # Add None check for document
            if backend.document is None:
                raise ValueError("Failed to save document: No active document.")
            backend.document.Save()
        
        # Generate summary
        summary = f"Batch formatting completed: {results['successful_operations']} successful, {results['failed_operations']} failed out of {results['total_operations']} operations."
        
        if results['failures']:
            summary += "\n\nFailed operations:\n"
            for failure in results['failures'][:5]:  # Show first 5 failures
                summary += f"- Operation {failure['operation_index']}: {failure['error']}\n"
            
            if len(results['failures']) > 5:
                summary += f"- ... and {len(results['failures']) - 5} more failures."
        
        return summary
    except Exception as e:
        return format_error_response(e)