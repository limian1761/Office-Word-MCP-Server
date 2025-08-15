"""
Document creation and manipulation tools for Word Document Server.
"""
import os
import json
import glob
from datetime import datetime
from typing import Dict, List, Optional, Any
from mcp.server.fastmcp.server import Context
from word_document_server.app import app
from word_document_server.utils.app_context import AppContext
from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension


def handle_com_error(e: Exception) -> str:
    """Handle COM-related errors consistently across all functions."""
    if "-2147417848" in str(e) or "disconnected" in str(e).lower():
        return "COM Error: The Word application has been disconnected. Please restart the server."
    return str(e)


@app.tool()
async def create_document(title: Optional[str] = None, author: Optional[str] = None, context: Context = None) -> str:
    """Create a new Word document with optional metadata."""
    # Check if there's an active document
    app_context = context.request_context.lifespan_context.get(AppContext)
    active_doc = app_context.get_active_document()
    if active_doc is not None:
        # Use the active document instead of creating a new one
        doc = active_doc
    else:
        return "No active document found"
    
    try:
        if title:
            doc.BuiltInDocumentProperties("Title").Value = title
        if author:
            doc.BuiltInDocumentProperties("Author").Value = author
        doc.Save()
        return f"Document {doc.Name} created successfully"
    except Exception as e:
        return f"Failed to create document: {handle_com_error(e)}"


@app.tool()
async def get_document_outline(context: Context) -> str:
    """Get the structure of a Word document."""
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        structure = []
        # Add debug information
        total_paragraphs = doc.Paragraphs.Count
        heading_paragraphs = 0
        
        for p in doc.Paragraphs:
            style_name = p.Style.NameLocal
            # Count paragraphs with heading styles (support both English and Chinese headings)
            if style_name.startswith("Heading") or style_name.startswith("标题"):
                heading_paragraphs += 1
                try:
                    # Extract level from either English or Chinese heading style
                    if style_name.startswith("Heading"):
                        level = int(style_name.split()[-1])
                    elif style_name.startswith("标题"):
                        # Extract numeric part from Chinese heading style (e.g., "标题 1" -> 1)
                        level = int(''.join(filter(str.isdigit, style_name)))
                    structure.append({"level": level, "text": p.Range.Text.strip()})
                except (ValueError, IndexError):
                    # Include malformed heading styles in debug info
                    structure.append({"level": "unknown", "text": p.Range.Text.strip(), "style_name": style_name})
                    continue
        
        # Add debug information to the response
        debug_info = {
            "total_paragraphs": total_paragraphs,
            "heading_paragraphs": heading_paragraphs,
            "structure": structure
        }
        return json.dumps(debug_info, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"Failed to get document outline: {handle_com_error(e)}"

@app.tool()
async def copy_document(context: Context, destination_filename: Optional[str] = None) -> str:
    """Create a copy of a Word document."""
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    if destination_filename:
        destination_filename = ensure_docx_extension(destination_filename)
    else:
        base, ext = os.path.splitext(doc.FullName)
        destination_filename = f"{base}_copy{ext}"

    is_writeable, error_message = check_file_writeable(destination_filename)
    if not is_writeable:
        return f"Cannot create copy: {error_message}"

    try:
        doc.SaveAs(os.path.abspath(destination_filename))
        return f"Document copied to {destination_filename}"
    except Exception as e:
        return f"Failed to copy document: {handle_com_error(e)}"

@app.tool()
async def merge_documents(context: Context, documents: List[str], output_filename: Optional[str] = None) -> str:
    """Merge multiple Word documents using COM."""
    app_context: AppContext = context.request_context.lifespan_context
    if not documents:
        return "No documents provided for merging"
    
    target_doc = app_context.get_active_document()
    if target_doc is None:
        return "No active document found to merge into"
    
    if output_filename:
        output_filename = ensure_docx_extension(output_filename)
    else:
        output_filename = "merged_document.docx"
    
    is_writeable, error_message = check_file_writeable(output_filename)
    if not is_writeable:
        return f"Cannot create merged document: {error_message}"
    
    try:
        # Process each document in the list
        for doc_path in documents:
            if not os.path.exists(doc_path):
                continue # Skip missing files
            
            # Insert a page break before each document except the first
            if doc_path != documents[0]:
                target_doc.Content.InsertBreak(7) # wdSectionBreakNextPage = 7
            
            # Insert the content of the document
            source_doc = target_doc.Application.Documents.Open(os.path.abspath(doc_path))
            source_range = source_doc.Content
            source_range.Copy()
            target_doc.Content.Paste()
            source_doc.Close()
        
        # Save the merged document
        target_doc.SaveAs(os.path.abspath(output_filename))
        return f"Documents merged successfully into {output_filename}"
    except Exception as e:
        return f"Failed to merge documents: {handle_com_error(e)}"
    finally:
        if target_doc:
            target_doc.Close(SaveChanges=0)

@app.tool()
async def get_document_xml_tool(context: Context) -> str:
    """Get the raw XML structure of a Word document using COM."""
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        xml_content = doc.WordOpenXML
        return xml_content
    except Exception as e:
        return f"Failed to get document XML: {handle_com_error(e)}"
    finally:
        # No need to close the active document
        pass

@app.tool()
async def get_all_paragraphs_tool(context: Context) -> str:
    """Get all paragraphs from a Word document."""
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        paragraphs = []
        for i, p in enumerate(doc.Paragraphs):
            paragraphs.append({
                "index": i,
                "text": p.Range.Text.strip(),
                "style": p.Style.NameLocal,
                "word_count": p.Range.Words.Count,
            })
        return json.dumps(paragraphs, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"Failed to get paragraphs: {handle_com_error(e)}"

@app.tool()
async def get_paragraphs_by_range_tool(context: Context, start_index: int = 0, end_index: Optional[int] = None) -> str:
    """Get paragraphs by range."""
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        total_paragraphs = doc.Paragraphs.Count
        
        if start_index < 0:
            start_index = 0
        if end_index is None or end_index > total_paragraphs:
            end_index = total_paragraphs
        
        if start_index >= end_index:
            return "Invalid range: start_index must be less than end_index"
        
        paragraphs = []
        for i in range(start_index + 1, end_index + 1):
            p = doc.Paragraphs(i)
            paragraphs.append({
                "index": i - 1,
                "text": p.Range.Text.strip(),
                "style": p.Style.NameLocal,
                "word_count": p.Range.Words.Count,
            })
        
        return json.dumps(paragraphs, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"Failed to get paragraphs by range: {handle_com_error(e)}"

@app.tool()
async def get_paragraphs_by_page_tool(context: Context, page_number: int = 1, page_size: int = 10) -> str:
    """Get paragraphs by page."""
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        total_paragraphs = doc.Paragraphs.Count
        
        if page_number < 1:
            page_number = 1
        if page_size < 1:
            page_size = 10
        
        start_index = (page_number - 1) * page_size
        end_index = start_index + page_size
        
        if start_index >= total_paragraphs:
            return f"Page {page_number} is beyond the document's paragraph count"
        
        if end_index > total_paragraphs:
            end_index = total_paragraphs
        
        paragraphs = []
        for i in range(start_index + 1, end_index + 1):
            p = doc.Paragraphs(i)
            paragraphs.append({
                "index": i - 1,
                "text": p.Range.Text.strip(),
                "style": p.Style.NameLocal,
                "word_count": p.Range.Words.Count,
            })
        
        return json.dumps(paragraphs, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"Failed to get paragraphs by page: {handle_com_error(e)}"

@app.tool()
async def analyze_paragraph_distribution(context: Context) -> str:
    """Analyze paragraph distribution."""
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        total_paragraphs = doc.Paragraphs.Count
        style_counts = {}
        word_counts = []
        
        for p in doc.Paragraphs:
            style_name = p.Style.NameLocal
            style_counts[style_name] = style_counts.get(style_name, 0) + 1
            word_counts.append(p.Range.Words.Count)
        
        if word_counts:
            avg_words = sum(word_counts) / len(word_counts)
            min_words = min(word_counts)
            max_words = max(word_counts)
        else:
            avg_words = min_words = max_words = 0
        
        result = {
            "total_paragraphs": total_paragraphs,
            "average_words_per_paragraph": round(avg_words, 2),
            "min_words_in_paragraph": min_words,
            "max_words_in_paragraph": max_words,
            "style_distribution": style_counts
        }
        
        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"Failed to analyze paragraph distribution: {handle_com_error(e)}"


@app.tool()
async def get_active_documents_info(context: Context) -> str:
    """Get information about all active Word documents."""
    app_context: AppContext = context.request_context.lifespan_context
    try:
        documents_info = []
        
        if app_context.word_app.Documents.Count == 0:
            return json.dumps([], indent=2, ensure_ascii=False)
        
        for i in range(app_context.word_app.Documents.Count):
            try:
                doc = app_context.word_app.Documents(i + 1)
                info = {
                    "name": doc.Name,
                    "path": doc.FullName,
                    "saved": doc.Saved,
                    "word_count": doc.Words.Count,
                    "paragraph_count": doc.Paragraphs.Count,
                    "page_count": doc.ComputeStatistics(2),  # wdStatisticPages = 2
                }
                documents_info.append(info)
            except Exception as doc_e:
                documents_info.append({
                    "error": f"Failed to access document {i+1}: {str(doc_e)}"
                })
        
        return json.dumps(documents_info, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"Failed to get active documents info: {handle_com_error(e)}"

@app.tool()
async def get_document_info(context: Context) -> str:
    """Get information about a Word document."""
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        properties = {
            "Title": doc.BuiltInDocumentProperties("Title").Value,
            "Author": doc.BuiltInDocumentProperties("Author").Value,
            "Subject": doc.BuiltInDocumentProperties("Subject").Value,
            "Keywords": doc.BuiltInDocumentProperties("Keywords").Value,
            "Comments": doc.BuiltInDocumentProperties("Comments").Value,
            "Last saved by": doc.BuiltInDocumentProperties("Last author").Value,
            "Revision number": doc.BuiltInDocumentProperties("Revision number").Value,
            "Word count": doc.Words.Count,
            "Paragraph count": doc.Paragraphs.Count,
            "Page count": doc.ComputeStatistics(2), # wdStatisticPages = 2
        }
        return json.dumps(properties, indent=2)
    except Exception as e:
        return f"Failed to get document info: {handle_com_error(e)}"

@app.tool()
async def list_opened_documents(context: Context) -> str:
    """List all currently opened Word documents."""
    app_context: AppContext = context.request_context.lifespan_context
    try:
        documents = []
        for i in range(app_context.word_app.Documents.Count):
            try:
                doc = app_context.word_app.Documents(i + 1)
                documents.append({
                    "name": doc.Name,
                    "path": doc.FullName,
                    "saved": doc.Saved
                })
            except Exception as doc_e:
                documents.append({
                    "error": f"Failed to access document {i+1}: {str(doc_e)}"
                })
        return json.dumps(documents, indent=2)
    except Exception as e:
        return f"Failed to list opened documents: {handle_com_error(e)}"

@app.tool()
async def set_active_document(context: Context, document_name: str) -> str:
    """Set the active document by name."""
    app_context: AppContext = context.request_context.lifespan_context
    try:
        # Try to find the document by name
        for i in range(app_context.word_app.Documents.Count):
            doc = app_context.word_app.Documents(i + 1)
            if doc.Name == document_name:
                doc.Activate()
                app_context.set_active_document(doc)
                return f"Document '{document_name}' is now active"
        
        # If not found by name, try to find by full path
        for i in range(app_context.word_app.Documents.Count):
            doc = app_context.word_app.Documents(i + 1)
            if doc.FullName == document_name:
                doc.Activate()
                return f"Document '{document_name}' is now active"
        
        return f"Document '{document_name}' not found"
    except Exception as e:
        return f"Failed to set active document: {handle_com_error(e)}"