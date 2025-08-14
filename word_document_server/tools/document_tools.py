"""
Document creation and manipulation tools for Word Document Server using COM.
"""
import os
import json
from typing import Dict, List, Optional, Any
from word_document_server.utils import com_utils
from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension, create_document_copy

async def create_document(title: Optional[str] = None, author: Optional[str] = None) -> str:
    """Create a new Word document with optional metadata using COM."""
    # Check if there's an active document
    active_doc = com_utils.get_active_document()
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
        return f"Failed to create document: {str(e)}"

async def get_document_info() -> str:
    """Get information about a Word document using COM."""
    # Check if there's an active document
    active_doc = com_utils.get_active_document()
    if active_doc is not None:
        # Use the active document
        doc = active_doc
    else:
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
        return f"Failed to get document info: {str(e)}"

async def get_document_text(max_chars: Optional[int] = None, 
                           include_tables: bool = True) -> str:
    """Extract all text from a Word document using COM."""
    # Check if there's an active document
    active_doc = com_utils.get_active_document()
    if active_doc is not None:
        # Use the active document
        doc = active_doc
    else:
        return "No active document found"
        
    try:
        text = doc.Content.Text
        if max_chars and len(text) > max_chars:
            text = text[:max_chars] + "..."
        return text
    except Exception as e:
        return f"Failed to extract text: {str(e)}"


async def get_document_outline() -> str:
    """Get the structure of a Word document using COM."""
    # Check if there's an active document
    active_doc = com_utils.get_active_document()
    if active_doc is not None:
        # Use the active document
        doc = active_doc
    else:
        return "No active document found"
    
    try:
        structure = []
        for p in doc.Paragraphs:
            style_name = p.Style.NameLocal
            if style_name.startswith("Heading"):
                try:
                    level = int(style_name.split()[-1])
                    structure.append({"level": level, "text": p.Range.Text.strip()})
                except (ValueError, IndexError):
                    continue # Ignore malformed heading styles
        return json.dumps(structure, indent=2)
    except Exception as e:
        return f"Failed to get document outline: {str(e)}"


async def copy_document(destination_filename: Optional[str] = None) -> str:
    """Create a copy of a Word document using COM."""
    # Check if there's an active document
    active_doc = com_utils.get_active_document()
    if active_doc is not None:
        # Use the active document as source
        doc = active_doc
    else:
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
        return f"Failed to copy document: {str(e)}"}]}}}

async def merge_documents(source_filenames: List[str], add_page_breaks: bool = True) -> str:
    """Merge multiple Word documents into a single document using COM."""
    # 检查是否有活动文档，这里不需要活动文档，直接创建新文档作为目标文档
    # 检查源文件是否存在
    missing_files = [f for f in source_filenames if not os.path.exists(ensure_docx_extension(f))]
    if missing_files:
        return f"Cannot merge. The following source files do not exist: {', '.join(missing_files)}"

    target_doc = None
    try:
        app = com_utils.get_word_app()
        target_doc = app.Documents.Add()
        
        for i, filename in enumerate(source_filenames):
            if i > 0 and add_page_breaks:
                target_doc.Content.InsertBreak(7)  # wdPageBreak = 7
            
            # 在文档末尾插入文件
            target_doc.Content.InsertFile(os.path.abspath(ensure_docx_extension(filename)))

        # 生成目标文件名，基于第一个源文件名
        base, ext = os.path.splitext(source_filenames[0])
        target_filename = f"{base}_merged{ext}"
        target_doc.SaveAs(os.path.abspath(target_filename))
        return f"Successfully merged {len(source_filenames)} documents into {target_filename}"
    except Exception as e:
        return f"Failed to merge documents: {str(e)}"
    finally:
        if target_doc:
            target_doc.Close(SaveChanges=0)


async def get_document_xml_tool() -> str:
    """Get the raw XML structure of a Word document using COM."""
    # Check if there's an active document
    active_doc = com_utils.get_active_document()
    if active_doc is not None:
        # Use the active document
        doc = active_doc
    else:
        return "No active document found"
    
    try:
        return doc.WordOpenXML
    except Exception as e:
        return f"Failed to get document XML: {str(e)}"
    finally:
        # No need to close the active document
        pass

async def get_all_paragraphs_tool() -> str:
    """Gets all paragraph content, returning a JSON with paragraph details."""
    # Check if there's an active document
    active_doc = com_utils.get_active_document()
    if active_doc is not None:
        # Use the active document
        doc = active_doc
    else:
        return "No active document found"
    
    try:
        paragraphs = []
        for i, p in enumerate(doc.Paragraphs):
            paragraphs.append({
                "index": i,
                "text": p.Range.Text.strip(),
                "style": p.Style.NameLocal
            })
        return json.dumps(paragraphs, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"Failed to get all paragraphs: {str(e)}"
    finally:
        # No need to close the active document
        pass

async def get_paragraphs_by_range_tool(start_index: int = 0, end_index: Optional[int] = None) -> str:
    """Gets content of a specified paragraph range."""
    # Check if there's an active document
    active_doc = com_utils.get_active_document()
    if active_doc is not None:
        # Use the active document
        doc = active_doc
    else:
        return "No active document found"
    
    try:
        para_count = doc.Paragraphs.Count
        if end_index is None:
            end_index = para_count
        
        start_index = max(0, start_index)
        end_index = min(end_index, para_count)

        paragraphs = []
        for i in range(start_index, end_index):
            p = doc.Paragraphs(i + 1) # COM is 1-based
            paragraphs.append({
                "index": i,
                "text": p.Range.Text.strip(),
                "style": p.Style.NameLocal
            })
        return json.dumps(paragraphs, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"Failed to get paragraphs by range: {str(e)}"
    finally:
        # No need to close the active document
        pass

async def get_paragraphs_by_page_tool(page_number: int = 1, page_size: int = 100) -> str:
    """Paginates through paragraph content."""
    start_index = (page_number - 1) * page_size
    end_index = start_index + page_size
    return await get_paragraphs_by_range_tool(start_index, end_index)

async def analyze_paragraph_distribution_tool() -> str:
    """Analyzes paragraph distribution and returns statistics."""
    # Check if there's an active document
    active_doc = com_utils.get_active_document()
    if active_doc is not None:
        # Use the active document
        doc = active_doc
    else:
        return "No active document found"
    
    try:
        stats = {
            "total_paragraphs": doc.Paragraphs.Count,
            "style_distribution": {}
        }
        for p in doc.Paragraphs:
            style = p.Style.NameLocal
            stats["style_distribution"][style] = stats["style_distribution"].get(style, 0) + 1
        return json.dumps(stats, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"Failed to analyze paragraph distribution: {str(e)}"
    finally:
        # No need to close the active document
        pass