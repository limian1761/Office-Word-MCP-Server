"""
Document utility functions for Word Document Server.
"""
import json
from typing import Dict, List, Any, Optional, Tuple
from docx import Document


def get_document_properties(doc_path: str) -> Dict[str, Any]:
    """Get properties of a Word document."""
    import os
    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}
    
    try:
        doc = Document(doc_path)
        core_props = doc.core_properties
        
        return {
            "title": core_props.title or "",
            "author": core_props.author or "",
            "subject": core_props.subject or "",
            "keywords": core_props.keywords or "",
            "created": str(core_props.created) if core_props.created else "",
            "modified": str(core_props.modified) if core_props.modified else "",
            "last_modified_by": core_props.last_modified_by or "",
            "revision": core_props.revision or 0,
            "page_count": len(doc.sections),
            "word_count": sum(len(paragraph.text.split()) for paragraph in doc.paragraphs),
            "paragraph_count": len(doc.paragraphs),
            "table_count": len(doc.tables)
        }
    except Exception as e:
        return {"error": f"Failed to get document properties: {str(e)}"}


def extract_document_text(doc_path: str, max_chars: Optional[int] = None, 
                         include_tables: bool = True) -> str:
    """
    Extract all text from a Word document with optional size limits.
    
    Args:
        doc_path: Path to the Word document
        max_chars: Maximum characters to return (None for no limit)
        include_tables: Whether to include table content
    """
    import os
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"
    
    try:
        doc = Document(doc_path)
        text_parts = []
        total_chars = 0
        
        # Extract paragraph text
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                text_parts.append(paragraph.text)
                total_chars += len(paragraph.text)
                if max_chars and total_chars > max_chars:
                    break
        
        # Extract table text if requested
        if include_tables:
            for table in doc.tables:
                table_text = []
                for row in table.rows:
                    row_text = []
                    for cell in row.cells:
                        cell_content = " ".join(p.text for p in cell.paragraphs if p.text.strip())
                        if cell_content:
                            row_text.append(cell_content)
                    if row_text:
                        table_text.append(" | ".join(row_text))
                
                if table_text:
                    text_parts.append("\n[Table]\n" + "\n".join(table_text))
                    total_chars += len(text_parts[-1])
                    if max_chars and total_chars > max_chars:
                        break
        
        result = "\n\n".join(text_parts)
        
        # Apply character limit if specified
        if max_chars and len(result) > max_chars:
            result = result[:max_chars] + "... [content truncated]"
        
        return result
        
    except Exception as e:
        return f"Failed to extract text: {str(e)}"


def extract_document_text_chunked(doc_path: str, chunk_size: int = 50000) -> Tuple[str, bool]:
    """
    Extract document text in chunks for large files.
    
    Args:
        doc_path: Path to the Word document
        chunk_size: Maximum characters per chunk
        
    Returns:
        Tuple of (text_content, is_truncated)
    """
    import os
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist", False
    
    try:
        doc = Document(doc_path)
        text_parts = []
        total_chars = 0
        is_truncated = False
        
        # Extract paragraph text
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                if total_chars + len(paragraph.text) > chunk_size:
                    remaining = chunk_size - total_chars
                    if remaining > 0:
                        text_parts.append(paragraph.text[:remaining])
                    is_truncated = True
                    break
                
                text_parts.append(paragraph.text)
                total_chars += len(paragraph.text)
        
        # Add table content if space permits
        if not is_truncated:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        cell_text = " ".join(p.text for p in cell.paragraphs if p.text.strip())
                        if cell_text:
                            if total_chars + len(cell_text) > chunk_size:
                                remaining = chunk_size - total_chars
                                if remaining > 0:
                                    text_parts.append(cell_text[:remaining])
                                is_truncated = True
                                break
                            text_parts.append(cell_text)
                            total_chars += len(cell_text)
                    if is_truncated:
                        break
                if is_truncated:
                    break
        
        result = "\n\n".join(text_parts)
        return result, is_truncated
        
    except Exception as e:
        return f"Failed to extract text: {str(e)}", False


def get_document_structure(doc_path: str) -> Dict[str, Any]:
    """Get the structure of a Word document."""
    import os
    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}
    
    try:
        doc = Document(doc_path)
        structure = {
            "paragraphs": [],
            "tables": []
        }
        
        # Get paragraphs
        for i, para in enumerate(doc.paragraphs):
            structure["paragraphs"].append({
                "index": i,
                "text": para.text[:100] + ("..." if len(para.text) > 100 else ""),
                "style": para.style.name if para.style else "Normal"
            })
        
        # Get tables
        for i, table in enumerate(doc.tables):
            table_data = {
                "index": i,
                "rows": len(table.rows),
                "columns": len(table.columns),
                "preview": []
            }
            
            # Get sample of table data
            max_rows = min(3, len(table.rows))
            for row_idx in range(max_rows):
                row_data = []
                max_cols = min(3, len(table.columns))
                for col_idx in range(max_cols):
                    try:
                        cell_text = table.cell(row_idx, col_idx).text
                        row_data.append(cell_text[:20] + ("..." if len(cell_text) > 20 else ""))
                    except IndexError:
                        row_data.append("N/A")
                table_data["preview"].append(row_data)
            
            structure["tables"].append(table_data)
        
        return structure
    except Exception as e:
        return {"error": f"Failed to get document structure: {str(e)}"}


def get_all_paragraphs(doc_path: str) -> Dict[str, Any]:
    """
    一次性获取所有段落内容
    
    Args:
        doc_path: Word文档路径
    
    Returns:
        包含所有段落信息的字典
    """
    import os
    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}
    
    try:
        doc = Document(doc_path)
        paragraphs = []
        
        for i, para in enumerate(doc.paragraphs):
            paragraphs.append({
                "index": i,
                "text": para.text,
                "style": para.style.name if para.style else "Normal",
                "runs": [
                    {
                        "text": run.text,
                        "bold": run.bold,
                        "italic": run.italic,
                        "underline": run.underline
                    }
                    for run in para.runs if run.text
                ]
            })
        
        return {
            "total_paragraphs": len(paragraphs),
            "paragraphs": paragraphs
        }
    except Exception as e:
        return {"error": f"Failed to get paragraphs: {str(e)}"}


def get_paragraphs_by_range(doc_path: str, start_index: int = 0, end_index: Optional[int] = None) -> Dict[str, Any]:
    """
    获取指定段落范围的内容
    
    Args:
        doc_path: Word文档路径
        start_index: 起始段落索引（包含）
        end_index: 结束段落索引（不包含），None表示到最后
    
    Returns:
        包含指定范围段落信息的字典
    """
    import os
    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}
    
    try:
        doc = Document(doc_path)
        paragraphs = []
        
        total_paragraphs = len(doc.paragraphs)
        
        # 验证索引范围
        if start_index < 0:
            start_index = 0
        if end_index is None:
            end_index = total_paragraphs
        elif end_index > total_paragraphs:
            end_index = total_paragraphs
        
        if start_index >= total_paragraphs:
            return {
                "error": f"Start index {start_index} exceeds total paragraphs {total_paragraphs}"
            }
        
        # 获取指定范围的段落
        for i in range(start_index, min(end_index, total_paragraphs)):
            para = doc.paragraphs[i]
            paragraphs.append({
                "index": i,
                "text": para.text,
                "style": para.style.name if para.style else "Normal",
                "runs": [
                    {
                        "text": run.text,
                        "bold": run.bold,
                        "italic": run.italic,
                        "underline": run.underline,
                        "font_size": run.font.size.pt if run.font.size else None
                    }
                    for run in para.runs if run.text
                ]
            })
        
        return {
            "total_paragraphs": total_paragraphs,
            "range": f"{start_index}-{min(end_index, total_paragraphs)}",
            "paragraphs": paragraphs
        }
    except Exception as e:
        return {"error": f"Failed to get paragraphs by range: {str(e)}"}


def get_paragraphs_by_page(doc_path: str, page_number: int, page_size: int = 100) -> Dict[str, Any]:
    """
    分页获取段落内容
    
    Args:
        doc_path: Word文档路径
        page_number: 页码（从1开始）
        page_size: 每页段落数量
    
    Returns:
        包含分页段落信息的字典
    """
    import os
    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}
    
    try:
        doc = Document(doc_path)
        total_paragraphs = len(doc.paragraphs)
        
        # 计算分页信息
        total_pages = (total_paragraphs + page_size - 1) // page_size
        
        if page_number < 1:
            page_number = 1
        if page_number > total_pages:
            page_number = total_pages
        
        start_index = (page_number - 1) * page_size
        end_index = min(start_index + page_size, total_paragraphs)
        
        paragraphs = []
        for i in range(start_index, end_index):
            para = doc.paragraphs[i]
            paragraphs.append({
                "index": i,
                "text": para.text,
                "style": para.style.name if para.style else "Normal",
                "character_count": len(para.text),
                "word_count": len(para.text.split())
            })
        
        return {
            "total_paragraphs": total_paragraphs,
            "total_pages": total_pages,
            "current_page": page_number,
            "page_size": page_size,
            "range": f"{start_index}-{end_index-1}",
            "paragraphs": paragraphs
        }
    except Exception as e:
        return {"error": f"Failed to get paragraphs by page: {str(e)}"}


def analyze_paragraph_distribution(doc_path: str) -> Dict[str, Any]:
    """
    分析段落分布情况
    
    Args:
        doc_path: Word文档路径
    
    Returns:
        段落统计分析信息
    """
    import os
    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}
    
    try:
        doc = Document(doc_path)
        
        if not doc.paragraphs:
            return {
                "total_paragraphs": 0,
                "empty_paragraphs": 0,
                "non_empty_paragraphs": 0,
                "styles": {},
                "character_stats": {
                    "min": 0,
                    "max": 0,
                    "avg": 0,
                    "total": 0
                }
            }
        
        paragraphs = doc.paragraphs
        total_paragraphs = len(paragraphs)
        
        # 统计分析
        empty_paragraphs = sum(1 for p in paragraphs if not p.text.strip())
        non_empty_paragraphs = total_paragraphs - empty_paragraphs
        
        # 样式统计
        styles = {}
        for para in paragraphs:
            style_name = para.style.name if para.style else "Normal"
            styles[style_name] = styles.get(style_name, 0) + 1
        
        # 字符统计
        char_counts = [len(p.text) for p in paragraphs if p.text.strip()]
        if char_counts:
            char_stats = {
                "min": min(char_counts),
                "max": max(char_counts),
                "avg": sum(char_counts) // len(char_counts),
                "total": sum(char_counts)
            }
        else:
            char_stats = {"min": 0, "max": 0, "avg": 0, "total": 0}
        
        return {
            "total_paragraphs": total_paragraphs,
            "empty_paragraphs": empty_paragraphs,
            "non_empty_paragraphs": non_empty_paragraphs,
            "styles": styles,
            "character_stats": char_stats,
            "word_stats": {
                "total_words": sum(len(p.text.split()) for p in paragraphs),
                "avg_words_per_paragraph": sum(len(p.text.split()) for p in paragraphs) // total_paragraphs if total_paragraphs > 0 else 0
            }
        }
    except Exception as e:
        return {"error": f"Failed to analyze paragraph distribution: {str(e)}"}


def find_paragraph_by_text(doc, text, partial_match=False):
    """
    Find paragraphs containing specific text.
    
    Args:
        doc: Document object
        text: Text to search for
        partial_match: If True, matches paragraphs containing the text; if False, matches exact text
        
    Returns:
        List of paragraph indices that match the criteria
    """
    matching_paragraphs = []
    
    for i, para in enumerate(doc.paragraphs):
        if partial_match and text in para.text:
            matching_paragraphs.append(i)
        elif not partial_match and para.text == text:
            matching_paragraphs.append(i)
            
    return matching_paragraphs


def find_and_replace_text(doc, old_text, new_text):
    """
    Find and replace text throughout the document.
    
    Args:
        doc: Document object
        old_text: Text to find
        new_text: Text to replace with
        
    Returns:
        Number of replacements made
    """
    count = 0
    
    # Search in paragraphs
    for para in doc.paragraphs:
        if old_text in para.text:
            for run in para.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)
                    count += 1
    
    # Search in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if old_text in para.text:
                        for run in para.runs:
                            if old_text in run.text:
                                run.text = run.text.replace(old_text, new_text)
                                count += 1
    
    return count


def get_document_xml(doc_path: str) -> str:
    """Extract and return the raw XML structure of the Word document (word/document.xml)."""
    import os
    import zipfile
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"
    try:
        with zipfile.ZipFile(doc_path) as docx_zip:
            with docx_zip.open('word/document.xml') as xml_file:
                return xml_file.read().decode('utf-8')
    except Exception as e:
        return f"Failed to extract XML: {str(e)}"


def insert_header_near_text(doc_path: str, target_text: str, header_title: str, position: str = 'after', header_style: str = 'Heading 1') -> str:
    """Insert a header (with specified style) before or after the first paragraph containing target_text."""
    import os
    from docx import Document
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"
    try:
        doc = Document(doc_path)
        found = False
        for i, para in enumerate(doc.paragraphs):
            if target_text in para.text:
                found = True
                # Create the new header paragraph with the specified style
                new_para = doc.add_paragraph(header_title, style=header_style)
                # Move the new paragraph to the correct position
                if position == 'before':
                    para._element.addprevious(new_para._element)
                else:
                    para._element.addnext(new_para._element)
                break
        
        if not found:
            return f"Text '{target_text}' not found in document"
        
        doc.save(doc_path)
        return f"Header '{header_title}' inserted {position} '{target_text}'"
    except Exception as e:
        return f"Failed to insert header: {str(e)}"


def insert_line_or_paragraph_near_text(doc_path: str, target_text: str, line_text: str, position: str = 'after', line_style: Optional[str] = None) -> str:
    """
    Insert a new line or paragraph (with specified or matched style) before or after the first paragraph containing target_text.
    """
    import os
    from docx import Document
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"
    try:
        doc = Document(doc_path)
        found = False
        for i, para in enumerate(doc.paragraphs):
            if target_text in para.text:
                found = True
                # Determine the style to use
                if line_style:
                    style = line_style
                else:
                    style = para.style.name if para.style else 'Normal'
                
                # Create the new paragraph
                new_para = doc.add_paragraph(line_text, style=style)
                
                # Move the new paragraph to the correct position
                if position == 'before':
                    para._element.addprevious(new_para._element)
                else:
                    para._element.addnext(new_para._element)
                break
        
        if not found:
            return f"Text '{target_text}' not found in document"
        
        doc.save(doc_path)
        return f"Line inserted {position} '{target_text}'"
    except Exception as e:
        return f"Failed to insert line: {str(e)}"
