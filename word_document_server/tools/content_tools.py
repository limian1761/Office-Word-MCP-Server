"""
Content tools for Word Document Server using COM.
"""
import os
from typing import List, Optional
from word_document_server.utils import com_utils
from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension

# Word constants
wdAlignParagraphLeft = 0
wdAlignParagraphCenter = 1
wdAlignParagraphRight = 2
wdLineSpaceSingle = 0
wdStory = 6

async def add_heading(text: str, level: int = 1) -> str:
    """Add a heading to a Word document using COM."""
    doc = None
    try:
        # 仅使用活动文档
        active_doc = com_utils.get_active_document()
        if active_doc:
            doc = active_doc
        else:
            return "没有活动文档"
        
        # Go to the end of the document
        selection = doc.Application.Selection
        selection.EndKey(Unit=wdStory)
        
        # Insert a new paragraph and apply heading style
        selection.InsertParagraphAfter()
        selection.Style = f"Heading {level}"
        selection.TypeText(text)
        doc.Save()
        return f"Heading '{text}' (level {level}) added to document"
    except Exception as e:
        return f"Failed to add heading: {str(e)}"
    finally:
        # 不需要关闭活动文档
        pass

async def add_paragraph(text: str, style: Optional[str] = None) -> str:
    """Add a paragraph to a Word document using COM."""
    doc = None
    try:
        # 仅使用活动文档
        active_doc = com_utils.get_active_document()
        if active_doc:
            doc = active_doc
        else:
            return "没有活动文档"
        
        selection = doc.Application.Selection
        selection.EndKey(Unit=wdStory)
        
        selection.TypeText(text)
        if style:
            try:
                selection.Style = style
            except Exception:
                # Style not found, use default
                pass
        doc.Save()
        return f"Paragraph added to document"
    except Exception as e:
        return f"Failed to add paragraph: {str(e)}"
    finally:
        # 不需要关闭活动文档
        pass

async def add_table(rows: int, cols: int, data: Optional[List[List[str]]] = None) -> str:
    """Add a table to a Word document using COM."""
    doc = None
    try:
        # 仅使用活动文档
        active_doc = com_utils.get_active_document()
        if active_doc:
            doc = active_doc
        else:
            return "没有活动文档"
        
        selection = doc.Application.Selection
        selection.EndKey(Unit=wdStory)
        selection.InsertParagraphAfter()
        
        table = doc.Tables.Add(selection.Range, rows, cols)
        table.Borders.Enable = True
        
        if data:
            for i, row_data in enumerate(data):
                for j, cell_text in enumerate(row_data):
                    if i < rows and j < cols:
                        table.Cell(i + 1, j + 1).Range.Text = str(cell_text)
        
        doc.Save()
        return f"Table ({rows}x{cols}) added to document"
    except Exception as e:
        return f"Failed to add table: {str(e)}"
    finally:
        # 不需要关闭活动文档
        pass

async def add_picture(image_path: str, width: Optional[float] = None) -> str:
    """Add an image to a Word document using COM."""
    abs_image_path = os.path.abspath(image_path)
    if not os.path.exists(abs_image_path):
        return f"Image file not found: {abs_image_path}"

    doc = None
    try:
        # 仅使用活动文档
        active_doc = com_utils.get_active_document()
        if active_doc:
            doc = active_doc
        else:
            return "没有活动文档"
        
        selection = doc.Application.Selection
        selection.EndKey(Unit=wdStory)
        selection.InsertParagraphAfter()
        
        shape = selection.InlineShapes.AddPicture(FileName=abs_image_path)
        if width:
            # Convert inches to points (1 inch = 72 points)
            shape.Width = width * 72
            shape.LockAspectRatio = -1 # msoTrue
        
        doc.Save()
        return f"Picture {image_path} added to document"
    except Exception as e:
        return f"Failed to add picture: {str(e)}"
    finally:
        # 不需要关闭活动文档
        pass

async def add_page_break() -> str:
    """Add a page break to the document using COM."""
    doc = None
    try:
        # 仅使用活动文档
        active_doc = com_utils.get_active_document()
        if active_doc:
            doc = active_doc
        else:
            return "没有活动文档"
        
        selection = doc.Application.Selection
        selection.EndKey(Unit=wdStory)
        selection.InsertBreak(Type=7) # wdPageBreak
        doc.Save()
        return f"Page break added to document."
    except Exception as e:
        return f"Failed to add page break: {str(e)}"
    finally:
        # 不需要关闭活动文档
        pass

async def delete_paragraph(paragraph_index: int) -> str:
    """Delete a paragraph from a document using COM."""
    doc = None
    try:
        # 仅使用活动文档
        active_doc = com_utils.get_active_document()
        if active_doc:
            doc = active_doc
        else:
            return "没有活动文档"
        
        if paragraph_index < 0 or paragraph_index >= doc.Paragraphs.Count:
            return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs."
        
        # COM is 1-based
        doc.Paragraphs(paragraph_index + 1).Range.Delete()
        doc.Save()
        return f"Paragraph at index {paragraph_index} deleted successfully."
    except Exception as e:
        return f"Failed to delete paragraph: {str(e)}"
    finally:
        # 不需要关闭活动文档
        pass

async def search_and_replace(find_text: str, replace_text: str) -> str:
    """Search for text and replace all occurrences using COM."""
    doc = None
    try:
        # 仅使用活动文档
        active_doc = com_utils.get_active_document()
        if active_doc:
            doc = active_doc
        else:
            return "没有活动文档"
        
        selection = doc.Application.Selection
        selection.HomeKey(Unit=wdStory)
        
        find = selection.Find
        find.ClearFormatting()
        find.Text = find_text
        find.Replacement.ClearFormatting()
        find.Replacement.Text = replace_text
        
        # wdReplaceAll = 2
        result = find.Execute(Replace=2)
        
        count = 0
        # The Execute method in a loop can be tricky. A simpler way for "count" is not directly available.
        # We'll rely on the fact that it ran. A more complex implementation could count replacements.
        if result:
            doc.Save()
            # We can't easily get the count of replacements from the return value.
            # We can just confirm the operation was attempted.
            return f"Search and replace for '{find_text}' with '{replace_text}' executed. Please verify the document."
        else:
            return f"No occurrences of '{find_text}' found."
            
    except Exception as e:
        return f"Failed to search and replace: {str(e)}"
    finally:
        # 不需要关闭活动文档
        pass