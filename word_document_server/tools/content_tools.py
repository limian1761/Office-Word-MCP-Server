"""
Content tools for Word Document Server.
"""
import os
from typing import List, Optional
from mcp.server.fastmcp.server import Context
from word_document_server.app import app
from word_document_server.utils.app_context import AppContext
from word_document_server.utils.com_utils import handle_com_error
from typing import Any, Callable, Optional

# Word constants
wdAlignParagraphLeft = 0
wdAlignParagraphCenter = 1
wdAlignParagraphRight = 2
wdLineSpaceSingle = 0
wdStory = 6

@app.tool()
def add_heading(context: Context, text: str, level: int = 1, paragraph_index: Optional[int] = None) -> str:
    """Add a heading to a Word document.
    
    Args:
        text: The text content of the heading
        level: The heading level (1-9)
        paragraph_index: Optional index where to insert the heading. 
                        If None, adds at the end of document.
    """
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        selection = doc.Application.Selection
        
        # If paragraph_index is specified, move to that position
        if paragraph_index is not None:
            if paragraph_index < 0 or paragraph_index > doc.Paragraphs.Count:
                return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs."
            # Move to the specified paragraph
            target_range = doc.Paragraphs(paragraph_index + 1).Range
            target_range.Select()
            selection.Collapse(Direction=0)  # Collapse to start of paragraph
        else:
            # Move to end of document
            selection.EndKey(Unit=wdStory)
        
        selection.InsertParagraphAfter()
        selection.Style = f"Heading {level}"
        selection.TypeText(text)
        doc.Save()
        return f"Heading '{text}' (level {level}) added to document"
    except Exception as e:
        return handle_com_error(e)

@app.tool()
def add_paragraph(context: Context, text: str, style: Optional[str] = None, paragraph_index: Optional[int] = None) -> str:
    """Add a paragraph to a Word document.
    
    Args:
        text: The text content of the paragraph
        style: Optional style to apply to the paragraph
        paragraph_index: Optional index where to insert the paragraph. 
                        If None, adds at the end of document.
    """
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        selection = doc.Application.Selection
        
        # If paragraph_index is specified, move to that position
        if paragraph_index is not None:
            if paragraph_index < 0 or paragraph_index > doc.Paragraphs.Count:
                return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs."
            # Move to the specified paragraph
            target_range = doc.Paragraphs(paragraph_index + 1).Range
            target_range.Select()
            selection.Collapse(Direction=0)  # Collapse to start of paragraph
        else:
            # Move to end of document
            selection.EndKey(Unit=wdStory)
        
        selection.TypeText(text)
        if style:
            try:
                selection.Style = style
            except Exception:
                pass
        selection.TypeParagraph()  # Add paragraph break
        doc.Save()
        return f"Paragraph added to document"
    except Exception as e:
        return handle_com_error(e)

@app.tool()
def add_table(context: Context, rows: int, cols: int, data: Optional[List[List[str]]] = None, paragraph_index: Optional[int] = None) -> str:
    """Add a table to a Word document.
    
    Args:
        rows: Number of rows in the table
        cols: Number of columns in the table
        data: Optional data to populate the table
        paragraph_index: Optional index where to insert the table. 
                        If None, adds at the end of document.
    """
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        selection = doc.Application.Selection
        
        # If paragraph_index is specified, move to that position
        if paragraph_index is not None:
            if paragraph_index < 0 or paragraph_index > doc.Paragraphs.Count:
                return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs."
            # Move to the specified paragraph
            target_range = doc.Paragraphs(paragraph_index + 1).Range
            target_range.Select()
            selection.Collapse(Direction=0)  # Collapse to start of paragraph
        else:
            # Move to end of document
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
        return handle_com_error(e)

@app.tool()
def add_picture(context: Context, image_path: str, width: Optional[float] = None, paragraph_index: Optional[int] = None) -> str:
    """Add an image to a Word document.
    
    Args:
        image_path: Path to the image file
        width: Optional width for the image
        paragraph_index: Optional index where to insert the image. 
                        If None, adds at the end of document.
    """
    app_context: AppContext = context.request_context.lifespan_context
    abs_image_path = os.path.abspath(image_path)
    if not os.path.exists(abs_image_path):
        return f"Image file not found: {abs_image_path}"

    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        selection = doc.Application.Selection
        
        # If paragraph_index is specified, move to that position
        if paragraph_index is not None:
            if paragraph_index < 0 or paragraph_index > doc.Paragraphs.Count:
                return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs."
            # Move to the specified paragraph
            target_range = doc.Paragraphs(paragraph_index + 1).Range
            target_range.Select()
            selection.Collapse(Direction=0)  # Collapse to start of paragraph
        else:
            # Move to end of document
            selection.EndKey(Unit=wdStory)
        
        selection.InsertParagraphAfter()
        
        shape = selection.InlineShapes.AddPicture(FileName=abs_image_path)
        if width:
            shape.Width = width * 72
            shape.LockAspectRatio = -1
        
        doc.Save()
        return f"Picture {image_path} added to document"
    except Exception as e:
        return handle_com_error(e)

@app.tool()
def select_paragraphs(context: Context, start_index: int, end_index: Optional[int] = None) -> str:
    """Select a range of paragraphs in the document.
    
    Args:
        start_index: The starting paragraph index (0-based)
        end_index: The ending paragraph index (0-based, inclusive). 
                  If None, only the start_index paragraph is selected.
    """
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        # Validate indices
        if start_index < 0 or start_index >= doc.Paragraphs.Count:
            return f"Invalid start index. Document has {doc.Paragraphs.Count} paragraphs."
        
        if end_index is not None:
            if end_index < 0 or end_index >= doc.Paragraphs.Count:
                return f"Invalid end index. Document has {doc.Paragraphs.Count} paragraphs."
            
            if end_index < start_index:
                return "End index must be greater than or equal to start index."
        
        # Get the selection object
        selection = doc.Application.Selection
        
        # Get the start paragraph range
        start_range = doc.Paragraphs(start_index + 1).Range
        
        # If end_index is None, select only the start paragraph
        if end_index is None:
            start_range.Select()
        else:
            # Get the end paragraph range
            end_range = doc.Paragraphs(end_index + 1).Range
            
            # Create a new range that spans from start to end
            combined_range = doc.Range(Start=start_range.Start, End=end_range.End)
            combined_range.Select()
        
        return f"Selected paragraphs from index {start_index} to {end_index if end_index is not None else start_index}"
    except Exception as e:
        return handle_com_error(e)


@app.tool()
def add_page_break(context: Context, paragraph_index: Optional[int] = None) -> str:
    """Add a page break to the document.
    
    Args:
        paragraph_index: Optional index where to insert the page break. 
                        If None, adds at the end of document.
    """
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        selection = doc.Application.Selection
        
        # If paragraph_index is specified, move to that position
        if paragraph_index is not None:
            if paragraph_index < 0 or paragraph_index > doc.Paragraphs.Count:
                return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs."
            # Move to the specified paragraph
            target_range = doc.Paragraphs(paragraph_index + 1).Range
            target_range.Select()
            selection.Collapse(Direction=0)  # Collapse to start of paragraph
        else:
            # Move to end of document
            selection.EndKey(Unit=wdStory)
        
        selection.InsertBreak(Type=7) # wdPageBreak
        doc.Save()
        return f"Page break added to document."
    except Exception as e:
        return handle_com_error(e)

@app.tool()
def delete_paragraph(context: Context, paragraph_index: int) -> str:
    """Delete a paragraph from a document."""
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        if paragraph_index < 0 or paragraph_index >= doc.Paragraphs.Count:
            return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs."
        
        doc.Paragraphs(paragraph_index + 1).Range.Delete()
        doc.Save()
        return f"Paragraph at index {paragraph_index} deleted successfully."
    except Exception as e:
        return handle_com_error(e)

@app.tool()
def replace_block_below_header(context: Context, header_text: str, new_paragraphs: list[str]) -> str:
    """Replace the block of paragraphs below a header, avoiding modification of TOC."""
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    try:
        found_header = False
        start_index = -1
        
        for i in range(doc.Paragraphs.Count):
            p = doc.Paragraphs(i + 1)
            # 考虑标题可能为中文，修改匹配逻辑，兼容中英文标题样式
            if (p.Style.NameLocal.startswith("Heading") or p.Style.NameLocal.startswith("标题")) and header_text in p.Range.Text.strip().replace('\u3000', ' ').replace('\xa0', ' '):
                found_header = True
                start_index = i + 1
                break
        
        if not found_header:
            return f"Header '{header_text}' not found"
        
        end_index = doc.Paragraphs.Count
        
        header_style = doc.Paragraphs(start_index).Style.NameLocal
        # 考虑标题可能为中文，修改匹配逻辑，兼容中英文标题样式
        if header_style.startswith("Heading") or header_style.startswith("标题"):
            try:
                # 尝试从英文或中文标题样式中提取级别
                if header_style.startswith("Heading"):
                    current_level = int(header_style.split()[-1])
                elif header_style.startswith("标题"):
                    current_level = int(''.join(filter(str.isdigit, header_style)))
                
                for i in range(start_index, doc.Paragraphs.Count):
                    p = doc.Paragraphs(i + 1)
                    style_name = p.Style.NameLocal
                    # 考虑标题可能为中文，修改匹配逻辑，兼容中英文标题样式
                    if style_name.startswith("Heading") or style_name.startswith("标题"):
                        try:
                            if style_name.startswith("Heading"):
                                level = int(style_name.split()[-1])
                            elif style_name.startswith("标题"):
                                level = int(''.join(filter(str.isdigit, style_name)))
                            
                            if level <= current_level:
                                end_index = i
                                break
                        except (ValueError, IndexError):
                            continue
            except (ValueError, IndexError):
                pass
        
        for i in range(end_index - 1, start_index, -1):
            doc.Paragraphs(i + 1).Range.Delete()
        
        selection = doc.Application.Selection
        header_range = doc.Paragraphs(start_index).Range
        header_range.Select()
        selection.Collapse(Direction=0)
        
        for paragraph_text in new_paragraphs:
            selection.TypeText(paragraph_text)
            selection.TypeParagraph()
        
        doc.Save()
        return f"Successfully replaced block below header '{header_text}'"
    except Exception as e:
        return handle_com_error(e)

@app.tool()
def search_and_replace(context: Context, find_text: str, replace_text: str, match_case: bool = False, match_whole_word: bool = False) -> str:
    """Search for text and replace all occurrences, considering formatting.
    
    This function preserves formatting during find and replace operations.
    """
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        selection = doc.Application.Selection
        selection.HomeKey(Unit=wdStory)
        
        find = selection.Find
        find.ClearFormatting()
        find.Text = find_text
        find.Replacement.ClearFormatting()
        find.Replacement.Text = replace_text
        
        # Set formatting options
        find.MatchCase = match_case
        find.MatchWholeWord = match_whole_word
        
        # Enable format preservation through font properties
        find.Font.Bold = False  # Reset font properties
        find.Replacement.Font.Bold = False
        
        result = find.Execute(
            FindText=find_text,
            ReplaceWith=replace_text,
            Replace=2,  # wdReplaceAll = 2
            Forward=True,
            Wrap=1,     # wdFindContinue = 1
            Format=True,
            MatchCase=match_case,
            MatchWholeWord=match_whole_word
        )
        
        if result:
            doc.Save()
            return f"Search and replace for '{find_text}' with '{replace_text}' executed."
        else:
            return f"No occurrences of '{find_text}' found."
            
    except Exception as e:
        return handle_com_error(e)

@app.tool()
def add_picture_caption(context: Context, caption_text: str, picture_index: Optional[int] = None, paragraph_index: Optional[int] = None) -> str:
    """Add a caption to a picture in a Word document.
    
    Args:
        caption_text: The text content of the caption
        picture_index: Optional index of the picture to caption (0-based). 
                      If None, captions the last added picture.
        paragraph_index: Optional index where to insert the caption. 
                        If None, adds after the picture or at end of document.
    """
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        selection = doc.Application.Selection
        
        # Handle picture selection
        shape = None
        if picture_index is not None:
            if picture_index < 0 or picture_index >= doc.InlineShapes.Count:
                return f"Invalid picture index. Document has {doc.InlineShapes.Count} pictures."
            # Get the specified picture
            shape = doc.InlineShapes(picture_index + 1)
        elif doc.InlineShapes.Count > 0:
            # Get the last picture if no index specified
            shape = doc.InlineShapes(doc.InlineShapes.Count)
        else:
            # No pictures found
            return "No pictures found in document"
        
        # Position the cursor correctly
        if paragraph_index is not None:
            # If paragraph_index is specified, move to that position
            if paragraph_index < 0 or paragraph_index > doc.Paragraphs.Count:
                return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs."
            # Move to the specified paragraph
            target_range = doc.Paragraphs(paragraph_index + 1).Range
            target_range.Select()
            selection.Collapse(Direction=0)  # Collapse to start of paragraph
        elif shape is not None:
            # If a picture was selected, position after the picture
            shape.Range.Select()
            selection.Collapse(Direction=0)  # Collapse to end of picture
        else:
            # Move to end of document if no specific position
            selection.EndKey(Unit=wdStory)
        
        # Insert paragraph break and caption text
        selection.TypeParagraph()
        selection.TypeText(f"图 {doc.InlineShapes.Count if picture_index is None else picture_index + 1} {caption_text}")
        selection.TypeParagraph()  # Add paragraph break after caption
        
        doc.Save()
        return f"Caption '{caption_text}' added to picture"
    except Exception as e:
        return handle_com_error(e)

@app.tool()
def replace_between_anchors(context: Context, start_anchor_text: str, new_paragraphs: list[str], end_anchor_text: str, new_paragraph_style: Optional[str] = None) -> str:
    """Replace all content between start and end anchor texts."""
    app_context: AppContext = context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        start_index = -1
        for i in range(doc.Paragraphs.Count):
            p = doc.Paragraphs(i + 1)
            text = p.Range.Text.strip()
            
            if start_anchor_text == text:
                start_index = i
                break
        
        if start_index == -1:
            return f"Start anchor '{start_anchor_text}' not found"
        
        end_index = -1
        for i in range(start_index + 1, doc.Paragraphs.Count):
            p = doc.Paragraphs(i + 1)
            text = p.Range.Text.strip()
            
            if end_anchor_text == text:
                end_index = i
                break
        
        if end_index == -1:
            return f"End anchor '{end_anchor_text}' not found after start anchor"
        
        for i in range(end_index - 1, start_index, -1):
            doc.Paragraphs(i + 1).Range.Delete()
        
        selection = doc.Application.Selection
        start_range = doc.Paragraphs(start_index + 1).Range
        start_range.Select()
        selection.Collapse(Direction=0)
        
        for paragraph_text in new_paragraphs:
            selection.TypeText(paragraph_text)
            if new_paragraph_style:
                try:
                    selection.Style = new_paragraph_style
                except Exception:
                    pass
            selection.TypeParagraph()
        
        doc.Save()
        return f"Successfully replaced content between anchors '{start_anchor_text}' and '{end_anchor_text}'"
    except Exception as e:
        return f"Failed to replace content between anchors: {str(e)}"

@app.tool()
def add_paragraph_numbering(context: Context, start_index: int = 0, end_index: Optional[int] = None, style: str = "Normal") -> str:
    """Add numbering to paragraphs in a Word document.
    
    Args:
        start_index: The starting paragraph index (0-based)
        end_index: The ending paragraph index (0-based, inclusive). 
                  If None, numbers paragraphs from start_index to the end of document.
        style: The style of paragraphs to number (e.g., "Normal", "Body Text")
    """
    app_context: AppContext = context.lifespan_context
    doc = app_context.get_active_document()
    if doc is None:
        return "No active document found"
    
    try:
        # Validate indices
        if start_index < 0 or start_index >= doc.Paragraphs.Count:
            return f"Invalid start index. Document has {doc.Paragraphs.Count} paragraphs."
        
        if end_index is not None:
            if end_index < 0 or end_index >= doc.Paragraphs.Count:
                return f"Invalid end index. Document has {doc.Paragraphs.Count} paragraphs."
            
            if end_index < start_index:
                return "End index must be greater than or equal to start index."
        
        # Set default end_index to last paragraph if not specified
        if end_index is None:
            end_index = doc.Paragraphs.Count - 1
        
        # Initialize numbering counter
        number = 1
        
        # Iterate through specified paragraphs
        for i in range(start_index, end_index + 1):
            paragraph = doc.Paragraphs(i + 1)
            
            # Check if paragraph has the specified style
            if paragraph.Style.NameLocal == style or style == "All":
                # Get the paragraph text
                text = paragraph.Range.Text.strip()
                
                # Skip empty paragraphs
                if not text:
                    continue
                
                # Add numbering to the paragraph
                numbered_text = f"{number}. {text}"
                paragraph.Range.Text = numbered_text
                number += 1
        
        doc.Save()
        return f"Numbering added to paragraphs from index {start_index} to {end_index} with style '{style}'"
    except Exception as e:
        return handle_com_error(e)