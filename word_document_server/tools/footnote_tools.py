"""
Footnote and endnote tools for Word Document Server using COM.
"""
import os
from typing import Optional
from word_document_server.utils import com_utils
from word_document_server.utils.file_utils import check_file_writeable, ensure_docx_extension

# Word numbering style constants
wdNoteNumberStyleArabic = 0
wdNoteNumberStyleUppercaseRoman = 2
wdNoteNumberStyleLowercaseRoman = 3
wdNoteNumberStyleUppercaseLetter = 4
wdNoteNumberStyleLowercaseLetter = 5
wdNoteNumberStyleSymbol = 9

async def add_footnote_to_document(paragraph_index: int, footnote_text: str) -> str:
    """Add a footnote to a specific paragraph in a Word document using COM."""
    doc = None
    try:
        doc = com_utils.get_active_document()
        if not doc:
            return "No active document found."

        if paragraph_index < 0 or paragraph_index >= doc.Paragraphs.Count:
            return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs."

        # Get the range at the end of the specified paragraph
        p_range = doc.Paragraphs(paragraph_index + 1).Range
        # Collapse the range to its end point to not replace the paragraph text
        p_range.Collapse(0) # wdCollapseEnd = 0

        # Add the footnote
        doc.Footnotes.Add(Range=p_range, Text=footnote_text)
        
        doc.Save()
        return f"Footnote added to paragraph {paragraph_index}"
    except Exception as e:
        return f"Failed to add footnote: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def add_endnote_to_document(paragraph_index: int, endnote_text: str) -> str:
    """Add an endnote to a specific paragraph in a Word document using COM."""
    doc = None
    try:
        doc = com_utils.get_active_document()
        if not doc:
            return "No active document found."

        if paragraph_index < 0 or paragraph_index >= doc.Paragraphs.Count:
            return f"Invalid paragraph index. Document has {doc.Paragraphs.Count} paragraphs."

        p_range = doc.Paragraphs(paragraph_index + 1).Range
        p_range.Collapse(0) # wdCollapseEnd

        doc.Endnotes.Add(Range=p_range, Text=endnote_text)
        
        doc.Save()
        return f"Endnote added to paragraph {paragraph_index}"
    except Exception as e:
        return f"Failed to add endnote: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def convert_footnotes_to_endnotes_in_document() -> str:
    """Convert all footnotes to endnotes in a Word document using COM."""
    doc = None
    try:
        doc = com_utils.get_active_document()
        if not doc:
            return "No active document found."

        if doc.Footnotes.Count == 0:
            return "No footnotes found to convert."
            
        doc.Footnotes.Convert()
        doc.Save()
        return f"Converted {doc.Endnotes.Count} footnotes to endnotes"
    except Exception as e:
        return f"Failed to convert footnotes to endnotes: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def customize_footnote_style(numbering_format: str = "1, 2, 3", 
                                  start_number: int = 1, font_name: Optional[str] = None,
                                  font_size: Optional[int] = None) -> str:
    """Customize footnote numbering and formatting in a Word document using COM."""
    doc = None
    try:
        doc = com_utils.get_active_document()
        if not doc:
            return "No active document found."

        fn_options = doc.FootnoteOptions
        
        format_map = {
            "1, 2, 3": wdNoteNumberStyleArabic,
            "i, ii, iii": wdNoteNumberStyleLowercaseRoman,
            "I, II, III": wdNoteNumberStyleUppercaseRoman,
            "a, b, c": wdNoteNumberStyleLowercaseLetter,
            "A, B, C": wdNoteNumberStyleUppercaseLetter,
            "symbols": wdNoteNumberStyleSymbol,
        }
        
        fn_options.NumberingRule = 0 # wdRestartPage = 1, wdRestartSection = 0, wdContinuous = 2
        fn_options.StartingNumber = start_number
        fn_options.NumberStyle = format_map.get(numbering_format, wdNoteNumberStyleArabic)

        # Style formatting
        try:
            fn_style = doc.Styles("Footnote Text")
            if font_name:
                fn_style.Font.Name = font_name
            if font_size:
                fn_style.Font.Size = font_size
        except Exception:
            return "Could not find 'Footnote Text' style to customize."

        doc.Save()
        return f"Footnote style and numbering customized"
    except Exception as e:
        return f"Failed to customize footnote style: {str(e)}"
    finally:
        if doc:
            doc.Close(SaveChanges=0)