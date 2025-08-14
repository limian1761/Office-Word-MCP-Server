"""
Comment extraction tools for Word Document Server using COM.
"""
import os
import json
from word_document_server.utils import com_utils
from word_document_server.utils.file_utils import ensure_docx_extension

def _comment_to_dict(comment):
    """Helper to convert a COM comment object to a dictionary."""
    return {
        "id": comment.Index,
        "author": comment.Author,
        "initials": comment.Initial,
        "date": str(comment.Date),
        "text": comment.Range.Text.strip(),
        "scope_text": comment.Scope.Text.strip(),
    }

async def get_all_comments(filename: str) -> str:
    """Extract all comments from a Word document using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({'success': False, 'error': f'Document {filename} does not exist'}, indent=2)

    doc = None
    try:
        doc = com_utils.open_document(filename)
        comments_list = [_comment_to_dict(c) for c in doc.Comments]
        return json.dumps({
            'success': True,
            'comments': comments_list,
            'total_comments': len(comments_list)
        }, indent=2)
    except Exception as e:
        return json.dumps({'success': False, 'error': f'Failed to extract comments: {str(e)}'}, indent=2)
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def get_comments_by_author(filename: str, author: str) -> str:
    """Extract comments from a specific author in a Word document using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({'success': False, 'error': f'Document {filename} does not exist'}, indent=2)
    if not author or not author.strip():
        return json.dumps({'success': False, 'error': 'Author name cannot be empty'}, indent=2)

    doc = None
    try:
        doc = com_utils.open_document(filename)
        author_comments = [
            _comment_to_dict(c) for c in doc.Comments if c.Author.lower() == author.lower()
        ]
        return json.dumps({
            'success': True,
            'author': author,
            'comments': author_comments,
            'total_comments': len(author_comments)
        }, indent=2)
    except Exception as e:
        return json.dumps({'success': False, 'error': f'Failed to extract comments: {str(e)}'}, indent=2)
    finally:
        if doc:
            doc.Close(SaveChanges=0)

async def get_comments_for_paragraph(filename: str, paragraph_index: int) -> str:
    """Extract comments for a specific paragraph in a Word document using COM."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return json.dumps({'success': False, 'error': f'Document {filename} does not exist'}, indent=2)
    if paragraph_index < 0:
        return json.dumps({'success': False, 'error': 'Paragraph index must be non-negative'}, indent=2)

    doc = None
    try:
        doc = com_utils.open_document(filename)
        if paragraph_index >= doc.Paragraphs.Count:
            return json.dumps({
                'success': False,
                'error': f'Paragraph index {paragraph_index} is out of range. Document has {doc.Paragraphs.Count} paragraphs.'
            }, indent=2)

        p = doc.Paragraphs(paragraph_index + 1) # COM is 1-based
        p_range = p.Range
        paragraph_text = p_range.Text.strip()

        para_comments = [
            _comment_to_dict(c) for c in doc.Comments 
            if c.Scope.InRange(p_range)
        ]

        return json.dumps({
            'success': True,
            'paragraph_index': paragraph_index,
            'paragraph_text': paragraph_text,
            'comments': para_comments,
            'total_comments': len(para_comments)
        }, indent=2)
    except Exception as e:
        return json.dumps({'success': False, 'error': f'Failed to extract comments: {str(e)}'}, indent=2)
    finally:
        if doc:
            doc.Close(SaveChanges=0)
