"""
Comment extraction tools for Word Document Server.
"""
import os
import json
from mcp.server.fastmcp.server import Context
from word_document_server.app import app
from word_document_server.utils.app_context import AppContext
from word_document_server.utils.com_utils import handle_com_error

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

@app.tool()
def get_all_comments(context: Context) -> str:
    """Extract all comments from a Word document."""
    app_context: AppContext = context.request_context.lifespan_context
    doc = app_context.get_active_document()
    if not doc:
        return json.dumps({'success': False, 'error': 'No active document found'}, indent=2)

    try:
        comments_list = [_comment_to_dict(c) for c in doc.Comments]
        return json.dumps({
            'success': True,
            'comments': comments_list,
            'total_comments': len(comments_list)
        }, indent=2)
    except Exception as e:
        return json.dumps({'success': False, 'error': handle_com_error(e)}, indent=2)

@app.tool()
def get_comments_by_author(context: Context, author: str) -> str:
    """Extract comments from a specific author in a Word document."""
    app_context: AppContext = context.request_context.lifespan_context
    if not author or not author.strip():
        return json.dumps({'success': False, 'error': 'Author name cannot be empty'}, indent=2)

    doc = app_context.get_active_document()
    if not doc:
        return json.dumps({'success': False, 'error': 'No active document found'}, indent=2)

    try:
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
        return json.dumps({'success': False, 'error': handle_com_error(e)}, indent=2)

@app.tool()
def get_comments_for_paragraph(context: Context, paragraph_index: int) -> str:
    """Extract comments for a specific paragraph in a Word document."""
    app_context: AppContext = context.request_context.lifespan_context
    if paragraph_index < 0:
        return json.dumps({'success': False, 'error': 'Paragraph index must be non-negative'}, indent=2)

    doc = app_context.get_active_document()
    if not doc:
        return json.dumps({'success': False, 'error': 'No active document found'}, indent=2)

    try:
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
        return json.dumps({'success': False, 'error': handle_com_error(e)}, indent=2)