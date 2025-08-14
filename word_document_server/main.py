"""
Main entry point for the Word Document MCP Server.
Acts as the central controller for the MCP server that handles Word document operations.
Supports multiple transports: stdio, sse, and streamable-http using standalone FastMCP.
"""

import os
import sys
from typing import Optional, List, Any
# Set required environment variable for FastMCP 2.8.1+
os.environ.setdefault('FASTMCP_LOG_LEVEL', 'INFO')
from fastmcp import FastMCP
from word_document_server.tools import (
    document_tools,
    content_tools,
    format_tools,
    protection_tools,
    footnote_tools,
    comment_tools
)
from word_document_server.tools.content_tools import replace_paragraph_block_below_header_tool
from word_document_server.tools.content_tools import replace_block_between_manual_anchors_tool

def get_transport_config():
    """
    Get transport configuration from environment variables.
    
    Returns:
        dict: Transport configuration with type, host, port, and other settings
    """
    # Default configuration
    config = {
        'transport': 'stdio',  # Default to stdio for backward compatibility
        'host': '0.0.0.0',
        'port': 8000,
        'path': '/mcp',
        'sse_path': '/sse'
    }
    
    # Override with environment variables if provided
    transport = os.getenv('MCP_TRANSPORT', 'stdio').lower()
    print(f"Transport: {transport}")
    # Validate transport type
    valid_transports = ['stdio', 'streamable-http', 'sse']
    if transport not in valid_transports:
        print(f"Warning: Invalid transport '{transport}'. Falling back to 'stdio'.")
        transport = 'stdio'
    
    config['transport'] = transport
    config['host'] = os.getenv('MCP_HOST', config['host'])
    config['port'] = int(os.getenv('MCP_PORT', config['port']))
    config['path'] = os.getenv('MCP_PATH', config['path'])
    config['sse_path'] = os.getenv('MCP_SSE_PATH', config['sse_path'])
    
    return config


def setup_logging(debug_mode):
    """
    Setup logging based on debug mode.
    
    Args:
        debug_mode (bool): Whether to enable debug logging
    """
    import logging
    
    if debug_mode:
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        print("Debug logging enabled")
    else:
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )


# Initialize FastMCP server
mcp = FastMCP("Word Document Server")


def register_tools():
    """Register all tools with the MCP server using FastMCP decorators."""
    
    # Document tools (create, copy, info, etc.)
    @mcp.tool(name="create_doc")
    def create_document(title: Optional[str] = None, author: Optional[str] = None):
        """Create a new Word document with optional metadata."""
        return document_tools.create_document(title, author)
    
    @mcp.tool()
    def copy_document(destination_filename: Optional[str] = None):
        """Create a copy of a Word document."""
        return document_tools.copy_document(destination_filename)
    
    @mcp.tool(name="doc_info")
    def get_document_info():
        """Get information about a Word document."""
        return document_tools.get_document_info()
    
    @mcp.tool(name="doc_text")
    def get_document_text():
        """Extract all text from a Word document."""
        return document_tools.get_document_text()
    
    @mcp.tool("doc_outline")
    def get_document_outline():
        """Get the structure of a Word document."""
        return document_tools.get_document_outline()
    
    @mcp.tool("list_docs")
    def list_available_documents(directory: str = "."):
        """List all .docx files in the specified directory."""
        return document_tools.list_available_documents(directory)
    
    @mcp.tool("list_active_docs")
    def get_active_documents_info():
        """Get information about all active Word documents."""
        from word_document_server.tools.document_tools import get_active_documents_info
        return get_active_documents_info()
    
    @mcp.tool("set_active_doc")
    def set_active_document():
        """Set the current active document."""
        from word_document_server.utils import com_utils
        try:
            doc = com_utils.get_active_document()
            if doc is None:
                return "No active document found"
            com_utils.set_active_document(doc)
            return f"Active document set to {doc.Name}"
        except Exception as e:
            return f"Failed to set active document: {str(e)}"
    
    @mcp.tool()
    def get_document_xml():
        """Get the raw XML structure of a Word document."""
        return document_tools.get_document_xml_tool()
    
    @mcp.tool("add_header")
    def insert_header_near_text(target_text: str = None, header_title: str = None, position: str = 'after', header_style: str = 'Heading 1', target_paragraph_index: int = None):
        """Insert a header (with specified style) before or after the target paragraph. Specify by text or paragraph index. Args: target_text (str, optional), header_title (str), position ('before' or 'after'), header_style (str, default 'Heading 1'), target_paragraph_index (int, optional)."""
        return content_tools.insert_header_near_text_tool(target_text, header_title, position, header_style, target_paragraph_index)
    
    @mcp.tool("add_line")
    def insert_line_or_paragraph_near_text(target_text: str = None, line_text: str = None, position: str = 'after', line_style: str = None, target_paragraph_index: int = None):
        """
        Insert a new line or paragraph (with specified or matched style) before or after the target paragraph. Specify by text or paragraph index. Args: target_text (str, optional), line_text (str), position ('before' or 'after'), line_style (str, optional), target_paragraph_index (int, optional).
        """
        return content_tools.insert_line_or_paragraph_near_text_tool(target_text, line_text, position, line_style, target_paragraph_index)
    
    @mcp.tool("add_list")
    def insert_numbered_list_near_text(target_text: str = None, list_items: list = None, position: str = 'after', target_paragraph_index: int = None):
        """Insert a numbered list before or after the target paragraph. Specify by text or paragraph index. Args: target_text (str, optional), list_items (list of str), position ('before' or 'after'), target_paragraph_index (int, optional)."""
        return content_tools.insert_numbered_list_near_text_tool(target_text, list_items, position, target_paragraph_index)
    # Content tools (paragraphs, headings, tables, etc.)

    @mcp.tool(name="add_para")
    def add_paragraph(text: str, style: Optional[str] = None):
        """Add a paragraph to a Word document."""
        return content_tools.add_paragraph(text, style)
    
    @mcp.tool(name="add_head")
    def add_heading(text: str, level: int = 1):
        """Add a heading to a Word document."""
        return content_tools.add_heading(text, level)
    
    @mcp.tool()
    def add_picture(image_path: str, width: float = None):
        """Add an image to a Word document."""
        return content_tools.add_picture(image_path, width)
    
    @mcp.tool(name="add_table")
    def add_table(rows: int, cols: int, data: Optional[List[List[str]]] = None):
        """Add a table to a Word document."""
        return content_tools.add_table(rows, cols, data)
    
    @mcp.tool()
    def add_page_break():
        """Add a page break to the document."""
        return content_tools.add_page_break()
    
    @mcp.tool()
    def delete_paragraph(paragraph_index: int):
        """Delete a paragraph from a document."""
        return content_tools.delete_paragraph(paragraph_index)
    
    @mcp.tool(name="find_replace")
    def search_and_replace(find_text: str, replace_text: str):
        """Search for text and replace all occurrences."""
        return content_tools.search_and_replace(find_text, replace_text)
    
    # Format tools (styling, text formatting, etc.)
    @mcp.tool("add_style")
    def create_custom_style(style_name: str, bold: Optional[bool] = None,
                          italic: Optional[bool] = None, font_size: Optional[int] = None,
                          font_name: Optional[str] = None, color: Optional[str] = None,
                          base_style: Optional[str] = None):
        """Create a custom style in the document."""
        return format_tools.create_custom_style(
            style_name, bold, italic, font_size, font_name, color, base_style
        )
    
    @mcp.tool(name="format_text")
    def format_text(paragraph_index: int, start_pos: int, end_pos: int,
                   bold: Optional[bool] = None, italic: Optional[bool] = None, underline: Optional[bool] = None,
                   color: Optional[str] = None, font_size: Optional[int] = None, font_name: Optional[str] = None):
        """Format a specific range of text within a paragraph."""
        return format_tools.format_text(
            paragraph_index, start_pos, end_pos, bold, italic, 
            underline, color, font_size, font_name
        )
    
    @mcp.tool("add_table_format")
    def format_table(table_index: int, has_header_row: Optional[bool] = None,
                    border_style: Optional[str] = None, shading: Optional[List[List[str]]] = None):
        """Format a table with borders, shading, and structure."""
        return format_tools.format_table(table_index, has_header_row, border_style, shading)
    
    # New table cell shading tools
    @mcp.tool("add_shading")
    def set_table_cell_shading(table_index: int, row_index: int, 
                              col_index: int, fill_color: str, pattern: str = "clear"):
        """Apply shading/filling to a specific table cell."""
        return format_tools.set_table_cell_shading(table_index, row_index, col_index, fill_color, pattern)
    
    @mcp.tool("add_alt_rows")
    def apply_table_alternating_rows(table_index: int, 
                                   color1: str = "FFFFFF", color2: str = "F2F2F2"):
        """Apply alternating row colors to a table for better readability."""
        return format_tools.apply_table_alternating_rows(table_index, color1, color2)
    
    @mcp.tool("add_header")
    def highlight_table_header(table_index: int, 
                             header_color: str = "4472C4", text_color: str = "FFFFFF"):
        """Apply special highlighting to table header row."""
        return format_tools.highlight_table_header(table_index, header_color, text_color)
    
    # Cell merging tools
    @mcp.tool("merge_cells")
    def merge_table_cells(table_index: int, start_row: int, start_col: int, 
                        end_row: int, end_col: int):
        """Merge cells in a rectangular area of a table."""
        return format_tools.merge_table_cells(table_index, start_row, start_col, end_row, end_col)
    
    @mcp.tool("merge_cells_horizontal")
    def merge_table_cells_horizontal(table_index: int, row_index: int, 
                                   start_col: int, end_col: int):
        """Merge cells horizontally in a single row."""
        return format_tools.merge_table_cells_horizontal(table_index, row_index, start_col, end_col)
    
    @mcp.tool("merge_cells_vertical")
    def merge_table_cells_vertical(table_index: int, col_index: int, 
                                 start_row: int, end_row: int):
        """Merge cells vertically in a single column."""
        return format_tools.merge_table_cells_vertical(table_index, col_index, start_row, end_row)
    
    # Cell alignment tools
    @mcp.tool("add_alignment")
    def set_table_cell_alignment(table_index: int, row_index: int, col_index: int,
                               horizontal: str = "left", vertical: str = "top"):
        """Set text alignment for a specific table cell."""
        return format_tools.set_table_cell_alignment(table_index, row_index, col_index, horizontal, vertical)
    
    @mcp.tool("add_alignment_all")
    def set_table_alignment_all(table_index: int, 
                              horizontal: str = "left", vertical: str = "top"):
        """Set text alignment for all cells in a table."""
        return format_tools.set_table_alignment_all(table_index, horizontal, vertical)
    
    # Protection tools
    @mcp.tool(name="protect_doc")
    def protect_document(password: str, protection_type: str = "readOnly"):
        """Add password protection to a Word document.\n\n        Args:\n            password (str): The password to protect the document\n            protection_type (str): Type of protection - 'readOnly', 'comments', 'trackedChanges', 'forms'\n        """
        return protection_tools.protect_document(password, protection_type)
    
    @mcp.tool()
    def unprotect_document(password: str):
        """Remove password protection from a Word document.\n\n        Args:\n            password (str): The password to remove protection\n        """
        return protection_tools.unprotect_document(password)
    
    # Footnote tools
    @mcp.tool(name="add_footnote")
    def add_footnote_to_document(paragraph_index: int, footnote_text: str):
        """Add a footnote to a specific paragraph in a Word document."""
        return footnote_tools.add_footnote_to_document(paragraph_index, footnote_text)
    
    @mcp.tool(name="add_endnote")
    def add_endnote_to_document(paragraph_index: int, endnote_text: str):
        """Add an endnote to a specific paragraph in a Word document."""
        return footnote_tools.add_endnote_to_document(paragraph_index, endnote_text)
    
    @mcp.tool(name="custom_footnote")
    def customize_footnote_style(numbering_format: str = "1, 2, 3",
                                start_number: int = 1, font_name: Optional[str] = None,
                                font_size: Optional[int] = None):
        """Customize footnote numbering and formatting in a Word document."""
        return footnote_tools.customize_footnote_style(
            numbering_format, start_number, font_name, font_size
        )
    
    # 新增段落读取工具
    @mcp.tool(name="all_paras")
    def get_all_paragraphs():
        """一次性获取所有段落内容，返回包含所有段落详细信息的JSON"""
        from word_document_server.tools.document_tools import get_all_paragraphs_tool
        return get_all_paragraphs_tool()
    
    @mcp.tool(name="paras_range")
    def get_paragraphs_by_range(start_index: int = 0, end_index: Optional[int] = None):
        """获取指定段落范围的内容，支持起始和结束索引"""
        from word_document_server.tools.document_tools import get_paragraphs_by_range_tool
        return get_paragraphs_by_range_tool(start_index, end_index)
    
    @mcp.tool(name="paras_page")
    def get_paragraphs_by_page(page_number: int = 1, page_size: int = 10):
        """分页获取段落内容，支持页码和每页数量"""
        from word_document_server.tools.document_tools import get_paragraphs_by_page_tool
        return get_paragraphs_by_page_tool(page_number, page_size)
    
    @mcp.tool(name="para_stats")
    def analyze_paragraph_distribution():
        """分析段落分布情况，返回统计信息"""
        from word_document_server.tools.document_tools import analyze_paragraph_distribution_tool
        return analyze_paragraph_distribution_tool()

    @mcp.tool(name="replace_below_header")
    def replace_paragraph_block_below_header(header_text: str, new_paragraphs: list, detect_block_end_fn=None):
        """Reemplaza el bloque de párrafos debajo de un encabezado, evitando modificar TOC."""
        return replace_paragraph_block_below_header_tool(header_text, new_paragraphs, detect_block_end_fn)

    @mcp.tool(name="replace_anchors")
    def replace_block_between_manual_anchors(start_anchor_text: str, new_paragraphs: list, end_anchor_text: str = None, match_fn=None, new_paragraph_style: str = None):
        """Replace all content between start_anchor_text and end_anchor_text (or next logical header if not provided)."""
        return replace_block_between_manual_anchors_tool(start_anchor_text, new_paragraphs, end_anchor_text, match_fn, new_paragraph_style)

    # Comment tools
    @mcp.tool(name="all_comments")
    def get_all_comments():
        """Extract all comments from a Word document."""
        return comment_tools.get_all_comments()
    
    @mcp.tool(name="comments_by_author")
    def get_comments_by_author(author: str):
        """Extract comments from a specific author in a Word document."""
        return comment_tools.get_comments_by_author(author)
    
    @mcp.tool(name="para_comments")
    def get_comments_for_paragraph(paragraph_index: int):
        """Extract comments for a specific paragraph in a Word document."""
        return comment_tools.get_comments_for_paragraph(paragraph_index)
    # New table column width tools
    @mcp.tool(name="set_col_width")
    def set_table_column_width(table_index: int, col_index: int, 
                              width: float, width_type: str = "points"):
        """Set the width of a specific table column."""
        return format_tools.set_table_column_width(table_index, col_index, width, width_type)

    @mcp.tool(name="set_cols_widths")
    def set_table_column_widths(table_index: int, widths: list, 
                               width_type: str = "points"):
        """Set the widths of multiple table columns."""
        return format_tools.set_table_column_widths(table_index, widths, width_type)

    @mcp.tool(name="set_table_width")
    def set_table_width(table_index: int, width: float, 
                       width_type: str = "points"):
        """Set the overall width of a table."""
        return format_tools.set_table_width(table_index, width, width_type)

    @mcp.tool(name="auto_fit_cols")
    def auto_fit_table_columns(table_index: int):
        """Set table columns to auto-fit based on content."""
        return format_tools.auto_fit_table_columns(table_index)

    # New table cell text formatting and padding tools
    @mcp.tool(name="format_cell")
    def format_table_cell_text(table_index: int, row_index: int, col_index: int,
                               text_content: str = None, bold: bool = None, italic: bool = None,
                               underline: bool = None, color: str = None, font_size: int = None,
                               font_name: str = None):
        """Format text within a specific table cell."""
        return format_tools.format_table_cell_text(table_index, row_index, col_index,
                                                   text_content, bold, italic, underline, color, font_size, font_name)

    @mcp.tool(name="cell_padding")
    def set_table_cell_padding(table_index: int, row_index: int, col_index: int,
                               top: float = None, bottom: float = None, left: float = None, 
                               right: float = None, unit: str = "points"):
        """Set padding/margins for a specific table cell."""
        return format_tools.set_table_cell_padding(table_index, row_index, col_index,
                                                   top, bottom, left, right, unit)



def run_server():
    """Run the Word Document MCP Server with configurable transport."""
    # Get transport configuration
    config = get_transport_config()
    
    # Setup logging
    # setup_logging(config['debug'])
    
    # Register all tools
    register_tools()
    
    # Print startup information
    transport_type = config['transport']
    print(f"Starting Word Document MCP Server with {transport_type} transport...")
    
    # if config['debug']:
    #     print(f"Configuration: {config}")
    
    try:
        if transport_type == 'stdio':
            # Run with stdio transport (default, backward compatible)
            print("Server running on stdio transport")
            mcp.run(transport='stdio')
            
        elif transport_type == 'streamable-http':
            # Run with streamable HTTP transport
            print(f"Server running on streamable-http transport at http://{config['host']}:{config['port']}{config['path']}")
            mcp.run(
                transport='streamable-http',
                host=config['host'],
                port=config['port'],
                path=config['path']
            )
            
        elif transport_type == 'sse':
            # Run with SSE transport
            print(f"Server running on SSE transport at http://{config['host']}:{config['port']}{config['sse_path']}")
            mcp.run(
                transport='sse',
                host=config['host'],
                port=config['port'],
                path=config['sse_path']
            )
            
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"Error starting server: {e}")
        if config['debug']:
            import traceback
            traceback.print_exc()
        sys.exit(1)
    
    return mcp


def main():
    """Main entry point for the server."""
    run_server()


if __name__ == "__main__":
    main()
