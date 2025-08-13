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
    extended_document_tools
)
def get_transport_config():
    """
    Get transport configuration from environment variables.
    
    Returns:
        dict: Transport configuration with type, host, port, and other settings
    """
    # Default configuration
    config = {
        'transport': 'stdio',  # Default to stdio for backward compatibility
        'host': '127.0.0.1',
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
    @mcp.tool()
    def create_document(filename: str, title: Optional[str] = None, author: Optional[str] = None):
        """Create a new Word document with optional metadata."""
        return document_tools.create_document(filename, title, author)
    
    @mcp.tool()
    def copy_document(source_filename: str, destination_filename: Optional[str] = None):
        """Create a copy of a Word document."""
        return document_tools.copy_document(source_filename, destination_filename)
    
    @mcp.tool()
    def get_document_info(filename: str):
        """Get information about a Word document."""
        return document_tools.get_document_info(filename)
    
    @mcp.tool()
    def get_document_text(filename: str):
        """Extract all text from a Word document."""
        return document_tools.get_document_text(filename)
    
    @mcp.tool()
    def get_document_outline(filename: str):
        """Get the structure of a Word document."""
        return document_tools.get_document_outline(filename)
    
    @mcp.tool()
    def list_available_documents(directory: str = "."):
        """List all .docx files in the specified directory."""
        return document_tools.list_available_documents(directory)
    
    @mcp.tool()
    def get_document_xml(filename: str):
        """Get the raw XML structure of a Word document."""
        return document_tools.get_document_xml_tool(filename)
    
    @mcp.tool()
    def insert_header_near_text(filename: str, target_text: str, header_title: str, position: str = 'after', header_style: str = 'Heading 1'):
        """Insert a header (with specified style) before or after the first paragraph containing target_text. Args: filename (str), target_text (str), header_title (str), position ('before' or 'after'), header_style (str, default 'Heading 1')."""
        return document_tools.insert_header_near_text_tool(filename, target_text, header_title, position, header_style)
    
    @mcp.tool()
    def insert_line_or_paragraph_near_text(filename: str, target_text: str, line_text: str, position: str = 'after', line_style: Optional[str] = None):
        """
        Insert a new line or paragraph (with specified or matched style) before or after the first paragraph containing target_text.
        Args: filename (str), target_text (str), line_text (str), position ('before' or 'after'), line_style (str, optional).
        """
        return document_tools.insert_line_or_paragraph_near_text_tool(filename, target_text, line_text, position, line_style)
    # Content tools (paragraphs, headings, tables, etc.)
    @mcp.tool()
    def add_paragraph(filename: str, text: str, style: Optional[str] = None):
        """Add a paragraph to a Word document."""
        return content_tools.add_paragraph(filename, text, style)
    
    @mcp.tool()
    def add_heading(filename: str, text: str, level: int = 1):
        """Add a heading to a Word document."""
        return content_tools.add_heading(filename, text, level)
    
    @mcp.tool()
    def add_picture(filename: str, image_path: str, width: float = None):
        """Add an image to a Word document."""
        return content_tools.add_picture(filename, image_path, width)
    
    @mcp.tool()
    def add_table(filename: str, rows: int, cols: int, data: Optional[List[List[str]]] = None):
        """Add a table to a Word document."""
        return content_tools.add_table(filename, rows, cols, data)
    
    @mcp.tool()
    def add_page_break(filename: str):
        """Add a page break to the document."""
        return content_tools.add_page_break(filename)
    
    @mcp.tool()
    def delete_paragraph(filename: str, paragraph_index: int):
        """Delete a paragraph from a document."""
        return content_tools.delete_paragraph(filename, paragraph_index)
    
    @mcp.tool()
    def search_and_replace(filename: str, find_text: str, replace_text: str):
        """Search for text and replace all occurrences."""
        return content_tools.search_and_replace(filename, find_text, replace_text)
    
    # Format tools (styling, text formatting, etc.)
    @mcp.tool()
    def create_custom_style(filename: str, style_name: str, bold: Optional[bool] = None,
                          italic: Optional[bool] = None, font_size: Optional[int] = None,
                          font_name: Optional[str] = None, color: Optional[str] = None,
                          base_style: Optional[str] = None):
        """Create a custom style in the document."""
        return format_tools.create_custom_style(
            filename, style_name, bold, italic, font_size, font_name, color, base_style
        )
    
    @mcp.tool()
    def format_text(filename: str, paragraph_index: int, start_pos: int, end_pos: int,
                   bold: Optional[bool] = None, italic: Optional[bool] = None, underline: Optional[bool] = None,
                   color: Optional[str] = None, font_size: Optional[int] = None, font_name: Optional[str] = None):
        """Format a specific range of text within a paragraph."""
        return format_tools.format_text(
            filename, paragraph_index, start_pos, end_pos, bold, italic, 
            underline, color, font_size, font_name
        )
    
    @mcp.tool()
    def format_table(filename: str, table_index: int, has_header_row: Optional[bool] = None,
                    border_style: Optional[str] = None, shading: Optional[List[List[str]]] = None):
        """Format a table with borders, shading, and structure."""
        return format_tools.format_table(filename, table_index, has_header_row, border_style, shading)
    
    # Protection tools
    @mcp.tool()
    def protect_document(filename: str, password: str):
        """Add password protection to a Word document."""
        return protection_tools.protect_document(filename, password)
    
    @mcp.tool()
    def unprotect_document(filename: str, password: str):
        """Remove password protection from a Word document."""
        return protection_tools.unprotect_document(filename, password)
    
    # Footnote tools
    @mcp.tool()
    def add_footnote_to_document(filename: str, paragraph_index: int, footnote_text: str):
        """Add a footnote to a specific paragraph in a Word document."""
        return footnote_tools.add_footnote_to_document(filename, paragraph_index, footnote_text)
    
    @mcp.tool()
    def add_endnote_to_document(filename: str, paragraph_index: int, endnote_text: str):
        """Add an endnote to a specific paragraph in a Word document."""
        return footnote_tools.add_endnote_to_document(filename, paragraph_index, endnote_text)
    
    @mcp.tool()
    def customize_footnote_style(filename: str, numbering_format: str = "1, 2, 3",
                                start_number: int = 1, font_name: Optional[str] = None,
                                font_size: Optional[int] = None):
        """Customize footnote numbering and formatting in a Word document."""
        return footnote_tools.customize_footnote_style(
            filename, numbering_format, start_number, font_name, font_size
        )
    
    # Extended document tools
    @mcp.tool()
    def get_paragraph_text_from_document(filename: str, paragraph_index: int):
        """Get text from a specific paragraph in a Word document."""
        return extended_document_tools.get_paragraph_text_from_document(filename, paragraph_index)
    
    @mcp.tool()
    def find_text_in_document(filename: str, text_to_find: str, match_case: bool = True,
                             whole_word: bool = False):
        """Find occurrences of specific text in a Word document."""
        return extended_document_tools.find_text_in_document(
            filename, text_to_find, match_case, whole_word
        )
    
    @mcp.tool()
    def convert_to_pdf(filename: str, output_filename: Optional[str] = None):
        """Convert a Word document to PDF format."""
        return extended_document_tools.convert_to_pdf(filename, output_filename)
    
    # 新增段落读取工具
    @mcp.tool()
    def get_all_paragraphs(filename: str):
        """一次性获取所有段落内容，返回包含所有段落详细信息的JSON"""
        from word_document_server.tools.document_tools import get_all_paragraphs_tool
        return get_all_paragraphs_tool(filename)
    
    @mcp.tool()
    def get_paragraphs_by_range(filename: str, start_index: int = 0, end_index: Optional[int] = None):
        """获取指定段落范围的内容，支持起始和结束索引"""
        from word_document_server.tools.document_tools import get_paragraphs_by_range_tool
        return get_paragraphs_by_range_tool(filename, start_index, end_index)
    
    @mcp.tool()
    def get_paragraphs_by_page(filename: str, page_number: int = 1, page_size: int = 100):
        """分页获取段落内容，支持页码和每页数量"""
        from word_document_server.tools.document_tools import get_paragraphs_by_page_tool
        return get_paragraphs_by_page_tool(filename, page_number, page_size)
    
    @mcp.tool()
    def analyze_paragraph_distribution(filename: str):
        """分析段落分布情况，返回统计信息"""
        from word_document_server.tools.document_tools import analyze_paragraph_distribution_tool
        return analyze_paragraph_distribution_tool(filename)


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