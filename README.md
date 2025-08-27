# Office-Word-MCP-Server## Installation## Usage with Claude for Desktop## Development## Troubleshooting## Contributing## License## Acknowledgments

- [Model Context Protocol](https://modelcontextprotocol.io/) for the protocol specification
- [python-docx](https://python-docx.readthedocs.io/) for Word document manipulation
- [MCP](https://github.com/modelcontextprotocol/python-sdk) for the Python MCP implementation

---

_Note: This server interacts with document files on your system. Always verify that requested operations are appropriate before confirming them in Claude for Desktop or other MCP clients._

This project is licensed under the MIT License - see the LICENSE file for details.

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Common Issues

1. **Missing Styles**

   - Some documents may lack required styles for heading and table operations
   - The server will attempt to create missing styles or use direct formatting
   - For best results, use templates with standard Word styles

2. **Permission Issues**

   - Ensure the server has permission to read/write to the document paths
   - Use the `copy_document` function to create editable copies of locked documents
   - Check file ownership and permissions if operations fail

3. **Image Insertion Problems**
   - Use absolute paths for image files
   - Verify image format compatibility (JPEG, PNG recommended)
   - Check image file size and permissions

4. **Table Formatting Issues**

   - **Cell index errors**: Ensure row and column indices are within table bounds (0-based indexing)
   - **Color format problems**: Use hex colors without '#' prefix (e.g., "FF0000" for red) or standard color names
   - **Padding unit confusion**: Specify "points" or "percent" explicitly when setting cell padding
   - **Column width conflicts**: Auto-fit may override manual column width settings
   - **Text formatting persistence**: Apply cell text formatting after setting cell content for best results

### Code Formatting and Static Analysis

This project uses several tools to maintain code quality and consistency:

- [Black](https://github.com/psf/black) for code formatting
- [isort](https://pycqa.github.io/isort/) for import sorting
- [mypy](http://mypy-lang.org/) for static type checking

To format the code and sort imports:

```bash
black word_document_server
isort word_document_server
```

To run static type checking:

```bash
mypy word_document_server
```

These tools help ensure code consistency and catch potential type-related errors before runtime.

### Configuration

#### Method 1: After Local Installation

1. After installation, add the server to your Claude for Desktop configuration file:

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "python",
      "args": ["/path/to/word_mcp_server.py"]
    }
  }
}
```

#### Method 2: Without Installation (Using uvx)

1. You can also configure Claude for Desktop to use the server without local installation by using the uvx package manager:

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "uvx",
      "args": ["--from", "office-word-mcp-server", "word_mcp_server"]
    }
  }
}
```

2. Configuration file locations:

   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`

3. Restart Claude for Desktop to load the configuration.

### Example Operations

Once configured, you can ask Claude to perform operations like:

- "Create a new document called 'report.docx' with a title page"
- "Add a heading and three paragraphs to my document"
- "Insert a 4x4 table with sales data"
- "Format the word 'important' in paragraph 2 to be bold and red"
- "Search and replace all instances of 'old term' with 'new term'"
- "Create a custom style for section headings"
- "Apply formatting to the table in my document"
- "Extract all comments from my document"
- "Filter comments by author"
- "Make the text in table cell (1,2) bold and blue with 14pt font"
- "Add 10 points of padding to all sides of the header cells"
- "Create a callout table with a blue checkmark icon and white text"
- "Set the first column width to 50 points and auto-fit the remaining columns"
- "Apply alternating row colors to make the table more readable"
- "Add a paragraph at the beginning of my document"
- "Insert a heading at position 5 in my document"
- "Add a table after paragraph 3"
- "Show me the outline of my document"
- "Add a caption to the first picture in my document"
- "Add numbering to all paragraphs in my document"
- "Add numbering to paragraphs 2 through 5"

### Installing via Smithery

To install Office Word Document Server for Claude Desktop automatically via [Smithery](https://smithery.ai/server/@GongRzhe/Office-Word-MCP-Server):

```bash
npx -y @smithery/cli install @GongRzhe/Office-Word-MCP-Server --client claude
```

### Prerequisites

- Python 3.8 or higher
- pip package manager

### Basic Installation

```bash
# Clone the repository
git clone https://github.com/GongRzhe/Office-Word-MCP-Server.git
cd Office-Word-MCP-Server

# Install dependencies
pip install -r requirements.txt
```

### Using uv for Development

This project supports [uv](https://docs.astral.sh/uv/), a fast Python package installer and resolver, for managing dependencies:

```bash
# Clone the repository
git clone https://github.com/GongRzhe/Office-Word-MCP-Server.git
cd Office-Word-MCP-Server

# Install dependencies with uv
uv pip install -r requirements.txt
```

### Using the Setup Script

Alternatively, you can use the provided setup script which handles:

- Checking prerequisites
- Setting up a virtual environment
- Installing dependencies
- Generating MCP configuration

```bash
python setup_mcp.py
```

[![smithery badge](https://smithery.ai/badge/@GongRzhe/Office-Word-MCP-Server)](https://smithery.ai/server/@GongRzhe/Office-Word-MCP-Server)

A Model Context Protocol (MCP) server for creating, reading, and manipulating Microsoft Word documents. This server enables AI assistants to work with Word documents through a standardized interface, providing rich document editing capabilities.

<a href="https://glama.ai/mcp/servers/@GongRzhe/Office-Word-MCP-Server">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/@GongRzhe/Office-Word-MCP-Server/badge" alt="Office Word Server MCP server" />
</a>

![](https://badge.mcpx.dev?type=server "MCP Server")

## Overview

Office-Word-MCP-Server implements the [Model Context Protocol](https://modelcontextprotocol.io/) to expose Word document operations as tools and resources. It serves as a bridge between AI assistants and Microsoft Word documents, allowing for document creation, content addition, formatting, and analysis.

The server features a modular architecture that separates concerns into core functionality, tools, and utilities, making it highly maintainable and extensible for future enhancements.

### Example

#### Pormpt

![image](https://github.com/user-attachments/assets/f49b0bcc-88b2-4509-bf50-995b9a40038c)

#### Output

![image](https://github.com/user-attachments/assets/ff64385d-3822-4160-8cdf-f8a484ccc01a)

## Features

### Document Protection Status
The system can detect and return the protection status of Word documents using the `get_protection_status()` method, which returns a JSON object with:
- `is_protected`: Boolean indicating if the document is protected
- `protection_type`: Human-readable protection type based on Microsoft's WdProtectionType enumeration

Protection types include:
- `No protection` (-1): Document is not protected
- `Allow only revisions` (0): Only revisions to existing content are allowed
- `Allow only comments` (1): Only comments can be added
- `Allow only form fields` (2): Content can only be added through form fields
- `Allow only reading` (3): Read-only access only

For implementation details, see: [WdProtectionType Enumeration](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdprotectiontype)

### Document Management

- Create new Word documents with metadata
- Extract text and analyze document structure
- View document properties and statistics
- List available documents in a directory
- Create copies of existing documents
- Merge multiple documents into a single document
- Convert Word documents to PDF format

### Quick Tools

- Add headings quickly with `add_heading_quick`
- Add paragraphs quickly with `add_paragraph_quick`
- Get document outline with `get_document_outline`

### Content Creation

- Add headings with different levels
- Insert paragraphs with optional styling
- Create tables with custom data
- Add images with proportional scaling
- Insert page breaks
- Add footnotes and endnotes to documents
- Convert footnotes to endnotes
- Customize footnote and endnote styling
- Create professional table layouts for technical documentation
- Design callout boxes and formatted content for instructional materials
- Build structured data tables for business reports with consistent styling

### Rich Text Formatting

- Format specific text sections (bold, italic, underline)
- Change text color and font properties
- Apply custom styles to text elements
- Search and replace text throughout documents
- Individual cell text formatting within tables
- Multiple formatting combinations for enhanced visual appeal
- Font customization with family and size control

### Table Formatting

- Format tables with borders and styles
- Create header rows with distinct formatting
- Apply cell shading and custom borders
- Structure tables for better readability
- Individual cell background shading with color support
- Alternating row colors for improved readability
- Enhanced header row highlighting with custom colors
- Cell text formatting with bold, italic, underline, color, font size, and font family
- Comprehensive color support with named colors and hex color codes
- Cell padding management with independent control of all sides
- Cell alignment (horizontal and vertical positioning)
- Cell merging (horizontal, vertical, and rectangular areas)
- Column width management with multiple units (points, percentage, auto-fit)
- Auto-fit capabilities for dynamic column sizing
- Professional callout table support with icon cells and styled content

### Comment Management

- Add comments to document elements with `add_comment`
- Retrieve all comments with `get_comments`
- Delete specific comments with `delete_comment`
- Delete all comments with `delete_all_comments`
- Edit existing comments with `edit_comment`
- Reply to comments with `reply_to_comment`
- Get complete comment thread with `get_comment_thread`

### Advanced Document Manipulation

- Delete paragraphs
- Create custom document styles
- Apply consistent formatting throughout documents
- Format specific ranges of text with detailed control
- Flexible padding units with support for points and percentage-based measurements
- Clear, readable table presentation with proper alignment and spacing

### Document Protection

- Add password protection to documents
- Implement restricted editing with editable sections
- Add digital signatures to documents
- Verify document authenticity and integrity

### Comment Extraction

- Extract all comments from a document
- Filter comments by author
- Get comments for specific paragraphs
- Access comment metadata (author, date, text)

## Installation

### Installing via Smithery

To install Office Word Document Server for Claude Desktop automatically via [Smithery](https://smithery.ai/server/@GongRzhe/Office-Word-MCP-Server):

```bash
npx -y @smithery/cli install @GongRzhe/Office-Word-MCP-Server --client claude
```

### Prerequisites

- Python 3.8 or higher
- pip package manager

### Basic Installation

```bash
# Clone the repository
git clone https://github.com/GongRzhe/Office-Word-MCP-Server.git
cd Office-Word-MCP-Server

# Install dependencies
pip install -r requirements.txt
```

### Using uv for Development

This project supports [uv](https://docs.astral.sh/uv/), a fast Python package installer and resolver, for managing dependencies:

```bash
# Clone the repository
git clone https://github.com/GongRzhe/Office-Word-MCP-Server.git
cd Office-Word-MCP-Server

# Install dependencies with uv
uv pip install -r requirements.txt
```

### Using the Setup Script

Alternatively, you can use the provided setup script which handles:

- Checking prerequisites
- Setting up a virtual environment
- Installing dependencies
- Generating MCP configuration

```bash
python setup_mcp.py
```

## Usage with Claude for Desktop

### Configuration

#### Method 1: After Local Installation

1. After installation, add the server to your Claude for Desktop configuration file:

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "python",
      "args": ["/path/to/word_mcp_server.py"]
    }
  }
}
```

#### Method 2: Without Installation (Using uvx)

1. You can also configure Claude for Desktop to use the server without local installation by using the uvx package manager:

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "uvx",
      "args": ["--from", "office-word-mcp-server", "word_mcp_server"]
    }
  }
}
```

2. Configuration file locations:

   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`

3. Restart Claude for Desktop to load the configuration.

### Example Operations

Once configured, you can ask Claude to perform operations like:

- "Create a new document called 'report.docx' with a title page"
- "Add a heading and three paragraphs to my document"
- "Insert a 4x4 table with sales data"
- "Format the word 'important' in paragraph 2 to be bold and red"
- "Search and replace all instances of 'old term' with 'new term'"
- "Create a custom style for section headings"
- "Apply formatting to the table in my document"
- "Extract all comments from my document"
- "Show me all comments by John Doe"
- "Get comments for paragraph 3"
- "Make the text in table cell (1,2) bold and blue with 14pt font"
- "Add 10 points of padding to all sides of the header cells"
- "Create a callout table with a blue checkmark icon and white text"
- "Set the first column width to 50 points and auto-fit the remaining columns"
- "Apply alternating row colors to make the table more readable"
- "Add a paragraph at the beginning of my document"
- "Insert a heading at position 5 in my document"
- "Add a table after paragraph 3"
- "Select the first 3 paragraphs of my document"
- "Select paragraphs from index 2 to 5"
- "Show me the outline of my document"
- "Add a caption to the first picture in my document"
- "Add a caption 'Sales Data Visualization' to picture 2"
- "Add numbering to all paragraphs in my document"
- "Add numbering to paragraphs 2 through 5"


## AI大模型客户端使用指南

本节详细介绍AI大模型客户端如何使用Office-Word-MCP-Server进行Word文档操作，包括工具调用方法、参数说明和使用示例。

### 使用流程概述

1. **连接服务器**：确保MCP服务器已启动并与AI大模型客户端正确连接
2. **打开文档**：首先使用`open_document`工具打开目标Word文档
3. **执行操作**：调用相应工具进行文档编辑、内容提取等操作
4. **关闭文档**：完成操作后使用`shutdown_word`关闭Word应用实例

### 核心工具列表及参数说明

#### 文档管理工具

```
# 打开Word文档，这是使用其他工具的前提
open_document(file_path: str) -> str
# 参数：file_path - .docx文件的绝对路径
# 返回值：确认消息或错误信息

# 关闭文档并关闭Word应用实例
shutdown_word() -> str
# 返回值：关闭结果确认消息

# 获取文档结构（所有标题）
get_document_structure() -> List[Dict[str, Any]]
# 返回值：标题列表，每项包含text和level

# 接受文档中的所有修订
accept_all_changes() -> str
# 返回值：操作结果确认消息
```

#### 内容操作工具

```
# 插入新段落
insert_paragraph(locator: Dict[str, Any], text: str, position: str = "after") -> str
# 参数：
#   locator - 定位锚点元素的定位器
#   text - 要插入的段落文本
#   position - 相对于锚点的位置（"before"或"after"）
# 返回值：操作结果确认消息

# 获取指定元素的文本或指定范围的文本
get_text(locator: Dict[str, Any] = None, start_pos: int = None, end_pos: int = None) -> str
# 参数：
#   locator - 可选，定位目标元素的定位器
#   start_pos - 可选，文本范围的起始位置（整数）
#   end_pos - 可选，文本范围的结束位置（整数）
# 返回值：元素文本内容或范围内的文本内容

# 替换指定元素的文本
replace_text(locator: Dict[str, Any], new_text: str) -> str
# 参数：
#   locator - 定位目标元素的定位器
#   new_text - 替换的新文本
# 返回值：操作结果确认消息

# 删除指定元素
delete_element(locator: Dict[str, Any]) -> str
# 参数：locator - 定位目标元素的定位器
# 返回值：操作结果确认消息

# 创建表格
create_table(locator: Dict[str, Any], rows: int, cols: int) -> str
# 参数：
#   locator - 定位表格插入位置的定位器
#   rows - 表格行数
#   cols - 表格列数
# 返回值：操作结果确认消息

# 创建项目符号列表
create_bulleted_list(locator: Dict[str, Any], items: List[str], position: str = "after") -> str
# 参数：
#   locator - 定位列表插入位置的定位器
#   items - 列表项内容列表
#   position - 相对于锚点的位置
# 返回值：操作结果确认消息
```

#### 表格操作工具

```
# 获取表格单元格文本
get_text_from_cell(locator: Dict[str, Any]) -> str
# 参数：locator - 定位表格单元格的定位器
# 返回值：单元格文本内容

# 设置表格单元格值
set_cell_value(locator: Dict[str, Any], text: str) -> str
# 参数：
#   locator - 定位表格单元格的定位器
#   text - 要设置的单元格文本
# 返回值：操作结果确认消息
```

#### 页眉页脚工具

```
# 设置文档页眉文本
set_header_text(text: str) -> str
# 参数：text - 页眉文本
# 返回值：操作结果确认消息

# 设置文档页脚文本
set_footer_text(text: str) -> str
# 参数：text - 页脚文本
# 返回值：操作结果确认消息
```

#### 格式设置工具

```
# 应用格式到指定元素
apply_format(locator: Dict[str, Any], formatting: Dict[str, Any]) -> str
# 参数：
#   locator - 定位目标元素的定位器
#   formatting - 格式设置字典，如{"bold": True, "alignment": "center"}
# 返回值：操作结果确认消息
```

### Locator使用详解

Locator是一个特殊的查询对象，用于精确定位Word文档中的元素。它支持多种定位方式和过滤条件，使AI大模型能够精确地找到需要操作的文档部分。

#### 基本结构

Locator是一个字典，必须包含`target`字段，`target`内部包含以下组件：
- `type`: 元素类型（如"paragraph", "table", "cell", "heading"等）
- `filters`: 过滤条件列表（可选）

对于相对定位，Locator还可以包含：
- `anchor`: 锚点元素的定位器（包含`type`和`filters`）
- `relation`: 相对位置描述

#### 支持的元素类型

- `paragraph`: 段落
- `table`: 表格
- `cell`: 表格单元格
- `heading`: 标题
- `image`: 图片
- `list`: 列表

#### 常用过滤器

- `contains_text`: 文本包含指定内容
- `text_matches_regex`: 文本匹配正则表达式
- `index_in_parent`: 元素在父容器中的索引位置
- `style`: 元素样式名称
- `is_bold`: 是否为粗体
- `row_index`: 表格行索引
- `column_index`: 表格列索引
- `is_list_item`: 是否为列表项

#### 相对定位

通过`relation`字段可以指定元素相对于其他元素的位置，如：
- `all_occurrences_within`
- `first_occurrence_after`
- `last_occurrence_before`

### 使用示例

下面是AI大模型客户端调用MCP工具的典型示例：

#### 1. 打开文档

```
# 调用示例
response = await mcp_client.call_tool(
    "open_document",
    {
        "file_path": "C:/Users/username/Documents/report.docx"
    }
)
print(response)  # "Active document set to: C:/Users/username/Documents/report.docx"
```

#### 2. 读取文档结构

```
# 获取文档标题结构
headings = await mcp_client.call_tool("get_document_outline", {})
print(headings)
# 输出示例: [{"text": "Introduction", "level": 1}, {"text": "Methods", "level": 1}, ...]
```

#### 3. 插入段落

```
# 在文档开头插入段落
response = await mcp_client.call_tool(
    "insert_paragraph",
    {
        "locator": {"target": {"type": "paragraph", "filters": [{"index_in_parent": 0}]}},
        "text": "This is a new paragraph at the beginning of the document.",
        "position": "before"
    }
)
print(response)  # "Successfully inserted paragraph."
```

#### 4. 查找并替换文本

```
# 查找包含特定文本的段落并替换
response = await mcp_client.call_tool(
    "replace_text",
    {
        "locator": {"target": {"type": "paragraph", "filters": [{"contains_text": "old information"}]}},
        "new_text": "Updated content with new information."
    }
)
print(response)  # "Successfully replaced text."
```

#### 5. 格式化文本

```
# 将第一段设置为粗体并居中对齐
response = await mcp_client.call_tool(
    "apply_format",
    {
        "locator": {"target": {"type": "paragraph", "filters": [{"index_in_parent": 0}]}},
        "formatting": {"bold": True, "alignment": "center"}
    }
)
print(response)  # "Formatting applied successfully."
```

#### 6. 操作表格

```
# 在文档末尾创建一个3x4的表格
response = await mcp_client.call_tool(
    "create_table",
    {
        "locator": {"target": {"type": "paragraph", "filters": [{"index_in_parent": -1}]}},  # 最后一个段落
        "rows": 3,
        "cols": 4
    }
)
print(response)  # "Successfully created table."

# 设置表格单元格值
response = await mcp_client.call_tool(
    "set_cell_value",
    {
        "locator": {
            "target": {
                "type": "cell", 
                "filters": [
                    {"table_index": 0},
                    {"row_index": 0},
                    {"column_index": 0}
                ]
            }
        },
        "text": "Header 1"
    }
)
print(response)  # "Successfully set cell value."
```

#### 7. 设置页眉页脚

```
# 设置页眉文本
response = await mcp_client.call_tool(
    "set_header_text",
    {"text": "Confidential Report - 2023"}
)
print(response)  # "Header text set successfully."

# 设置页脚文本
response = await mcp_client.call_tool(
    "set_footer_text",
    {"text": "Page [Page] of [Pages]"}
)
print(response)  # "Footer text set successfully."
```

#### 8. 关闭文档

```
# 关闭Word应用
response = await mcp_client.call_tool("shutdown_word", {})
print(response)  # "Word application shut down successfully."
```

### 最佳实践

1. **会话管理**：在一个会话中完成一组相关操作，最后调用`shutdown_word`释放资源
2. **错误处理**：检查工具返回值，处理可能的错误信息
3. **定位策略**：使用精确的locator避免误操作
4. **路径处理**：确保提供的文件路径是绝对路径
5. **文档检查**：使用`open_document`前确认文件存在且格式正确
6. **资源释放**：长时间不使用时关闭Word应用实例

## Development

### Code Formatting and Static Analysis

This project uses several tools to maintain code quality and consistency:

- [Black](https://github.com/psf/black) for code formatting
- [isort](https://pycqa.github.io/isort/) for import sorting
- [mypy](http://mypy-lang.org/) for static type checking

To format the code and sort imports:

```bash
black word_document_server
isort word_document_server
```

To run static type checking:

```bash
mypy word_document_server
```

These tools help ensure code consistency and catch potential type-related errors before runtime.

## API Reference

### Document Creation and Properties

```
create_document(filename, title=None, author=None)
get_document_info(filename)
get_all_text(filename)
get_document_outline(filename)
list_opened_documents()
copy_document(source_filename, destination_filename=None)
convert_to_pdf(filename, output_filename=None)
```

### Content Addition

```
add_heading(filename, text, level=1, paragraph_index=None)
add_paragraph(filename, text, style=None, paragraph_index=None)
add_table(filename, rows, cols, data=None, paragraph_index=None)
add_picture(filename, image_path, width=None, paragraph_index=None)
add_page_break(filename, paragraph_index=None)
select_paragraphs(filename, start_index, end_index=None)
add_picture_caption(filename, caption_text, picture_index=None, paragraph_index=None)
add_paragraph_numbering(filename, start_index=0, end_index=None, style="Normal")
```

### Content Extraction

```
get_paragraph_text_from_document(filename, paragraph_index)
find_text_in_document(filename, text_to_find, match_case=True, whole_word=False)
```

### Text Formatting

```
format_text(filename, paragraph_index, start_pos, end_pos, bold=None,
            italic=None, underline=None, color=None, font_size=None, font_name=None)
search_and_replace(filename, find_text, replace_text)
delete_paragraph(filename, paragraph_index)
create_custom_style(filename, style_name, bold=None, italic=None,
                    font_size=None, font_name=None, color=None, base_style=None)
```

### Table Formatting

```
format_table(filename, table_index, has_header_row=None,
             border_style=None, shading=None)
set_table_cell_shading(filename, table_index, row_index, col_index, 
                      fill_color, pattern="clear")
apply_table_alternating_rows(filename, table_index, 
                            color1="FFFFFF", color2="F2F2F2")
highlight_table_header(filename, table_index, 
                      header_color="4472C4", text_color="FFFFFF")

# Cell merging tools
merge_table_cells(filename, table_index, start_row, start_col, end_row, end_col)
merge_table_cells_horizontal(filename, table_index, row_index, start_col, end_col)
merge_table_cells_vertical(filename, table_index, col_index, start_row, end_row)

# Cell alignment tools
set_table_cell_alignment(filename, table_index, row_index, col_index,
                        horizontal="left", vertical="top")
set_table_alignment_all(filename, table_index, 
                       horizontal="left", vertical="top")

# Cell text formatting tools
format_table_cell_text(filename, table_index, row_index, col_index,
                      text_content=None, bold=None, italic=None, underline=None,
                      color=None, font_size=None, font_name=None)

# Cell padding tools
set_table_cell_padding(filename, table_index, row_index, col_index,
                      top=None, bottom=None, left=None, right=None, unit="points")

# Column width management
set_table_column_width(filename, table_index, col_index, width, width_type="points")
set_table_column_widths(filename, table_index, widths, width_type="points")
set_table_width(filename, table_index, width, width_type="points")
auto_fit_table_columns(filename, table_index)
```

### Comment Extraction

```
get_all_comments(filename)
get_comments_by_author(filename, author)
get_comments_for_paragraph(filename, paragraph_index)
```

## Troubleshooting

### Common Issues

1. **Missing Styles**

   - Some documents may lack required styles for heading and table operations
   - The server will attempt to create missing styles or use direct formatting
   - For best results, use templates with standard Word styles

2. **Permission Issues**

   - Ensure the server has permission to read/write to the document paths
   - Use the `copy_document` function to create editable copies of locked documents
   - Check file ownership and permissions if operations fail

3. **Image Insertion Problems**
   - Use absolute paths for image files
   - Verify image format compatibility (JPEG, PNG recommended)
   - Check image file size and permissions

4. **Table Formatting Issues**

   - **Cell index errors**: Ensure row and column indices are within table bounds (0-based indexing)
   - **Color format problems**: Use hex colors without '#' prefix (e.g., "FF0000" for red) or standard color names
   - **Padding unit confusion**: Specify "points" or "percent" explicitly when setting cell padding
   - **Column width conflicts**: Auto-fit may override manual column width settings
   - **Text formatting persistence**: Apply cell text formatting after setting cell content for best results

### Debugging

Enable detailed logging by setting the environment variable:

```bash
export MCP_DEBUG=1  # Linux/macOS
set MCP_DEBUG=1     # Windows
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- [Model Context Protocol](https://modelcontextprotocol.io/) for the protocol specification
- [python-docx](https://python-docx.readthedocs.io/) for Word document manipulation
- [MCP](https://github.com/modelcontextprotocol/python-sdk) for the Python MCP implementation

---

_Note: This server interacts with document files on your system. Always verify that requested operations are appropriate before confirming them in Claude for Desktop or other MCP clients._
