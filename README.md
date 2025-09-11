# Word Document MCP Server

`word_docx_tools` implements the [Model Context Protocol](https://modelcontextprotocol.io/) to expose Word document operations as tools and resources. It serves as a bridge between AI assistants and Microsoft Word documents, allowing for direct document creation, content addition, formatting, and analysis through COM interface with full Office object access.

## Features

- Create and manipulate Word documents
- Text operations (insert, replace, format)
- Table operations
- Image operations
- Comment operations
- Object operations
- Style operations
- Advanced selector engine for targeting document elements

## Installation

### Prerequisites

- Windows 10/11
- Microsoft Word (2016 or later)
- Python 3.11+

### Install from PyPI

```bash
pip install word_docx_tools
```

### Install from Source

```bash
git clone <repository-url>
cd word_docx_tools
pip install -e .
```

## Usage

### Run as MCP Server

```bash
word_docx_tools
```

### Development Mode

```bash
cd word_docx_tools
python -m word_docx_tools.main
```

Or:

```bash
python -m word_docx_tools
```

### HTTP Mode

```bash
word_docx_tools_http
```

This starts the server on `http://0.0.0.0:8000`.

## Docker

### Build

```bash
docker build -t word_docx_tools .
```

### Run

```bash
docker run -it --rm word_docx_tools
```

## Development

### Code Structure

```
word_docx_tools/
├── main.py              # Entry point
├── mcp_service/         # MCP service integration
├── selector/            # Document selector engine
├── operations/          # Low-level document operations
├── tools/               # MCP tools exposed to clients
├── com_backend/         # COM interface handling
└── utils/               # Utility functions
```

### Code Quality

```bash
black word_docx_tools
isort word_docx_tools
mypy word_docx_tools
```

### Testing

```bash
python -m pytest tests/
```

## Configuration

### MCP Server Configuration

MCP server configuration can be specified in your AI assistant's configuration file. For Claude Desktop, this is typically located at:

- Windows: `%APPDATA%\Claude\claude_desktop_config.json`

Example configuration:
```json
{
  "mcpServers": {
    "word-docx-tools": {
      "command": "word_docx_tools"
    }
  }
}
```

### Project Configuration

MCP configuration is defined in `mcp-config.json`:

```json
{
    "mcpServers": {
        "word_docx_tools": {
            "command": "word_docx_tools",
            "args": []
        }
    }
}
```

## Example Operations

Once configured, you can ask your AI assistant to perform operations like:

- "Create a new document called 'report.docx' with direct Office object access"
- "Add a heading and three paragraphs to my document with direct Office formatting"
- "Insert a 4x4 table with sales data and directly manipulate Office table objects"
- "Format the word 'important' in paragraph 2 to be bold and red through direct Office object access"
- "Add a comment to the first paragraph and directly manage Office comment objects"
- "Insert an image after paragraph 3 and directly control Office image properties with full object access"

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Model Context Protocol](https://modelcontextprotocol.io/) for the protocol specification
- [pywin32](https://pypi.org/project/pywin32/) for COM integration with Word
- [MCP Python SDK](https://github.com/modelcontextprotocol/python-sdk) for the Python MCP implementation
- [Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server.git) as a reference project

