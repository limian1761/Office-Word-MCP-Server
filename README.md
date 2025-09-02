// DirectOfficeWordMCP

A Model Context Protocol (MCP) server for directly manipulating Microsoft Word documents via COM interface with direct Office object access. This server enables AI assistants to work with Word documents through a standardized interface, providing rich document editing capabilities with direct object access.

## Overview

DirectOfficeWordMCP implements the [Model Context Protocol](https://modelcontextprotocol.io/) to expose Word document operations as tools and resources. It serves as a bridge between AI assistants and Microsoft Word documents, allowing for direct document creation, content addition, formatting, and analysis through COM interface with full Office object access.

This version has been significantly restructured and enhanced with:

- Complete COM backend integration for native Word operations with direct Office object access
- Advanced document selector system for precise element targeting in Office documents
- Modular architecture separating concerns into core functionality, operations, and tools
- Support for multiple transport protocols (stdio, HTTP, SSE)
- Docker containerization support
- Enhanced error handling and logging

## Features

### Document Operations
- Create, open, save, and close Word documents
- Get document information and properties
- Document protection and encryption
- Print document support

### Text Operations
- Insert, replace, and format text
- Paragraph manipulation
- Text search and replace
- Character and paragraph formatting

### Table Operations
- Create and manipulate tables
- Cell content and formatting
- Row and column insertion
- Table styling and formatting

### Image Operations
- Insert images with positioning
- Image resizing and formatting
- Caption management

### Comment Operations
- Add, edit, and delete comments
- Comment threading and replies
- Author-based filtering

### Style Operations
- Apply formatting to elements
- Font and paragraph styling
- Custom style creation

### Object Operations
- Bookmark management
- Hyperlink creation
- Citation handling

## Installation

### Prerequisites
- Windows operating system (required for COM integration with Word)
- Python 3.11 or higher
- Microsoft Word installation

### Installation Options

#### Option 1: Direct Installation
```bash
# Clone the repository
git clone <repository-url>
cd Office-Word-MCP-Server

# Install dependencies
pip install -r requirements.txt
```

#### Option 2: Using uv (Recommended)
```bash
# Clone the repository
git clone https://github.com/GongRzhe/Office-Word-MCP-Server.git
cd Office-Word-MCP-Server

# Install with uv
uv pip install -r requirements.txt
```

#### Option 3: Using the Setup Script
```bash
python setup_mcp.py
```

The setup script will:
- Check prerequisites
- Set up a virtual environment
- Install dependencies
- Generate MCP configuration

## Usage

### Running the Server

#### Via Direct Execution
```bash
python -m word_document_server.main
```

#### Via Installed Script
```bash
directofficeword_mcp
```

### Transport Options

The server supports multiple transport protocols:

1. **STDIO** (default): For local AI assistant integration
2. **HTTP**: For web-based deployments
3. **SSE** (Server-Sent Events): For compatibility scenarios

Configure the transport method using command-line arguments or environment variables.

## Deployment

See [DEPLOYMENT.md](DEPLOYMENT.md) for detailed deployment instructions including Docker deployment, environment configuration, and best practices.

### Docker Deployment

The server can be deployed as a Docker container:

```bash
# Build the image
docker build -t office-word-mcp-server .

# Run the container
docker run -it --rm office-word-mcp-server
```

Note: Docker deployment has limitations with COM objects and may require additional configuration for full Word integration.

## Development

### Code Structure
```
word_document_server/
├── com_backend/       # COM integration layer
├── mcp_service/       # MCP protocol implementation
├── operations/        # Core document operations
├── selector/          # Element selection system
├── tools/             # MCP tool implementations
└── utils/             # Utility functions
```

### Code Quality Tools
- [Black](https://github.com/psf/black) for code formatting
- [isort](https://pycqa.github.io/isort/) for import sorting
- [mypy](http://mypy-lang.org/) for static type checking

Run these tools before committing:
```bash
black word_document_server
isort word_document_server
mypy word_document_server
```

## Configuration

MCP server configuration can be specified in your AI assistant's configuration file. For Claude Desktop, this is typically located at:

- Windows: `%APPDATA%\Claude\claude_desktop_config.json`

Example configuration:
```json
{
  "mcpServers": {
    "directofficeword-mcp": {
      "command": "directofficeword_mcp"
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

---