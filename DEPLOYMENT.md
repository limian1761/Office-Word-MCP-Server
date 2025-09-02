# Deployment Guide

This guide provides instructions for deploying the Word Document MCP Server in various environments.

## Deployment Options

### 1. Local Installation (Recommended for Windows with Word)

For the best experience with full Word COM integration, install directly on a Windows machine with Microsoft Word:

```bash
# Clone the repository
git clone https://github.com/GongRzhe/Office-Word-MCP-Server.git
cd Office-Word-MCP-Server

# Install with pip
pip install .

# Run the server
directofficeword_mcp
```

### 2. Docker Deployment

Docker deployment is possible but has limitations with COM objects and Word integration.

#### Linux Container Deployment

```bash
# Build the image
docker build -t word-mcp-server .

# Run with stdio (for direct integration)
docker run -it --rm word-mcp-server

# Run with HTTP server
docker run -d -p 8000:8000 word-mcp-server
```

#### Windows Container Deployment

For full Word integration, Windows containers are required:

```bash
# Build Windows container (requires Windows host)
docker build -t word-mcp-server --target windows-base .

# Run Windows container with proper configuration
docker run -d -p 8000:8000 word-mcp-server
```

**Note**: Windows container deployment with full COM support requires additional configuration and is not fully supported in all environments.

### 3. Docker Compose Deployment

Use the provided docker-compose.yml for easier deployment:

```bash
# Development setup
docker-compose up word-mcp-dev

# Production setup
docker-compose up word-mcp-prod
```

## Configuration

### Environment Variables

The server can be configured using environment variables:

- `MCP_TRANSPORT`: Transport protocol (stdio, http, sse)
- `HOST`: Host address for HTTP/SSE transport
- `PORT`: Port number for HTTP/SSE transport

Example:
```bash
MCP_TRANSPORT=http HOST=0.0.0.0 PORT=8000 directofficeword_mcp
```

### MCP Client Configuration

Configure your MCP client (e.g., Claude Desktop) to connect to the server:

```json
{
  "mcpServers": {
    "word-mcp": {
      "command": "directofficeword_mcp"
    }
  }
}
```

For HTTP transport:
```json
{
  "mcpServers": {
    "word-mcp": {
      "url": "http://localhost:8000/mcp"
    }
  }
}
```

## Limitations

### Docker & Container Limitations

1. **COM Object Integration**: Docker containers have limited access to Windows COM objects, which are required for Word integration.

2. **GUI Applications**: Word is a GUI application that may not function properly in containerized environments.

3. **Licensing**: Microsoft Word licensing in containers may have restrictions.

### Performance Considerations

1. **COM Object Overhead**: Creating and destroying COM objects has overhead. Consider keeping connections alive for better performance.

2. **Memory Usage**: Word instances consume significant memory. Monitor resource usage in production.

## Best Practices

### For Development

1. Use local installation for full functionality
2. Run tests to verify Word integration
3. Use development mode with auto-reload for faster iteration

### For Production

1. Use local Windows installation with Word for full functionality
2. Monitor resource usage
3. Implement proper error handling and logging
4. Consider using process isolation for multiple concurrent users
5. Implement authentication and authorization if exposing over network

### For Containerized Environments

1. Understand the limitations of COM object access
2. Consider using the server in stdio mode with a wrapper
3. Monitor for memory leaks in long-running containers
4. Use appropriate restart policies

## Troubleshooting

### Common Issues

1. **COM Object Errors**: Ensure Word is properly installed and accessible
2. **Permission Issues**: Run with appropriate permissions for COM access
3. **Port Conflicts**: Check that configured ports are available
4. **Path Issues**: Use absolute paths for document operations

### Logs and Monitoring

Enable logging for troubleshooting:

```bash
# Enable debug logging
export LOG_LEVEL=DEBUG
directofficeword_mcp
```

## Security Considerations

1. **File System Access**: The server can access any file the running process has permissions to access
2. **Network Exposure**: When using HTTP/SSE transport, ensure proper network security
3. **Input Validation**: Validate all inputs to prevent injection attacks
4. **Process Isolation**: Consider running in isolated environments for untrusted inputs