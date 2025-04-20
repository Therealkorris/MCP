# MCP-Visio: Microsoft Visio Integration for AI Agents

MCP-Visio is a server that implements the Model Context Protocol (MCP) to enable AI agents to interact with Microsoft Visio diagrams. It provides tools for analyzing, modifying, and working with Visio diagrams programmatically.

## Features

- **Analyze Visio Diagrams**: Extract information about shapes, connections, and text
- **Modify Diagrams**: Add, update, or delete shapes and connections
- **Verify Connections**: Check connections between shapes
- **Ollama Integration**: Leverage local language models for diagram analysis and suggestions

## Architecture

- **MCP Server**: Implements the Model Context Protocol for AI integration
- **Visio Service**: Handles all Visio-related operations via COM automation
- **Ollama Service**: Provides AI capabilities using local models

## Prerequisites

- Windows 10/11 (required for Microsoft Visio)
- Python 3.11+ 
- Microsoft Visio (licensed version)
- Ollama (optional, for AI features)

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/mcp-visio.git
   cd mcp-visio
   ```

2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Create a `.env` file based on `env.example`:
   ```
   cp env.example .env
   ```

4. Configure your environment variables in the `.env` file

## Running the Server

### Local Mode

Run the server with Server-Sent Events (SSE) transport:

```
python src/main.py --transport sse --host 0.0.0.0 --port 8050
```

Or use STDIO transport:

```
python src/main.py --transport stdio
```

### Docker Mode

Build and run using Docker Compose:

```
docker-compose up -d
```

This will start the MCP-Visio server in a Docker container, connecting to the local Visio instance on your Windows host.

## Integration with MCP Clients

### SSE Configuration

Once you have the server running with SSE transport, you can connect to it using this configuration:

```json
{
  "mcpServers": {
    "visio": {
      "transport": "sse",
      "serverUrl": "http://localhost:8050/sse"
    }
  }
}
```

### STDIO Configuration

For Claude Desktop, Windsurf, or other MCP clients, use this configuration:

```json
{
  "mcpServers": {
    "visio": {
      "command": "python",
      "args": ["path/to/mcp-visio/src/main.py"],
      "env": {
        "TRANSPORT": "stdio",
        "VISIO_SERVICE_HOST": "localhost",
        "VISIO_SERVICE_PORT": "8051",
        "OLLAMA_API_URL": "http://localhost:11434",
        "OLLAMA_MODEL": "llama3"
      }
    }
  }
}
```

## Testing

Run the test client to verify server functionality:

```
python test_client.py
```

## Tools Available

The MCP-Visio server provides the following tools to AI agents:

1. **analyze_visio_diagram**: Get detailed information about a Visio diagram
2. **modify_visio_diagram**: Make changes to a Visio diagram
3. **get_active_document**: Get information about the currently open Visio document
4. **verify_connections**: Check connections between shapes in a diagram

## License

[MIT License](LICENSE)

## Acknowledgements

- Based on the Model Context Protocol specification by Anthropic
- Inspired by various MCP server implementations in the community 