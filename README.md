# MCP-Visio Server

A powerful server that integrates Microsoft Visio with AI assistants through MCP (Machine Communication Protocol).

## Architecture

The MCP-Visio system consists of two main components:

1. **MCP Server**: A containerized service that provides API endpoints for AI assistants to interact with Visio diagrams.
   
2. **Visio Relay Service**: A local service that runs on the Windows host with Visio installed and provides communication between the MCP Server and the Visio application.

```
  ┌───────────────┐          ┌───────────────┐         ┌───────────────┐
  │    AI Agent   │  <--->   │   MCP Server  │  <--->  │ Visio Relay   │  <--->  Microsoft Visio
  │  (Claude/GPT) │          │  (Container)  │         │   Service     │         (Windows Host)
  └───────────────┘          └───────────────┘         └───────────────┘
```

## Features

- **Diagram Analysis**: Extract shapes, connections, text, and layout information
- **Diagram Modification**: Add, update, or delete shapes and connections
- **Active Document Access**: Work with the current open Visio document
- **Connection Verification**: Verify connections between shapes
- **Document Management**: Create new diagrams, save diagrams
- **Shape Management**: Get detailed shape information
- **AI Integration**: Integrate diagram analysis with AI models

## Prerequisites

- Windows OS with Microsoft Visio installed
- Docker Desktop for Windows
- Python 3.8+ installed on the host
- Network connectivity between components

## Setup Instructions

### 1. Clone the Repository

```bash
git clone https://github.com/Therealkorris/mcp-visio.git
cd mcp-visio
```

### 2. Start the MCP-Visio System

The system is now managed through a single batch file:

```bash
run_mcp_visio.bat
```

This will present you with a menu to:
1. Start everything (both relay and MCP server)
2. Start only the Visio relay service
3. Start only the MCP Docker container
4. Run tests
5. Stop all services

### 3. Testing the Connection

To test if everything is working correctly:

```bash
run_mcp_visio.bat test
```

or select option 4 from the menu.

## API Features

The MCP-Visio server provides the following tools for AI assistants:

- `analyze_visio_diagram`: Analyze shapes, connections, and text in a diagram
- `modify_visio_diagram`: Add, update, or delete shapes and connections
- `get_active_document`: Get information about the current Visio document
- `verify_connections`: Verify connections between shapes
- `create_new_diagram`: Create a new Visio diagram from a template
- `save_diagram`: Save the current diagram to disk
- `get_available_stencils`: Get a list of available Visio stencils
- `get_shapes_on_page`: Get detailed information about all shapes on a page
- `export_diagram`: Export a diagram to another format (PNG, JPG, PDF)

## Usage Examples

### Analyzing a Diagram

```python
import requests

response = requests.post("http://localhost:8050", json={
    "jsonrpc": "2.0",
    "id": 1,
    "method": "analyze_visio_diagram",
    "params": {
        "file_path": "active",
        "analysis_type": "all"
    }
})

result = response.json()
print(result)
```

### Adding a Shape

```python
import requests

response = requests.post("http://localhost:8050", json={
    "jsonrpc": "2.0",
    "id": 2,
    "method": "modify_visio_diagram",
    "params": {
        "file_path": "active",
        "operation": "add_shape",
        "shape_data": {
            "master_name": "Rectangle",
            "stencil_name": "Basic Shapes.vss",
            "page_index": 1,
            "x": 4.0,
            "y": 4.0,
            "text": "New Shape"
        }
    }
})

result = response.json()
print(result)
```

## File Structure

```
mcp-visio/
├── src/                          # Server source code
│   ├── services/                 # Service implementations
│   │   ├── visio_service.py      # Visio service implementation
│   │   ├── mcp_service.py        # MCP service implementation
│   │   └── ollama_service.py     # AI integration service
├── visio-files/                  # Directory for Visio files
├── Dockerfile.windows            # Docker configuration for Windows
├── host_visio_relay.py           # Visio relay service
├── run_mcp_visio.bat             # Main execution script
├── test_visio_connection.py      # Test script
├── visio_api.py                  # API definitions
└── README.md                     # This documentation
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 