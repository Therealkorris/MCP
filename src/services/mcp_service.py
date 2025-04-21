"""
MCP service implementation for Visio.
Handles all MCP JSON-RPC message processing.
"""

import json
import logging
import traceback
from typing import Dict, Any, List, Optional

from services.visio_service import VisioService
from services.ollama_service import OllamaService

# Configure logging
logger = logging.getLogger(__name__)

# Initialize services
visio_service = VisioService()
ollama_service = OllamaService()

# Define MCP tools
TOOLS = [
    {
        "name": "analyze_visio_diagram",
        "description": "Analyze a Visio diagram to extract information about shapes, connections, and layout",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string", 
                    "description": "File path to the Visio diagram or 'active' to use the active document"
                },
                "analysis_type": {
                    "type": "string",
                    "description": "Type of analysis to perform",
                    "enum": ["structure", "connections", "text", "all"],
                    "default": "all"
                }
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "modify_visio_diagram",
        "description": "Modify a Visio diagram by adding, updating, or deleting shapes and connections",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string", 
                    "description": "File path to the Visio diagram or 'active' to use the active document"
                },
                "operation": {
                    "type": "string",
                    "description": "Type of modification to perform",
                    "enum": ["add_shape", "update_shape", "delete_shape", "add_connection", "add_connector", "delete_connection"]
                },
                "shape_data": {
                    "type": "object",
                    "description": "Data for the shape operation",
                    "oneOf": [
                        {
                            "title": "Add Shape Data",
                            "type": "object",
                            "properties": {
                                "master_name": {
                                    "type": "string",
                                    "description": "Name of the master shape to add (e.g., 'Rectangle', 'Circle', 'Ellipse', 'Triangle', 'Diamond')"
                                },
                                "stencil_name": {
                                    "type": "string",
                                    "description": "Name of the stencil containing the master shape (e.g., 'Basic Shapes.vss')",
                                    "default": "Basic Shapes.vss"
                                },
                                "position": {
                                    "type": "object",
                                    "description": "Position of the shape",
                                    "properties": {
                                        "x": {
                                            "type": "number",
                                            "description": "X-coordinate of the shape"
                                        },
                                        "y": {
                                            "type": "number",
                                            "description": "Y-coordinate of the shape"
                                        }
                                    }
                                },
                                "size": {
                                    "type": "object",
                                    "description": "Size of the shape",
                                    "properties": {
                                        "width": {
                                            "type": "number",
                                            "description": "Width of the shape"
                                        },
                                        "height": {
                                            "type": "number",
                                            "description": "Height of the shape"
                                        }
                                    }
                                },
                                "text": {
                                    "type": "string",
                                    "description": "Text to display in the shape"
                                },
                                "page_index": {
                                    "type": "integer",
                                    "description": "Index of the page to add the shape to",
                                    "default": 1
                                }
                            },
                            "required": ["master_name"]
                        },
                        {
                            "title": "Update Shape Data",
                            "type": "object",
                            "properties": {
                                "shape_id": {
                                    "type": "integer",
                                    "description": "ID of the shape to update"
                                },
                                "text": {
                                    "type": "string",
                                    "description": "New text for the shape"
                                },
                                "position": {
                                    "type": "object",
                                    "description": "New position of the shape",
                                    "properties": {
                                        "x": {
                                            "type": "number",
                                            "description": "X-coordinate of the shape"
                                        },
                                        "y": {
                                            "type": "number",
                                            "description": "Y-coordinate of the shape"
                                        }
                                    }
                                },
                                "size": {
                                    "type": "object",
                                    "description": "New size of the shape",
                                    "properties": {
                                        "width": {
                                            "type": "number",
                                            "description": "Width of the shape"
                                        },
                                        "height": {
                                            "type": "number",
                                            "description": "Height of the shape"
                                        }
                                    }
                                },
                                "page_index": {
                                    "type": "integer",
                                    "description": "Index of the page containing the shape",
                                    "default": 1
                                }
                            },
                            "required": ["shape_id"]
                        },
                        {
                            "title": "Delete Shape Data",
                            "type": "object",
                            "properties": {
                                "shape_id": {
                                    "type": "integer",
                                    "description": "ID of the shape to delete"
                                },
                                "page_index": {
                                    "type": "integer",
                                    "description": "Index of the page containing the shape",
                                    "default": 1
                                }
                            },
                            "required": ["shape_id"]
                        },
                        {
                            "title": "Add Connector Data",
                            "type": "object",
                            "properties": {
                                "from_shape_id": {
                                    "type": "integer",
                                    "description": "ID of the shape to connect from"
                                },
                                "to_shape_id": {
                                    "type": "integer",
                                    "description": "ID of the shape to connect to"
                                },
                                "text": {
                                    "type": "string",
                                    "description": "Text to display on the connector"
                                },
                                "page_index": {
                                    "type": "integer",
                                    "description": "Index of the page containing the shapes",
                                    "default": 1
                                }
                            },
                            "required": ["from_shape_id", "to_shape_id"]
                        },
                        {
                            "title": "Delete Connection Data",
                            "type": "object",
                            "properties": {
                                "connector_id": {
                                    "type": "integer",
                                    "description": "ID of the connector to delete"
                                },
                                "page_index": {
                                    "type": "integer",
                                    "description": "Index of the page containing the connector",
                                    "default": 1
                                }
                            },
                            "required": ["connector_id"]
                        }
                    ]
                }
            },
            "required": ["file_path", "operation", "shape_data"]
        }
    },
    {
        "name": "get_active_document",
        "description": "Get information about the currently active Visio document",
        "parameters": {
            "type": "object",
            "properties": {}
        }
    },
    {
        "name": "verify_connections",
        "description": "Verify connections between shapes in a Visio diagram",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string", 
                    "description": "File path to the Visio diagram or 'active' to use the active document"
                },
                "shape_ids": {
                    "type": "array",
                    "description": "List of shape IDs to verify connections for",
                    "items": {
                        "type": "string"
                    }
                }
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "create_new_diagram",
        "description": "Create a new Visio diagram from a template",
        "parameters": {
            "type": "object",
            "properties": {
                "template": {
                    "type": "string",
                    "description": "Template to use for the new diagram",
                    "default": "Basic.vst"
                },
                "save_path": {
                    "type": "string",
                    "description": "Path where the new diagram should be saved"
                }
            }
        }
    },
    {
        "name": "save_diagram",
        "description": "Save the current Visio diagram",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to save the diagram, or 'active' to save the currently active document"
                }
            }
        }
    },
    {
        "name": "get_available_stencils",
        "description": "Get a list of available Visio stencils",
        "parameters": {
            "type": "object",
            "properties": {}
        }
    },
    {
        "name": "get_shapes_on_page",
        "description": "Get detailed information about all shapes on a page",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to the Visio diagram or 'active' to use the active document",
                    "default": "active"
                },
                "page_index": {
                    "type": "integer",
                    "description": "Index of the page to get shapes from",
                    "default": 1
                }
            }
        }
    },
    {
        "name": "export_diagram",
        "description": "Export a Visio diagram to another format (PNG, JPG, PDF, SVG)",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to the Visio diagram or 'active' to use the active document",
                    "default": "active"
                },
                "format": {
                    "type": "string",
                    "description": "Export format",
                    "enum": ["png", "jpg", "pdf", "svg"],
                    "default": "png"
                },
                "output_path": {
                    "type": "string",
                    "description": "Path to save the exported file"
                }
            }
        }
    },
    {
        "name": "get_available_masters",
        "description": "Get a list of available master shapes from all open stencils",
        "parameters": {
            "type": "object",
            "properties": {}
        }
    }
]

async def process_message(message: Dict[str, Any]) -> Dict[str, Any]:
    """Process an MCP message and return the response."""
    try:
        message_id = message.get("id")
        method = message.get("method")
        params = message.get("params", {})
        
        logger.debug(f"Processing message: {json.dumps(message, indent=2)}")
        
        # Basic validation
        if "jsonrpc" not in message or message.get("jsonrpc") != "2.0":
            logger.warning(f"Invalid jsonrpc version: {message.get('jsonrpc')}")
            return {
                "jsonrpc": "2.0",
                "id": message_id,
                "error": {
                    "code": -32600,
                    "message": "Invalid Request: jsonrpc version must be 2.0"
                }
            }
        
        # Handle get_client_info method (required for MCP)
        if method == "get_client_info":
            client_info = params.get("client_info", {})
            logger.info(f"Client connected: {json.dumps(client_info, indent=2)}")
            
            response = {
                "jsonrpc": "2.0",
                "id": message_id,
                "result": {
                    "server_info": {
                        "name": "MCP-Visio Server",
                        "version": "1.0.0"
                    },
                    "capabilities": {
                        "supports_tool_calls": True
                    },
                    "tools": TOOLS
                }
            }
            logger.debug(f"get_client_info response: {json.dumps(response, indent=2)}")
            return response
        
        # Handle ping method
        elif method == "ping":
            response = {
                "jsonrpc": "2.0",
                "id": message_id,
                "result": "pong"
            }
            logger.debug("ping response: pong")
            return response
        
        # Handle analyze_visio_diagram method
        elif method == "analyze_visio_diagram":
            file_path = params.get("file_path")
            analysis_type = params.get("analysis_type", "all")
            
            if not file_path:
                return {
                    "jsonrpc": "2.0",
                    "id": message_id,
                    "error": {
                        "code": -32602,
                        "message": "Invalid params: file_path is required"
                    }
                }
            
            result = visio_service.analyze_diagram(file_path, analysis_type)
            response = {
                "jsonrpc": "2.0",
                "id": message_id,
                "result": result
            }
            logger.debug(f"analyze_visio_diagram response completed")
            return response
        
        # Handle modify_visio_diagram method
        elif method == "modify_visio_diagram":
            file_path = params.get("file_path")
            operation = params.get("operation")
            shape_data = params.get("shape_data", {})
            
            if not file_path or not operation:
                return {
                    "jsonrpc": "2.0",
                    "id": message_id,
                    "error": {
                        "code": -32602,
                        "message": "Invalid params: file_path and operation are required"
                    }
                }
            
            result = visio_service.modify_diagram(file_path, operation, shape_data)
            response = {
                "jsonrpc": "2.0",
                "id": message_id,
                "result": result
            }
            logger.debug(f"modify_visio_diagram response completed")
            return response
        
        # Handle get_active_document method
        elif method == "get_active_document":
            result = visio_service.get_active_document()
            response = {
                "jsonrpc": "2.0",
                "id": message_id,
                "result": result
            }
            logger.debug(f"get_active_document response completed")
            return response
        
        # Handle verify_connections method
        elif method == "verify_connections":
            file_path = params.get("file_path")
            shape_ids = params.get("shape_ids", [])
            
            if not file_path:
                return {
                    "jsonrpc": "2.0",
                    "id": message_id,
                    "error": {
                        "code": -32602,
                        "message": "Invalid params: file_path is required"
                    }
                }
            
            result = visio_service.verify_connections(file_path, shape_ids)
            response = {
                "jsonrpc": "2.0",
                "id": message_id,
                "result": result
            }
            logger.debug(f"verify_connections response completed")
            return response

        # Handle create_new_diagram method
        elif method == "create_new_diagram":
            template = params.get("template", "Basic.vst")
            save_path = params.get("save_path")
            
            result = visio_service.create_new_diagram(template, save_path)
            response = {
                "jsonrpc": "2.0",
                "id": message_id,
                "result": result
            }
            logger.debug(f"create_new_diagram response completed")
            return response
        
        # Handle save_diagram method
        elif method == "save_diagram":
            file_path = params.get("file_path")
            
            result = visio_service.save_diagram(file_path)
            response = {
                "jsonrpc": "2.0",
                "id": message_id,
                "result": result
            }
            logger.debug(f"save_diagram response completed")
            return response
        
        # Handle get_available_stencils method
        elif method == "get_available_stencils":
            result = visio_service.get_available_stencils()
            response = {
                "jsonrpc": "2.0",
                "id": message_id,
                "result": result
            }
            logger.debug(f"get_available_stencils response completed")
            return response
        
        # Handle get_available_masters method
        elif method == "get_available_masters":
            result = visio_service.get_available_masters()
            response = {
                "jsonrpc": "2.0",
                "id": message_id,
                "result": result
            }
            logger.debug(f"get_available_masters response completed")
            return response
        
        # Handle get_shapes_on_page method
        elif method == "get_shapes_on_page":
            file_path = params.get("file_path", "active")
            page_index = params.get("page_index", 1)
            
            result = visio_service.get_shapes_on_page(file_path, page_index)
            response = {
                "jsonrpc": "2.0",
                "id": message_id,
                "result": result
            }
            logger.debug(f"get_shapes_on_page response completed")
            return response
        
        # Handle export_diagram method
        elif method == "export_diagram":
            file_path = params.get("file_path", "active")
            format = params.get("format", "png")
            output_path = params.get("output_path")
            
            result = visio_service.export_diagram(file_path, format, output_path)
            response = {
                "jsonrpc": "2.0",
                "id": message_id,
                "result": result
            }
            logger.debug(f"export_diagram response completed")
            return response
        
        # Handle method not found
        else:
            logger.warning(f"Method not found: {method}")
            return {
                "jsonrpc": "2.0",
                "id": message_id,
                "error": {
                    "code": -32601,
                    "message": f"Method '{method}' not found"
                }
            }
    
    except Exception as e:
        logger.exception(f"Error processing message: {e}")
        logger.error(traceback.format_exc())
        return {
            "jsonrpc": "2.0",
            "id": message.get("id"),
            "error": {
                "code": -32603,
                "message": f"Internal error: {str(e)}"
            }
        } 