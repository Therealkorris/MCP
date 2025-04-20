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
                    "description": "File path to the Visio diagram"
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
                    "description": "File path to the Visio diagram"
                },
                "operation": {
                    "type": "string",
                    "description": "Type of modification to perform",
                    "enum": ["add_shape", "update_shape", "delete_shape", "add_connection", "delete_connection"]
                },
                "shape_data": {
                    "type": "object",
                    "description": "Data for the shape operation"
                }
            },
            "required": ["file_path", "operation"]
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
                    "description": "File path to the Visio diagram"
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