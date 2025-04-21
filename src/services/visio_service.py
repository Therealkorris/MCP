"""
Visio service for MCP-Visio server.
Provides functionality to interact with Microsoft Visio using a relay service on the host.
"""

import os
import sys
import json
import logging
import traceback
import requests
from typing import Dict, Any, List, Optional
from pathlib import Path

# Configure logging
logger = logging.getLogger(__name__)

class VisioService:
    """Service for interacting with Microsoft Visio."""
    
    def __init__(self):
        """Initialize the Visio service."""
        self.is_connected = False
        self.host = os.getenv("VISIO_SERVICE_HOST", "host.docker.internal")
        self.port = os.getenv("VISIO_SERVICE_PORT", "8051")
        self.api_url = f"http://{self.host}:{self.port}"
        
        # Try to connect to the relay service
        self.connect_to_visio()
    
    def connect_to_visio(self) -> bool:
        """Connect to Microsoft Visio via the relay service."""
        try:
            # Check if the relay service is running
            logger.info(f"Checking relay service at {self.api_url}/health")
            response = requests.get(f"{self.api_url}/health", timeout=5)
            
            if response.status_code == 200:
                # Try to connect to Visio via the relay
                connect_response = requests.get(f"{self.api_url}/connect", timeout=5)
                
                if connect_response.status_code == 200:
                    self.is_connected = True
                    logger.info("Connected to Visio via relay service")
                    return True
                else:
                    logger.error(f"Failed to connect to Visio via relay service: {connect_response.text}")
                    self.is_connected = False
                    return False
            else:
                logger.error(f"Relay service health check failed: {response.status_code}")
                self.is_connected = False
                return False
                
        except Exception as e:
            logger.error(f"Failed to connect to relay service: {e}")
            logger.error(traceback.format_exc())
            self.is_connected = False
            return False
    
    def get_active_document(self) -> Dict[str, Any]:
        """Get information about the active Visio document."""
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"error": "Not connected to Visio"}
            
            # Get active document info from relay service
            response = requests.get(f"{self.api_url}/active-document", timeout=10)
            
            if response.status_code == 200:
                result = response.json()
                
                # Handle no document case
                if result.get("status") == "no_document":
                    return {
                        "status": "no_document",
                        "message": result.get("message", "No Visio document is currently open")
                    }
                
                # Return the data
                return {
                    "status": "success",
                    **result.get("data", {})
                }
            else:
                return {"error": f"Failed to get active document: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error getting active document: {e}")
            logger.error(traceback.format_exc())
            return {"error": str(e)}
    
    def analyze_diagram(self, file_path: str, analysis_type: str = "all") -> Dict[str, Any]:
        """
        Analyze a Visio diagram to extract information via the relay service.
        
        Args:
            file_path: Path to the Visio diagram
            analysis_type: Type of analysis to perform (structure, connections, text, all)
        
        Returns:
            Dictionary with analysis results
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"error": "Not connected to Visio"}
            
            # Normalize file path for container environment
            file_path = self._normalize_file_path(file_path)
            
            # Call relay service to analyze diagram
            data = {
                "file_path": file_path,
                "analysis_type": analysis_type
            }
            
            response = requests.post(f"{self.api_url}/analyze-diagram", json=data, timeout=30)
            
            if response.status_code == 200:
                result = response.json()
                
                if result.get("status") == "error":
                    return {
                        "status": "error",
                        "message": result.get("message", "Unknown error")
                    }
                
                # Return the data
                return {
                    "status": "success",
                    **result.get("data", {})
                }
            else:
                return {"error": f"Failed to analyze diagram: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error analyzing diagram: {e}")
            logger.error(traceback.format_exc())
            return {"error": str(e)}
    
    def modify_diagram(self, file_path: str, operation: str, shape_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Modify a Visio diagram.
        
        Args:
            file_path: Path to the Visio diagram
            operation: Type of modification (add_shape, update_shape, delete_shape, etc.)
            shape_data: Data for the shape operation
        
        Returns:
            Dictionary with operation results
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"error": "Not connected to Visio"}
            
            # Normalize file path for container environment
            file_path = self._normalize_file_path(file_path)
            
            # Check if file exists
            if not os.path.exists(file_path):
                return {
                    "status": "error",
                    "message": f"File not found: {file_path}"
                }
            
            # Open the document if not already open
            try:
                doc = self.app.Documents(Path(file_path).name)
                logger.info(f"Using already open document: {file_path}")
            except:
                # Open the document
                doc = self.app.Documents.Open(file_path)
                logger.info(f"Opened document: {file_path}")
            
            # Determine which page to use
            page_index = shape_data.get("page_index", 1)
            page = doc.Pages.Item(page_index)
            
            # Handle different operations
            if operation == "add_shape":
                return self._add_shape(page, shape_data)
            elif operation == "update_shape":
                return self._update_shape(page, shape_data)
            elif operation == "delete_shape":
                return self._delete_shape(page, shape_data)
            elif operation == "add_connection":
                return self._add_connection(page, shape_data)
            elif operation == "delete_connection":
                return self._delete_connection(page, shape_data)
            else:
                return {
                    "status": "error",
                    "message": f"Unknown operation: {operation}"
                }
        
        except Exception as e:
            logger.error(f"Error modifying diagram: {e}")
            logger.error(traceback.format_exc())
            return {"error": str(e)}
    
    def verify_connections(self, file_path: str, shape_ids: List[str] = None) -> Dict[str, Any]:
        """
        Verify connections between shapes in a Visio diagram.
        
        Args:
            file_path: Path to the Visio diagram
            shape_ids: List of shape IDs to verify connections for
        
        Returns:
            Dictionary with verification results
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"error": "Not connected to Visio"}
            
            # Normalize file path for container environment
            file_path = self._normalize_file_path(file_path)
            
            # Run full diagram analysis to get connections
            analysis = self.analyze_diagram(file_path, "connections")
            
            if "error" in analysis:
                return analysis
            
            # Filter connections for specified shapes if needed
            result = {
                "status": "success",
                "file_path": file_path,
                "connections": []
            }
            
            # Extract connections from analysis
            for page in analysis["pages"]:
                page_connections = page.get("connections", [])
                
                # Filter by shape IDs if provided
                if shape_ids:
                    filtered_connections = [
                        conn for conn in page_connections
                        if (str(conn.get("from_shape_id")) in shape_ids or 
                            str(conn.get("to_shape_id")) in shape_ids)
                    ]
                    result["connections"].extend(filtered_connections)
                else:
                    result["connections"].extend(page_connections)
            
            return result
        
        except Exception as e:
            logger.error(f"Error verifying connections: {e}")
            logger.error(traceback.format_exc())
            return {"error": str(e)}
    
    def _get_shapes_info(self, page) -> List[Dict[str, Any]]:
        """Get information about shapes on a page."""
        shapes_info = []
        
        for i in range(1, page.Shapes.Count + 1):
            shape = page.Shapes.Item(i)
            
            # Skip connectors as they're handled separately
            if shape.OneD:
                continue
            
            # Get basic shape info
            shape_info = {
                "id": shape.ID,
                "name": shape.Name,
                "text": shape.Text if hasattr(shape, "Text") else "",
                "type": shape.Type,
                "master": shape.Master.Name if shape.Master else "None",
                "position": {
                    "x": shape.Cells("PinX").Result(""),
                    "y": shape.Cells("PinY").Result("")
                },
                "size": {
                    "width": shape.Cells("Width").Result(""),
                    "height": shape.Cells("Height").Result("")
                }
            }
            
            shapes_info.append(shape_info)
        
        return shapes_info
    
    def _get_connections_info(self, page) -> List[Dict[str, Any]]:
        """Get information about connections on a page."""
        connections_info = []
        
        for i in range(1, page.Shapes.Count + 1):
            shape = page.Shapes.Item(i)
            
            # Only process connectors
            if shape.OneD:
                # Get connection points
                begin_connect = shape.Cells("BeginX").Result("")
                end_connect = shape.Cells("EndX").Result("")
                
                try:
                    from_shape = shape.Connects.Item(1).FromSheet
                    to_shape = shape.Connects.Item(shape.Connects.Count).ToSheet
                    
                    connection_info = {
                        "id": shape.ID,
                        "name": shape.Name,
                        "text": shape.Text if hasattr(shape, "Text") else "",
                        "from_shape_id": from_shape.ID,
                        "from_shape_name": from_shape.Name,
                        "to_shape_id": to_shape.ID,
                        "to_shape_name": to_shape.Name
                    }
                    
                    connections_info.append(connection_info)
                except:
                    # Some connectors might not be connected properly
                    pass
        
        return connections_info
    
    def _get_text_info(self, page) -> List[Dict[str, Any]]:
        """Get text information from shapes on a page."""
        text_info = []
        
        for i in range(1, page.Shapes.Count + 1):
            shape = page.Shapes.Item(i)
            
            if hasattr(shape, "Text") and shape.Text:
                text_item = {
                    "shape_id": shape.ID,
                    "shape_name": shape.Name,
                    "text": shape.Text
                }
                text_info.append(text_item)
        
        return text_info
    
    def _add_shape(self, page, shape_data: Dict[str, Any]) -> Dict[str, Any]:
        """Add a shape to a page."""
        try:
            # Extract shape data
            master_name = shape_data.get("master_name")
            x = shape_data.get("x", 4.0)
            y = shape_data.get("y", 4.0)
            text = shape_data.get("text", "")
            
            # Get stencil if specified
            stencil_name = shape_data.get("stencil_name", "Basic Shapes.vss")
            
            # Try to open the stencil
            try:
                stencil = self.app.Documents(stencil_name)
            except:
                # Open the stencil if not already open
                try:
                    stencil_path = os.path.join(self.app.GetBuiltInStencilFile(0, 0), stencil_name)
                    stencil = self.app.Documents.OpenEx(stencil_path, 0)
                except:
                    # Try fallback for built-in stencils
                    try:
                        stencil = self.app.Documents.OpenStencil(stencil_name)
                    except:
                        return {
                            "status": "error",
                            "message": f"Could not open stencil: {stencil_name}"
                        }
            
            # Get the master shape
            try:
                master = stencil.Masters.ItemU(master_name)
            except:
                return {
                    "status": "error",
                    "message": f"Master shape not found: {master_name} in {stencil_name}"
                }
            
            # Drop the shape
            shape = page.Drop(master, x, y)
            
            # Set text if provided
            if text:
                shape.Text = text
            
            # Return information about the new shape
            return {
                "status": "success",
                "operation": "add_shape",
                "shape_id": shape.ID,
                "shape_name": shape.Name
            }
        
        except Exception as e:
            logger.error(f"Error adding shape: {e}")
            logger.error(traceback.format_exc())
            return {"error": str(e)}
    
    def _update_shape(self, page, shape_data: Dict[str, Any]) -> Dict[str, Any]:
        """Update a shape on a page."""
        try:
            # Extract shape data
            shape_id = shape_data.get("shape_id")
            text = shape_data.get("text")
            x = shape_data.get("x")
            y = shape_data.get("y")
            
            if not shape_id:
                return {
                    "status": "error",
                    "message": "shape_id is required for update_shape operation"
                }
            
            # Find the shape
            shape = None
            for i in range(1, page.Shapes.Count + 1):
                s = page.Shapes.Item(i)
                if s.ID == shape_id:
                    shape = s
                    break
            
            if not shape:
                return {
                    "status": "error",
                    "message": f"Shape not found with ID: {shape_id}"
                }
            
            # Update text if provided
            if text is not None:
                shape.Text = text
            
            # Update position if provided
            if x is not None and y is not None:
                shape.Cells("PinX").SetResult(x, 0)
                shape.Cells("PinY").SetResult(y, 0)
            
            # Return success
            return {
                "status": "success",
                "operation": "update_shape",
                "shape_id": shape.ID,
                "shape_name": shape.Name
            }
        
        except Exception as e:
            logger.error(f"Error updating shape: {e}")
            logger.error(traceback.format_exc())
            return {"error": str(e)}
    
    def _delete_shape(self, page, shape_data: Dict[str, Any]) -> Dict[str, Any]:
        """Delete a shape from a page."""
        try:
            # Extract shape data
            shape_id = shape_data.get("shape_id")
            
            if not shape_id:
                return {
                    "status": "error",
                    "message": "shape_id is required for delete_shape operation"
                }
            
            # Find and delete the shape
            for i in range(1, page.Shapes.Count + 1):
                shape = page.Shapes.Item(i)
                if shape.ID == shape_id:
                    shape_name = shape.Name
                    shape.Delete()
                    return {
                        "status": "success",
                        "operation": "delete_shape",
                        "shape_id": shape_id,
                        "shape_name": shape_name
                    }
            
            return {
                "status": "error",
                "message": f"Shape not found with ID: {shape_id}"
            }
        
        except Exception as e:
            logger.error(f"Error deleting shape: {e}")
            logger.error(traceback.format_exc())
            return {"error": str(e)}
    
    def _add_connection(self, page, shape_data: Dict[str, Any]) -> Dict[str, Any]:
        """Add a connection between two shapes."""
        try:
            # Extract shape data
            from_shape_id = shape_data.get("from_shape_id")
            to_shape_id = shape_data.get("to_shape_id")
            
            if not from_shape_id or not to_shape_id:
                return {
                    "status": "error",
                    "message": "from_shape_id and to_shape_id are required for add_connection operation"
                }
            
            # Find the shapes
            from_shape = None
            to_shape = None
            
            for i in range(1, page.Shapes.Count + 1):
                shape = page.Shapes.Item(i)
                if shape.ID == from_shape_id:
                    from_shape = shape
                if shape.ID == to_shape_id:
                    to_shape = shape
            
            if not from_shape:
                return {
                    "status": "error",
                    "message": f"From shape not found with ID: {from_shape_id}"
                }
            
            if not to_shape:
                return {
                    "status": "error",
                    "message": f"To shape not found with ID: {to_shape_id}"
                }
            
            # Create the connector
            connector = page.Shapes.Item(page.Drop(self.app.ConnectorToolDataObject, 0, 0))
            
            # Connect the shapes
            connector.CellsU("BeginX").GlueTo(from_shape.CellsU("PinX"))
            connector.CellsU("EndX").GlueTo(to_shape.CellsU("PinX"))
            
            # Set connector text if provided
            if "text" in shape_data:
                connector.Text = shape_data["text"]
            
            # Return success
            return {
                "status": "success",
                "operation": "add_connection",
                "connector_id": connector.ID,
                "connector_name": connector.Name,
                "from_shape_id": from_shape.ID,
                "to_shape_id": to_shape.ID
            }
        
        except Exception as e:
            logger.error(f"Error adding connection: {e}")
            logger.error(traceback.format_exc())
            return {"error": str(e)}
    
    def _delete_connection(self, page, shape_data: Dict[str, Any]) -> Dict[str, Any]:
        """Delete a connection (connector shape)."""
        try:
            # Extract shape data
            connector_id = shape_data.get("connector_id")
            
            if not connector_id:
                return {
                    "status": "error",
                    "message": "connector_id is required for delete_connection operation"
                }
            
            # Find and delete the connector
            for i in range(1, page.Shapes.Count + 1):
                shape = page.Shapes.Item(i)
                if shape.ID == connector_id and shape.OneD:
                    shape_name = shape.Name
                    shape.Delete()
                    return {
                        "status": "success",
                        "operation": "delete_connection",
                        "connector_id": connector_id,
                        "connector_name": shape_name
                    }
            
            return {
                "status": "error",
                "message": f"Connector not found with ID: {connector_id}"
            }
        
        except Exception as e:
            logger.error(f"Error deleting connection: {e}")
            logger.error(traceback.format_exc())
            return {"error": str(e)}
    
    def _normalize_file_path(self, file_path: str) -> str:
        """
        Normalize file path for container environment.
        
        Args:
            file_path: Original file path
            
        Returns:
            Normalized file path
        """
        # If path already exists, return it
        if os.path.exists(file_path):
            return file_path
        
        # Check if this is a relative path that should be in the visio-files directory
        if not os.path.isabs(file_path):
            # Try in the visio-files directory
            container_path = os.path.join("C:\\visio-files", file_path)
            if os.path.exists(container_path):
                logger.info(f"Using container path: {container_path}")
                return container_path
        
        # For paths with forward slashes, convert to backslashes
        normalized_path = file_path.replace('/', '\\')
        if os.path.exists(normalized_path):
            return normalized_path
        
        # For paths without drive letter in Windows, assume C:
        if sys.platform == "win32" and not ':' in normalized_path:
            windows_path = f"C:{normalized_path}" if normalized_path.startswith('\\') else f"C:\\{normalized_path}"
            if os.path.exists(windows_path):
                return windows_path
            
        # Return original if all else fails
        return file_path 