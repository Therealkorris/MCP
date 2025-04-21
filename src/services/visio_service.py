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
from typing import Dict, Any, List, Optional, Tuple
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
                    return {"status": "error", "message": "Not connected to Visio"}
            
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
                return {"status": "error", "message": f"Failed to get active document: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error getting active document: {e}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)}
    
    def analyze_diagram(self, file_path: str, analysis_type: str = "all") -> Dict[str, Any]:
        """
        Analyze a Visio diagram to extract information via the relay service.
        
        Args:
            file_path: Path to the Visio diagram or 'active' to use the active document
            analysis_type: Type of analysis to perform (structure, connections, text, all)
        
        Returns:
            Dictionary with analysis results
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}
            
            # Handle special 'active' keyword for current document - pass it directly to host
            if file_path.lower() != 'active':
                # Only normalize non-active file paths
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
                return {"status": "error", "message": f"Failed to analyze diagram: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error analyzing diagram: {e}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)}
    
    def modify_diagram(self, file_path: str, operation: str, shape_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Modify a Visio diagram via the relay service.
        
        Args:
            file_path: Path to the Visio diagram or 'active' to use the active document
            operation: Type of modification (add_shape, update_shape, delete_shape, etc.)
            shape_data: Data for the shape operation
        
        Returns:
            Dictionary with operation results
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}
            
            # Handle special 'active' keyword for current document - pass it directly to host
            if file_path.lower() != 'active':
                # Only normalize non-active file paths
                file_path = self._normalize_file_path(file_path)
            
            # Call relay service to modify diagram
            data = {
                "file_path": file_path,
                "operation": operation,
                "shape_data": shape_data
            }
            
            response = requests.post(f"{self.api_url}/modify-diagram", json=data, timeout=30)
            
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
                return {"status": "error", "message": f"Failed to modify diagram: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error modifying diagram: {e}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)}
    
    def verify_connections(self, file_path: str, shape_ids: List[str] = None) -> Dict[str, Any]:
        """
        Verify connections between shapes in a Visio diagram.
        
        Args:
            file_path: Path to the Visio diagram or 'active' to use the active document
            shape_ids: Optional list of shape IDs to filter connections
            
        Returns:
            Dictionary with verification results
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}
            
            # Handle special 'active' keyword for current document - pass it directly to host
            if file_path.lower() != 'active':
                # Only normalize non-active file paths
                file_path = self._normalize_file_path(file_path)
            
            # Call relay service to verify connections
            data = {
                "file_path": file_path,
                "shape_ids": shape_ids
            }
            
            response = requests.post(f"{self.api_url}/verify-connections", json=data, timeout=30)
            
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
                return {"status": "error", "message": f"Failed to verify connections: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error verifying connections: {e}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)}
    
    def create_new_diagram(self, template: str = "Basic.vst", save_path: str = None) -> Dict[str, Any]:
        """
        Create a new Visio diagram.
        
        Args:
            template: Template to use for the new diagram
            save_path: Path to save the new diagram
            
        Returns:
            Dictionary with creation results
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}
            
            # Call relay service to create new diagram
            data = {
                "template": template,
                "save_path": save_path
            }
            
            response = requests.post(f"{self.api_url}/create-diagram", json=data, timeout=30)
            
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
                return {"status": "error", "message": f"Failed to create diagram: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error creating diagram: {e}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)}
    
    def save_diagram(self, file_path: str = None) -> Dict[str, Any]:
        """
        Save a Visio diagram.
        
        Args:
            file_path: Path to save the diagram, or 'active' to save the active document
            
        Returns:
            Dictionary with save results
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}
            
            # Use 'active' as default if not specified
            if not file_path:
                file_path = 'active'
            
            # Handle special 'active' keyword for current document - pass it directly to host
            if file_path.lower() != 'active':
                # Only normalize non-active file paths
                file_path = self._normalize_file_path(file_path)
            
            # Call relay service to save diagram
            data = {
                "file_path": file_path
            }
            
            response = requests.post(f"{self.api_url}/save-diagram", json=data, timeout=30)
            
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
                return {"status": "error", "message": f"Failed to save diagram: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error saving diagram: {e}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)}
    
    def get_available_stencils(self) -> Dict[str, Any]:
        """
        Get list of available stencils.
        
        Returns:
            Dictionary with list of stencils
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}
            
            # Call relay service to get stencils
            response = requests.get(f"{self.api_url}/available-stencils", timeout=30)
            
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
                return {"status": "error", "message": f"Failed to get stencils: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error getting stencils: {e}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)}
    
    def get_shapes_on_page(self, file_path: str = 'active', page_index: int = 1) -> Dict[str, Any]:
        """
        Get information about all shapes on a page.
        
        Args:
            file_path: Path to the Visio diagram or 'active' to use the active document
            page_index: Index of the page to get shapes from (1-based)
            
        Returns:
            Dictionary with shapes information
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}
            
            # Handle special 'active' keyword for current document - pass it directly to host
            if file_path.lower() != 'active':
                # Only normalize non-active file paths
                file_path = self._normalize_file_path(file_path)
            
            # Call relay service to get shapes
            data = {
                "file_path": file_path,
                "page_index": page_index
            }
            
            response = requests.post(f"{self.api_url}/get-shapes", json=data, timeout=30)
            
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
                return {"status": "error", "message": f"Failed to get shapes: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error getting shapes: {e}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)}
    
    def export_diagram(self, file_path: str = 'active', format: str = 'png', output_path: str = None) -> Dict[str, Any]:
        """
        Export a Visio diagram to another format.
        
        Args:
            file_path: Path to the Visio diagram or 'active' to use the active document
            format: Format to export to (png, jpg, pdf, svg)
            output_path: Path to save the exported file
            
        Returns:
            Dictionary with export results
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}
            
            # Handle special 'active' keyword for current document - pass it directly to host
            if file_path.lower() != 'active':
                # Only normalize non-active file paths
                file_path = self._normalize_file_path(file_path)
            
            # Call relay service to export diagram
            data = {
                "file_path": file_path,
                "format": format,
                "output_path": output_path
            }
            
            response = requests.post(f"{self.api_url}/export-diagram", json=data, timeout=60)
            
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
                return {"status": "error", "message": f"Failed to export diagram: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error exporting diagram: {e}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)}
    
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
    
    def get_available_masters(self) -> Dict[str, Any]:
        """
        Get a list of available master shapes from all open stencils.
        
        Returns:
            Dictionary with stencils and their master shapes
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}
            
            # Call relay service to get available masters
            response = requests.get(f"{self.api_url}/available-masters", timeout=30)
            
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
                return {"status": "error", "message": f"Failed to get available masters: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error getting available masters: {e}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)}
            
    def image_to_diagram(self, image_path: str, output_path: str = None, detection_level: str = "standard") -> Dict[str, Any]:
        """
        Convert an image (screenshot, photo, etc.) to a Visio diagram.
        
        This method analyzes an image and attempts to create a Visio diagram that represents
        the shapes, connections and layout found in the image. Currently this is a framework
        for future AI-based image analysis capabilities.
        
        Args:
            image_path: Path to the input image file
            output_path: Path where the resulting Visio diagram should be saved
            detection_level: Level of detail for shape detection (simple, standard, detailed)
            
        Returns:
            Dictionary with operation results
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}

            # Normalize file path for container environment
            image_path = self._normalize_file_path(image_path)
            
            # Validate image file exists
            if not os.path.exists(image_path):
                return {"status": "error", "message": f"Image file not found: {image_path}"}
                
            # Set default output path if not provided
            if not output_path:
                # Create output in same directory as input with .vsdx extension
                base_name = os.path.basename(image_path)
                name_without_ext = os.path.splitext(base_name)[0]
                output_dir = os.path.dirname(image_path)
                output_path = os.path.join(output_dir, f"{name_without_ext}_diagram.vsdx")
            
            # Normalize output path
            output_path = self._normalize_file_path(output_path)
            
            # For the initial implementation, we'll create a simple diagram with placeholder elements
            # First, create a new diagram
            new_diagram = self.create_new_diagram("BASIC_M.vssx", output_path)
            if new_diagram.get("status") != "success":
                return {"status": "error", "message": f"Failed to create new diagram: {new_diagram.get('message')}"}
                
            # Get the active document for adding shapes
            doc_info = self.get_active_document()
            if doc_info.get("status") != "success":
                return {"status": "error", "message": f"Failed to access active document: {doc_info.get('message')}"}
            
            # Log that we're starting image analysis (placeholder for actual image analysis)
            logger.info(f"Starting image analysis of {image_path} with detection level: {detection_level}")
            
            # Placeholder for actual image analysis results
            # In a future implementation, this would use computer vision to detect shapes
            mock_analysis_results = {
                "shapes": [
                    {"type": "Rectangle", "x": 3, "y": 3, "width": 2, "height": 1.5, "text": "Detected Rectangle", "fill_color": "blue"},
                    {"type": "Circle", "x": 7, "y": 3, "width": 1.5, "height": 1.5, "text": "Detected Circle", "fill_color": "red"},
                    {"type": "Diamond", "x": 5, "y": 7, "width": 2, "height": 1.5, "text": "Detected Diamond", "fill_color": "green"}
                ],
                "connections": [
                    {"from_index": 0, "to_index": 1, "text": "Connection 1", "line_pattern": "dashed"},
                    {"from_index": 1, "to_index": 2, "text": "Connection 2", "line_pattern": "dotted"}
                ]
            }
            
            # Track created shapes and their IDs
            created_shapes = []
            
            # Add detected shapes to the diagram
            for shape_info in mock_analysis_results["shapes"]:
                shape_data = {
                    "master_name": shape_info["type"],
                    "position": {"x": shape_info["x"], "y": shape_info["y"]},
                    "size": {"width": shape_info["width"], "height": shape_info["height"]},
                    "text": shape_info["text"],
                    "fill_color": shape_info["fill_color"]
                }
                
                # Add the shape
                shape_result = self.modify_diagram(output_path, "add_shape", shape_data)
                if shape_result.get("status") == "success":
                    created_shapes.append({
                        "index": len(created_shapes),
                        "id": shape_result.get("shape_id"),
                        "name": shape_result.get("shape_name")
                    })
                    logger.info(f"Added shape: {shape_info['type']} at ({shape_info['x']}, {shape_info['y']})")
                else:
                    logger.warning(f"Failed to add shape: {shape_result.get('message')}")
            
            # Add connections between shapes
            for conn_info in mock_analysis_results["connections"]:
                # Skip if shape indices are invalid
                if conn_info["from_index"] >= len(created_shapes) or conn_info["to_index"] >= len(created_shapes):
                    logger.warning(f"Invalid shape indices for connection: {conn_info}")
                    continue
                
                # Get the shape IDs
                from_shape_id = created_shapes[conn_info["from_index"]]["id"]
                to_shape_id = created_shapes[conn_info["to_index"]]["id"]
                
                # Create connector data
                connector_data = {
                    "from_shape_id": from_shape_id,
                    "to_shape_id": to_shape_id,
                    "text": conn_info.get("text", ""),
                    "line_pattern": conn_info.get("line_pattern", "solid")
                }
                
                # Add the connector
                conn_result = self.modify_diagram(output_path, "add_connector", connector_data)
                if conn_result.get("status") == "success":
                    logger.info(f"Added connection from shape {from_shape_id} to {to_shape_id}")
                else:
                    logger.warning(f"Failed to add connection: {conn_result.get('message')}")
            
            # Save the diagram
            save_result = self.save_diagram(output_path)
            if save_result.get("status") != "success":
                logger.warning(f"Failed to save diagram: {save_result.get('message')}")
            
            # Return success with diagram info
            return {
                "status": "success",
                "message": f"Created diagram from image {image_path}",
                "diagram_path": output_path,
                "detected_elements": {
                    "shapes": len(mock_analysis_results["shapes"]),
                    "connections": len(mock_analysis_results["connections"])
                },
                "note": "This is currently using mock data. Future versions will use AI-based image analysis."
            }
            
        except Exception as e:
            logger.error(f"Error in image to diagram conversion: {e}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)} 