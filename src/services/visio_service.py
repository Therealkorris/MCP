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
            
            # Handle special 'active' keyword for current document
            if file_path.lower() == 'active':
                # Get the active document info first
                active_doc = self.get_active_document()
                if active_doc.get("status") != "success":
                    return {"status": "error", "message": "No active document or unable to access it"}
                
                # Use the full path from active document
                file_path = active_doc.get("full_path", "")
                if not file_path:
                    return {"status": "error", "message": "Unable to determine path of active document"}
            else:
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
            
            # Handle special 'active' keyword for current document
            if file_path.lower() == 'active':
                # Get the active document info first
                active_doc = self.get_active_document()
                if active_doc.get("status") != "success":
                    return {"status": "error", "message": "No active document or unable to access it"}
                
                # Use the full path from active document
                file_path = active_doc.get("full_path", "")
                if not file_path:
                    return {"status": "error", "message": "Unable to determine path of active document"}
            else:
                # Normalize file path for container environment
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
            shape_ids: List of shape IDs to verify connections for
        
        Returns:
            Dictionary with verification results
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}
            
            # Handle special 'active' keyword for current document
            if file_path and file_path.lower() == 'active':
                # Get the active document info first
                active_doc = self.get_active_document()
                if active_doc.get("status") != "success":
                    return {"status": "error", "message": "No active document or unable to access it"}
                
                # Use the full path from active document
                file_path = active_doc.get("full_path", "")
                if not file_path:
                    return {"status": "error", "message": "Unable to determine path of active document"}
            else:
                # Normalize file path for container environment
                file_path = self._normalize_file_path(file_path)
            
            # Call relay service to verify connections
            data = {
                "file_path": file_path,
                "shape_ids": shape_ids or []
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
            
            # Handle special 'active' keyword for current document
            if not file_path or file_path.lower() == 'active':
                file_path = 'active'
            
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
            page_index: Index of the page to get shapes from
            
        Returns:
            Dictionary with shapes information
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}
            
            # Handle special 'active' keyword for current document
            if file_path.lower() == 'active':
                # Get the active document info first
                active_doc = self.get_active_document()
                if active_doc.get("status") != "success":
                    return {"status": "error", "message": "No active document or unable to access it"}
                
                # Use the full path from active document
                file_path = active_doc.get("full_path", "")
                if not file_path:
                    return {"status": "error", "message": "Unable to determine path of active document"}
            else:
                # Normalize file path for container environment
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
            format: Export format (png, jpg, pdf, svg)
            output_path: Path to save the exported file
            
        Returns:
            Dictionary with export results
        """
        try:
            if not self.is_connected:
                if not self.connect_to_visio():
                    return {"status": "error", "message": "Not connected to Visio"}
            
            # Handle special 'active' keyword for current document
            if file_path.lower() == 'active':
                # Get the active document info first
                active_doc = self.get_active_document()
                if active_doc.get("status") != "success":
                    return {"status": "error", "message": "No active document or unable to access it"}
                
                # Use the full path from active document
                file_path = active_doc.get("full_path", "")
                if not file_path:
                    return {"status": "error", "message": "Unable to determine path of active document"}
            else:
                # Normalize file path for container environment
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
            Dictionary with list of available master shapes by stencil
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
                return {"status": "error", "message": f"Failed to get masters: {response.text}"}
        
        except Exception as e:
            logger.error(f"Error getting masters: {e}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)} 