"""
Visio Relay Service for MCP-Visio

This script runs on a Windows host with Visio installed and provides a REST API
for the containerized MCP service to interact with Visio.

Run this script on the Windows host before starting the MCP container.
"""

import os
import sys
import json
import logging
import win32com.client
from typing import Dict, Any, List, Optional
from pathlib import Path
from fastapi import FastAPI, HTTPException, Body
from pydantic import BaseModel
import uvicorn

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("visio_relay.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("visio-relay")

# Create FastAPI app
app = FastAPI(title="Visio Relay Service")

# Global Visio application object
visio_app = None

# Data models
class DiagramRequest(BaseModel):
    file_path: str
    analysis_type: str = "all"

class ModifyRequest(BaseModel):
    file_path: str
    operation: str
    shape_data: Dict[str, Any]

class ConnectionsRequest(BaseModel):
    file_path: str
    shape_ids: List[str] = None

class CreateDiagramRequest(BaseModel):
    template: str = "Basic.vst"
    save_path: Optional[str] = None

class SaveDiagramRequest(BaseModel):
    file_path: Optional[str] = None

class ShapesRequest(BaseModel):
    file_path: str
    page_index: int = 1

class ExportRequest(BaseModel):
    file_path: str
    format: str = "png"
    output_path: Optional[str] = None

def connect_to_visio():
    """Connect to Microsoft Visio application."""
    global visio_app
    try:
        if visio_app is None:
            logger.info("Connecting to Visio application...")
            visio_app = win32com.client.Dispatch("Visio.Application")
            visio_app.Visible = True
            logger.info("Connected to Visio application")
        return True
    except Exception as e:
        logger.error(f"Failed to connect to Visio: {e}")
        return False

def normalize_file_path(file_path: str) -> str:
    """Normalize file path for Windows."""
    # Handle special 'active' keyword
    if file_path.lower() == 'active':
        return 'active'
        
    # Convert forward slashes to backslashes
    path = file_path.replace('/', '\\')
    
    # Add drive letter if missing
    if not ':' in path:
        path = f"C:\\{path}" if not path.startswith('\\') else f"C:{path}"
    
    return path

@app.get("/health")
async def health_check():
    """Check health of the relay service."""
    return {"status": "healthy", "service": "Visio Relay"}

@app.get("/connect")
async def connect():
    """Connect to Visio application."""
    if connect_to_visio():
        return {"status": "success", "message": "Connected to Visio"}
    else:
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")

@app.get("/active-document")
async def get_active_document():
    """Get information about the active Visio document."""
    if not connect_to_visio():
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")
    
    try:
        # Check if there's an active document
        if visio_app.Documents.Count == 0:
            return {"status": "no_document", "message": "No Visio document is currently open"}
        
        active_doc = visio_app.ActiveDocument
        
        # Get basic document info
        doc_info = {
            "name": active_doc.Name,
            "path": active_doc.Path,
            "full_path": os.path.join(active_doc.Path, active_doc.Name),
            "pages_count": active_doc.Pages.Count,
            "saved": active_doc.Saved,
            "readonly": active_doc.ReadOnly,
            "pages": []
        }
        
        # Get pages info
        for i in range(1, active_doc.Pages.Count + 1):
            page = active_doc.Pages.Item(i)
            page_info = {
                "name": page.Name,
                "index": i,
                "shapes_count": page.Shapes.Count,
                "is_foreground": page.Background == 0
            }
            doc_info["pages"].append(page_info)
        
        # Get open documents
        open_docs = []
        for i in range(1, visio_app.Documents.Count + 1):
            doc = visio_app.Documents.Item(i)
            open_docs.append({
                "name": doc.Name,
                "path": doc.Path,
                "full_path": os.path.join(doc.Path, doc.Name),
                "is_active": doc.Name == active_doc.Name
            })
        
        doc_info["open_documents"] = open_docs
        
        return {"status": "success", "data": doc_info}
    
    except Exception as e:
        logger.error(f"Error getting active document: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/analyze-diagram")
async def analyze_diagram(request: DiagramRequest):
    """
    Analyze a Visio diagram to extract information.
    """
    if not connect_to_visio():
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")
    
    try:
        file_path = request.file_path
        analysis_type = request.analysis_type
        
        # Handle special 'active' keyword
        if file_path.lower() == 'active':
            doc = visio_app.ActiveDocument
            if not doc:
                return {"status": "error", "message": "No active document"}
        else:
            # Normalize path
            file_path = normalize_file_path(file_path)
            
            # Check if file exists
            if not os.path.exists(file_path):
                return {"status": "error", "message": f"File not found: {file_path}"}
            
            # Open the document
            try:
                doc = visio_app.Documents.Open(file_path)
            except Exception as e:
                return {"status": "error", "message": f"Failed to open document: {str(e)}"}
        
        # Prepare result
        result = {
            "name": doc.Name,
            "path": doc.Path,
            "full_path": os.path.join(doc.Path, doc.Name),
            "pages_count": doc.Pages.Count,
            "pages": []
        }
        
        # Process each page
        for i in range(1, doc.Pages.Count + 1):
            page = doc.Pages.Item(i)
            page_data = {
                "name": page.Name,
                "index": i,
                "shapes_count": page.Shapes.Count,
                "shapes": [],
                "connections": [],
                "text_elements": []
            }
            
            # Get shapes data based on analysis type
            if analysis_type in ["structure", "all"]:
                # Get shapes
                for j in range(1, page.Shapes.Count + 1):
                    shape = page.Shapes.Item(j)
                    
                    # Skip connectors as they're handled separately
                    if shape.OneD:
                        continue
                    
                    # Get basic shape info
                    shape_data = {
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
                    
                    page_data["shapes"].append(shape_data)
            
            # Get connections data
            if analysis_type in ["connections", "all"]:
                # Get connectors
                for j in range(1, page.Shapes.Count + 1):
                    shape = page.Shapes.Item(j)
                    
                    # Only process connectors
                    if shape.OneD:
                        try:
                            # Try to get connection information
                            from_shape = None
                            to_shape = None
                            
                            # Check connects collection
                            if shape.Connects.Count > 0:
                                try:
                                    # Get first connection (from)
                                    first_connect = shape.Connects.Item(1)
                                    from_shape = {
                                        "id": first_connect.FromSheet.ID,
                                        "name": first_connect.FromSheet.Name
                                    }
                                    
                                    # Get last connection (to)
                                    last_connect = shape.Connects.Item(shape.Connects.Count)
                                    to_shape = {
                                        "id": last_connect.ToSheet.ID,
                                        "name": last_connect.ToSheet.Name
                                    }
                                except:
                                    # Some connectors might not have proper connections
                                    pass
                            
                            connection_data = {
                                "id": shape.ID,
                                "name": shape.Name,
                                "text": shape.Text if hasattr(shape, "Text") else "",
                                "from_shape": from_shape,
                                "to_shape": to_shape,
                                "type": "connector" 
                            }
                            
                            page_data["connections"].append(connection_data)
                        except Exception as conn_err:
                            logger.warning(f"Error processing connector {shape.Name}: {conn_err}")
            
            # Get text elements
            if analysis_type in ["text", "all"]:
                # Get all text elements
                for j in range(1, page.Shapes.Count + 1):
                    shape = page.Shapes.Item(j)
                    
                    if hasattr(shape, "Text") and shape.Text:
                        text_data = {
                            "shape_id": shape.ID,
                            "shape_name": shape.Name,
                            "text": shape.Text
                        }
                        page_data["text_elements"].append(text_data)
            
            # Add page data to result
            result["pages"].append(page_data)
        
        # Close document if it was opened for analysis
        if file_path.lower() != 'active':
            doc.Close()
        
        return {"status": "success", "data": result}
    
    except Exception as e:
        logger.error(f"Error analyzing diagram: {e}")
        return {"status": "error", "message": str(e)}

@app.post("/modify-diagram")
async def modify_diagram(request: ModifyRequest):
    """
    Modify a Visio diagram.
    """
    if not connect_to_visio():
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")
    
    try:
        file_path = request.file_path
        operation = request.operation
        shape_data = request.shape_data
        
        # Handle special 'active' keyword
        if file_path.lower() == 'active':
            doc = visio_app.ActiveDocument
            if not doc:
                return {"status": "error", "message": "No active document"}
            was_opened = False
        else:
            # Normalize path
            file_path = normalize_file_path(file_path)
            
            # Check if file exists
            if not os.path.exists(file_path):
                return {"status": "error", "message": f"File not found: {file_path}"}
            
            # Open the document
            try:
                # Check if already open
                try:
                    doc = visio_app.Documents(os.path.basename(file_path))
                    was_opened = True
                except:
                    doc = visio_app.Documents.Open(file_path)
                    was_opened = False
            except Exception as e:
                return {"status": "error", "message": f"Failed to open document: {str(e)}"}
        
        # Determine which page to use
        page_index = shape_data.get("page_index", 1)
        if page_index > doc.Pages.Count:
            return {"status": "error", "message": f"Page index {page_index} out of range (document has {doc.Pages.Count} pages)"}
        
        page = doc.Pages.Item(page_index)
        
        # Perform the requested operation
        result = {"status": "success", "operation": operation}
        
        if operation == "add_shape":
            # Add a shape to the page
            master_name = shape_data.get("master_name")
            if not master_name:
                return {"status": "error", "message": "master_name is required for add_shape operation"}
            
            x = shape_data.get("x", 4.0)
            y = shape_data.get("y", 4.0)
            text = shape_data.get("text", "")
            
            # Get stencil if specified
            stencil_name = shape_data.get("stencil_name", "Basic Shapes.vss")
            
            # Try to open the stencil
            try:
                try:
                    stencil = visio_app.Documents(stencil_name)
                except:
                    # Open the stencil if not already open
                    try:
                        stencil_path = os.path.join(visio_app.GetBuiltInStencilFile(0, 0), stencil_name)
                        stencil = visio_app.Documents.OpenEx(stencil_path, 0)
                    except:
                        # Try fallback for built-in stencils
                        stencil = visio_app.Documents.OpenStencil(stencil_name)
            except Exception as e:
                return {"status": "error", "message": f"Could not open stencil: {stencil_name} - {str(e)}"}
            
            # Get the master shape
            try:
                master = stencil.Masters.ItemU(master_name)
            except Exception as e:
                return {"status": "error", "message": f"Master shape not found: {master_name} in {stencil_name} - {str(e)}"}
            
            # Drop the shape
            shape = page.Drop(master, x, y)
            
            # Set text if provided
            if text:
                shape.Text = text
            
            # Update result with shape info
            result.update({
                "shape_id": shape.ID,
                "shape_name": shape.Name
            })
            
        elif operation == "update_shape":
            # Update an existing shape
            shape_id = shape_data.get("shape_id")
            if not shape_id:
                return {"status": "error", "message": "shape_id is required for update_shape operation"}
            
            # Find the shape
            shape = None
            for i in range(1, page.Shapes.Count + 1):
                s = page.Shapes.Item(i)
                if s.ID == shape_id:
                    shape = s
                    break
            
            if not shape:
                return {"status": "error", "message": f"Shape not found with ID: {shape_id}"}
            
            # Update text if provided
            if "text" in shape_data:
                shape.Text = shape_data["text"]
            
            # Update position if provided
            if "x" in shape_data and "y" in shape_data:
                shape.Cells("PinX").SetResult(shape_data["x"], 0)
                shape.Cells("PinY").SetResult(shape_data["y"], 0)
            
            # Update size if provided
            if "width" in shape_data:
                shape.Cells("Width").SetResult(shape_data["width"], 0)
            
            if "height" in shape_data:
                shape.Cells("Height").SetResult(shape_data["height"], 0)
            
            # Update result with shape info
            result.update({
                "shape_id": shape.ID,
                "shape_name": shape.Name
            })
            
        elif operation == "delete_shape":
            # Delete a shape
            shape_id = shape_data.get("shape_id")
            if not shape_id:
                return {"status": "error", "message": "shape_id is required for delete_shape operation"}
            
            # Find and delete the shape
            for i in range(1, page.Shapes.Count + 1):
                shape = page.Shapes.Item(i)
                if shape.ID == shape_id:
                    shape_name = shape.Name
                    shape.Delete()
                    
                    # Update result
                    result.update({
                        "shape_id": shape_id,
                        "shape_name": shape_name
                    })
                    break
            else:
                return {"status": "error", "message": f"Shape not found with ID: {shape_id}"}
            
        elif operation == "add_connection":
            # Add a connection between two shapes
            from_shape_id = shape_data.get("from_shape_id")
            to_shape_id = shape_data.get("to_shape_id")
            
            if not from_shape_id or not to_shape_id:
                return {"status": "error", "message": "from_shape_id and to_shape_id are required for add_connection operation"}
            
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
                return {"status": "error", "message": f"From shape not found with ID: {from_shape_id}"}
            
            if not to_shape:
                return {"status": "error", "message": f"To shape not found with ID: {to_shape_id}"}
            
            # Create the connector
            connector = page.Shapes.Item(page.Drop(visio_app.ConnectorToolDataObject, 0, 0))
            
            # Connect the shapes
            connector.CellsU("BeginX").GlueTo(from_shape.CellsU("PinX"))
            connector.CellsU("EndX").GlueTo(to_shape.CellsU("PinX"))
            
            # Set connector text if provided
            if "text" in shape_data:
                connector.Text = shape_data["text"]
            
            # Update result
            result.update({
                "connector_id": connector.ID,
                "connector_name": connector.Name,
                "from_shape_id": from_shape.ID,
                "to_shape_id": to_shape.ID
            })
            
        elif operation == "delete_connection":
            # Delete a connection
            connector_id = shape_data.get("connector_id")
            
            if not connector_id:
                return {"status": "error", "message": "connector_id is required for delete_connection operation"}
            
            # Find and delete the connector
            for i in range(1, page.Shapes.Count + 1):
                shape = page.Shapes.Item(i)
                if shape.ID == connector_id and shape.OneD:
                    shape_name = shape.Name
                    shape.Delete()
                    
                    # Update result
                    result.update({
                        "connector_id": connector_id,
                        "connector_name": shape_name
                    })
                    break
            else:
                return {"status": "error", "message": f"Connector not found with ID: {connector_id}"}
        
        else:
            return {"status": "error", "message": f"Unknown operation: {operation}"}
        
        # Save the document if it was modified
        if not was_opened and file_path.lower() != 'active':
            doc.Save()
            doc.Close()
        
        return {"status": "success", "data": result}
    
    except Exception as e:
        logger.error(f"Error modifying diagram: {e}")
        return {"status": "error", "message": str(e)}

@app.post("/verify-connections")
async def verify_connections(request: ConnectionsRequest):
    """
    Verify connections between shapes in a Visio diagram.
    """
    if not connect_to_visio():
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")
    
    try:
        file_path = request.file_path
        shape_ids = request.shape_ids or []
        
        # Handle special 'active' keyword
        if file_path.lower() == 'active':
            doc = visio_app.ActiveDocument
            if not doc:
                return {"status": "error", "message": "No active document"}
            was_opened = False
        else:
            # Normalize path
            file_path = normalize_file_path(file_path)
            
            # Check if file exists
            if not os.path.exists(file_path):
                return {"status": "error", "message": f"File not found: {file_path}"}
            
            # Open the document
            try:
                # Check if already open
                try:
                    doc = visio_app.Documents(os.path.basename(file_path))
                    was_opened = True
                except:
                    doc = visio_app.Documents.Open(file_path)
                    was_opened = False
            except Exception as e:
                return {"status": "error", "message": f"Failed to open document: {str(e)}"}
        
        # Prepare result
        result = {
            "file_path": file_path,
            "connections": []
        }
        
        # Check each page for connections
        for i in range(1, doc.Pages.Count + 1):
            page = doc.Pages.Item(i)
            
            # Process each shape that is a connector
            for j in range(1, page.Shapes.Count + 1):
                shape = page.Shapes.Item(j)
                
                # Only process connectors
                if shape.OneD:
                    try:
                        # Process connections if they involve the requested shapes
                        if shape.Connects.Count > 0:
                            from_shape = None
                            to_shape = None
                            
                            # Get from and to shapes
                            try:
                                first_connect = shape.Connects.Item(1)
                                from_shape_id = first_connect.FromSheet.ID
                                from_shape_name = first_connect.FromSheet.Name
                                
                                last_connect = shape.Connects.Item(shape.Connects.Count)
                                to_shape_id = last_connect.ToSheet.ID
                                to_shape_name = last_connect.ToSheet.Name
                                
                                # Check if this connection involves requested shapes
                                if not shape_ids or str(from_shape_id) in shape_ids or str(to_shape_id) in shape_ids:
                                    connection = {
                                        "connector_id": shape.ID,
                                        "connector_name": shape.Name,
                                        "text": shape.Text if hasattr(shape, "Text") else "",
                                        "from_shape_id": from_shape_id,
                                        "from_shape_name": from_shape_name,
                                        "to_shape_id": to_shape_id,
                                        "to_shape_name": to_shape_name,
                                        "page_name": page.Name,
                                        "page_index": i
                                    }
                                    result["connections"].append(connection)
                            except Exception as conn_err:
                                logger.warning(f"Error processing connector connections: {conn_err}")
                    except Exception as e:
                        logger.warning(f"Error verifying connection for shape {shape.Name}: {e}")
        
        # Close document if it was opened for verification
        if not was_opened and file_path.lower() != 'active':
            doc.Close()
        
        return {"status": "success", "data": result}
    
    except Exception as e:
        logger.error(f"Error verifying connections: {e}")
        return {"status": "error", "message": str(e)}

@app.post("/create-diagram")
async def create_diagram(request: CreateDiagramRequest):
    """
    Create a new Visio diagram.
    """
    if not connect_to_visio():
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")
    
    try:
        template = request.template
        save_path = request.save_path
        
        # Create new document from template
        doc = visio_app.Documents.Add(template)
        
        # Save if path provided
        if save_path:
            save_path = normalize_file_path(save_path)
            doc.SaveAs(save_path)
        
        # Return information about the new document
        result = {
            "name": doc.Name,
            "path": doc.Path,
            "full_path": os.path.join(doc.Path, doc.Name),
            "pages_count": doc.Pages.Count
        }
        
        return {"status": "success", "data": result}
    
    except Exception as e:
        logger.error(f"Error creating diagram: {e}")
        return {"status": "error", "message": str(e)}

@app.post("/save-diagram")
async def save_diagram(request: SaveDiagramRequest):
    """
    Save a Visio diagram.
    """
    if not connect_to_visio():
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")
    
    try:
        file_path = request.file_path
        
        # Handle special 'active' keyword or None
        if not file_path or file_path.lower() == 'active':
            doc = visio_app.ActiveDocument
            if not doc:
                return {"status": "error", "message": "No active document"}
            
            # Save the document
            doc.Save()
            
            result = {
                "name": doc.Name,
                "path": doc.Path,
                "full_path": os.path.join(doc.Path, doc.Name)
            }
        else:
            # Normalize path
            file_path = normalize_file_path(file_path)
            
            # Find document
            try:
                # Check if already open
                doc = visio_app.Documents(os.path.basename(file_path))
            except:
                return {"status": "error", "message": f"Document not found: {os.path.basename(file_path)}"}
            
            # Save the document
            doc.SaveAs(file_path)
            
            result = {
                "name": doc.Name,
                "path": doc.Path,
                "full_path": os.path.join(doc.Path, doc.Name)
            }
        
        return {"status": "success", "data": result}
    
    except Exception as e:
        logger.error(f"Error saving diagram: {e}")
        return {"status": "error", "message": str(e)}

@app.get("/available-stencils")
async def get_available_stencils():
    """
    Get a list of available Visio stencils.
    """
    if not connect_to_visio():
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")
    
    try:
        # Get built-in stencil directories
        stencil_dir = visio_app.GetBuiltInStencilFile(0, 0)
        
        # List stencils in directory
        stencils = []
        
        if os.path.exists(stencil_dir):
            for file in os.listdir(stencil_dir):
                if file.endswith('.vss') or file.endswith('.vssx'):
                    stencils.append({
                        "name": file,
                        "path": os.path.join(stencil_dir, file)
                    })
        
        # Check for open stencils
        open_stencils = []
        for i in range(1, visio_app.Documents.Count + 1):
            doc = visio_app.Documents.Item(i)
            # Check if document is a stencil
            if '.vss' in doc.Name or '.vssx' in doc.Name:
                open_stencils.append({
                    "name": doc.Name,
                    "path": doc.Path,
                    "full_path": os.path.join(doc.Path, doc.Name),
                    "is_open": True
                })
        
        result = {
            "built_in_stencils": stencils,
            "open_stencils": open_stencils
        }
        
        return {"status": "success", "data": result}
    
    except Exception as e:
        logger.error(f"Error getting stencils: {e}")
        return {"status": "error", "message": str(e)}

@app.post("/get-shapes")
async def get_shapes_on_page(request: ShapesRequest):
    """
    Get information about all shapes on a page.
    """
    if not connect_to_visio():
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")
    
    try:
        file_path = request.file_path
        page_index = request.page_index
        
        # Handle special 'active' keyword
        if file_path.lower() == 'active':
            doc = visio_app.ActiveDocument
            if not doc:
                return {"status": "error", "message": "No active document"}
            was_opened = False
        else:
            # Normalize path
            file_path = normalize_file_path(file_path)
            
            # Check if file exists
            if not os.path.exists(file_path):
                return {"status": "error", "message": f"File not found: {file_path}"}
            
            # Open the document
            try:
                # Check if already open
                try:
                    doc = visio_app.Documents(os.path.basename(file_path))
                    was_opened = True
                except:
                    doc = visio_app.Documents.Open(file_path)
                    was_opened = False
            except Exception as e:
                return {"status": "error", "message": f"Failed to open document: {str(e)}"}
        
        # Check page index
        if page_index > doc.Pages.Count:
            return {"status": "error", "message": f"Page index {page_index} out of range (document has {doc.Pages.Count} pages)"}
        
        page = doc.Pages.Item(page_index)
        
        # Prepare result
        result = {
            "document_name": doc.Name,
            "page_name": page.Name,
            "page_index": page_index,
            "shapes_count": page.Shapes.Count,
            "shapes": []
        }
        
        # Get all shapes on the page
        for i in range(1, page.Shapes.Count + 1):
            shape = page.Shapes.Item(i)
            
            # Get shape properties
            shape_data = {
                "id": shape.ID,
                "name": shape.Name,
                "text": shape.Text if hasattr(shape, "Text") else "",
                "is_connector": shape.OneD,
                "position": {
                    "x": shape.Cells("PinX").Result(""),
                    "y": shape.Cells("PinY").Result("")
                }
            }
            
            # Add shape-specific data
            if shape.OneD:
                # Connector-specific data
                try:
                    if shape.Connects.Count > 0:
                        # Get connection points
                        connections = []
                        for j in range(1, shape.Connects.Count + 1):
                            connect = shape.Connects.Item(j)
                            conn_data = {
                                "from_sheet": connect.FromSheet.Name,
                                "from_sheet_id": connect.FromSheet.ID,
                                "to_sheet": connect.ToSheet.Name,
                                "to_sheet_id": connect.ToSheet.ID
                            }
                            connections.append(conn_data)
                        
                        shape_data["connections"] = connections
                except Exception as e:
                    shape_data["connections_error"] = str(e)
            else:
                # Regular shape-specific data
                shape_data.update({
                    "type": shape.Type,
                    "master": shape.Master.Name if shape.Master else "None",
                    "size": {
                        "width": shape.Cells("Width").Result(""),
                        "height": shape.Cells("Height").Result("")
                    }
                })
            
            result["shapes"].append(shape_data)
        
        # Close document if it was opened for analysis
        if not was_opened and file_path.lower() != 'active':
            doc.Close()
        
        return {"status": "success", "data": result}
    
    except Exception as e:
        logger.error(f"Error getting shapes: {e}")
        return {"status": "error", "message": str(e)}

@app.post("/export-diagram")
async def export_diagram(request: ExportRequest):
    """
    Export a Visio diagram to another format.
    """
    if not connect_to_visio():
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")
    
    try:
        file_path = request.file_path
        export_format = request.format.lower()
        output_path = request.output_path
        
        # Handle special 'active' keyword
        if file_path.lower() == 'active':
            doc = visio_app.ActiveDocument
            if not doc:
                return {"status": "error", "message": "No active document"}
            was_opened = False
        else:
            # Normalize path
            file_path = normalize_file_path(file_path)
            
            # Check if file exists
            if not os.path.exists(file_path):
                return {"status": "error", "message": f"File not found: {file_path}"}
            
            # Open the document
            try:
                # Check if already open
                try:
                    doc = visio_app.Documents(os.path.basename(file_path))
                    was_opened = True
                except:
                    doc = visio_app.Documents.Open(file_path)
                    was_opened = False
            except Exception as e:
                return {"status": "error", "message": f"Failed to open document: {str(e)}"}
        
        # Generate output path if not provided
        if not output_path:
            base_name = os.path.splitext(doc.Name)[0]
            output_path = f"{doc.Path}\\{base_name}.{export_format}"
        else:
            output_path = normalize_file_path(output_path)
        
        # Ensure directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # Export based on format
        if export_format == "png":
            # Export as PNG
            doc.ExportAsFixedFormat(1, output_path, 1, 0)
        elif export_format == "jpg" or export_format == "jpeg":
            # Export as JPG
            doc.ExportAsFixedFormat(2, output_path, 1, 0)
        elif export_format == "pdf":
            # Export as PDF
            doc.ExportAsFixedFormat(1, output_path, 1, 0)
        elif export_format == "svg":
            # Export as SVG
            doc.ExportAsFixedFormat(5, output_path, 1, 0)
        else:
            return {"status": "error", "message": f"Unsupported export format: {export_format}"}
        
        # Close document if it was opened for export
        if not was_opened and file_path.lower() != 'active':
            doc.Close()
        
        result = {
            "document_name": doc.Name,
            "format": export_format,
            "output_path": output_path
        }
        
        return {"status": "success", "data": result}
    
    except Exception as e:
        logger.error(f"Error exporting diagram: {e}")
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    logger.info("Starting Visio Relay Service")
    
    # Try to connect to Visio at startup
    connect_to_visio()
    
    # Run the FastAPI app
    uvicorn.run(app, host="0.0.0.0", port=8051) 