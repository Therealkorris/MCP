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
            
            # Get position information, use proper defaults
            position = shape_data.get("position", {})
            x = position.get("x", 4.0)
            y = position.get("y", 4.0)
            text = shape_data.get("text", "")
            
            # Get size information if provided
            size = shape_data.get("size", {})
            width = size.get("width")
            height = size.get("height")
            
            # Get stencil if specified
            stencil_name = shape_data.get("stencil_name", "Basic Shapes.vss")
            
            # Try to open the stencil
            try:
                try:
                    # Try to get already open stencil first
                    stencil = visio_app.Documents(stencil_name)
                    logger.info(f"Using already open stencil: {stencil_name}")
                except:
                    # Try different methods to open the stencil
                    try:
                        # First try built-in stencil folder
                        stencil_path = os.path.join(visio_app.GetBuiltInStencilFile(0, 0), stencil_name)
                        if os.path.exists(stencil_path):
                            stencil = visio_app.Documents.OpenEx(stencil_path, 0)
                            logger.info(f"Opened stencil from built-in path: {stencil_path}")
                        else:
                            # Try standard method
                            stencil = visio_app.Documents.OpenStencil(stencil_name)
                            logger.info(f"Opened stencil using OpenStencil: {stencil_name}")
                    except Exception as stencil_err:
                        logger.warning(f"Error opening stencil {stencil_name}: {stencil_err}")
                        # Fallback to Basic_U.vss which should always be available
                        try:
                            stencil = visio_app.Documents.OpenStencil("Basic_U.vss")
                            logger.info(f"Fallback to Basic_U.vss stencil")
                        except:
                            # Last resort - try each open document to find a stencil
                            for i in range(1, visio_app.Documents.Count + 1):
                                doc_item = visio_app.Documents.Item(i)
                                if ".vss" in doc_item.Name or ".vssx" in doc_item.Name:
                                    stencil = doc_item
                                    logger.info(f"Using already open stencil as fallback: {doc_item.Name}")
                                    break
                            else:
                                return {"status": "error", "message": "Could not find any usable stencil"}
            except Exception as e:
                return {"status": "error", "message": f"Could not open any stencil: {str(e)}"}
            
            # Get the master shape - try different naming patterns
            try:
                try:
                    # Try exact name match first
                    master = stencil.Masters.ItemU(master_name)
                except:
                    # Try pattern matching to find similar shape names
                    found = False
                    for i in range(1, stencil.Masters.Count + 1):
                        m = stencil.Masters.Item(i)
                        # Check for similar names (case insensitive, partial match)
                        if master_name.lower() in m.Name.lower():
                            master = m
                            found = True
                            logger.info(f"Found similar master shape: {m.Name} for request: {master_name}")
                            break
                    
                    if not found:
                        # Try common basic shapes as fallback
                        basic_shapes = ["Rectangle", "Square", "Circle", "Ellipse", "Triangle", "Diamond", "Pentagon", "Hexagon"]
                        for shape_name in basic_shapes:
                            try:
                                master = stencil.Masters.ItemU(shape_name)
                                found = True
                                logger.info(f"Using fallback basic shape: {shape_name}")
                                break
                            except:
                                continue
                        
                        if not found:
                            # Last resort: use the first master in the stencil
                            master = stencil.Masters.Item(1)
                            logger.info(f"Using first available master shape: {master.Name}")
            except Exception as e:
                return {"status": "error", "message": f"Master shape not found: {master_name} in {stencil_name} - {str(e)}"}
            
            # Drop the shape
            shape = page.Drop(master, x, y)
            
            # Set text if provided
            if text:
                shape.Text = text
            
            # Set custom size if provided - using a completely new approach
            try:
                if width is not None:
                    # Try direct assignment first
                    try:
                        shape.Width = float(width)
                    except:
                        logger.info(f"Direct width assignment failed, trying different method")
                        try:
                            shape.CellsSRC(1, 1, 2).Formula = f"{width}"
                        except:
                            logger.warning(f"Failed to set width using any method")
                            
                if height is not None:
                    # Try direct assignment first
                    try:
                        shape.Height = float(height)
                    except:
                        logger.info(f"Direct height assignment failed, trying different method")
                        try:
                            shape.CellsSRC(1, 1, 3).Formula = f"{height}"
                        except:
                            logger.warning(f"Failed to set height using any method")
            except Exception as size_err:
                logger.warning(f"Error setting shape size: {size_err}")
            
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
            
        elif operation == "add_connector":
            # Add a connector between two shapes (more reliable method)
            from_shape_id = shape_data.get("from_shape_id")
            to_shape_id = shape_data.get("to_shape_id")
            
            if not from_shape_id or not to_shape_id:
                return {"status": "error", "message": "from_shape_id and to_shape_id are required for add_connector operation"}
            
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
            
            # Try to get the Dynamic Connector master
            try:
                # Try different methods to get connector master
                connector_master = None
                try:
                    # Try to find a stencil with connectors
                    for i in range(1, visio_app.Documents.Count + 1):
                        doc_item = visio_app.Documents.Item(i)
                        if ".vss" in doc_item.Name or ".vssx" in doc_item.Name:
                            # Check masters in this stencil
                            for j in range(1, doc_item.Masters.Count + 1):
                                m = doc_item.Masters.Item(j)
                                if "connector" in m.Name.lower() or "dynamic" in m.Name.lower():
                                    connector_master = m
                                    logger.info(f"Using connector master: {m.Name} from {doc_item.Name}")
                                    break
                            if connector_master:
                                break
                except:
                    pass
                
                if not connector_master:
                    # Fallback to opening the basic stencil
                    try:
                        basic_stencil = visio_app.Documents.OpenStencil("Basic_U.vss")
                        connector_master = basic_stencil.Masters.ItemU("Dynamic connector")
                    except:
                        # Last resort: create with ConnectorToolDataObject
                        connector = page.Drop(visio_app.ConnectorToolDataObject, 0, 0)
                        has_master = False
                else:
                    # Drop the connector with master
                    connector = page.Drop(connector_master, 0, 0)
                    has_master = True
                
                # Connect the shapes - different methods based on how connector was created
                if has_master:
                    # Use proper begin/end connect methods
                    begin_cell = connector.Cells("BeginX")
                    end_cell = connector.Cells("EndX")
                    
                    # Get position of shapes
                    from_pos_x = from_shape.Cells("PinX").Result("")
                    from_pos_y = from_shape.Cells("PinY").Result("")
                    to_pos_x = to_shape.Cells("PinX").Result("")
                    to_pos_y = to_shape.Cells("PinY").Result("")
                    
                    # Configure connector begin and end points to be near the shapes
                    connector.Cells("BeginX").Result[""] = from_pos_x
                    connector.Cells("BeginY").Result[""] = from_pos_y
                    connector.Cells("EndX").Result[""] = to_pos_x
                    connector.Cells("EndY").Result[""] = to_pos_y
                    
                    # Connect the endpoints to the shapes
                    connector.CellsSRC(7, 0, 2).GlueTo(from_shape.CellsSRC(7, 0, 0))
                    connector.CellsSRC(7, 1, 2).GlueTo(to_shape.CellsSRC(7, 0, 0))
                else:
                    # Use alternate method
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
                
            except Exception as e:
                logger.error(f"Error creating connector: {e}")
                return {"status": "error", "message": f"Failed to create connector: {str(e)}"}
            
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
        
        # Try to create new document with requested template
        try:
            # First check if the template is already an open stencil
            template_found = False
            for i in range(1, visio_app.Documents.Count + 1):
                doc = visio_app.Documents.Item(i)
                if template.lower() in doc.Name.lower() and ('.vss' in doc.Name or '.vssx' in doc.Name or '.vst' in doc.Name):
                    template = doc.Name
                    template_found = True
                    logger.info(f"Using already open template: {template}")
                    break
            
            if not template_found:
                # Check if template path is specified or just a filename
                if os.path.dirname(template) and os.path.exists(template):
                    # Full path provided
                    logger.info(f"Using full path template: {template}")
                else:
                    # Try to find template in standard locations
                    template_dir = visio_app.GetBuiltInStencilFile(0, 0)
                    if template_dir and os.path.exists(template_dir):
                        possible_template = os.path.join(template_dir, template)
                        if os.path.exists(possible_template):
                            template = possible_template
                            logger.info(f"Found template in built-in directory: {template}")
                    
            # Create new document from template
            doc = None
            try:
                doc = visio_app.Documents.Add(template)
            except Exception as template_error:
                logger.warning(f"Error using template {template}: {template_error}")
                
                # Try fallback templates
                fallback_templates = ["BASIC_M.vssx", "Basic_U.vss", ""]  # Empty string means blank document
                for fallback in fallback_templates:
                    try:
                        if fallback:
                            logger.info(f"Trying fallback template: {fallback}")
                            doc = visio_app.Documents.Add(fallback)
                        else:
                            logger.info("Using blank document as fallback")
                            doc = visio_app.Documents.Add("")
                        break
                    except Exception as fallback_error:
                        logger.warning(f"Fallback template {fallback} failed: {fallback_error}")
                        continue
            
            if not doc:
                return {"status": "error", "message": "Could not create document with any template"}
            
            # Save if path provided
            if save_path:
                save_path = normalize_file_path(save_path)
                doc.SaveAs(save_path)
            
            # Return information about the new document
            result = {
                "name": doc.Name,
                "path": doc.Path,
                "full_path": os.path.join(doc.Path, doc.Name) if doc.Path else doc.Name,
                "pages_count": doc.Pages.Count
            }
            
            return {"status": "success", "data": result}
        except Exception as doc_error:
            logger.error(f"Error creating or saving document: {doc_error}")
            return {"status": "error", "message": f"Failed to create document: {str(doc_error)}"}
    
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
        # Only focus on already open stencils to avoid file system issues
        open_stencils = []
        
        # First get all currently open stencils
        for i in range(1, visio_app.Documents.Count + 1):
            try:
                doc = visio_app.Documents.Item(i)
                # Check if the document is a stencil
                if doc.Name.lower().endswith(('.vss', '.vssx')):
                    stencil_info = {
                        "name": doc.Name,
                        "is_open": True,
                        "masters_count": doc.Masters.Count
                    }
                    open_stencils.append(stencil_info)
            except Exception as e:
                logger.warning(f"Error processing document at index {i}: {e}")
                continue
        
        # Common stencil names to suggest (without attempting to access them)
        suggested_stencils = [
            {"name": "BASIC_M.vssx", "type": "basic"},
            {"name": "ARROWS_M.vssx", "type": "arrows"},
            {"name": "Basic_U.vss", "type": "basic"},
            {"name": "Backgrounds.vssx", "type": "backgrounds"},
            {"name": "Borders.vssx", "type": "borders"},
            {"name": "Connectors.vssx", "type": "connectors"},
            {"name": "CONNEC_M.vssx", "type": "connectors"}
        ]
        
        # Return the information without any file system access
        result = {
            "open_stencils": open_stencils,
            "suggested_stencils": suggested_stencils,
            "open_stencils_count": len(open_stencils)
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

@app.get("/available-masters")
async def get_available_masters():
    """
    Get a list of available masters from all open stencils.
    """
    if not connect_to_visio():
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")
    
    try:
        masters_by_stencil = {}
        
        # Check open stencils first
        for i in range(1, visio_app.Documents.Count + 1):
            doc = visio_app.Documents.Item(i)
            
            # Check if document is a stencil
            if '.vss' in doc.Name or '.vssx' in doc.Name:
                stencil_masters = []
                
                # Get all masters in this stencil
                for j in range(1, doc.Masters.Count + 1):
                    master = doc.Masters.Item(j)
                    stencil_masters.append({
                        "name": master.Name,
                        "id": j,
                        "type": master.Type,
                        "one_d": master.OneD
                    })
                
                masters_by_stencil[doc.Name] = stencil_masters
        
        # If no open stencils found, try to open basic stencil
        if not masters_by_stencil:
            try:
                basic_stencil = visio_app.Documents.OpenStencil("Basic_U.vss")
                stencil_masters = []
                
                # Get all masters in this stencil
                for j in range(1, basic_stencil.Masters.Count + 1):
                    master = basic_stencil.Masters.Item(j)
                    stencil_masters.append({
                        "name": master.Name,
                        "id": j,
                        "type": master.Type,
                        "one_d": master.OneD
                    })
                
                masters_by_stencil["Basic_U.vss"] = stencil_masters
                
                # Close stencil after we're done
                basic_stencil.Close()
            except Exception as e:
                logger.warning(f"Error opening basic stencil: {e}")
        
        return {"status": "success", "data": {"masters_by_stencil": masters_by_stencil}}
    
    except Exception as e:
        logger.error(f"Error getting masters: {e}")
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    logger.info("Starting Visio Relay Service")
    
    # Try to connect to Visio at startup
    connect_to_visio()
    
    # Set up monitored directories for auto-reload
    # This will automatically detect changes to Python files and restart the server
    watch_dirs = [".", "src", "src/services"]
    reload_dirs = [os.path.abspath(dir) for dir in watch_dirs]
    
    logger.info(f"Auto-reload enabled. Monitoring directories: {watch_dirs}")
    logger.info("Make changes to host_visio_relay.py, mcp_service.py, or visio_service.py to trigger reload")
    
    # Create a main module for uvicorn to import
    if not os.path.exists("__init__.py"):
        with open("__init__.py", "w") as f:
            pass  # Create an empty __init__.py file
    
    # For auto-reload to work, we need to run as module
    if os.path.basename(sys.argv[0]) == "host_visio_relay.py":
        # Run the FastAPI app with auto-reload enabled using the import string format
        import uvicorn
        uvicorn.run("host_visio_relay:app", host="0.0.0.0", port=8051, reload=True, reload_dirs=reload_dirs) 