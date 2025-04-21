from fastapi import FastAPI, Body
from pydantic import BaseModel
from typing import Dict, Any, List, Optional
import sys
import os

# Ensure the services directory is in the path
sys.path.append(os.path.join(os.path.dirname(__file__), "src"))

app = FastAPI(title="Visio API")

# Define your data models
class AnalyzeDiagramRequest(BaseModel):
    file_path: Optional[str] = "active"
    analysis_type: Optional[str] = "all"

class ModifyDiagramRequest(BaseModel):
    file_path: Optional[str] = "active"
    operation: str
    shape_data: Dict[str, Any]

class VerifyConnectionsRequest(BaseModel):
    file_path: Optional[str] = "active"
    shape_ids: Optional[List[str]] = None

class CreateDiagramRequest(BaseModel):
    template: Optional[str] = "Basic.vst"
    save_path: Optional[str] = None

class SaveDiagramRequest(BaseModel):
    file_path: Optional[str] = None

class ShapesRequest(BaseModel):
    file_path: Optional[str] = "active"
    page_index: Optional[int] = 1

class ExportDiagramRequest(BaseModel):
    file_path: Optional[str] = "active"
    format: Optional[str] = "png"
    output_path: Optional[str] = None

class ImageToDiagramRequest(BaseModel):
    image_path: str
    output_path: Optional[str] = None
    detection_level: Optional[str] = "standard"  # standard, detailed, simple

# Define your endpoints with proper operation_id values
@app.post("/analyze-diagram", operation_id="analyze_visio_diagram")
async def analyze_diagram(request: AnalyzeDiagramRequest):
    """
    Analyze a Visio diagram to extract information about shapes, connections, and layout
    """
    # Import here to avoid circular imports
    from services.visio_service import VisioService
    visio_service = VisioService()
    result = visio_service.analyze_diagram(request.file_path, request.analysis_type)
    return result

@app.post("/modify-diagram", operation_id="modify_visio_diagram")
async def modify_diagram(request: ModifyDiagramRequest):
    """
    Modify a Visio diagram by adding, updating, or deleting shapes and connections
    """
    from services.visio_service import VisioService
    visio_service = VisioService()
    result = visio_service.modify_diagram(request.file_path, request.operation, request.shape_data)
    return result

@app.get("/active-document", operation_id="get_active_document")
async def get_active_document():
    """
    Get information about the currently active Visio document
    """
    from services.visio_service import VisioService
    visio_service = VisioService()
    result = visio_service.get_active_document()
    return result

@app.post("/verify-connections", operation_id="verify_connections")
async def verify_connections(request: VerifyConnectionsRequest):
    """
    Verify connections between shapes in a Visio diagram
    """
    from services.visio_service import VisioService
    visio_service = VisioService()
    result = visio_service.verify_connections(request.file_path, request.shape_ids)
    return result

@app.post("/create-diagram", operation_id="create_new_diagram")
async def create_diagram(request: CreateDiagramRequest):
    """
    Create a new Visio diagram from a template
    """
    from services.visio_service import VisioService
    visio_service = VisioService()
    result = visio_service.create_new_diagram(request.template, request.save_path)
    return result

@app.post("/save-diagram", operation_id="save_diagram")
async def save_diagram(request: SaveDiagramRequest):
    """
    Save the current Visio diagram
    """
    from services.visio_service import VisioService
    visio_service = VisioService()
    result = visio_service.save_diagram(request.file_path)
    return result

@app.get("/available-stencils", operation_id="get_available_stencils")
async def get_available_stencils():
    """
    Get a list of available Visio stencils
    """
    from services.visio_service import VisioService
    visio_service = VisioService()
    result = visio_service.get_available_stencils()
    return result

@app.post("/get-shapes", operation_id="get_shapes_on_page")
async def get_shapes_on_page(request: ShapesRequest):
    """
    Get detailed information about all shapes on a page
    """
    from services.visio_service import VisioService
    visio_service = VisioService()
    result = visio_service.get_shapes_on_page(request.file_path, request.page_index)
    return result

@app.post("/export-diagram", operation_id="export_diagram")
async def export_diagram(request: ExportDiagramRequest):
    """
    Export a Visio diagram to another format (PNG, JPG, PDF, SVG)
    """
    from services.visio_service import VisioService
    visio_service = VisioService()
    result = visio_service.export_diagram(request.file_path, request.format, request.output_path)
    return result

@app.get("/available-masters", operation_id="get_available_masters")
async def get_available_masters():
    """
    Get a list of available master shapes from all open stencils
    """
    from services.visio_service import VisioService
    visio_service = VisioService()
    result = visio_service.get_available_masters()
    return result

@app.post("/image-to-diagram", operation_id="image_to_diagram")
async def image_to_diagram(request: ImageToDiagramRequest):
    """
    Convert an image (screenshot, photo, etc.) to a Visio diagram
    """
    from services.visio_service import VisioService
    visio_service = VisioService()
    result = visio_service.image_to_diagram(request.image_path, request.output_path, request.detection_level)
    return result 