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
    file_path: str
    analysis_type: Optional[str] = "all"

class ModifyDiagramRequest(BaseModel):
    file_path: str
    operation: str
    shape_data: Dict[str, Any]

class VerifyConnectionsRequest(BaseModel):
    file_path: str
    shape_ids: Optional[List[str]] = None

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