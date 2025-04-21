"""
Visio Relay Service

This script runs on the Windows host and provides a REST API to interact with
Visio documents. The Docker container communicates with this relay service
to access Visio functionality.

Usage:
    python visio_relay.py
"""

import os
import sys
import json
import logging
import traceback
from pathlib import Path
from typing import Dict, Any, List

import win32com.client
import pythoncom
from fastapi import FastAPI, HTTPException, Body
from pydantic import BaseModel
import uvicorn

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("visio_relay.log")
    ]
)
logger = logging.getLogger(__name__)

# Create FastAPI app
app = FastAPI(title="Visio Relay Service")

# Global Visio app instance
visio_app = None

# Models for API
class VisioDocument(BaseModel):
    file_path: str = None
    analysis_type: str = "all"
    page_index: int = 1

class VisioResponse(BaseModel):
    status: str
    message: str = None
    data: Dict[str, Any] = None

# Connect to Visio
def connect_to_visio():
    """Connect to Microsoft Visio application."""
    global visio_app
    
    try:
        if visio_app is not None:
            return True
            
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Try to connect to an existing Visio instance
        try:
            logger.info("Attempting to connect to existing Visio instance...")
            visio_app = win32com.client.GetActiveObject("Visio.Application")
            logger.info("Connected to existing Visio instance")
        except:
            # Create a new Visio instance
            logger.info("No existing Visio instance found, creating new instance...")
            visio_app = win32com.client.Dispatch("Visio.Application")
            visio_app.Visible = True
            logger.info("Created new Visio instance")
        
        return True
    except Exception as e:
        logger.error(f"Failed to connect to Visio: {e}")
        logger.error(traceback.format_exc())
        return False

# API Endpoints
@app.get("/")
def root():
    """Root endpoint with basic info."""
    return {"message": "Visio Relay Service is running"}

@app.get("/health")
def health_check():
    """Health check endpoint."""
    return {"status": "healthy"}

@app.get("/connect")
def connect():
    """Connect to Visio."""
    if connect_to_visio():
        return VisioResponse(status="success", message="Connected to Visio")
    else:
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")

@app.get("/active-document")
def get_active_document():
    """Get information about the active document."""
    if not connect_to_visio():
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")
    
    try:
        # Check if there is an active document
        if visio_app.Documents.Count == 0:
            return VisioResponse(
                status="no_document",
                message="No Visio document is currently open"
            )
        
        # Get active document
        doc = visio_app.ActiveDocument
        
        # Get basic document info
        doc_info = {
            "name": doc.Name,
            "path": doc.Path,
            "full_path": os.path.join(doc.Path, doc.Name) if doc.Path else doc.Name,
            "pages_count": doc.Pages.Count,
            "saved": doc.Saved,
            "readonly": doc.ReadOnly,
            "pages": [],
            "open_documents": []
        }
        
        # Get information about all open documents
        for i in range(1, visio_app.Documents.Count + 1):
            open_doc = visio_app.Documents.Item(i)
            doc_info["open_documents"].append({
                "name": open_doc.Name,
                "path": open_doc.Path,
                "full_path": os.path.join(open_doc.Path, open_doc.Name) if open_doc.Path else open_doc.Name,
                "is_active": (open_doc.Name == doc.Name)
            })
        
        # Get information about pages
        for i in range(1, doc.Pages.Count + 1):
            page = doc.Pages.Item(i)
            page_info = {
                "name": page.Name,
                "index": i,
                "shapes_count": page.Shapes.Count,
                "is_foreground": page.Background == 0
            }
            doc_info["pages"].append(page_info)
        
        return VisioResponse(status="success", data=doc_info)
    
    except Exception as e:
        logger.error(f"Error getting active document: {e}")
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/analyze-diagram")
def analyze_diagram(doc: VisioDocument):
    """Analyze a Visio diagram."""
    if not connect_to_visio():
        raise HTTPException(status_code=500, detail="Failed to connect to Visio")
    
    try:
        file_path = doc.file_path
        analysis_type = doc.analysis_type
        
        # If no file path is provided, use the active document
        if not file_path:
            if visio_app.Documents.Count == 0:
                return VisioResponse(
                    status="error",
                    message="No Visio document is open and no file path provided"
                )
            doc_obj = visio_app.ActiveDocument
        else:
            # Check if file exists
            if not os.path.exists(file_path):
                return VisioResponse(
                    status="error",
                    message=f"File not found: {file_path}"
                )
            
            # Try to open the document
            try:
                doc_obj = visio_app.Documents.Open(file_path)
                logger.info(f"Opened document: {file_path}")
            except Exception as e:
                return VisioResponse(
                    status="error",
                    message=f"Failed to open document: {e}"
                )
        
        # Prepare result structure
        result = {
            "name": doc_obj.Name,
            "pages": []
        }
        
        # Analyze each page
        for i in range(1, doc_obj.Pages.Count + 1):
            page = doc_obj.Pages.Item(i)
            page_info = {
                "name": page.Name,
                "index": i,
                "shapes_count": page.Shapes.Count,
                "shapes": [],
                "connections": [],
                "text": []
            }
            
            # Get shapes
            for j in range(1, page.Shapes.Count + 1):
                shape = page.Shapes.Item(j)
                
                # Skip connectors for the shapes list
                if not shape.OneD:
                    shape_info = {
                        "id": shape.ID,
                        "name": shape.Name,
                        "text": shape.Text if hasattr(shape, "Text") else "",
                        "type": shape.Type
                    }
                    page_info["shapes"].append(shape_info)
                else:
                    # Process connectors
                    try:
                        connection_info = {
                            "id": shape.ID,
                            "name": shape.Name,
                            "text": shape.Text if hasattr(shape, "Text") else ""
                        }
                        page_info["connections"].append(connection_info)
                    except:
                        pass
                
                # Process text
                if hasattr(shape, "Text") and shape.Text:
                    text_info = {
                        "shape_id": shape.ID,
                        "shape_name": shape.Name,
                        "text": shape.Text
                    }
                    page_info["text"].append(text_info)
            
            result["pages"].append(page_info)
        
        return VisioResponse(status="success", data=result)
    
    except Exception as e:
        logger.error(f"Error analyzing diagram: {e}")
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))

# Run the service
if __name__ == "__main__":
    logger.info("Starting Visio Relay Service...")
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", 8051))
    logger.info(f"Service will be available at http://{host}:{port}")
    uvicorn.run(app, host=host, port=port) 