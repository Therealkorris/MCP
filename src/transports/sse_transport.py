"""
Server-Sent Events (SSE) transport for the MCP-Visio server.
"""

import json
import logging
import asyncio
import traceback
from typing import Dict, Any, List, Optional

from fastapi import FastAPI, Request, Response
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

from services.mcp_service import process_message

# Configure logging
logger = logging.getLogger(__name__)

# Create FastAPI app
app = FastAPI(title="MCP-Visio")

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allow all origins (for development)
    allow_credentials=True,
    allow_methods=["*"],  # Allow all methods
    allow_headers=["*"],  # Allow all headers
)

# Request logging middleware
@app.middleware("http")
async def log_requests(request: Request, call_next):
    """Log all incoming requests and outgoing responses."""
    # Log request details
    logger.debug(f"Request: {request.method} {request.url}")
    
    # Process the request
    try:
        response = await call_next(request)
        logger.debug(f"Response status: {response.status_code}")
        return response
    except Exception as e:
        logger.error(f"Error during request processing: {e}")
        logger.error(traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={"detail": str(e)}
        )

# Health check endpoint for monitoring
@app.get("/health")
async def health_check():
    """Simple health check endpoint."""
    return JSONResponse(
        content={"status": "healthy", "service": "MCP-Visio"}
    )

# SSE endpoint for MCP
@app.route("/sse", methods=["GET", "POST"])
async def sse_endpoint(request: Request):
    """SSE endpoint for MCP JSON-RPC messages."""
    try:
        # Handle POST requests - process MCP messages
        if request.method == "POST":
            # Parse the JSON-RPC request
            data = await request.json()
            logger.debug(f"Received MCP request: {json.dumps(data, indent=2)}")
            
            # Process the MCP message
            response = await process_message(data)
            
            # Return the response
            return JSONResponse(content=response)
        
        # Handle GET requests - keep connection open for SSE
        else:
            async def event_stream():
                while True:
                    # Just keep the connection alive
                    await asyncio.sleep(60)
                    yield ""
                    
            return StreamingResponse(
                event_stream(),
                media_type="text/event-stream",
                headers={
                    "Content-Type": "text/event-stream",
                    "Cache-Control": "no-cache",
                    "Connection": "keep-alive",
                }
            )
    except Exception as e:
        logger.error(f"Error in SSE endpoint: {e}")
        logger.error(traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={"detail": str(e)}
        )

def run_sse_transport(host: str, port: int):
    """Run the SSE transport."""
    uvicorn.run(app, host=host, port=port) 