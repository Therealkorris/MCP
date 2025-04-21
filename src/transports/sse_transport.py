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
@app.get("/mcp")
async def sse_endpoint_get(request: Request):
    """SSE endpoint for MCP - GET handler."""
    try:
        # Set response headers
        response = StreamingResponse(
            content=event_generator(),
            media_type="text/event-stream",
            headers={
                "Content-Type": "text/event-stream",
                "Cache-Control": "no-cache, no-transform",
                "Connection": "keep-alive",
                "Access-Control-Allow-Origin": "*",
                "X-Accel-Buffering": "no"
            }
        )
        
        # Add CORS headers
        response.headers["Access-Control-Allow-Credentials"] = "true"
        response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "*"
        
        return response
    except Exception as e:
        logger.error(f"Error in SSE GET endpoint: {e}")
        logger.error(traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={"detail": str(e)}
        )

async def event_generator():
    """Generate SSE events."""
    # Send initial connection established event
    yield "data: {\"type\":\"connection_established\"}\n\n"
    
    # Send heartbeat events
    counter = 0
    while True:
        await asyncio.sleep(3)
        counter += 1
        yield f"data: {{\"type\":\"heartbeat\",\"count\":{counter}}}\n\n"

@app.post("/mcp")
async def sse_endpoint_post(request: Request):
    """SSE endpoint for MCP - POST handler."""
    try:
        # Parse the JSON-RPC request
        data = await request.json()
        logger.debug(f"Received MCP request: {json.dumps(data, indent=2)}")
        
        # Process the MCP message
        response = await process_message(data)
        
        # Return the response
        return JSONResponse(content=response)
    except Exception as e:
        logger.error(f"Error in SSE POST endpoint: {e}")
        logger.error(traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={"detail": str(e)}
        )

@app.options("/mcp")
async def options_mcp(request: Request):
    """Handle OPTIONS requests for CORS preflight."""
    headers = {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        "Access-Control-Allow-Headers": "*",
        "Access-Control-Allow-Credentials": "true",
        "Access-Control-Max-Age": "86400"  # 24 hours
    }
    return Response(status_code=204, headers=headers)

def run_sse_transport(host: str, port: int, debug: bool = False):
    """Run the SSE transport with optional auto-reload."""
    import os
    from pathlib import Path
    
    if debug:
        # Set up monitored directories for auto-reload
        watch_dirs = [".", "src", "src/services", "src/transports"]
        reload_dirs = [os.path.abspath(dir) for dir in watch_dirs]
        
        logger.info(f"Auto-reload enabled for MCP server. Monitoring directories: {watch_dirs}")
        logger.info("Make changes to Python files to trigger reload")
        
        # Ensure __init__.py files exist for proper module imports
        for dir_path in ["src", "src/services", "src/transports"]:
            init_file = os.path.join(dir_path, "__init__.py")
            if not os.path.exists(init_file):
                with open(init_file, "w") as f:
                    pass  # Create empty __init__.py file
        
        # Run with auto-reload enabled using import string format
        uvicorn.run("src.transports.sse_transport:app", host=host, port=port, reload=True, reload_dirs=reload_dirs)
    else:
        # Run in standard mode without auto-reload
        logger.info("Running in standard mode (no auto-reload)")
        uvicorn.run(app, host=host, port=port) 