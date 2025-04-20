"""
MCP Server for Visio using fastapi_mcp.
This module creates an MCP server from the Visio API.
"""

import os
import logging
import argparse
from dotenv import load_dotenv
from fastapi_mcp import FastApiMCP
import uvicorn

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("mcp_visio.log")
    ]
)
logger = logging.getLogger(__name__)

def parse_args():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(description="MCP-Visio Server")
    
    parser.add_argument(
        "--host",
        type=str,
        default=os.getenv("HOST", "0.0.0.0"),
        help="Host to bind to"
    )
    parser.add_argument(
        "--port",
        type=int,
        default=os.getenv("PORT", 8050),
        help="Port to listen on"
    )
    
    return parser.parse_args()

def main():
    """Main entry point for the MCP-Visio server."""
    args = parse_args()
    
    try:
        # Import the Visio API app
        from visio_api import app
        
        # Create the MCP server
        logger.info("Creating MCP server from Visio API")
        mcp = FastApiMCP(app)
        
        # Mount the MCP server
        logger.info("Mounting MCP server at /mcp")
        mcp.mount()
        
        # Start the server
        logger.info(f"Starting server on {args.host}:{args.port}")
        uvicorn.run(app, host=args.host, port=args.port)
    
    except Exception as e:
        logger.error(f"Error starting MCP server: {e}")
        raise

if __name__ == "__main__":
    main() 