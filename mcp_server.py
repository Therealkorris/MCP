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

# Import the Visio API app at module level
try:
    from visio_api import app as visio_app
    
    # Create the MCP server at module level
    mcp = FastApiMCP(visio_app)
    
    # Mount the MCP server
    mcp.mount()
    
    # Expose the app for uvicorn
    app = visio_app
    
except ImportError:
    logger.warning("Visio API could not be imported at module level. Will try again in main().")
    app = None

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
    
    # Add debug mode flag
    parser.add_argument(
        "--debug",
        action="store_true",
        default=os.getenv("DEBUG", "false").lower() == "true",
        help="Enable debug mode with auto-reload"
    )
    
    return parser.parse_args()

def main():
    """Main entry point for the MCP-Visio server."""
    args = parse_args()
    
    try:
        global app
        # If app wasn't successfully imported at module level, try again
        if app is None:
            # Import the Visio API app
            from visio_api import app as visio_app
            
            # Create the MCP server
            logger.info("Creating MCP server from Visio API")
            mcp = FastApiMCP(visio_app)
            
            # Mount the MCP server
            logger.info("Mounting MCP server at /mcp")
            mcp.mount()
            
            # Set the global app
            app = visio_app
        else:
            logger.info("Using Visio API app imported at module level")
        
        # Start the server with optional auto-reload
        logger.info(f"Starting server on {args.host}:{args.port}")
        
        if args.debug:
            # Set up auto-reload
            import os
            from pathlib import Path
            
            # Define directories to watch for changes
            watch_dirs = [".", "src", "src/services"]
            reload_dirs = [os.path.abspath(dir) for dir in watch_dirs]
            
            logger.info("Debug mode enabled with auto-reload")
            logger.info(f"Monitoring directories: {watch_dirs}")
            
            # Create __init__.py files if they don't exist
            if not os.path.exists("__init__.py"):
                with open("__init__.py", "w") as f:
                    pass
            
            # Run with auto-reload enabled using import string
            uvicorn.run("mcp_server:app", host=args.host, port=args.port, reload=True, reload_dirs=reload_dirs)
        else:
            # Run in standard mode
            logger.info("Running in standard mode (no auto-reload)")
            uvicorn.run(app, host=args.host, port=args.port)
    
    except Exception as e:
        logger.error(f"Error starting MCP server: {e}")
        raise

if __name__ == "__main__":
    main() 