"""
Main entry point for the MCP-Visio server.
This server implements the Model Context Protocol to allow AI to interact with Microsoft Visio.
"""

import os
import sys
import json
import logging
import argparse
from dotenv import load_dotenv
from pathlib import Path

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

# Import transports
from transports.sse_transport import run_sse_transport
from transports.stdio_transport import run_stdio_transport

def parse_args():
    """Parse command-line arguments with environment variable fallbacks."""
    parser = argparse.ArgumentParser(description="MCP-Visio Server")
    
    # Transport config
    parser.add_argument(
        "--transport", 
        type=str,
        default=os.getenv("TRANSPORT", "sse"),
        choices=["sse", "stdio"],
        help="Transport protocol (sse or stdio)"
    )
    
    # SSE transport config
    parser.add_argument(
        "--host",
        type=str,
        default=os.getenv("HOST", "0.0.0.0"),
        help="Host to bind to when using SSE transport"
    )
    parser.add_argument(
        "--port",
        type=int,
        default=os.getenv("PORT", 8050),
        help="Port to listen on when using SSE transport"
    )
    
    # Debug mode
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
    
    if args.debug:
        logger.info("Debug mode enabled with auto-reload")
    
    # Determine which transport to use
    if args.transport == "sse":
        logger.info(f"Starting MPC-Visio server with SSE transport on {args.host}:{args.port}")
        run_sse_transport(args.host, args.port, debug=args.debug)
    elif args.transport == "stdio":
        logger.info("Starting MPC-Visio server with STDIO transport")
        run_stdio_transport()
    else:
        logger.error(f"Unsupported transport: {args.transport}")
        sys.exit(1)

if __name__ == "__main__":
    main() 